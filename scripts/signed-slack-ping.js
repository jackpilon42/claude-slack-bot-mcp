#!/usr/bin/env node
/**
 * POST a valid Slack-style signed event_callback (app_mention "ping") to your local bot.
 * Use when the Slack app Request URL points elsewhere and you still want to verify
 * index.js + SLACK_SIGNING_SECRET + outbound chat.postMessage on this machine.
 * The payload sets `_local_post_to_channel` so the bot omits `thread_ts` (the synthetic `ts` is not a real Slack message).
 *
 * Usage (bot must be running: npm run start):
 *   node scripts/signed-slack-ping.js [CHANNEL_ID]
 *
 * Env: SLACK_SIGNING_SECRET, SLACK_BOT_TOKEN (for default channel lookup optional)
 *      SLACK_TEST_CHANNEL — default channel if arg omitted
 *      SLACK_TEST_USER_ID — author user id (default U0000000000)
 *      SLACK_EVENTS_URL   — default http://127.0.0.1:3000/slack/events
 */
/* eslint-disable no-console */
require('dotenv').config();
const crypto = require('crypto');

async function getBotUserId() {
  const fromEnv = String(process.env.SLACK_BOT_USER_ID || '').trim();
  if (fromEnv) return fromEnv;
  const res = await fetch('https://slack.com/api/auth.test', {
    headers: { Authorization: `Bearer ${process.env.SLACK_BOT_TOKEN}` },
  });
  const data = await res.json().catch(() => ({}));
  if (!data.ok) throw new Error(`auth.test failed: ${data.error || res.status}`);
  return String(data.user_id || '').trim();
}

async function main() {
  const secret = String(process.env.SLACK_SIGNING_SECRET || '').trim();
  if (!secret) {
    console.error('Missing SLACK_SIGNING_SECRET in .env');
    process.exit(1);
  }
  const channel =
    String(process.argv[2] || process.env.SLACK_TEST_CHANNEL || '').trim() ||
    (() => {
      console.error('Pass channel id: node scripts/signed-slack-ping.js C0123… or set SLACK_TEST_CHANNEL in .env');
      process.exit(1);
    })();
  const botUserId = await getBotUserId();
  const humanUser = String(process.env.SLACK_TEST_USER_ID || 'U0000000000').trim();
  const target = String(process.env.SLACK_EVENTS_URL || 'http://127.0.0.1:3000/slack/events').trim();

  const ts = Math.floor(Date.now() / 1000);
  const eventTs = `${ts}.${String(Math.floor(Math.random() * 1e6)).padStart(6, '0')}`;
  const bodyObj = {
    token: 'ZZZZZZWWW',
    team_id: 'T0000000000',
    api_app_id: 'A0000000000',
    event: {
      type: 'app_mention',
      user: humanUser,
      text: `<@${botUserId}> ping`,
      ts: eventTs,
      channel,
      event_ts: eventTs,
      /** Synthetic ts is not a real Slack message — tell the bot to post replies in the channel root. */
      _local_post_to_channel: true,
    },
    type: 'event_callback',
    event_id: `EvTest${Date.now()}`,
    event_time: ts,
  };
  const rawBody = JSON.stringify(bodyObj);
  const sigBase = `v0:${ts}:${rawBody}`;
  const sig = `v0=${crypto.createHmac('sha256', secret).update(sigBase).digest('hex')}`;

  const res = await fetch(target, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'X-Slack-Request-Timestamp': String(ts),
      'X-Slack-Signature': sig,
    },
    body: rawBody,
  });
  const text = await res.text();
  console.log('HTTP', res.status, res.statusText);
  console.log(text.slice(0, 500));
  if (!res.ok) process.exit(1);
}

main().catch((e) => {
  console.error(e?.message || e);
  process.exit(1);
});
