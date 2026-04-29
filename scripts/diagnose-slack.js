#!/usr/bin/env node
/* eslint-disable no-console */
require('dotenv').config();
const { WebClient } = require('@slack/web-api');

async function main() {
  const token = process.env.SLACK_BOT_TOKEN;
  if (!token) {
    console.error('FAIL: SLACK_BOT_TOKEN is missing in .env');
    process.exit(1);
  }

  const channelArg = process.argv[2];
  const channel = channelArg || process.env.DIAGNOSE_SLACK_CHANNEL || '';
  const slack = new WebClient(token);

  console.log('Slack Diagnose');
  console.log('-------------');

  try {
    const auth = await slack.auth.test();
    console.log(`OK: auth.test succeeded (team=${auth.team}, bot_user_id=${auth.user_id})`);
  } catch (err) {
    const code = err?.data?.error || err?.code || err?.message || 'unknown_error';
    console.error(`FAIL: auth.test failed: ${code}`);
    process.exit(1);
  }

  if (!channel) {
    console.log('');
    console.log('SKIP: No channel provided for chat.postMessage test.');
    console.log('Run one of these:');
    console.log('  npm run diagnose:slack -- C1234567890');
    console.log('  DIAGNOSE_SLACK_CHANNEL=C1234567890 npm run diagnose:slack');
    process.exit(0);
  }

  try {
    const result = await slack.chat.postMessage({
      channel,
      text: 'Slack diagnostics test message from your bot.',
    });
    console.log(`OK: chat.postMessage succeeded (channel=${channel}, ts=${result.ts})`);
  } catch (err) {
    const code = err?.data?.error || err?.code || err?.message || 'unknown_error';
    console.error(`FAIL: chat.postMessage failed: ${code}`);
    process.exit(1);
  }
}

main().catch((err) => {
  const code = err?.data?.error || err?.code || err?.message || 'unknown_error';
  console.error(`FAIL: diagnostics crashed: ${code}`);
  process.exit(1);
});
