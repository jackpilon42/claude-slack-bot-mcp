const express = require('express');
const Anthropic = require('@anthropic-ai/sdk');
const { WebClient } = require('@slack/web-api');
const OpenAI = require('openai');
const crypto = require('crypto');
const fs = require('fs');
const path = require('path');
const { AsyncLocalStorage } = require('async_hooks');
require('dotenv').config();
const { shouldHandleAsProposalPipeline, handleProposalPipeline } = require('./proposal-pipeline');

const app = express();
app.use(
  express.json({
    verify: (req, _res, buf) => {
      req.rawBody = buf.toString('utf8');
    },
  })
);

/** Every POST (any path): proves inbound HTTP reached this process. Slack must POST to /slack/events on this port. */
app.use((req, res, next) => {
  if (req.method === 'POST') {
    const host = String(req.headers.host || '');
    const fwd = String(req.headers['x-forwarded-for'] || '').slice(0, 60);
    console.log('[http] POST', req.originalUrl || req.url, 'host=', host, fwd ? `xff=${fwd}` : '');
  }
  next();
});

const client = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });
const slack = new WebClient(process.env.SLACK_BOT_TOKEN, { timeout: 8000 });
/** Set from SLACK_BOT_USER_ID or first successful auth.test — used to ignore the bot's own outbound messages. */
let slackBotUserId = String(process.env.SLACK_BOT_USER_ID || '').trim() || null;
const openai = process.env.OPENAI_API_KEY ? new OpenAI({ apiKey: process.env.OPENAI_API_KEY }) : null;
const REQUIRED_ENV_VARS = ['SLACK_BOT_TOKEN', 'SLACK_SIGNING_SECRET', 'ANTHROPIC_API_KEY'];
const SMARTSHEET_BASE_URL = 'https://api.smartsheet.com/2.0';
const BUILDING_CONNECTED_BASE_URL = 'https://developer.api.autodesk.com/construction/buildingconnected/v2';
const AUTODESK_TOKEN_URL = 'https://developer.api.autodesk.com/authentication/v2/token';
const AUTODESK_AUTHORIZE_URL = 'https://developer.api.autodesk.com/authentication/v2/authorize';
const DEFAULT_AUTODESK_SCOPES = 'data:read';
/** Smartsheet bulk row update limit (rows per PUT /sheets/{id}/rows). */
const SMARTSHEET_ROW_PUT_BATCH = 100;
const SMARTSHEET_SHEET_PAGE_SIZE = 5000;
const SMARTSHEET_NON_CLEARABLE_COLUMN_TYPES = new Set([
  'AUTO_NUMBER',
  'SYSTEM_AUTO_NUMBER',
  'CREATED_DATE',
  'MODIFIED_DATE',
  'CREATED_BY',
  'MODIFIED_BY',
  'PREDECESSOR',
  'DURATION',
]);
const smartsheetSheetContext = new AsyncLocalStorage();
/** Slack channel id -> Smartsheet sheet id (numeric string). */
const slackChannelActiveSheetId = new Map();
/** Smartsheet sheet id -> Map(lower column title -> column meta). */
const smartsheetColumnsCacheBySheetId = new Map();
const RESTART_THIS_BOT_SLACK_MARKDOWN = ` When the user should restart this Node bot (after code or .env changes), always include this exact markdown so they can copy the command (opening \`\`\` on its own line, one line inside, closing fence on its own line):
\`\`\`bash
cd /Users/jackpilon/claude-slack-bot-mcp && npm run startup
\`\`\`
Say to press Ctrl+C in the bot terminal first if something is already listening on port 3000.`;
const ASSISTANT_SYSTEM_PROMPT =
  'You are Transform Sim Slack Assistant. Help with concise, practical answers. ' +
  'You are already speaking inside Slack through this bot. Never claim you cannot message on Slack. ' +
  'If asked to send a message to someone, explain that you can only reply in this thread/channel unless explicitly given tool support.' +
  RESTART_THIS_BOT_SLACK_MARKDOWN;
const ASSISTANT_SYSTEM_PROMPT_WITH_SMARTSHEET =
  ASSISTANT_SYSTEM_PROMPT +
  ' This bot is connected to Smartsheet via API on the server side. ' +
  'Never say you cannot access Smartsheet or external systems in general. ' +
  'If the user asks to change the sheet, tell them to use the real column titles from the sheet (exact spelling) and, for updates, the values to match/set. ' +
  'Formatting requests can mention a column title plus a row number range and a color. ' +
  'They can name which Smartsheet to use in plain English, e.g. "In Job Sim sheet, …" or "Switch to the Ops tracker sheet", when multiple sheets are configured. ' +
  'They can also use explicit commands like `smartsheet help`.';

/** When true, do not route the message through Smartsheet natural-language handlers. */
function mentionsBuildingConnectedIntent(text) {
  const lower = String(text || '').toLowerCase();
  if (/\bbuilding\s*connected\b/.test(lower) || /\bbuildingconnected\b/.test(lower)) return true;
  if (/\bbid\s*board\b/.test(lower) && /\b(autodesk|building\s*connected)\b/.test(lower)) return true;
  if (/\b(bc)\s+(api|projects?|bids?|oauth)\b/.test(lower)) return true;
  return false;
}

/** Restated questions often omit the product name; avoid Smartsheet for obvious BC-style discovery. */
function looksLikeGeographicOrBidDiscovery(text) {
  const lower = String(text || '').toLowerCase();
  if (/\bsmartsheet\b/.test(lower)) return false;
  const geoMile =
    /\b(within|inside|radius|around|near)\b[\s\S]{0,40}\b\d+\s*(?:mi|mile|miles|km|kilometers?)\b/i.test(lower) ||
    /\b\d+\s*(?:mi|mile|miles|km|kilometers?)\b[\s\S]{0,40}\b(of|from)\b/i.test(lower);
  const geoPlace =
    /\b(same\s+)?county(\s+as|\s+for|\s+near)?\b/i.test(lower) ||
    /\bmetro(politan)?\s+area\b|\bzip\s*code\b/i.test(lower) ||
    /\b(near|around|close to|outside|just)\s+[a-z][a-z0-9\s]{2,35}\b/i.test(lower) ||
    /\bin\s+[a-z][a-z0-9\s]{2,35},\s*[a-z]{2}\b/i.test(lower) ||
    /\b(?:jobs?|bids?|projects?|work|anything)\s+in\s+[a-z][a-z0-9]{2,30}\b/i.test(lower);
  const geo = geoMile || geoPlace;
  const scope =
    /\b(project|projects|bid|bids|bidding|rfp|rfqs?|gc|subcontractor|preconstruction|posted)\b/i.test(lower) ||
    /\b(roofing|hvac|electrical|plumbing|concrete|drywall|framing|paving|demolition|excavation)\b/i.test(lower) ||
    /\b(available|opportunit|leads?|work)\b/i.test(lower);
  return Boolean(geo && scope);
}

/** "List my Autodesk / BuildingConnected projects" without slash-commands — coworker-style. */
function looksLikeCasualProjectListingAsk(text) {
  const lower = String(text || '').toLowerCase();
  if (/\bsmartsheet\b/.test(lower)) return false;
  if (!/\b(building\s*connected|buildingconnected|autodesk|\bbc\b|bid\s*board)\b/i.test(lower)) return false;
  if (!/\b(project|projects|bid|bids)\b/i.test(lower)) return false;
  return /\b(list|show|see|pull up|open|what\s+(?:are|do|is)|which|everything|all(\s+the)?|anything|any)\b/i.test(lower);
}

function userTextLooksLikeBcDiscoveryOrCasualList(text) {
  return (
    mentionsBuildingConnectedIntent(text) ||
    looksLikeGeographicOrBidDiscovery(text) ||
    looksLikeCasualProjectListingAsk(text)
  );
}

function skipSmartsheetNaturalLanguageForMessage(text) {
  return userTextLooksLikeBcDiscoveryOrCasualList(text);
}

function appendBuildingConnectedAssistantGuidance(systemPrompt) {
  if (!buildingConnectedConfigured()) return systemPrompt;
  const connected = hasValidAutodeskAccessToken();
  let extra =
    ' Autodesk BuildingConnected (APS) is integrated on this server. ' +
    'Do NOT tell the user to use Smartsheet for BuildingConnected data unless they explicitly say they export or mirror BC data into a sheet. ';
  if (connected) {
    extra +=
      'CRITICAL: OAuth is CONNECTED on this server right now. Do NOT tell the user to run `bc connect` unless they report auth errors or the server was restarted and lost the session. ' +
      'They can confirm with `bc auth status`. **Coworker-style:** users ask in plain English (cities, counties, trades, "list my Autodesk projects"); this bot may already fetch BuildingConnected data on the server and paste results—do **not** tell them they must memorize commands like `bc projects`. Optional `bc …` commands remain for power users. ' +
      'BuildingConnected here is **API-only** (no browser / UI scraping). Do NOT promise automatic geo filters beyond keyword matching on API payloads. ' +
      'If their account hits BC_PRO_SUBSCRIPTION_REQUIRED when listing projects, explain the API refused project list reads—not Smartsheet.';
  } else {
    extra +=
      'OAuth is NOT connected on this server right now: the user should run `bc connect` in Slack once and approve in the browser. Do not substitute Smartsheet for BC.';
  }
  return systemPrompt + extra;
}

/** One-line facts for the model so it cannot contradict live server state. */
function buildingConnectedUserMessagePrefix() {
  if (!buildingConnectedConfigured()) return '';
  const connected = hasValidAutodeskAccessToken();
  return `[BuildingConnected: OAuth ${connected ? 'CONNECTED' : 'NOT_CONNECTED — user should run bc connect'}]`;
}

function assistantReplyIncorrectlyDemandsBcConnect(reply) {
  const lower = String(reply || '').toLowerCase().replace(/\u2019/g, "'");
  if (/\bbc\s+connect\b/.test(lower)) return true;
  if (/[`'"]\s*bc\s+connect\b/.test(lower)) return true;
  if (/slack:\s*bc\s+connect/.test(lower)) return true;
  if (/\bconnect your (?:building\s*connected\s*)?account\b/.test(lower)) return true;
  if (/\bneed to connect\b/.test(lower) && /building|oauth|autodesk/.test(lower)) return true;
  if (/\byou(?:'ll| will) need to connect\b/.test(lower)) return true;
  if (/\bfirst you(?:'ll| will) need to connect\b/.test(lower)) return true;
  if (/approve.*oauth|oauth.*approv/.test(lower) && /browser|window/.test(lower)) return true;
  return false;
}

/**
 * True if this process still holds OAuth tokens from `bc connect` (access and/or refresh).
 * Wider than hasValidAutodeskAccessToken() so we still scrub bad "run bc connect" replies when the
 * access token is in its last minute or just expired but refresh is present.
 */
function oauthSessionPresentInMemory() {
  return Boolean(autodeskTokenCache?.accessToken || autodeskTokenCache?.refreshToken);
}

/** Model sometimes ignores CONNECTED facts; enforce correct server state in Slack. */
function replaceFalseBcConnectDemand(userText, reply) {
  if (!oauthSessionPresentInMemory()) return reply;
  if (!assistantReplyIncorrectlyDemandsBcConnect(reply)) return reply;
  const discovery = userTextLooksLikeBcDiscoveryOrCasualList(userText);
  let body =
    'BuildingConnected **is already signed in** on this bot (tokens are stored on disk after `bc connect`, so restarts keep the session). Check `bc auth status`. You do **not** need `bc connect` again unless auth failed or you deleted the token file.\n\n';
  if (discovery) {
    body +=
      'Geographic distance (e.g. 50 miles) is **not** computed in the bot—API search is keyword-only on JSON. ' +
      'Ask again in plain language (city, county, trade) when Autodesk allows listing. **BC_PRO_SUBSCRIPTION_REQUIRED** means Autodesk blocked project listing over the API for this account.';
  } else {
    body += 'Ask in plain language about BuildingConnected, or use `bc help` for technical commands.';
  }
  return body;
}
const REFUSAL_PATTERNS = [
  /not able to send direct messages/i,
  /don'?t have the ability to initiate communications/i,
  /cannot message users directly/i,
  /access external messaging systems/i,
  /don'?t have the ability to directly access/i,
  /cannot directly access/i,
  /i can only see and respond to messages in this slack conversation/i,
];
const SUPPORTED_AUDIO_EXTENSIONS = ['m4a', 'mp3', 'wav', 'ogg', 'webm', 'mp4'];
/** Cached `formats.color` array from GET /serverinfo (Smartsheet cell format uses palette indices, not raw hex objects). */
let smartsheetFormatsColorTableCache = null;
let autodeskTokenCache = null;
let autodeskOauthStateCache = null;

function getAutodeskTokenStorePath() {
  const explicit = String(process.env.AUTODESK_TOKEN_STORE_PATH || '').trim();
  if (explicit) return explicit;
  return path.join(__dirname, '.autodesk-tokens.json');
}

function persistAutodeskTokenCache() {
  if (!autodeskTokenCache?.accessToken && !autodeskTokenCache?.refreshToken) return;
  try {
    const payload = JSON.stringify({
      accessToken: autodeskTokenCache.accessToken,
      refreshToken: autodeskTokenCache.refreshToken,
      expiresAtMs: autodeskTokenCache.expiresAtMs,
      tokenType: autodeskTokenCache.tokenType,
      obtainedAtMs: autodeskTokenCache.obtainedAtMs,
    });
    fs.writeFileSync(getAutodeskTokenStorePath(), payload, { encoding: 'utf8', mode: 0o600 });
  } catch (err) {
    console.warn('Could not persist Autodesk tokens to disk:', err?.message || err);
  }
}

function loadAutodeskTokenCacheFromDisk() {
  if (!process.env.AUTODESK_CLIENT_ID || !process.env.AUTODESK_CLIENT_SECRET) return;
  const storePath = getAutodeskTokenStorePath();
  try {
    if (!fs.existsSync(storePath)) return;
    const data = JSON.parse(fs.readFileSync(storePath, 'utf8'));
    const accessToken = String(data?.accessToken || '').trim();
    const refreshToken = String(data?.refreshToken || '').trim();
    const expiresAtMs = Number(data?.expiresAtMs) || 0;
    if (!accessToken && !refreshToken) return;
    autodeskTokenCache = {
      accessToken: accessToken || null,
      refreshToken: refreshToken || null,
      expiresAtMs,
      tokenType: String(data?.tokenType || 'Bearer'),
      obtainedAtMs: Number(data?.obtainedAtMs) || Date.now(),
    };
    console.log('Loaded Autodesk OAuth tokens from disk (survives bot restart).');
  } catch (err) {
    console.warn('Could not load Autodesk tokens from disk:', err?.message || err);
  }
}

function clearAutodeskTokenStoreFile() {
  try {
    const storePath = getAutodeskTokenStorePath();
    if (fs.existsSync(storePath)) fs.unlinkSync(storePath);
  } catch (_) {
    /* ignore */
  }
}
const SHEET_INTENT_SYSTEM_PROMPT = `You classify user messages for Smartsheet edits connected to this Slack bot.

Return ONE JSON object only. No markdown, no prose.

Schema:
{
  "action": "add" | "update" | "format" | "clear_sheet" | "read_cell" | "none",
  "confidence": number,
  "fields": { "<column title>": "<value>" },
  "row_id": number | null,
  "match_column": string | null,
  "match_value": string | null,
  "format": {
    "column": "<column title>",
    "row_start": number,
    "row_end": number,
    "background_color": "#RRGGBB"
  } | null,
  "read": {
    "column_index": number | null,
    "column_title": string | null,
    "row_number": number | null
  } | null,
  "clarification": string | null
}

Rules:
- Use column titles EXACTLY as given in the column list (case and spelling).
- action "add": user wants a new row. Put all cells to set in fields.
- action "update": user wants to change existing row(s). Put changed cells in fields.
- action "format": user wants cell formatting (background color) for a column across a row number range.
- action "read_cell": user asks what VALUE is in one cell (contents / what is in / show me) at a numeric row and either "column N" (1-based from left) OR a column title. Set read.row_number and read.column_index when they say "column 3 row 1" (column_index=3, row_number=1). Set read.column_title when they name a real column title. The answer is the cell value, NOT the column header name unless the cell literally contains that text.
- If the user wants green (or highlighted) cells turned back to white / no fill, use action "format" with background_color "#FFFFFF" and the narrowest row range + column you can infer; if they mean the whole sheet, set row_start to 1 and row_end to the last row number you can infer from context, otherwise set clarification.
- action "clear_sheet": user wants to wipe / empty most cells and "start fresh" (clear all editable cells in every row). NOT for deleting rows unless they explicitly say delete/remove rows.
- If user gives a numeric Smartsheet row id, set row_id.
- If user references a row indirectly ("where Task is X"), set match_column and match_value to UNIQUE identify it.
- If the user clearly mentions Smartsheet/sheet/row/column/status but you cannot map to columns, set clarification with ONE question listing the exact column titles they can use.
- If ambiguous OR you are not confident this is a sheet edit, set action "none", confidence <= 0.4, fields {}.
- If you need one specific detail, set clarification to ONE short question.
- confidence 0..1.

Examples:
User: "Add a row: Task=Foo, Status=Open" -> add, high confidence.
User: "What's the weather?" -> none, low confidence.
User: "Empty every cell and start over with a clean sheet" -> clear_sheet, high confidence.
User: "Change the green cells to white again" -> format, background_color "#FFFFFF", include column + row range if possible; if they clearly mean every green cell in the sheet, row_start 1 row_end large enough to cover sheet rows.
User: "Change Bid Out to Bid Accepted for the Boston job" -> update: find the row whose cells mention Boston, then set the cell that currently equals Bid Out to Bid Accepted (use match_column/match_value if needed).
User: "What is in column 3 row 1" -> read_cell, read: { column_index: 3, row_number: 1 }, high confidence.`;

/** Bumped when shipping notable behavior changes; echoed by /health/slack. */
const AGENT_CODE_BUILD = 'transform-sim-mcp-branch-v14-smartsheet-bold-emphasize-nl';

function fetchTimeoutSignal(ms) {
  if (typeof AbortSignal !== 'undefined' && typeof AbortSignal.timeout === 'function') {
    return AbortSignal.timeout(ms);
  }
  const c = new AbortController();
  setTimeout(() => c.abort(), ms);
  return c.signal;
}

function verifySlack(req) {
  const sig = req.headers['x-slack-signature'];
  const ts = req.headers['x-slack-request-timestamp'];
  if (!sig || !ts || !process.env.SLACK_SIGNING_SECRET || !req.rawBody) return false;
  const hmac = crypto.createHmac('sha256', process.env.SLACK_SIGNING_SECRET);
  hmac.update(`v0:${ts}:${req.rawBody}`);
  return sig === `v0=${hmac.digest('hex')}`;
}

/** Why verifySlack failed; for logs only (never log secrets or raw body). */
function slackSignatureRejectReason(req) {
  if (!req.headers['x-slack-signature']) return 'missing X-Slack-Signature';
  if (!req.headers['x-slack-request-timestamp']) return 'missing X-Slack-Request-Timestamp';
  if (!process.env.SLACK_SIGNING_SECRET) return 'SLACK_SIGNING_SECRET unset';
  if (!req.rawBody) {
    return 'rawBody missing (express.json verify must set req.rawBody for signing — check middleware order)';
  }
  return 'signature mismatch (wrong SLACK_SIGNING_SECRET, or body altered before verify)';
}

function normalizeSlackText(text) {
  if (!text) return '';
  // Remove Slack user mentions: <@U123ABC> or <@U123ABC|display name> (pipe form is common in channels).
  return text.replace(/<@[^>]+>/g, '').replace(/\s+/g, ' ').trim();
}

/** When Slack omits or truncates `event.text`, Block Kit / rich_text may still carry the words (common for mentions). */
function slackPlainTextFromBlocks(event) {
  const blocks = event?.blocks;
  if (!Array.isArray(blocks)) return '';
  const parts = [];
  for (const b of blocks) {
    if (!b || typeof b !== 'object') continue;
    if (b.type === 'section' && typeof b.text?.text === 'string') parts.push(b.text.text);
    if (b.type === 'header' && typeof b.text?.text === 'string') parts.push(b.text.text);
    if (b.type === 'rich_text' && Array.isArray(b.elements)) {
      for (const el of b.elements) {
        if (el?.type === 'rich_text_section' && Array.isArray(el.elements)) {
          for (const bit of el.elements) {
            if (bit?.type === 'text' && typeof bit.text === 'string') parts.push(bit.text);
            if (bit?.type === 'user' && bit.user_id) parts.push(`<@${bit.user_id}>`);
          }
        }
      }
    }
  }
  return parts.join(' ');
}

/**
 * Plain user message text: Slack often repeats the same words in `event.text` and Block Kit /
 * rich_text blocks. Joining both blindly duplicates the message; merge so each phrase appears once.
 */
function slackUserMessagePlainText(event) {
  const textRaw = String(event.text || '').trim();
  const blocksRaw = String(slackPlainTextFromBlocks(event) || '').trim();
  if (!blocksRaw) return normalizeSlackText(textRaw);
  if (!textRaw) return normalizeSlackText(blocksRaw);
  const a = normalizeSlackText(textRaw);
  const b = normalizeSlackText(blocksRaw);
  if (a === b) return a;
  if (b.length >= 8 && a.includes(b)) return a;
  if (a.length >= 8 && b.includes(a)) return b;
  return normalizeSlackText(`${textRaw} ${blocksRaw}`.trim());
}

/**
 * Thread target for outbound replies. Local signed tests set `_local_post_to_channel` because the
 * synthetic `ts` is not a real Slack message — threading under it hides replies from the main timeline.
 */
function slackReplyThreadTs(event) {
  if (!event || event._local_post_to_channel) return undefined;
  const t = event.thread_ts || event.ts;
  return t == null ? undefined : String(t);
}

const SLACK_THREAD_REPLIES_LIMIT = Math.min(
  100,
  Math.max(5, Number(process.env.SLACK_THREAD_REPLIES_LIMIT || 40) || 40)
);
const SLACK_THREAD_CONTEXT_MAX_CHARS = Math.min(
  32000,
  Math.max(1500, Number(process.env.SLACK_THREAD_CONTEXT_MAX_CHARS || 14000) || 14000)
);

/**
 * Fetches prior messages in the same Slack thread so follow-ups ("roofing") inherit context
 * ("Stanislaus County") without repeating. Requires `conversations.replies` OAuth scope.
 */
async function fetchSlackThreadPriorTranscript({ channel, threadRootTs, currentMessageTs }) {
  if (!channel || !threadRootTs || !currentMessageTs) return '';
  try {
    const res = await slack.conversations.replies({
      channel: String(channel),
      ts: String(threadRootTs),
      limit: SLACK_THREAD_REPLIES_LIMIT,
      inclusive: true,
    });
    if (!res.ok) {
      console.warn('[slack] conversations.replies not ok:', res.error || res);
      return '';
    }
    const msgs = Array.isArray(res.messages) ? res.messages : [];
    const lines = [];
    for (const m of msgs) {
      if (String(m.ts) === String(currentMessageTs)) continue;
      const raw = [m.text, slackPlainTextFromBlocks(m)].filter(Boolean).join(' ');
      const line = normalizeSlackText(raw);
      if (!line) continue;
      const fromBot =
        (slackBotUserId && m.user && String(m.user) === String(slackBotUserId)) ||
        Boolean(m.bot_id && !m.user);
      const label = fromBot ? 'Assistant' : 'User';
      lines.push(`${label}: ${line}`);
    }
    let block = lines.join('\n');
    if (!block) return '';
    if (block.length > SLACK_THREAD_CONTEXT_MAX_CHARS) {
      block = `…(older thread truncated)\n${block.slice(-SLACK_THREAD_CONTEXT_MAX_CHARS)}`;
    }
    return (
      'Earlier messages in this Slack thread (oldest to newest; use for context):\n' +
      `${block}\n` +
      '---\n' +
      'Latest message from the user:\n'
    );
  } catch (e) {
    console.warn('[slack] conversations.replies failed:', e?.message || e);
    return '';
  }
}

function shouldReplaceRefusal(text) {
  return REFUSAL_PATTERNS.some((pattern) => pattern.test(text || ''));
}

function fallbackHelpfulReply(userText) {
  const lower = (userText || '').toLowerCase();
  if (lower.includes('introduce') || lower.includes('intro')) {
    return "Hi, I'm your Transform Sim Slack assistant. I can help with drafting messages, summaries, brainstorming, and quick answers right here in this thread.";
  }
  if (lower.includes('message') || lower.includes('dm') || lower.includes('direct')) {
    return 'I can help you draft exactly what to send. Share who it is for and your goal, and I will write a ready-to-send message.';
  }
  return 'I am here and working in Slack. Tell me what you want to do, and I will help directly in this thread.';
}

function isAudioSlackFile(file) {
  const mimeType = file?.mimetype || '';
  const extension = (file?.filetype || '').toLowerCase();
  return mimeType.startsWith('audio/') || SUPPORTED_AUDIO_EXTENSIONS.includes(extension);
}

async function downloadSlackFile(url) {
  const response = await fetch(url, {
    headers: { Authorization: `Bearer ${process.env.SLACK_BOT_TOKEN}` },
  });
  if (!response.ok) {
    throw new Error(`Slack file download failed: ${response.status}`);
  }
  const arrayBuffer = await response.arrayBuffer();
  return Buffer.from(arrayBuffer);
}

async function transcribeAudioBuffer(buffer, filename, mimeType) {
  if (!openai) {
    throw new Error('OPENAI_API_KEY is required for audio transcription');
  }
  const file = new File([buffer], filename || 'audio-message.m4a', {
    type: mimeType || 'application/octet-stream',
  });
  const transcription = await openai.audio.transcriptions.create({
    file,
    model: 'whisper-1',
  });
  return transcription.text;
}

function extractAnthropicAssistantText(message) {
  const blocks = message?.content;
  if (!Array.isArray(blocks)) return '';
  const parts = [];
  for (const block of blocks) {
    if (block && block.type === 'text' && typeof block.text === 'string') parts.push(block.text);
  }
  return parts.join('\n').trim();
}

async function generateAssistantReply(userText, options = {}) {
  const { threadPriorBlock = '' } = options;
  const bodyForModel = threadPriorBlock ? `${threadPriorBlock}${userText}` : userText;
  let systemPrompt = smartsheetConfigured()
    ? ASSISTANT_SYSTEM_PROMPT_WITH_SMARTSHEET
    : ASSISTANT_SYSTEM_PROMPT;
  systemPrompt = appendBuildingConnectedAssistantGuidance(systemPrompt);
  if (threadPriorBlock) {
    systemPrompt +=
      ' The user message may include a transcript of earlier Slack thread messages before the latest line; treat that transcript as authoritative context for follow-up questions.';
  }
  const prefix = buildingConnectedUserMessagePrefix();
  const augmentedUser = prefix ? `${prefix}\n\n${bodyForModel}` : bodyForModel;
  const message = await client.messages.create(
    {
      model: 'claude-sonnet-4-20250514',
      system: systemPrompt,
      max_tokens: 1024,
      temperature: 0,
      messages: [{ role: 'user', content: augmentedUser }],
    },
    { signal: fetchTimeoutSignal(120_000) }
  );
  const replyText = extractAnthropicAssistantText(message) || message.content?.[0]?.text || '';
  const trimmed = String(replyText || '').trim();
  let finalText = trimmed || fallbackHelpfulReply(userText);
  if (shouldReplaceRefusal(finalText)) finalText = fallbackHelpfulReply(userText);
  return replaceFalseBcConnectDemand(bodyForModel, finalText);
}

function buildingConnectedConfigured() {
  return Boolean(process.env.AUTODESK_CLIENT_ID && process.env.AUTODESK_CLIENT_SECRET);
}

function getAutodeskRedirectUri(req = null) {
  const explicit = String(process.env.AUTODESK_REDIRECT_URI || '').trim();
  if (explicit) return explicit;
  const publicBase = String(process.env.PUBLIC_BASE_URL || '').trim().replace(/\/+$/, '');
  if (publicBase) return `${publicBase}/autodesk/callback`;
  if (req && req.headers) {
    const host = req.headers['x-forwarded-host'] || req.headers.host;
    const proto = req.headers['x-forwarded-proto'] || req.protocol || 'http';
    if (host) return `${proto}://${host}/autodesk/callback`;
  }
  return 'http://localhost:3000/autodesk/callback';
}

function getBotBaseUrl() {
  const publicBase = String(process.env.PUBLIC_BASE_URL || '').trim().replace(/\/+$/, '');
  if (publicBase) return publicBase;
  const redirect = String(process.env.AUTODESK_REDIRECT_URI || '').trim();
  if (redirect) {
    try {
      const u = new URL(redirect);
      return `${u.protocol}//${u.host}`;
    } catch {
      return '';
    }
  }
  return '';
}

function getAutodeskScopes() {
  return String(process.env.AUTODESK_SCOPES || DEFAULT_AUTODESK_SCOPES)
    .trim()
    .replace(/\s+/g, ' ');
}

function buildAutodeskBasicAuthHeader() {
  const creds = `${process.env.AUTODESK_CLIENT_ID}:${process.env.AUTODESK_CLIENT_SECRET}`;
  return `Basic ${Buffer.from(creds, 'utf8').toString('base64')}`;
}

async function exchangeAutodeskToken(formParams) {
  const response = await fetch(AUTODESK_TOKEN_URL, {
    method: 'POST',
    headers: {
      Authorization: buildAutodeskBasicAuthHeader(),
      'Content-Type': 'application/x-www-form-urlencoded',
    },
    body: new URLSearchParams(formParams).toString(),
    signal: fetchTimeoutSignal(15000),
  });
  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`Autodesk OAuth error ${response.status}: ${errorText}`);
  }
  return response.json();
}

function setAutodeskTokenCache(tokenData) {
  const now = Date.now();
  const accessToken = String(tokenData?.access_token || '').trim();
  const newRefresh = String(tokenData?.refresh_token || '').trim();
  const previousRefresh = autodeskTokenCache?.refreshToken || null;
  const refreshToken = newRefresh || previousRefresh || null;
  const expiresInSec = Number(tokenData?.expires_in || 0);
  if (!accessToken) throw new Error('Autodesk OAuth did not return access_token.');
  autodeskTokenCache = {
    accessToken,
    refreshToken,
    expiresAtMs: now + Math.max(60, expiresInSec) * 1000,
    tokenType: String(tokenData?.token_type || 'Bearer'),
    obtainedAtMs: now,
  };
  persistAutodeskTokenCache();
  return autodeskTokenCache.accessToken;
}

function hasValidAutodeskAccessToken() {
  const now = Date.now();
  return Boolean(autodeskTokenCache?.accessToken && autodeskTokenCache?.expiresAtMs > now + 60_000);
}

function createAutodeskOauthState() {
  const state = crypto.randomBytes(24).toString('hex');
  autodeskOauthStateCache = { state, createdAtMs: Date.now() };
  return state;
}

function isValidAutodeskOauthState(receivedState) {
  const cached = autodeskOauthStateCache;
  if (!cached || !cached.state || !receivedState) return false;
  const maxAgeMs = 10 * 60 * 1000;
  if (Date.now() - cached.createdAtMs > maxAgeMs) return false;
  return cached.state === String(receivedState);
}

function getAutodeskAuthorizeUrl(req = null) {
  const state = createAutodeskOauthState();
  const redirectUri = getAutodeskRedirectUri(req);
  const params = new URLSearchParams({
    response_type: 'code',
    client_id: process.env.AUTODESK_CLIENT_ID || '',
    redirect_uri: redirectUri,
    scope: getAutodeskScopes(),
    state,
  });
  return `${AUTODESK_AUTHORIZE_URL}?${params.toString()}`;
}

async function exchangeAutodeskAuthorizationCode(code, redirectUri) {
  const tokenData = await exchangeAutodeskToken({
    grant_type: 'authorization_code',
    code: String(code || ''),
    redirect_uri: redirectUri,
  });
  return setAutodeskTokenCache(tokenData);
}

async function refreshAutodeskAccessToken(refreshToken) {
  const tokenData = await exchangeAutodeskToken({
    grant_type: 'refresh_token',
    refresh_token: String(refreshToken || ''),
    scope: getAutodeskScopes(),
  });
  return setAutodeskTokenCache(tokenData);
}

async function getAutodeskAccessToken() {
  if (!buildingConnectedConfigured()) {
    throw new Error(
      'AUTODESK_CLIENT_ID and AUTODESK_CLIENT_SECRET are required for BuildingConnected integration.'
    );
  }
  if (hasValidAutodeskAccessToken()) {
    return autodeskTokenCache.accessToken;
  }
  if (!autodeskTokenCache?.refreshToken) {
    throw new Error(
      'BuildingConnected requires user authorization. Run `bc connect` in Slack, open the link, authorize Autodesk, then retry.'
    );
  }
  try {
    return await refreshAutodeskAccessToken(autodeskTokenCache.refreshToken);
  } catch (err) {
    autodeskTokenCache = null;
    clearAutodeskTokenStoreFile();
    throw new Error(
      `Autodesk token refresh failed. Run \`bc connect\` again. Details: ${String(err?.message || err)}`
    );
  }
}

async function buildingConnectedRequest(path, options = {}) {
  const token = await getAutodeskAccessToken();
  const response = await fetch(`${BUILDING_CONNECTED_BASE_URL}${path}`, {
    method: options.method || 'GET',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json',
      ...(options.headers || {}),
    },
    body: options.body ? JSON.stringify(options.body) : undefined,
    signal: fetchTimeoutSignal(20000),
  });
  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`BuildingConnected API error ${response.status}: ${errorText}`);
  }
  return response.json();
}

async function buildingConnectedRequestWithStatus(path, options = {}) {
  try {
    const data = await buildingConnectedRequest(path, options);
    return { ok: true, data };
  } catch (err) {
    const message = String(err?.message || err);
    const statusMatch = message.match(/error\s+(\d{3})/i);
    const status = statusMatch ? Number(statusMatch[1]) : null;
    return { ok: false, error: err, status, message };
  }
}

function projectsArrayFromBcApiResult(apiResult) {
  if (Array.isArray(apiResult)) return apiResult;
  if (Array.isArray(apiResult?.results)) return apiResult.results;
  if (Array.isArray(apiResult?.data)) return apiResult.data;
  return [];
}

function summarizeBuildingConnectedProjects(apiResult, limit = 10) {
  const records = projectsArrayFromBcApiResult(apiResult);
  if (records.length === 0) {
    return 'No projects were returned.';
  }
  const lines = records.slice(0, Math.max(1, limit)).map((project) => {
    const id = project?.id || project?.projectId || 'unknown-id';
    const name = project?.name || project?.projectName || 'Untitled project';
    const location = project?.location?.city || project?.city || project?.location || '';
    return location ? `- ${name} (${id}) — ${location}` : `- ${name} (${id})`;
  });
  const extra = records.length > lines.length ? `\n…and ${records.length - lines.length} more.` : '';
  return `BuildingConnected projects:\n${lines.join('\n')}${extra}`;
}

const BC_TRADE_TERMS = [
  'roofing',
  'hvac',
  'electrical',
  'plumbing',
  'concrete',
  'drywall',
  'framing',
  'paving',
  'demolition',
  'excavation',
  'masonry',
  'steel',
  'landscaping',
  'painting',
  'insulation',
  'flooring',
  'glazing',
  'waterproofing',
];

const BC_SEARCH_STOPWORDS = new Set([
  'building',
  'connected',
  'buildingconnected',
  'transform',
  'slack',
  'assistant',
  'please',
  'could',
  'would',
  'should',
  'within',
  'inside',
  'around',
  'radius',
  'miles',
  'mile',
  'kilometers',
  'kilometer',
  'projects',
  'project',
  'posted',
  'bidding',
  'preconstruction',
  'find',
  'show',
  'list',
  'what',
  'which',
  'are',
  'there',
  'any',
  'some',
  'help',
  'with',
  'about',
  'near',
  'from',
  'that',
  'this',
  'your',
  'have',
  'does',
  'will',
  'just',
  'only',
  'like',
  'into',
  'using',
  'through',
  'california',
  'same',
  'county',
  'region',
  'available',
  'area',
  'smartsheet',
  'the',
  'a',
  'an',
  'give',
  'can',
  'you',
  'me',
  'for',
  'not',
]);

/** Drop phrase-shaped junk (e.g. "the same county as modesto") from keyword search. */
function filterGarbageBcNeedles(needles) {
  const out = new Set();
  for (const raw of needles) {
    const s = String(raw).toLowerCase().trim();
    if (s.length < 3 || s.length > 42) continue;
    const words = s.split(/\s+/).filter(Boolean);
    if (words.length > 3) continue;
    if (words.length > 1 && /\b(the|same|county|around|within|that|this|give|your|our|any|all)\b/.test(s)) {
      continue;
    }
    if (BC_SEARCH_STOPWORDS.has(s)) continue;
    if (/^(the|a|an|is|are|can|how|what|where|when|did|does|will)$/i.test(s)) continue;
    out.add(s);
  }
  return out;
}

function extractBcSearchNeedles(userText) {
  const raw = String(userText || '');
  const lower = raw.toLowerCase();
  const needles = new Set();
  for (const t of BC_TRADE_TERMS) {
    if (new RegExp(`\\b${t}\\b`, 'i').test(lower)) needles.add(t);
  }
  const withinM = /\bwithin\s+\d+\s*(?:mi|mile|miles|km)?\s+of\s+([a-z][a-z\s,.-]{2,60}?)(?=\s*[,.?!]|$|\s+and\s|\s+or\s|\s+within|\s+for\b)/i.exec(
    lower
  );
  if (withinM?.[1]) {
    const city = withinM[1]
      .replace(/\s*,\s*(ca|california)\s*$/i, '')
      .replace(/,\s*ca\b\s*$/i, '')
      .replace(/\s+/g, ' ')
      .trim();
    if (city.length >= 4) needles.add(city);
  }
  const nearM = /\bnear\s+([a-z][a-z\s,.-]{2,60}?)(?=\s*[,.?!]|$|\s+and\s|\s+within|\s+for\b)/i.exec(lower);
  if (nearM?.[1]) {
    const city = nearM[1]
      .replace(/\s*,\s*(ca|california)\s*$/i, '')
      .replace(/\s+/g, ' ')
      .trim();
    if (city.length >= 4 && !BC_SEARCH_STOPWORDS.has(city)) needles.add(city);
  }
  const countyAs = /\b(?:same\s+)?county\s+as\s+([a-z0-9][a-z0-9\s,.-]{2,50}?)(?=\s*[,.?!]|$|\s+and\s|\s+for\b)/i.exec(lower);
  if (countyAs?.[1]) {
    const city = countyAs[1]
      .replace(/\s*,\s*(ca|california)\s*$/i, '')
      .replace(/,\s*ca\b\s*$/i, '')
      .replace(/\s+/g, ' ')
      .trim();
    if (city.length >= 3 && !BC_SEARCH_STOPWORDS.has(city)) needles.add(city);
  }
  // "in Fresno, CA" — at most three short words before the comma (avoid "in the same county as modesto, ca").
  const inCityState =
    /\bin\s+([a-z][a-z0-9]+(?:\s+[a-z][a-z0-9]+){0,2}),\s*(?:ca|california|tx|texas|ny|new\s+york|fl|florida|wa|washington|or|oregon)\b/i.exec(
      lower
    );
  if (inCityState?.[1]) {
    const city = inCityState[1].replace(/\s+/g, ' ').trim();
    if (
      city.length >= 3 &&
      !BC_SEARCH_STOPWORDS.has(city) &&
      !/\b(the|same|county|around|within|that|this)\b/.test(city)
    ) {
      needles.add(city);
    }
  }
  const jobInPlace = /\b(?:jobs?|bids?|projects?|work)\s+in\s+([a-z][a-z0-9]{2,30})\b/i.exec(lower);
  if (jobInPlace?.[1]) {
    const w = String(jobInPlace[1]).toLowerCase();
    if (!/^(the|a|an|my|our|all|any|same|this|that|your)$/i.test(w) && !BC_SEARCH_STOPWORDS.has(w)) needles.add(w);
  }
  const caps = raw.match(/\b[A-Z][a-z]{3,}\b/g) || [];
  for (const w of caps) {
    const lw = w.toLowerCase();
    if (['Building', 'Connected', 'Transform', 'Sim', 'Slack', 'Same', 'County', 'Available'].includes(w)) continue;
    if (BC_SEARCH_STOPWORDS.has(lw)) continue;
    needles.add(lw);
  }
  for (const n of [...needles]) {
    if (BC_SEARCH_STOPWORDS.has(String(n).toLowerCase())) needles.delete(n);
  }
  return filterGarbageBcNeedles(needles);
}

async function buildingConnectedKeywordSearchFromNeedles(needlesArr) {
  const arr = [...new Set(needlesArr.map((n) => String(n).toLowerCase().trim()).filter(Boolean))];
  if (!arr.length) {
    return '*BuildingConnected search (server)*\nNo keywords to search.';
  }
  const pr = await buildingConnectedRequestWithStatus('/projects');
  if (!pr.ok) {
    const m = pr.message || '';
    let head;
    if (m.includes('403') && m.includes('BC_PRO_SUBSCRIPTION_REQUIRED')) {
      const kw =
        arr.length && arr.every((s) => s.length <= 32)
          ? `I would have searched for: ${arr.join(', ')}.`
          : 'Your account blocked automated project listing over the API.';
      head =
        '*BuildingConnected*\n' +
        'Autodesk returned **`BC_PRO_SUBSCRIPTION_REQUIRED`**: this workspace cannot list BuildingConnected projects through the API, so I cannot run a live keyword search from the server.\n\n' +
        `${kw}\n\n`;
    } else {
      head =
        `*BuildingConnected search (server)*\nCould not load projects via API: ${m.slice(0, 700)}\n\n` +
        `Keywords requested: ${arr.join(', ')}.\n\n`;
    }
    head += '_This bot uses the BuildingConnected API only (no web UI / Playwright fallback)._';
    return head;
  }
  const records = projectsArrayFromBcApiResult(pr.data);
  if (!records.length) {
    return '*BuildingConnected search (server)*\nAPI returned zero projects.';
  }
  const matched = records.filter((p) => {
    const blob = JSON.stringify(p).toLowerCase();
    return arr.every((n) => blob.includes(n));
  });
  const lines = matched.slice(0, 15).map((project) => {
    const id = project?.id || project?.projectId || 'unknown-id';
    const name = project?.name || project?.projectName || 'Untitled project';
    const location = project?.location?.city || project?.city || project?.location || '';
    return location ? `- ${name} (${id}) — ${location}` : `- ${name} (${id})`;
  });
  let msg = `*BuildingConnected keyword search (server)*\nKeywords (all must match somewhere in the project JSON): ${arr.join(
    ', '
  )}\n`;
  msg += `Scanned ${records.length} project(s), ${matched.length} matched.\n\n`;
  msg += matched.length
    ? lines.join('\n')
    : '_No matches._ Try naming a city or trade to search again, or ask to list your Autodesk / BuildingConnected projects in plain English.';
  if (matched.length > 15) msg += `\n…(showing first 15 of ${matched.length})`;
  msg +=
    '\n\n_This is text matching on the API payload, not GIS distance. Say what city, county, or trade you care about and we will search again._';
  return msg;
}

async function handleBuildingConnectedDiscoveryAnswer(userText, options = {}) {
  const { threadPriorBlock = '' } = options;
  const combined = threadPriorBlock ? `${threadPriorBlock}${userText}` : userText;
  if (!buildingConnectedConfigured()) return null;
  if (!userTextLooksLikeBcDiscoveryOrCasualList(combined)) return null;
  try {
    await getAutodeskAccessToken();
  } catch {
    return null;
  }
  const casualList = looksLikeCasualProjectListingAsk(userText);
  const needles = extractBcSearchNeedles(combined);
  if (needles.size === 0) {
    if (casualList) {
      const result = await buildingConnectedRequest('/projects');
      return summarizeBuildingConnectedProjects(result, 15);
    }
    if (looksLikeGeographicOrBidDiscovery(combined)) {
      const result = await buildingConnectedRequest('/projects');
      const head = summarizeBuildingConnectedProjects(result, 12);
      return `${head}\n\n_I could not pull a city or trade out of that message to narrow the list. Reply with a place name (e.g. Modesto) or a trade (e.g. electrical) and I will search again._`;
    }
    return null;
  }
  return buildingConnectedKeywordSearchFromNeedles([...needles]);
}

/** Strip punctuation glued to Slack / natural-language words so `bc projects, what…` still matches `projects`. */
function stripBcCliGlue(s) {
  return String(s || '')
    .trim()
    .replace(/^[`'"([{]+/, '')
    .replace(/[`'",.:;!?)\]}]+$/g, '')
    .trim();
}

function normalizeBcCliToken(s) {
  return stripBcCliGlue(s).toLowerCase();
}

async function handleBuildingConnectedCommand(userText) {
  const trimmed = String(userText || '').trim();
  if (!/^bc\b|^buildingconnected\b/i.test(trimmed)) return null;
  const parts = trimmed.split(/\s+/);
  const command = normalizeBcCliToken(parts[0]);
  const action = normalizeBcCliToken(parts[1]);
  if (!buildingConnectedConfigured()) {
    return (
      'BuildingConnected is not configured. Set `AUTODESK_CLIENT_ID` and `AUTODESK_CLIENT_SECRET` in `.env`, then restart the bot.'
    );
  }
  try {
    if (!action || action === 'help') {
      return (
        'BuildingConnected commands:\n' +
        '- bc connect\n' +
        '- bc auth status\n' +
        '- bc projects [limit]\n' +
        '- bc search <keywords…>  (all keywords must match text in the project JSON)\n' +
        '- bc web status   (shows that UI/web fallback is disabled in this bot)\n' +
        '- bc project PROJECT_ID\n' +
        '- bc whoami\n' +
        'Notes: configure AUTODESK_CLIENT_ID, AUTODESK_CLIENT_SECRET, AUTODESK_REDIRECT_URI (or PUBLIC_BASE_URL), and optional AUTODESK_SCOPES. ' +
        'This deployment is **API-only** for BuildingConnected (no Playwright / page scraping).'
      );
    }
    if (action === 'connect') {
      const base = getBotBaseUrl();
      const connectUrl = base ? `${base}/autodesk/connect` : 'http://localhost:3000/autodesk/connect';
      const isLocalHost = /localhost|127\.0\.0\.1/i.test(connectUrl);
      let msg =
        '**BuildingConnected — sign in with Autodesk (3-legged OAuth)**\n\n' +
        'Open this link **in a browser on the same computer where the bot is running** (`npm start`):\n' +
        `${connectUrl}\n\n`;
      if (isLocalHost) {
        msg +=
          '**Why Slack shows `localhost`:** only that machine can open it. Slack on your phone or another PC cannot reach your laptop’s `localhost`.\n\n' +
          '**To use the link from any device:** expose port 3000 with a public HTTPS URL, then set it in `.env` and restart the bot:\n' +
          '1. Run ngrok (example): `ngrok http 3000`\n' +
          '2. Set `PUBLIC_BASE_URL=https://<your-ngrok-host>` (no trailing slash)\n' +
          '3. If `.env` has `AUTODESK_REDIRECT_URI=http://localhost:...`, delete it or set it to `https://<your-ngrok-host>/autodesk/callback` (that variable overrides `PUBLIC_BASE_URL`).\n' +
          '4. In the Autodesk app, add **Callback URL** `https://<your-ngrok-host>/autodesk/callback` (exact match)\n' +
          '5. Restart the bot, run `bc connect` again — this reply will show the https URL instead of localhost.\n';
      } else {
        msg +=
          'Confirm this same base URL is allowed as an Autodesk OAuth **Callback URL** ending in `/autodesk/callback`.\n';
      }
      return msg;
    }
    if (action === 'auth' && normalizeBcCliToken(parts[2]) === 'status') {
      if (!autodeskTokenCache?.accessToken) {
        return 'Autodesk auth: not connected yet. Run `bc connect` and complete the browser authorization.';
      }
      const expiresInSec = Math.max(0, Math.floor((autodeskTokenCache.expiresAtMs - Date.now()) / 1000));
      return `Autodesk auth: connected. Access token expires in ~${expiresInSec}s.`;
    }
    if (action === 'web') {
      const { formatWebStatus } = require('./bc-web-fallback');
      const sub = normalizeBcCliToken(parts[2]);
      if (!sub || sub === 'status') return formatWebStatus();
      if (sub === 'search') {
        return (
          '*BuildingConnected*\n' +
          'Web / UI search is **disabled** in this bot (API only). Use `bc search <keywords>` when the API project list works, or ask in plain language in a thread.'
        );
      }
      return 'Usage: `bc web status` (web search is disabled).';
    }
    if (action === 'search') {
      const tokens = parts.slice(2).map((p) => normalizeBcCliToken(p)).filter(Boolean);
      if (!tokens.length) {
        return 'Usage: `bc search roofing modesto` — space-separated keywords; **all** must match somewhere in each project payload.';
      }
      return await buildingConnectedKeywordSearchFromNeedles(tokens);
    }
    if (action === 'projects') {
      const limit = Number(stripBcCliGlue(parts[2] || '') || 10);
      const safeLimit = Number.isFinite(limit) && limit > 0 ? Math.min(limit, 25) : 10;
      const result = await buildingConnectedRequest('/projects');
      return summarizeBuildingConnectedProjects(result, safeLimit);
    }
    if (action === 'project') {
      const projectId = stripBcCliGlue(parts[2] || '');
      if (!projectId) return 'Usage: `bc project PROJECT_ID`';
      const result = await buildingConnectedRequest(`/projects/${encodeURIComponent(projectId)}`);
      return `Project ${projectId}:\n${JSON.stringify(result, null, 2).slice(0, 3500)}`;
    }
    if (action === 'whoami') {
      const me = await buildingConnectedRequest('/users/me');
      return `Connected Autodesk user:\n${JSON.stringify(me, null, 2).slice(0, 2000)}`;
    }
    return command === 'bc' ? 'Unknown `bc` command. Try `bc help`.' : 'Unknown command. Try `buildingconnected help`.';
  } catch (err) {
    return `BuildingConnected error: ${err?.message || err}`;
  }
}

function safeJsonParse(text) {
  try {
    return JSON.parse(text);
  } catch {
    return null;
  }
}

function extractJsonObject(text) {
  if (!text) return null;
  const trimmed = String(text).trim();
  const direct = safeJsonParse(trimmed);
  if (direct) return direct;
  const start = trimmed.indexOf('{');
  const end = trimmed.lastIndexOf('}');
  if (start >= 0 && end > start) {
    return safeJsonParse(trimmed.slice(start, end + 1));
  }
  return null;
}

function getSmartsheetNameToIdMap() {
  const raw = process.env.SMARTSHEET_SHEET_MAP || process.env.SMARTSHEET_SHEETS_JSON || '{}';
  try {
    const obj = typeof raw === 'string' ? JSON.parse(raw) : {};
    const map = new Map();
    for (const [k, v] of Object.entries(obj)) {
      const id = String(v).trim();
      const name = String(k).trim().toLowerCase();
      if (name && id) map.set(name, id);
    }
    return map;
  } catch {
    return new Map();
  }
}

function getActiveSmartsheetSheetId() {
  const store = smartsheetSheetContext.getStore();
  if (store && store.sheetId) {
    const sid = String(store.sheetId).trim();
    if (sid) return sid;
  }
  const fb = String(process.env.SMARTSHEET_SHEET_ID || '').trim();
  if (fb) return fb;
  const m = getSmartsheetNameToIdMap();
  if (m.size === 1) return Array.from(m.values())[0];
  return '';
}

function invalidateSmartsheetColumnCache() {
  const sid = getActiveSmartsheetSheetId();
  if (sid) smartsheetColumnsCacheBySheetId.delete(sid);
}

function resolveSheetSelection(userText, channelId) {
  const text = String(userText || '').trim();
  const nameMap = getSmartsheetNameToIdMap();

  if (
    /\b(?:use|switch\s+back\s+to)\s+default\s+sheet\b/i.test(text.toLowerCase()) ||
    /\breset\s+smartsheet\s+sheet\b/i.test(text.toLowerCase())
  ) {
    slackChannelActiveSheetId.delete(channelId);
    const def = String(process.env.SMARTSHEET_SHEET_ID || '').trim();
    return { sheetId: def, label: 'default', cleared: true, switched: true };
  }

  let namedLabel = null;
  const leadingIntro = /^\s*(?:in|on|use)\s+(?:the\s+)?(.+?)\s+(?:sheet|spreadsheet)(?:\s*[:,])?\s*/i.exec(text);
  if (leadingIntro?.[1]) {
    namedLabel = leadingIntro[1].trim().replace(/^["']|["']$/g, '').replace(/\s+/g, ' ');
  }
  // "in AI Test Jack, create …" — no word "sheet"; only treat as sheet switch if name matches the map.
  if (!namedLabel) {
    const commaIntro = /^\s*(?:in|on|use)\s+(?:the\s+)?([^,\n]{1,160}?)\s*,\s*\S/i.exec(text);
    if (commaIntro?.[1]) {
      const candidate = commaIntro[1].trim().replace(/^["']|["']$/g, '').replace(/\s+/g, ' ');
      const key = candidate.toLowerCase();
      let id = nameMap.get(key);
      if (!id) {
        for (const [k, v] of nameMap.entries()) {
          if (key.includes(k) || k.includes(key)) {
            id = v;
            break;
          }
        }
      }
      if (id) namedLabel = candidate;
    }
  }
  if (!namedLabel) {
    const openLead = /^\s*open\s+(?:the\s+)?(.+?)\s+(?:sheet|spreadsheet)\b/i.exec(text);
    if (openLead?.[1]) {
      namedLabel = openLead[1].trim().replace(/^["']|["']$/g, '').replace(/\s+/g, ' ');
    }
  }
  if (!namedLabel) {
    const sw = text.match(/\bswitch\s+to\s+(?:the\s+)?(.+?)(?:\s+(?:sheet|spreadsheet))?\b/i);
    if (sw?.[1]) {
      namedLabel = sw[1].trim().replace(/^["']|["']$/g, '').replace(/\s+/g, ' ');
    }
  }

  if (namedLabel) {
    const key = namedLabel.toLowerCase();
    let id = nameMap.get(key);
    if (!id) {
      for (const [k, v] of nameMap.entries()) {
        if (key.includes(k) || k.includes(key)) {
          id = v;
          break;
        }
      }
    }
    if (!id) {
      const def = String(process.env.SMARTSHEET_SHEET_ID || '').trim();
      return { sheetId: def, unknownName: true, label: namedLabel, switched: false };
    }
    slackChannelActiveSheetId.set(channelId, id);
    return { sheetId: id, label: namedLabel, switched: true, cleared: false };
  }

  const remembered = slackChannelActiveSheetId.get(channelId);
  if (remembered) return { sheetId: remembered, label: 'remembered', switched: false };

  const def = String(process.env.SMARTSHEET_SHEET_ID || '').trim();
  return { sheetId: def, switched: false };
}

function hasInlineSheetIntro(text) {
  return /^\s*(?:in|on|use)\s+(?:the\s+)?.+?\s+(?:sheet|spreadsheet)[,:\s]+/i.test(String(text || ''));
}

function hasSheetOperationIntent(text) {
  return /\b(change|update|add|remove|delete|format|clear|empty|set\s+row|smartsheet|finished|completed|done|replace|background|cell|column|row|rows|contents?|what\s+is|what\s+are|value|read|show|tell\s+me|bold|unbold|italic|underline|increase|decrease|increment|decrement|multiply|times)\b/i.test(
    String(text || '')
  );
}

function smartsheetConfigured() {
  if (!process.env.SMARTSHEET_API_TOKEN) return false;
  if (process.env.SMARTSHEET_SHEET_ID) return true;
  return getSmartsheetNameToIdMap().size > 0;
}

/** True when the user means one column (or cells in one column), not the whole grid. */
function looksLikeClearSingleColumnCellsScope(text) {
  const lower = String(text || '').toLowerCase();
  if (!/\bcolumn\s*\d+\b/.test(lower)) return false;
  if (
    /\b(empty|clear|wipe|erase)\b[\s\S]{0,120}\b(all|every)\b[\s\S]{0,80}\bcells?\b[\s\S]{0,120}\bcolumn\s*\d+/i.test(
      lower
    )
  ) {
    return true;
  }
  if (/\b(empty|clear|wipe|erase)\b[\s\S]{0,60}\bcells?\b[\s\S]{0,100}\bcolumn\s*\d+/i.test(lower)) {
    return true;
  }
  if (
    /\b(empty|clear|wipe|erase)\b[\s\S]{0,100}\bcolumn\s*\d+\b[\s\S]{0,120}\b(except|excluding|leave|but\s+keep)/i.test(
      lower
    )
  ) {
    return true;
  }
  if (/\b(empty|clear|wipe|erase)\b\s+column\s*\d+\b/i.test(lower)) return true;
  return false;
}

function looksLikeClearWholeSheetRequest(text) {
  const lower = (text || '').toLowerCase();
  if (looksLikeClearSingleColumnCellsScope(text)) return false;
  if (
    /\b(empty|clear|wipe|erase)\b[\s\S]{0,120}\b(every|all)\b[\s\S]{0,60}\bcells?\b/.test(lower) ||
    /\b(every|all)\b[\s\S]{0,60}\bcells?\b[\s\S]{0,120}\b(in the sheet|in this sheet|on the sheet|in my sheet)\b/.test(
      lower
    )
  ) {
    return true;
  }
  if (/\bstart over\b[\s\S]{0,80}\bclean\b[\s\S]{0,50}\bsheet\b/.test(lower)) {
    return true;
  }
  if (
    /\b(clean|fresh)\b[\s\S]{0,50}\bsheet\b/.test(lower) &&
    /\b(start over|from scratch|empty|reset|all cells|every cell|clear all)\b/.test(lower)
  ) {
    return true;
  }
  if (
    /\bclear\b[\s\S]{0,30}\b(all|every)\b[\s\S]{0,40}\bcells?\b/.test(lower) ||
    /\bclear\b[\s\S]{0,40}\b(the\s+)?whole\b[\s\S]{0,25}\b(sheet|grid)\b/.test(lower)
  ) {
    return true;
  }
  if (/\breset\b[\s\S]{0,40}\b(the\s+)?(sheet|data)\b/.test(lower) && /\b(all|everything|whole)\b/.test(lower)) {
    return true;
  }
  return false;
}

function looksLikeResetGreenCellsToWhiteRequest(text) {
  const lower = (text || '').toLowerCase();
  const hasGreen = /\bgreen\b/.test(lower);
  const hasWhiteTarget =
    /\bto\s+white\b/.test(lower) ||
    /\bto\s+be\s+white\b/.test(lower) ||
    /\bwhite\s+again\b/.test(lower) ||
    /\b(be|become)\s+white\b/.test(lower) ||
    /\binto\s+white\b/.test(lower);
  const hasCellsSurface =
    /\bcell(s)?\b/.test(lower) || /\b(background|highlight|fill)\b/.test(lower) || /\b(colou?r)\b/.test(lower);
  return Boolean(hasGreen && hasWhiteTarget && hasCellsSurface);
}

function looksLikeSheetTopic(text) {
  const lower = (text || '').toLowerCase();
  const hasRowRange = /\brows?\s*\d+\s*[-–]\s*\d+\b/.test(lower);
  return (
    looksLikeClearWholeSheetRequest(text) ||
    lower.includes('smartsheet') ||
    lower.includes('smart sheet') ||
    /\bsheet\b/.test(lower) ||
    lower.includes('sheet id') ||
    lower.includes('row id') ||
    /\b(update|add|delete|remove|insert)\b.*\b(row|rows)\b/.test(lower) ||
    /\b(row|rows)\b.*\b(update|add|delete|remove|insert)\b/.test(lower) ||
    /\bcell(s)?\b/.test(lower) ||
    /\b(background|highlight)\b/.test(lower) ||
    (/\b(format|formatting)\b/.test(lower) && /\b(row|rows|column)\b/.test(lower)) ||
    (/\b(color|colour)\b/.test(lower) && /\b(row|rows|column|cell)\b/.test(lower)) ||
    (hasRowRange && /\b(column|cell|background|highlight|format|color|colour)\b/.test(lower)) ||
    (hasRowRange && /\bcolumn\s*\d+\b/.test(lower)) ||
    /\bcolumn\s*\d+\b[\s\S]{0,60}\brow\s*\d+\b/.test(lower) ||
    /\brow\s*\d+\b[\s\S]{0,60}\bcolumn\s*\d+\b/.test(lower) ||
    /\b(what|contents?|value|show|read)\b[\s\S]{0,80}\b(cell|column|row)\b/.test(lower) ||
    (/\bbold\b|\bunbold\b|\bnot\s+bold\b|\bremove\s+bold\b/.test(lower) &&
      (/\bcolumn\s*\d+\b/.test(lower) || /\brow\s*\d+\b/.test(lower)))
  );
}

function parseRowNumberRange(text) {
  const lower = (text || '').toLowerCase();
  const patterns = [
    /\brows?\s*(\d+)\s*[-–]\s*(\d+)\b/,
    /\brow\s*(\d+)\s+to\s+row\s*(\d+)\b/,
    /\brows?\s*(\d+)\s+through\s+(\d+)\b/,
    /\brow\s*(\d+)\s+through\s+row\s*(\d+)\b/,
    /\bfrom\s+row\s*(\d+)\s+to\s+row\s*(\d+)\b/,
    /\bbetween\s+row\s*(\d+)\s+and\s+row\s*(\d+)\b/,
  ];
  for (const re of patterns) {
    const m = lower.match(re);
    if (!m) continue;
    const start = Number(m[1]);
    const end = Number(m[2]);
    if (!Number.isFinite(start) || !Number.isFinite(end)) continue;
    if (start <= 0 || end <= 0) continue;
    return start <= end ? { start, end } : { start: end, end: start };
  }
  const single = lower.match(/\brow\s+(\d+)\b(?!\s+to\s+row\s*\d)(?!\s+through\s+row\s*\d)/);
  if (single) {
    const n = Number(single[1]);
    if (Number.isFinite(n) && n > 0) return { start: n, end: n };
  }
  return null;
}

function parseBackgroundColorChangeRequest(text) {
  const lower = String(text || '').toLowerCase();
  const hasSurface = /\bcell(s)?\b/.test(lower) || /\b(background|highlight|fill|color|colour)\b/.test(lower);
  if (!hasSurface) return null;
  const colors = [];
  const colorTokens = ['green', 'red', 'blue', 'yellow', 'orange', 'white', 'black', 'purple', 'pink', 'gray', 'grey', 'brown'];
  for (const token of colorTokens) {
    const re = new RegExp(`\\b${token}\\b`, 'g');
    let m;
    while ((m = re.exec(lower)) !== null) {
      const hex = normalizeColorToHex(token);
      if (hex) colors.push({ token, hex, idx: m.index });
    }
  }
  const hexRe = /#([0-9a-f]{6})\b/gi;
  let hm;
  while ((hm = hexRe.exec(lower)) !== null) {
    colors.push({ token: `#${hm[1]}`, hex: `#${hm[1]}`.toUpperCase(), idx: hm.index });
  }
  const wantsBlankTarget =
    /\b(blank|no\s+fill|no\s+colou?r|without\s+colou?r|clear\s+fill|remove\s+fill|transparent)\b/.test(lower) ||
    /\b(remove|clear|strip)\b[\s\S]{0,25}\b(background|fill|colou?r|highlight)\b/.test(lower);
  if (colors.length === 0) return null;
  const target = wantsBlankTarget ? '#FFFFFF' : colors[colors.length - 1].hex;
  let source = null;
  for (const c of colors) {
    if (c.hex !== target) source = c.hex;
  }
  if (!source || source === target) return null;
  return { sourceHex: source, targetHex: target };
}

function normalizeColorToHex(colorText) {
  const t = String(colorText || '').trim().toLowerCase();
  if (!t) return null;
  if (t === 'green') return '#00FF00';
  if (t === 'red') return '#FF0000';
  if (t === 'blue') return '#0000FF';
  if (t === 'yellow') return '#FFFF00';
  if (t === 'orange') return '#FFA500';
  if (t === 'white') return '#FFFFFF';
  if (t === 'black') return '#000000';
  if (t === 'purple') return '#800080';
  if (t === 'pink') return '#FFC0CB';
  if (t === 'gray' || t === 'grey') return '#808080';
  if (t === 'brown') return '#8B4513';
  if (t === 'blank' || t === 'none' || t === 'no fill' || t === 'no color' || t === 'no colour' || t === 'transparent')
    return '#FFFFFF';
  if (/^#[0-9a-f]{6}$/i.test(t)) return t.toUpperCase();
  return null;
}

function describeAppliedSmartsheetBackground(requestedHex, paletteHexRaw) {
  const req = normalizeColorToHex(requestedHex) || String(requestedHex || '').trim();
  const palTrim = typeof paletteHexRaw === 'string' ? paletteHexRaw.trim() : '';
  const pal = /^#[0-9a-f]{6}$/i.test(palTrim) ? palTrim.toUpperCase() : palTrim;
  if (!pal) return req;
  const palNorm = normalizeColorToHex(pal);
  if (palNorm && req && palNorm === req) return req;
  return `${pal} (nearest Smartsheet palette match for ${req})`;
}

function inferBackgroundColorFromText(text) {
  const lower = (text || '').toLowerCase();
  const hex = lower.match(/#([0-9a-f]{6})\b/i);
  if (hex) return `#${hex[1]}`.toUpperCase();
  const tokens = [
    'green',
    'red',
    'blue',
    'yellow',
    'orange',
    'white',
    'black',
    'purple',
    'pink',
    'gray',
    'grey',
    'brown',
  ];
  for (const token of tokens) {
    if (new RegExp(`\\b${token}\\b`).test(lower)) {
      const h = normalizeColorToHex(token);
      if (h) return h;
    }
  }
  return null;
}

function inferColumnTitleFromText(text, columnsByTitle) {
  const titles = Array.from(columnsByTitle.values()).map((c) => c.title);
  const lower = (text || '').toLowerCase();
  for (const title of titles) {
    if (lower.includes(title.toLowerCase())) return title;
  }
  const m = lower.match(/\bcolumn\s*(\d+)\b/);
  if (m) {
    const idx = Number(m[1]);
    if (Number.isFinite(idx) && idx > 0 && idx <= titles.length) return titles[idx - 1];
  }
  return null;
}

async function getRowIdsByRowNumbers(rowStart, rowEnd) {
  const sheet = await getSmartsheetSheet();
  const rows = sheet.rows || [];
  const byNumber = new Map();
  for (const row of rows) {
    const n = row.rowNumber;
    if (!n) continue;
    byNumber.set(Number(n), row.id);
  }
  const ids = [];
  for (let n = rowStart; n <= rowEnd; n += 1) {
    const id = byNumber.get(n);
    if (!id) return { ok: false, missingRowNumber: n, ids: [] };
    ids.push(id);
  }
  return { ok: true, ids };
}

/** Turn `run:bc projects` / `Slack:bc connect` into readable Slack text (colon + bc was glued by the model). */
function polishSlackMarkdownCommands(text) {
  return String(text || '').replace(
    /([:;])\s*bc\s+((?:auth\s+status|connect|projects|whoami|help|project)\b)/gi,
    (_m, punct, cmd) => `${punct}\n\n\`bc ${cmd}\``
  );
}

/**
 * Slack mrkdwn uses *bold*, not **bold**. Claude often emits **…**; normalize so replies look intentional.
 * Then fix glued `run:bc …` command hints.
 */
function formatSlackBotMessageText(raw) {
  let t = String(raw || '').replace(/\u00a0/g, ' ');
  let i = 0;
  while (/\*\*[^*]+\*\*/.test(t) && i < 40) {
    t = t.replace(/\*\*([^*]+)\*\*/, '*$1*');
    i += 1;
  }
  t = polishSlackMarkdownCommands(t);
  return t;
}

async function postSlackMessage(payload) {
  const text = formatSlackBotMessageText(String(payload.text || ''));
  const channel = String(payload.channel || '');
  const jsonBody = {
    channel,
    text,
    ...(payload.thread_ts ? { thread_ts: String(payload.thread_ts) } : {}),
  };
  const postOnce = async (body) => {
    const res = await fetch('https://slack.com/api/chat.postMessage', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json; charset=utf-8',
        Authorization: `Bearer ${process.env.SLACK_BOT_TOKEN || ''}`,
      },
      body: JSON.stringify(body),
      signal: fetchTimeoutSignal(8000),
    });
    const data = await res.json().catch(() => ({}));
    if (!data.ok) {
      throw new Error(String(data.error || `chat.postMessage failed (HTTP ${res.status})`));
    }
  };
  try {
    await postOnce(jsonBody);
  } catch (err) {
    const code = String(err?.message || err || '');
    console.error('Slack chat.postMessage failed:', code, {
      channel,
      hadThreadTs: Boolean(payload.thread_ts),
    });
    const retryWithoutThread =
      Boolean(payload.thread_ts) &&
      (code.includes('thread_not_found') || code.includes('invalid_thread_ts'));
    if (retryWithoutThread) {
      console.warn('postSlackMessage: retrying without thread_ts after', code);
      await postOnce({ channel, text });
    } else {
      throw err;
    }
  }
  console.log(
    `Posted Slack reply (channel=${payload.channel}, thread_ts=${payload.thread_ts || 'none'})`
  );
}

/**
 * Posts a short Slack message if `asyncWork` is still running after `ms` (default 5s).
 * Always call `.cancel()` in `finally` when the main reply path finishes so we do not leak timers.
 */
function startSlackSlowWorkHintTimer({ channel, thread_ts, ms, text }) {
  const raw = ms != null ? ms : process.env.SLACK_SLOW_WORK_HINT_MS;
  const delay = Math.max(500, Number(raw) || 5000);
  const hint =
    text ||
    '_Still working on that — thanks for waiting. Sheets, external APIs, or the model can take more than a few seconds._';
  let done = false;
  const t = setTimeout(() => {
    if (done) return;
    postSlackMessage({ channel, text: hint, thread_ts }).catch((err) => {
      console.warn('Slack slow-work hint failed:', err?.message || err);
    });
  }, delay);
  return {
    cancel() {
      done = true;
      clearTimeout(t);
    },
  };
}

function formatFieldSummary(pairs) {
  return pairs.map((p) => `${p.key}=${p.value}`).join(' | ');
}

function formatConfirmation(action, rowId, pairs) {
  const summary = formatFieldSummary(pairs);
  if (action === 'add') {
    return `Done — added a new Smartsheet row (row id ${rowId}).\n${summary}`;
  }
  return `Done — updated Smartsheet row ${rowId}.\n${summary}`;
}

function normalizeHealthCheckInput(userText) {
  return String(userText || '')
    .replace(/[\u200B-\u200D\uFEFF]/g, '')
    .replace(/`/g, '')
    .toLowerCase()
    .trim()
    .replace(/[!?.…]+$/u, '')
    .trim();
}

function getBasicHealthReply(userText) {
  const normalized = normalizeHealthCheckInput(userText);
  const quickChecks = new Set(['ping', 'hello', 'hi', 'test', 'are you there']);
  if (quickChecks.has(normalized)) {
    return 'I am live in Slack and receiving your messages.';
  }
  // After stripping <@…> mentions, text can look like "Transform Sim ping" or "Transform Sim Bot ping".
  const tokens = normalized.split(/\s+/).filter(Boolean);
  if (tokens.length < 1 || tokens[tokens.length - 1] !== 'ping') return null;
  const prefix = tokens.slice(0, -1);
  const stop = new Set(['how', 'why', 'what', 'when', 'where', 'who', 'which']);
  if (prefix.some((w) => stop.has(w))) return null;
  if (prefix.length <= 4) {
    return 'I am live in Slack and receiving your messages.';
  }
  return null;
}

function parseFieldPairs(rawText) {
  return rawText
    .split('|')
    .map((part) => part.trim())
    .filter(Boolean)
    .map((pair) => {
      const idx = pair.indexOf('=');
      if (idx < 1) return null;
      const key = pair.slice(0, idx).trim();
      const value = pair.slice(idx + 1).trim();
      if (!key || !value) return null;
      return { key, value };
    })
    .filter(Boolean);
}

async function smartsheetRequest(path, options = {}) {
  if (!process.env.SMARTSHEET_API_TOKEN) {
    throw new Error('SMARTSHEET_API_TOKEN is not configured.');
  }
  const response = await fetch(`${SMARTSHEET_BASE_URL}${path}`, {
    method: options.method || 'GET',
    headers: {
      Authorization: `Bearer ${process.env.SMARTSHEET_API_TOKEN}`,
      'Content-Type': 'application/json',
    },
    body: options.body ? JSON.stringify(options.body) : undefined,
  });
  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`Smartsheet API error ${response.status}: ${errorText}`);
  }
  return response.json();
}

function hexStringToRgbComponents(hex) {
  const m = /^#?([0-9a-f]{6})$/i.exec(String(hex || '').trim());
  if (!m) return null;
  const n = parseInt(m[1], 16);
  return { r: (n >> 16) & 255, g: (n >> 8) & 255, b: n & 255 };
}

function rgbSquaredDistance(a, b) {
  if (!a || !b) return Infinity;
  return (a.r - b.r) ** 2 + (a.g - b.g) ** 2 + (a.b - b.b) ** 2;
}

/** Position 9 in the 0..16 descriptor is background color; see Smartsheet cell-formatting guide. */
function buildSmartsheetBackgroundFormatDescriptor(colorIndex) {
  const parts = Array.from({ length: 17 }, () => '');
  parts[9] = String(colorIndex);
  return parts.join(',');
}

function parseSmartsheetFormatDescriptorParts(format) {
  if (format == null || format === '') return null;
  if (typeof format === 'number') return null;
  const raw = String(format).trim();
  if (!raw.includes(',')) return null;
  const parts = raw.split(',');
  while (parts.length < 17) parts.push('');
  return parts.slice(0, 17);
}

function getBackgroundPaletteIndexFromCellFormat(format) {
  const parts = parseSmartsheetFormatDescriptorParts(format);
  if (!parts) return null;
  const v = parts[9];
  if (v === '' || v == null) return null;
  const n = Number(v);
  return Number.isFinite(n) ? n : null;
}

function mergeSmartsheetFormatDescriptorBackground(format, newBgPaletteIndex) {
  const parts = parseSmartsheetFormatDescriptorParts(format) || Array.from({ length: 17 }, () => '');
  const merged = parts.slice(0, 17);
  while (merged.length < 17) merged.push('');
  merged[9] = String(newBgPaletteIndex);
  return merged.join(',');
}

/** Smartsheet format descriptor index 2 = bold (1 = on, empty = default / not bold). */
function mergeSmartsheetFormatDescriptorBold(format, wantBold) {
  const parts = parseSmartsheetFormatDescriptorParts(format) || Array.from({ length: 17 }, () => '');
  const merged = parts.slice(0, 17);
  while (merged.length < 17) merged.push('');
  merged[2] = wantBold ? '1' : '';
  return merged.join(',');
}

function isGreenishBackgroundPaletteIndex(colorTable, idx) {
  if (!Array.isArray(colorTable) || idx == null || idx < 0 || idx >= colorTable.length) return false;
  const entry = colorTable[idx];
  if (typeof entry !== 'string' || !/^#[0-9a-f]{6}$/i.test(entry.trim())) return false;
  const rgb = hexStringToRgbComponents(entry.trim());
  if (!rgb) return false;
  const pureGreen = hexStringToRgbComponents('#00FF00');
  if (pureGreen && rgbSquaredDistance(rgb, pureGreen) < 55_000) return true;
  if (rgb.g >= 140 && rgb.g > rgb.r + 25 && rgb.g > rgb.b + 25) return true;
  return false;
}

function pickBackgroundColorIndexFromTable(colorTable, desiredHex) {
  const normalizedDesired = normalizeColorToHex(desiredHex);
  const rgbDesired = hexStringToRgbComponents(normalizedDesired || String(desiredHex || '').trim());
  if (!Array.isArray(colorTable) || !rgbDesired) return -1;
  for (let i = 0; i < colorTable.length; i += 1) {
    const entry = colorTable[i];
    if (typeof entry !== 'string') continue;
    const trimmed = entry.trim();
    if (!/^#[0-9a-f]{6}$/i.test(trimmed)) continue;
    if (normalizeColorToHex(trimmed) === normalizedDesired) return i;
  }
  let bestIdx = -1;
  let bestD = Infinity;
  for (let i = 0; i < colorTable.length; i += 1) {
    const entry = colorTable[i];
    if (typeof entry !== 'string') continue;
    const trimmed = entry.trim();
    if (!/^#[0-9a-f]{6}$/i.test(trimmed)) continue;
    const rgb = hexStringToRgbComponents(trimmed);
    if (!rgb) continue;
    const d = rgbSquaredDistance(rgbDesired, rgb);
    if (d < bestD) {
      bestD = d;
      bestIdx = i;
    }
  }
  return bestIdx;
}

async function getSmartsheetFormatsColorTable() {
  if (smartsheetFormatsColorTableCache) return smartsheetFormatsColorTableCache;
  const info = await smartsheetRequest('/serverinfo');
  const formats = info.formats || info.result?.formats;
  const colorTable = formats?.color;
  if (!Array.isArray(colorTable)) {
    throw new Error('Smartsheet serverinfo response missing formats.color');
  }
  smartsheetFormatsColorTableCache = colorTable;
  return smartsheetFormatsColorTableCache;
}

async function smartsheetFormatDescriptorForBackgroundHex(desiredHex) {
  const colorTable = await getSmartsheetFormatsColorTable();
  const idx = pickBackgroundColorIndexFromTable(colorTable, desiredHex);
  if (idx < 0) {
    throw new Error(`Could not map "${desiredHex}" to a Smartsheet palette background color.`);
  }
  return { descriptor: buildSmartsheetBackgroundFormatDescriptor(idx), paletteIndex: idx, paletteHex: colorTable[idx] };
}

async function getSmartsheetSheetForCellValues() {
  const id = getActiveSmartsheetSheetId();
  if (!id) {
    throw new Error('SMARTSHEET_SHEET_ID is not configured (set SMARTSHEET_SHEET_ID or SMARTSHEET_SHEET_MAP).');
  }
  return smartsheetRequest(`/sheets/${id}?include=objectValue`);
}

async function getSmartsheetSheetForCellValuesIncludeFormat() {
  const id = getActiveSmartsheetSheetId();
  if (!id) {
    throw new Error('SMARTSHEET_SHEET_ID is not configured (set SMARTSHEET_SHEET_ID or SMARTSHEET_SHEET_MAP).');
  }
  return smartsheetRequest(`/sheets/${id}?include=objectValue,format`);
}

function findCellOnRow(row, columnId) {
  const target = Number(columnId);
  for (const cell of row.cells || []) {
    if (Number(cell.columnId) === target) return cell;
  }
  return null;
}

/** Smartsheet requires `value` on each cell in PUT /rows; preserve existing or use a type-safe empty. */
function pickCellValueForBackgroundUpdate(columnMeta, existingCell) {
  if (existingCell && Object.prototype.hasOwnProperty.call(existingCell, 'value')) {
    return existingCell.value;
  }
  const t = String(columnMeta?.type || 'TEXT_NUMBER');
  if (t === 'CHECKBOX') return false;
  if (t === 'MULTI_PICKLIST' || t === 'MULTI_CONTACT_LIST') return [];
  if (t === 'PICKLIST' || t === 'CONTACT_LIST' || t === 'USER_REF') return null;
  return '';
}

async function applyColumnBackgroundForRows(rowIds, columnTitle, columnsByTitle, colorHex) {
  const col = columnsByTitle.get(String(columnTitle).toLowerCase());
  if (!col) throw new Error(`Unknown column title: ${columnTitle}`);
  const { descriptor, paletteHex } = await smartsheetFormatDescriptorForBackgroundHex(colorHex);
  const sheet = await getSmartsheetSheetForCellValues();
  const rowById = new Map();
  for (const row of sheet.rows || []) {
    rowById.set(row.id, row);
    rowById.set(String(row.id), row);
  }
  const updates = rowIds.map((id) => {
    const row = rowById.get(id) ?? rowById.get(String(id)) ?? rowById.get(Number(id));
    if (!row) {
      throw new Error(
        `Could not load row id ${id} from the sheet (it may be beyond the first page of results).`
      );
    }
    const existing = findCellOnRow(row, col.id);
    const value = pickCellValueForBackgroundUpdate(col, existing);
    return {
      id,
      cells: [{ columnId: col.id, value, format: descriptor }],
    };
  });
  await smartsheetRequest(`/sheets/${getActiveSmartsheetSheetId()}/rows`, {
    method: 'PUT',
    body: updates,
  });
  return { paletteHex };
}

function isPaletteIndexNearColorHex(colorTable, idx, sourceHex) {
  if (!Array.isArray(colorTable) || idx == null || idx < 0 || idx >= colorTable.length) return false;
  const entry = colorTable[idx];
  if (typeof entry !== 'string' || !/^#[0-9a-f]{6}$/i.test(entry.trim())) return false;
  const rgb = hexStringToRgbComponents(entry.trim());
  const src = hexStringToRgbComponents(sourceHex);
  if (!rgb || !src) return false;
  return rgbSquaredDistance(rgb, src) < 45_000;
}

async function executeResetBackgroundColor(columnsByTitle, sourceHex, targetHex) {
  const sheetId = getActiveSmartsheetSheetId();
  const colorTable = await getSmartsheetFormatsColorTable();
  const { paletteIndex: targetIdx, paletteHex: appliedTargetHex } =
    await smartsheetFormatDescriptorForBackgroundHex(targetHex);
  const columnList = Array.from(columnsByTitle.values());
  const clearableById = new Map();
  for (const col of columnList) {
    if (isSmartsheetColumnClearable(col) && col.id != null) clearableById.set(Number(col.id), col);
  }
  const rowUpdates = new Map();
  let scannedCells = 0;
  let matchedCells = 0;
  let page = 1;
  const maxPages = 500;
  for (;;) {
    if (page > maxPages) {
      throw new Error(`Sheet row pagination exceeded ${maxPages} pages during background scan.`);
    }
    const path = `/sheets/${sheetId}?page=${page}&pageSize=${SMARTSHEET_SHEET_PAGE_SIZE}&include=objectValue,format`;
    const sheet = await smartsheetRequest(path);
    const rows = sheet.rows || [];
    if (rows.length === 0) break;
    for (const row of rows) {
      for (const cell of row.cells || []) {
        const colId = Number(cell.columnId);
        if (!clearableById.has(colId)) continue;
        scannedCells += 1;
        const bgIdx = getBackgroundPaletteIndexFromCellFormat(cell.format);
        if (bgIdx == null) continue;
        if (!isPaletteIndexNearColorHex(colorTable, bgIdx, sourceHex)) continue;
        const colMeta = clearableById.get(colId);
        const newFormat = mergeSmartsheetFormatDescriptorBackground(cell.format, targetIdx);
        const value = pickCellValueForBackgroundUpdate(colMeta, cell);
        const rid = row.id;
        if (!rowUpdates.has(rid)) rowUpdates.set(rid, { id: rid, cells: [] });
        const bucket = rowUpdates.get(rid).cells;
        const existingIdx = bucket.findIndex((c) => Number(c.columnId) === colId);
        const nextCell = { columnId: colId, value, format: newFormat };
        if (existingIdx >= 0) bucket[existingIdx] = nextCell;
        else bucket.push(nextCell);
        matchedCells += 1;
      }
    }
    if (rows.length < SMARTSHEET_SHEET_PAGE_SIZE) break;
    page += 1;
  }
  invalidateSmartsheetColumnCache();
  if (matchedCells === 0) {
    return (
      `I did not find any ${sourceHex} (or close) background cells to change in the loaded part of this sheet (scanned ${scannedCells} editable cells). ` +
      `If that color is only in part of the sheet, say the column and rows, e.g. \`Set Column2 background ${targetHex} for rows 1-20\`.`
    );
  }
  const payloads = Array.from(rowUpdates.values());
  for (let i = 0; i < payloads.length; i += SMARTSHEET_ROW_PUT_BATCH) {
    const batch = payloads.slice(i, i + SMARTSHEET_ROW_PUT_BATCH);
    await smartsheetRequest(`/sheets/${sheetId}/rows`, { method: 'PUT', body: batch });
  }
  return `Done — changed ${matchedCells} cell background(s) from ${sourceHex} (or nearest palette match) to ${describeAppliedSmartsheetBackground(targetHex, appliedTargetHex)} across ${payloads.length} row(s) (editable columns only).`;
}

async function executeResetGreenishBackgroundsToWhite(columnsByTitle) {
  const sheetId = getActiveSmartsheetSheetId();
  const colorTable = await getSmartsheetFormatsColorTable();
  const { paletteIndex: whiteIdx } = await smartsheetFormatDescriptorForBackgroundHex('#FFFFFF');
  const columnList = Array.from(columnsByTitle.values());
  const clearableById = new Map();
  for (const col of columnList) {
    if (isSmartsheetColumnClearable(col) && col.id != null) clearableById.set(Number(col.id), col);
  }
  const rowUpdates = new Map();
  let scannedCells = 0;
  let matchedCells = 0;
  let page = 1;
  const maxPages = 500;
  for (;;) {
    if (page > maxPages) {
      throw new Error(`Sheet row pagination exceeded ${maxPages} pages during background scan.`);
    }
    const path = `/sheets/${sheetId}?page=${page}&pageSize=${SMARTSHEET_SHEET_PAGE_SIZE}&include=objectValue,format`;
    const sheet = await smartsheetRequest(path);
    const rows = sheet.rows || [];
    if (rows.length === 0) break;
    for (const row of rows) {
      for (const cell of row.cells || []) {
        const colId = Number(cell.columnId);
        if (!clearableById.has(colId)) continue;
        scannedCells += 1;
        const bgIdx = getBackgroundPaletteIndexFromCellFormat(cell.format);
        if (bgIdx == null) continue;
        if (!isGreenishBackgroundPaletteIndex(colorTable, bgIdx)) continue;
        const colMeta = clearableById.get(colId);
        const newFormat = mergeSmartsheetFormatDescriptorBackground(cell.format, whiteIdx);
        const value = pickCellValueForBackgroundUpdate(colMeta, cell);
        const rid = row.id;
        if (!rowUpdates.has(rid)) rowUpdates.set(rid, { id: rid, cells: [] });
        const bucket = rowUpdates.get(rid).cells;
        const existingIdx = bucket.findIndex((c) => Number(c.columnId) === colId);
        const nextCell = { columnId: colId, value, format: newFormat };
        if (existingIdx >= 0) bucket[existingIdx] = nextCell;
        else bucket.push(nextCell);
        matchedCells += 1;
      }
    }
    if (rows.length < SMARTSHEET_SHEET_PAGE_SIZE) break;
    page += 1;
  }
  invalidateSmartsheetColumnCache();
  if (matchedCells === 0) {
    return (
      'I did not find any green background cells to reset in the loaded part of this sheet ' +
      `(scanned ${scannedCells} editable cells). ` +
      'If the green is only in part of the sheet, say the column name and row numbers, e.g. `Set Column2 background white for rows 1-20`.'
    );
  }
  const payloads = Array.from(rowUpdates.values());
  for (let i = 0; i < payloads.length; i += SMARTSHEET_ROW_PUT_BATCH) {
    const batch = payloads.slice(i, i + SMARTSHEET_ROW_PUT_BATCH);
    await smartsheetRequest(`/sheets/${sheetId}/rows`, { method: 'PUT', body: batch });
  }
  return `Done — reset ${matchedCells} green-tinted cell background(s) to white across ${payloads.length} row(s) (editable columns only).`;
}

function isSmartsheetColumnClearable(col) {
  if (!col || col.id == null) return false;
  if (col.formula != null && String(col.formula).trim() !== '') return false;
  if (SMARTSHEET_NON_CLEARABLE_COLUMN_TYPES.has(String(col.type || ''))) return false;
  return true;
}

async function executeClearEntireSheet() {
  const sheetId = getActiveSmartsheetSheetId();
  const allRows = [];
  let columns = null;
  let page = 1;
  const maxPages = 500;
  for (;;) {
    if (page > maxPages) {
      throw new Error(
        `Sheet row pagination exceeded ${maxPages} pages (${SMARTSHEET_SHEET_PAGE_SIZE} rows/page). Refusing unbounded clear.`
      );
    }
    const path = `/sheets/${sheetId}?page=${page}&pageSize=${SMARTSHEET_SHEET_PAGE_SIZE}&include=objectValue`;
    const sheet = await smartsheetRequest(path);
    if (!columns && sheet.columns) columns = sheet.columns;
    const rows = sheet.rows || [];
    if (rows.length === 0) break;
    allRows.push(...rows);
    if (rows.length < SMARTSHEET_SHEET_PAGE_SIZE) break;
    page += 1;
  }
  if (!columns || columns.length === 0) {
    return 'This sheet has no column definitions.';
  }
  const clearable = columns.filter(isSmartsheetColumnClearable);
  if (clearable.length === 0) {
    return 'No user-editable columns found (only system or formula columns). Nothing was changed.';
  }
  if (allRows.length === 0) {
    return 'The sheet has no data rows to clear.';
  }
  const cellTemplates = clearable.map((col) => ({
    columnId: col.id,
    value: pickCellValueForBackgroundUpdate(col, null),
  }));
  for (let i = 0; i < allRows.length; i += SMARTSHEET_ROW_PUT_BATCH) {
    const batch = allRows.slice(i, i + SMARTSHEET_ROW_PUT_BATCH);
    const body = batch.map((row) => ({
      id: row.id,
      cells: cellTemplates.map((c) => ({ columnId: c.columnId, value: c.value })),
    }));
    await smartsheetRequest(`/sheets/${sheetId}/rows`, { method: 'PUT', body });
  }
  invalidateSmartsheetColumnCache();
  return (
    `Done — cleared all editable cells in ${allRows.length} row(s) across ${clearable.length} column(s). ` +
    'System and formula columns were skipped. Existing rows were kept (only cell values cleared).'
  );
}

async function getSmartsheetColumns() {
  const sheetId = getActiveSmartsheetSheetId();
  if (!sheetId) {
    throw new Error('SMARTSHEET_SHEET_ID is not configured (set SMARTSHEET_SHEET_ID or SMARTSHEET_SHEET_MAP).');
  }
  const cached = smartsheetColumnsCacheBySheetId.get(sheetId);
  if (cached) return cached;
  const sheet = await smartsheetRequest(`/sheets/${sheetId}`);
  const byTitle = new Map();
  for (const col of sheet.columns || []) {
    byTitle.set(String(col.title).toLowerCase(), col);
  }
  smartsheetColumnsCacheBySheetId.set(sheetId, byTitle);
  return byTitle;
}

async function getSmartsheetSheet() {
  const sheetId = getActiveSmartsheetSheetId();
  if (!sheetId) {
    throw new Error('SMARTSHEET_SHEET_ID is not configured (set SMARTSHEET_SHEET_ID or SMARTSHEET_SHEET_MAP).');
  }
  return smartsheetRequest(`/sheets/${sheetId}`);
}

async function getAllSmartsheetRowsAndColumns() {
  const sheetId = getActiveSmartsheetSheetId();
  if (!sheetId) {
    throw new Error('SMARTSHEET_SHEET_ID is not configured (set SMARTSHEET_SHEET_ID or SMARTSHEET_SHEET_MAP).');
  }
  const allRows = [];
  let columns = null;
  let page = 1;
  const maxPages = 500;
  for (;;) {
    if (page > maxPages) {
      throw new Error(
        `Sheet row pagination exceeded ${maxPages} pages (${SMARTSHEET_SHEET_PAGE_SIZE} rows/page).`
      );
    }
    const path = `/sheets/${sheetId}?page=${page}&pageSize=${SMARTSHEET_SHEET_PAGE_SIZE}&include=objectValue`;
    const sheet = await smartsheetRequest(path);
    if (!columns && sheet.columns) columns = sheet.columns;
    const rows = sheet.rows || [];
    if (rows.length === 0) break;
    allRows.push(...rows);
    if (rows.length < SMARTSHEET_SHEET_PAGE_SIZE) break;
    page += 1;
  }
  return { columns: columns || [], rows: allRows };
}

function cellText(cell) {
  const raw = cell?.displayValue ?? cell?.value;
  if (raw == null) return '';
  return String(raw).trim();
}

function pickStatusColumn(columnsByTitle) {
  const cols = Array.from(columnsByTitle.values());
  const preferred = cols.find((c) => /\bstatus\b/i.test(String(c.title || '')));
  if (preferred) return preferred;
  return cols.find((c) => /\b(state|job status|progress)\b/i.test(String(c.title || ''))) || null;
}

function summarizeRowsForClarification(rows, columnsById) {
  return rows
    .slice(0, 6)
    .map((row) => {
      const pieces = [];
      for (const cell of row.cells || []) {
        const txt = cellText(cell);
        if (!txt) continue;
        const col = columnsById.get(Number(cell.columnId));
        const title = col?.title || `Column ${cell.columnId}`;
        if (/primary column|site|location|city|status/i.test(String(title))) {
          pieces.push(`${title}: ${txt}`);
        }
      }
      const summary = pieces.length > 0 ? pieces.join(' | ') : `row ${row.rowNumber || row.id}`;
      return `- row ${row.rowNumber || row.id}: ${summary}`;
    })
    .join('\n');
}

function parseContainsReplaceRequest(text) {
  const m = String(text || '').match(
    /\b(?:change|replace|update)\s+(?:the\s+)?(?:cell|value)\s+that\s+contains\s+["“]?(.+?)["”]?\s+(?:to|with)\s+["“]?(.+?)["”]?$/i
  );
  if (!m) return null;
  const fromValue = String(m[1] || '').trim();
  const toValue = String(m[2] || '').trim();
  if (!fromValue || !toValue) return null;
  return { fromValue, toValue };
}

function stripSheetIntroClause(text) {
  return String(text || '')
    .replace(/^\s*(?:in|on|use)\s+(?:the\s+)?[^,]+?\s+sheet[,:\s]+/i, '')
    .trim();
}

function parseChangeValueForJobLocation(text) {
  const cleaned = stripSheetIntroClause(text);
  const m = cleaned.match(/\bchange\s+(.+?)\s+to\s+(.+?)\s+for\s+(?:the\s+)?(.+?)(?:\s+job)?\s*$/i);
  if (!m) return null;
  const fromValue = String(m[1] || '')
    .trim()
    .replace(/^['"]|['"]$/g, '');
  const toValue = String(m[2] || '')
    .trim()
    .replace(/^['"]|['"]$/g, '');
  let location = String(m[3] || '')
    .trim()
    .replace(/^['"]|['"]$/g, '');
  if (!fromValue || !toValue || !location) return null;
  location = location.replace(/\s+/g, ' ').trim();
  return { fromValue, toValue, location };
}

function parseJobFinishedLocation(text) {
  const lower = String(text || '').toLowerCase();
  if (!/\b(finished|completed|done|wrapped up|closed)\b/.test(lower)) return null;
  const m =
    lower.match(/\b(?:job|work|task)\s+(?:in|at)\s+(.+?)(?:[.!?]|$)/i) ||
    lower.match(/\b(?:finished|completed|done)\b[\s\S]{0,25}\b(?:in|at)\s+(.+?)(?:[.!?]|$)/i);
  if (!m) return null;
  const location = String(m[1] || '')
    .replace(/\s+/g, ' ')
    .trim();
  if (!location) return null;
  return { location, nextStatus: 'Done' };
}

async function executeContainsReplaceRequest(userText, columnsByTitle) {
  const parsed = parseContainsReplaceRequest(userText);
  if (!parsed) return null;
  const { rows, columns } = await getAllSmartsheetRowsAndColumns();
  const columnsById = new Map((columns || []).map((c) => [Number(c.id), c]));
  const term = parsed.fromValue.toLowerCase();
  const matches = [];
  for (const row of rows) {
    for (const cell of row.cells || []) {
      const txt = cellText(cell);
      if (!txt) continue;
      if (!txt.toLowerCase().includes(term)) continue;
      const col = columnsById.get(Number(cell.columnId));
      if (!col || !isSmartsheetColumnClearable(col)) continue;
      matches.push({ row, cell, col });
    }
  }
  if (matches.length === 0) {
    return `I could not find any cell containing "${parsed.fromValue}".`;
  }
  if (matches.length > 1) {
    const uniqueRows = Array.from(new Map(matches.map((m) => [m.row.id, m.row])).values());
    return (
      `I found multiple cells containing "${parsed.fromValue}", so I need clarification before changing to "${parsed.toValue}".\n` +
      summarizeRowsForClarification(uniqueRows, columnsById)
    );
  }
  const one = matches[0];
  await smartsheetRequest(`/sheets/${getActiveSmartsheetSheetId()}/rows`, {
    method: 'PUT',
    body: [{ id: one.row.id, cells: [{ columnId: one.col.id, value: parsed.toValue }] }],
  });
  return `Done — updated row ${one.row.rowNumber || one.row.id}, "${one.col.title}" from "${cellText(
    one.cell
  )}" to "${parsed.toValue}".`;
}

async function executeImplicitJobFinishedUpdate(userText, columnsByTitle) {
  const parsed = parseJobFinishedLocation(userText);
  if (!parsed) return null;
  const statusColumn = pickStatusColumn(columnsByTitle);
  if (!statusColumn) return 'I could not find a Status column. Please tell me which column tracks status.';
  const { rows, columns } = await getAllSmartsheetRowsAndColumns();
  const columnsById = new Map((columns || []).map((c) => [Number(c.id), c]));
  const target = parsed.location.toLowerCase();
  const matchedRows = [];
  for (const row of rows) {
    for (const cell of row.cells || []) {
      const txt = cellText(cell);
      if (!txt) continue;
      if (txt.toLowerCase().includes(target)) {
        matchedRows.push(row);
        break;
      }
    }
  }
  if (matchedRows.length === 0) {
    return `I could not find any row for "${parsed.location}".`;
  }
  if (matchedRows.length > 1) {
    return (
      `I found multiple jobs matching "${parsed.location}". Which one should I mark ${parsed.nextStatus}?\n` +
      summarizeRowsForClarification(matchedRows, columnsById)
    );
  }
  const row = matchedRows[0];
  await smartsheetRequest(`/sheets/${getActiveSmartsheetSheetId()}/rows`, {
    method: 'PUT',
    body: [{ id: row.id, cells: [{ columnId: statusColumn.id, value: parsed.nextStatus }] }],
  });
  return `Done — marked row ${row.rowNumber || row.id} as ${parsed.nextStatus}.`;
}

async function executeChangeValueForJobLocation(userText, columnsByTitle) {
  const parsed = parseChangeValueForJobLocation(userText);
  if (!parsed) return null;
  const { rows, columns } = await getAllSmartsheetRowsAndColumns();
  const columnsById = new Map((columns || []).map((c) => [Number(c.id), c]));
  const target = parsed.location.toLowerCase().trim();
  const matchedRows = [];
  for (const row of rows) {
    for (const cell of row.cells || []) {
      const txt = cellText(cell);
      if (!txt) continue;
      if (txt.toLowerCase().includes(target)) {
        matchedRows.push(row);
        break;
      }
    }
  }
  if (matchedRows.length === 0) {
    return `I could not find any row containing "${parsed.location}".`;
  }
  if (matchedRows.length > 1) {
    return (
      `I found multiple rows matching "${parsed.location}". Which job did you mean?\n` +
      summarizeRowsForClarification(matchedRows, columnsById)
    );
  }
  const row = matchedRows[0];
  const fromNorm = parsed.fromValue.toLowerCase().trim();
  const candidates = [];
  for (const cell of row.cells || []) {
    const col = columnsById.get(Number(cell.columnId));
    if (!col || !isSmartsheetColumnClearable(col)) continue;
    const txt = cellText(cell);
    if (String(txt).toLowerCase().trim() === fromNorm) {
      candidates.push({ cell, col, txt });
    }
  }
  if (candidates.length === 0) {
    return (
      `I found row ${row.rowNumber || row.id} for "${parsed.location}", but no cell equals "${parsed.fromValue}". ` +
      'Name the column title exactly (use `smartsheet columns` in Slack if needed), or paste the row id.'
    );
  }
  if (candidates.length > 1) {
    const titles = candidates.map((c) => c.col.title).join(', ');
    return `Multiple columns on that row equal "${parsed.fromValue}": ${titles}. Which column should become "${parsed.toValue}"?`;
  }
  const one = candidates[0];
  await smartsheetRequest(`/sheets/${getActiveSmartsheetSheetId()}/rows`, {
    method: 'PUT',
    body: [{ id: row.id, cells: [{ columnId: one.col.id, value: parsed.toValue }] }],
  });
  invalidateSmartsheetColumnCache();
  return `Done — row ${row.rowNumber || row.id}, "${one.col.title}": "${parsed.fromValue}" → "${parsed.toValue}".`;
}

function toSmartsheetCells(pairs, columnsByTitle) {
  const cells = [];
  const missingColumns = [];
  for (const pair of pairs) {
    const col = columnsByTitle.get(pair.key.toLowerCase());
    if (!col) {
      missingColumns.push(pair.key);
      continue;
    }
    cells.push({
      columnId: col.id,
      value: pair.value,
    });
  }
  return { cells, missingColumns };
}

async function handleSmartsheetCommand(userText) {
  const trimmed = userText.trim();
  if (!trimmed.toLowerCase().startsWith('smartsheet')) return null;
  const parts = trimmed.split(' ');
  const action = (parts[1] || '').toLowerCase();
  if (!action || action === 'help') {
    return (
      'Smartsheet commands:\n' +
      '- smartsheet add Col A=Value A | Col B=Value B\n' +
      '- smartsheet update ROW_ID Col A=New Value | Col B=New Value\n' +
      '- smartsheet format COLUMN_TITLE ROW_START ROW_END #RRGGBB\n' +
      '- smartsheet bold COLUMN_NUMBER ROW_NUMBER   (1-based column + row; cell text weight)\n' +
      '- smartsheet unbold COLUMN_NUMBER ROW_NUMBER\n' +
      '- smartsheet clear all   (empty every editable cell; keeps rows)\n' +
      '- smartsheet picklist COLUMN_NUMBER   (Yes/No `PICKLIST` on that 1-based column; not primary)\n' +
      '- smartsheet columns\n' +
      'Notes: column names must match your Smartsheet column titles exactly.\n' +
      'Multi-sheet: set env SMARTSHEET_SHEET_MAP as JSON like {"Job Sim":"SHEET_ID_1","Ops":"SHEET_ID_2"}; then say e.g. "Switch to Job Sim sheet" or start a message with "In Ops sheet, …".'
    );
  }

  try {
    const columnsByTitle = await getSmartsheetColumns();
    if (action === 'columns') {
      const titles = Array.from(columnsByTitle.values())
        .map((c) => c.title)
        .join(', ');
      return `Smartsheet columns: ${titles}`;
    }
    if (action === 'clear') {
      const scope = (parts[2] || '').toLowerCase();
      if (scope === 'all' || scope === 'sheet' || parts.length < 3) {
        return await executeClearEntireSheet();
      }
      return 'Usage: `smartsheet clear all` — clears every editable cell in the configured sheet (rows stay).';
    }
    if (action === 'picklist') {
      const partsSp = trimmed.split(/\s+/);
      const colNum = Number(partsSp[2]);
      if (!Number.isFinite(colNum) || colNum < 1) {
        return 'Usage: `smartsheet picklist 4` — sets column 4 (1-based from the left) to a Yes/No dropdown. Not allowed on the primary column.';
      }
      try {
        return await executeSmartsheetColumnPicklistYesNo(colNum);
      } catch (e) {
        return `Smartsheet error: ${String(e?.message || e).slice(0, 900)}`;
      }
    }
    if (action === 'bold' || action === 'unbold') {
      const partsSp = trimmed.split(/\s+/);
      const colNum = Number(partsSp[2]);
      const rowNum = Number(partsSp[3]);
      if (!Number.isFinite(colNum) || colNum < 1 || !Number.isFinite(rowNum) || rowNum < 1) {
        return 'Usage: `smartsheet bold 1 7` or `smartsheet unbold 1 7` — 1-based column index and Smartsheet row number.';
      }
      try {
        return await executeSmartsheetCellBoldAtCoordinates(colNum, rowNum, action === 'bold');
      } catch (e) {
        return `Smartsheet error: ${String(e?.message || e).slice(0, 900)}`;
      }
    }
    if (action === 'format') {
      const parts = trimmed.split(/\s+/);
      const colTitle = parts[2];
      const rowStart = Number(parts[3]);
      const rowEnd = Number(parts[4]);
      const color = normalizeColorToHex(parts[5] || '');
      if (!colTitle || !Number.isFinite(rowStart) || !Number.isFinite(rowEnd) || !color) {
        return 'Invalid format. Example: smartsheet format Column2 1 5 #00FF00';
      }
      const mapped = await getRowIdsByRowNumbers(rowStart, rowEnd);
      if (!mapped.ok) {
        return `Could not find Smartsheet row number ${mapped.missingRowNumber} in this sheet (maybe fewer rows than requested).`;
      }
      const { paletteHex } = await applyColumnBackgroundForRows(mapped.ids, colTitle, columnsByTitle, color);
      return `Done — set background ${describeAppliedSmartsheetBackground(color, paletteHex)} for "${colTitle}" on row numbers ${rowStart}-${rowEnd}.`;
    }
    if (action === 'add') {
      const fieldsText = trimmed.replace(/^smartsheet\s+add\s+/i, '');
      const pairs = parseFieldPairs(fieldsText);
      if (pairs.length === 0) {
        return 'Invalid format. Example: smartsheet add Task=Call client | Status=In Progress';
      }
      const { cells, missingColumns } = toSmartsheetCells(pairs, columnsByTitle);
      if (missingColumns.length > 0) {
        return `Unknown column(s): ${missingColumns.join(', ')}`;
      }
      const result = await smartsheetRequest(`/sheets/${getActiveSmartsheetSheetId()}/rows`, {
        method: 'POST',
        body: [{ toBottom: true, cells }],
      });
      const rowId = result.result?.[0]?.id || 'created';
      return `Added row to Smartsheet (row id: ${rowId}).`;
    }

    if (action === 'update') {
      const match = trimmed.match(/^smartsheet\s+update\s+([0-9]+)\s+(.+)$/i);
      if (!match) {
        return 'Invalid format. Example: smartsheet update 123456789 Task=Follow up | Status=Done';
      }
      const rowId = Number(match[1]);
      const pairs = parseFieldPairs(match[2]);
      if (pairs.length === 0) {
        return 'No fields found to update.';
      }
      const { cells, missingColumns } = toSmartsheetCells(pairs, columnsByTitle);
      if (missingColumns.length > 0) {
        return `Unknown column(s): ${missingColumns.join(', ')}`;
      }
      await smartsheetRequest(`/sheets/${getActiveSmartsheetSheetId()}/rows`, {
        method: 'PUT',
        body: [{ id: rowId, cells }],
      });
      return `Updated row ${rowId} in Smartsheet.`;
    }

    return 'Unknown smartsheet command. Try: smartsheet help';
  } catch (err) {
    return `Smartsheet error: ${err?.message || err}`;
  }
}

async function interpretNaturalLanguageSmartsheetIntent(userText, columnsByTitle) {
  const columnList = Array.from(columnsByTitle.values())
    .map((col) => col.title)
    .join(', ');
  const message = await client.messages.create({
    model: 'claude-sonnet-4-20250514',
    temperature: 0,
    max_tokens: 500,
    system: SHEET_INTENT_SYSTEM_PROMPT,
    messages: [
      {
        role: 'user',
        content: `Column titles (use these exactly):\n${columnList}\n\nUser message:\n${userText}`,
      },
    ],
  });
  const raw = message.content?.[0]?.text?.trim() || '';
  return extractJsonObject(raw);
}

/** 1-based column index + row when the message names a column title before `row N` (e.g. "Primary Column row 10"). */
function resolveSmartsheetColumnTitleRowCoords(text, columnsOrdered) {
  if (!columnsOrdered || columnsOrdered.length === 0) return null;
  const lower = String(text || '').toLowerCase().replace(/\s+/g, ' ').trim();
  const titles = Array.from(
    new Set(columnsOrdered.map((c) => String(c.title || '').trim()).filter(Boolean))
  );
  titles.sort((a, b) => b.length - a.length);
  for (const title of titles) {
    const tl = title.toLowerCase();
    const escaped = tl
      .replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
      .split(/\s+/)
      .join('\\s+');
    const re = new RegExp(`\\b${escaped}\\s+row\\s*(\\d+)\\b`, 'i');
    const m = lower.match(re);
    if (m) {
      const row = Number(m[1]);
      const colIdx = columnsOrdered.findIndex((c) => String(c.title || '').trim() === title);
      if (colIdx >= 0 && row >= 1) return { col: colIdx + 1, row };
    }
  }
  return null;
}

function mightNeedNamedColumnForNumericDelta(text) {
  if (parseSmartsheetColumnRowRead(text)) return false;
  const lower = String(text || '').toLowerCase();
  return (
    /\b(increase|decrease|increment|decrement|bump\s+up|bump\s+down)\b/i.test(lower) ||
    /\badd\s+\d+(?:\.\d+)?\s+to\b/i.test(lower) ||
    /\bsubtract\s+\d/i.test(lower)
  );
}

/** 1-based column index + Smartsheet row number, e.g. "column 3 row 1" / "row 2 column 5". */
function parseSmartsheetColumnRowRead(text) {
  const lower = String(text || '').toLowerCase();
  const colFirst = lower.match(/\bcolumn\s*(\d+)\b[\s\S]{0,120}?\brow\s*(\d+)\b/i);
  if (colFirst) {
    const col = Number(colFirst[1]);
    const row = Number(colFirst[2]);
    if (col >= 1 && row >= 1) return { col, row };
  }
  const rowFirst = lower.match(/\brow\s*(\d+)\b[\s\S]{0,120}?\bcolumn\s*(\d+)\b/i);
  if (rowFirst) {
    const row = Number(rowFirst[1]);
    const col = Number(rowFirst[2]);
    if (col >= 1 && row >= 1) return { col, row };
  }
  const colAbbr = lower.match(/\bcol\.?\s*(\d+)\b[\s\S]{0,120}?\brow\s*(\d+)\b/i);
  if (colAbbr) {
    const col = Number(colAbbr[1]);
    const row = Number(colAbbr[2]);
    if (col >= 1 && row >= 1) return { col, row };
  }
  return null;
}

/** Plain English: "make column 1 row 7 bold" / "emphasize column 2 row 3" / "unbold column 2 row 3" (1-based column index). */
function parseSmartsheetCellBoldRequest(text) {
  const lower = String(text || '').toLowerCase();
  const boldOff =
    /\bunbold\b/.test(lower) ||
    /\bnot\s+bold\b/.test(lower) ||
    /\bremove\s+bold\b/.test(lower) ||
    /\bno\s+bold\b/.test(lower);
  const boldOn =
    !boldOff &&
    (/\bbold\b/.test(lower) ||
      /\bbolder\b/.test(lower) ||
      /\bemphasize\b/.test(lower) ||
      (/\bmake\b/.test(lower) && /\bbold\b/.test(lower)));
  if (!boldOn && !boldOff) return null;
  const coords = parseSmartsheetColumnRowRead(text);
  if (!coords) return null;
  return { col: coords.col, row: coords.row, bold: Boolean(boldOn) };
}

/**
 * Bump or lower a numeric cell (reads current value, writes the new number): "increase column 1 row 10 by 1",
 * "increase Primary Column row 10 by 2" (title must match the sheet when not using `column N`).
 */
function parseSmartsheetCellNumericDeltaRequest(text, columnsOrdered) {
  const lower = String(text || '').toLowerCase().replace(/\s+/g, ' ').trim();
  if (/\badd\s+row\b/i.test(lower)) return null;
  let coords = parseSmartsheetColumnRowRead(text);
  if (!coords && columnsOrdered && columnsOrdered.length > 0) {
    coords = resolveSmartsheetColumnTitleRowCoords(text, columnsOrdered);
  }
  if (!coords) return null;

  const addTo = lower.match(/\badd\s+(\d+(?:\.\d+)?)\s+to\b/);
  if (addTo) {
    const delta = Number(addTo[1]);
    return Number.isFinite(delta) ? { col: coords.col, row: coords.row, delta } : null;
  }

  if (/\bincrease\b|\bincrement\b|\bbump\s+up\b/.test(lower)) {
    const by = lower.match(/\bby\s+(-?\d+(?:\.\d+)?)\b/);
    const delta = by ? Number(by[1]) : 1;
    return Number.isFinite(delta) ? { col: coords.col, row: coords.row, delta } : null;
  }
  if (/\bdecrease\b|\bdecrement\b|\bbump\s+down\b/.test(lower)) {
    const by = lower.match(/\bby\s+(-?\d+(?:\.\d+)?)\b/);
    let mag = by ? Number(by[1]) : 1;
    if (!Number.isFinite(mag)) return null;
    mag = Math.abs(mag);
    return { col: coords.col, row: coords.row, delta: -mag };
  }

  const subFrom = lower.match(/\bsubtract\s+(\d+(?:\.\d+)?)\s+from\b/);
  if (subFrom) {
    const delta = -Number(subFrom[1]);
    return Number.isFinite(delta) ? { col: coords.col, row: coords.row, delta } : null;
  }

  return null;
}

/**
 * Same-sheet copy with scale: "make column 2 row 10 = to column 1 row 10 x 2",
 * "have column 2 row 10 equal to column 1 row 10 * 2",
 * "set column 2 row 10 to column 1 row 10 times 2.5".
 */
function parseSmartsheetCellCopyScaledFromCellRequest(text) {
  const lower = String(text || '').toLowerCase().replace(/\s+/g, ' ').trim();
  if (/\b(what|which|how\s+much|show|read|tell\s+me)\b/i.test(lower)) return null;

  const withVerb = lower.match(
    /\b(?:make|set|have)\s+column\s*(\d+)\s+row\s*(\d+)\s*(?:=\s*to|=\s*|equal\s+(?:to\s+)?|to\s+)(?:the\s+)?(?:value\s+of\s+)?(?:column|col\.?)\s*(\d+)\s+row\s*(\d+)\s*(?:\*|times|\bx)\s*(\d+(?:\.\d+)?)\b/i
  );
  if (withVerb) {
    const dstCol = Number(withVerb[1]);
    const dstRow = Number(withVerb[2]);
    const srcCol = Number(withVerb[3]);
    const srcRow = Number(withVerb[4]);
    const mult = Number(withVerb[5]);
    if ([dstCol, dstRow, srcCol, srcRow, mult].every((n) => Number.isFinite(n) && n > 0)) {
      return { dstCol, dstRow, srcCol, srcRow, mult };
    }
    return null;
  }

  const bare = lower.match(
    /\bcolumn\s*(\d+)\s+row\s*(\d+)\s*(?:=\s*to|=|equal\s+(?:to\s+)?)\s*(?:the\s+)?(?:value\s+of\s+)?(?:column|col\.?)\s*(\d+)\s+row\s*(\d+)\s*(?:\*|times|\bx)\s*(\d+(?:\.\d+)?)\b/i
  );
  if (bare) {
    const dstCol = Number(bare[1]);
    const dstRow = Number(bare[2]);
    const srcCol = Number(bare[3]);
    const srcRow = Number(bare[4]);
    const mult = Number(bare[5]);
    if ([dstCol, dstRow, srcCol, srcRow, mult].every((n) => Number.isFinite(n) && n > 0)) {
      return { dstCol, dstRow, srcCol, srcRow, mult };
    }
  }
  return null;
}

/** Only treat column/row coordinates as a cell *read* — not "create dropdown in column 4 row 1". */
function looksLikeSmartsheetColumnRowReadQuestion(text) {
  const lower = String(text || '').toLowerCase();
  if (
    /\b(create|add|make|insert|delete|remove|rename|convert|change|update|set|format|clear|empty|wipe|duplicate|move|hide)\b/i.test(
      lower
    )
  ) {
    return false;
  }
  if (
    /\b(drop[\s-]?down|dropdown|picklist|multi-?pick|contact\s+list|checkbox|formula|column\s+type|field\s+type|data\s+validation|yes\s*\/\s*no)\b/i.test(
      lower
    )
  ) {
    return false;
  }
  if (
    /\b(increase|decrease|increment|decrement|bump\s+up|bump\s+down)\b/i.test(lower) ||
    /\badd\s+\d+(?:\.\d+)?\s+to\b/i.test(lower) ||
    (/\bsubtract\s+\d/i.test(lower) && /\bfrom\b/.test(lower))
  ) {
    return false;
  }
  if (
    /\b(what|which|whose|contents?|how\s+much|show|read|tell\s+me|give\s+me|value\s+of)\b/i.test(lower)
  ) {
    return true;
  }
  const t = lower.replace(/\s+/g, ' ').trim();
  return (
    /^(?:what\s+is\s+)?(?:the\s+)?(?:value|contents?)\s+(?:of|in)\s+column\s*\d+\s+row\s*\d+\??$/i.test(t) ||
    /^column\s*\d+\s*,?\s*row\s*\d+\??$/i.test(t) ||
    /^row\s*\d+\s*,?\s*column\s*\d+\??$/i.test(t)
  );
}

async function readSmartsheetCellAtCoordinates(colIndex1Based, rowNumber) {
  const sheet = await getSmartsheetSheetForCellValues();
  const columns = sheet.columns || [];
  if (colIndex1Based > columns.length) {
    return `This sheet only has ${columns.length} column(s); column ${colIndex1Based} is out of range.`;
  }
  const colMeta = columns[colIndex1Based - 1];
  const colTitle = String(colMeta?.title || '').trim() || `Column ${colIndex1Based}`;
  const targetRow = (sheet.rows || []).find((r) => Number(r.rowNumber) === rowNumber);
  if (!targetRow) {
    return `I could not find **row ${rowNumber}** on this sheet (no row with that row number in the data returned from Smartsheet).`;
  }
  const cell = findCellOnRow(targetRow, colMeta.id);
  const value = cell ? cellText(cell) : '';
  const shown = value === '' ? '*(empty)*' : value;
  return `**${colTitle}** (column ${colIndex1Based}), row **${rowNumber}**: ${shown}`;
}

/** Set a single cell by 1-based column index and Smartsheet row number (e.g. after PICKLIST conversion). */
async function setSmartsheetCellValueAtColumnRow(columnIndex1Based, rowNumber, value) {
  const sheet = await getSmartsheetSheetForCellValues();
  const columns = sheet.columns || [];
  if (columnIndex1Based < 1 || columnIndex1Based > columns.length) {
    return `Could not set the cell: column ${columnIndex1Based} is out of range.`;
  }
  const colMeta = columns[columnIndex1Based - 1];
  const targetRow = (sheet.rows || []).find((r) => Number(r.rowNumber) === Number(rowNumber));
  if (!targetRow) {
    return `There is no **row ${rowNumber}** on this sheet yet — add that row in Smartsheet, then set column ${columnIndex1Based} to **${value}**.`;
  }
  await smartsheetRequest(`/sheets/${getActiveSmartsheetSheetId()}/rows`, {
    method: 'PUT',
    body: [{ id: targetRow.id, cells: [{ columnId: colMeta.id, value }] }],
  });
  invalidateSmartsheetColumnCache();
  const title = String(colMeta.title || `Column${columnIndex1Based}`).trim();
  return `Set **${title}** (column ${columnIndex1Based}), **row ${rowNumber}**, to **${value}**.`;
}

async function executeSmartsheetCellNumericDelta(columnIndex1Based, rowNumber, delta) {
  if (!Number.isFinite(delta)) {
    return 'I need a numeric amount (for example **by 1** or **by 0.5**).';
  }
  const sheet = await getSmartsheetSheetForCellValues();
  const columns = sheet.columns || [];
  if (columnIndex1Based < 1 || columnIndex1Based > columns.length) {
    return `This sheet has ${columns.length} column(s); column ${columnIndex1Based} is out of range.`;
  }
  const colMeta = columns[columnIndex1Based - 1];
  if (!isSmartsheetColumnClearable(colMeta)) {
    return `Column **${colMeta.title || columnIndex1Based}** is not user-editable here (system or formula column).`;
  }
  const targetRow = (sheet.rows || []).find((r) => Number(r.rowNumber) === Number(rowNumber));
  if (!targetRow) {
    return `I could not find **row ${rowNumber}** on this sheet.`;
  }
  const cell = findCellOnRow(targetRow, colMeta.id);
  const raw = cell ? cellText(cell) : '';
  const trimmed = String(raw).trim();
  let cur;
  if (trimmed === '') {
    cur = 0;
  } else {
    cur = Number(trimmed.replace(/,/g, ''));
    if (!Number.isFinite(cur)) {
      return `That cell is not a plain number (**${raw}**). Put a number in the cell first, or clear the formula and enter a fixed value.`;
    }
  }
  const next = cur + delta;
  await smartsheetRequest(`/sheets/${getActiveSmartsheetSheetId()}/rows`, {
    method: 'PUT',
    body: [{ id: targetRow.id, cells: [{ columnId: colMeta.id, value: next }] }],
  });
  invalidateSmartsheetColumnCache();
  const title = String(colMeta.title || `Column${columnIndex1Based}`).trim();
  const verb = delta >= 0 ? 'increased' : 'decreased';
  return `**${title}** (column ${columnIndex1Based}), row **${rowNumber}**: was **${trimmed === '' ? '(empty → 0)' : raw}**, ${verb} by **${Math.abs(delta)}** → set to **${next}**.`;
}

async function executeSmartsheetCellCopyScaledFromCell(spec) {
  const { dstCol, dstRow, srcCol, srcRow, mult } = spec;
  if (!Number.isFinite(mult)) {
    return 'Invalid multiplier for that cell update.';
  }
  const sheet = await getSmartsheetSheetForCellValues();
  const columns = sheet.columns || [];
  for (const [label, idx] of [
    ['destination', dstCol],
    ['source', srcCol],
  ]) {
    if (idx < 1 || idx > columns.length) {
      return `This sheet has ${columns.length} column(s); ${label} column ${idx} is out of range.`;
    }
  }
  const dstMeta = columns[dstCol - 1];
  if (!isSmartsheetColumnClearable(dstMeta)) {
    return `Column **${dstMeta.title || dstCol}** is not user-editable here (system or formula column).`;
  }
  const srcMeta = columns[srcCol - 1];
  const srcRowObj = (sheet.rows || []).find((r) => Number(r.rowNumber) === Number(srcRow));
  if (!srcRowObj) {
    return `I could not find **source row ${srcRow}** on this sheet.`;
  }
  const dstRowObj = (sheet.rows || []).find((r) => Number(r.rowNumber) === Number(dstRow));
  if (!dstRowObj) {
    return `I could not find **destination row ${dstRow}** on this sheet.`;
  }
  const srcCell = findCellOnRow(srcRowObj, srcMeta.id);
  const raw = srcCell ? cellText(srcCell) : '';
  const trimmed = String(raw).trim();
  let cur;
  if (trimmed === '') {
    cur = 0;
  } else {
    cur = Number(trimmed.replace(/,/g, ''));
    if (!Number.isFinite(cur)) {
      return `Source cell is not a plain number (**${raw}**). Put a number in column ${srcCol} row ${srcRow} first.`;
    }
  }
  const next = cur * mult;
  await smartsheetRequest(`/sheets/${getActiveSmartsheetSheetId()}/rows`, {
    method: 'PUT',
    body: [{ id: dstRowObj.id, cells: [{ columnId: dstMeta.id, value: next }] }],
  });
  invalidateSmartsheetColumnCache();
  const dstTitle = String(dstMeta.title || `Column${dstCol}`).trim();
  return `Set **${dstTitle}** (column ${dstCol}), row **${dstRow}**, to **${next}** (column ${srcCol} row ${srcRow} was **${trimmed === '' ? '(empty → 0)' : raw}**, × **${mult}**).`;
}

/** Set bold on/off for one cell (Smartsheet format descriptor index 2). */
async function executeSmartsheetCellBoldAtCoordinates(columnIndex1Based, rowNumber, wantBold) {
  const sheet = await getSmartsheetSheetForCellValuesIncludeFormat();
  const columns = sheet.columns || [];
  if (columnIndex1Based < 1 || columnIndex1Based > columns.length) {
    return `This sheet has ${columns.length} column(s); column ${columnIndex1Based} is out of range.`;
  }
  const colMeta = columns[columnIndex1Based - 1];
  if (!isSmartsheetColumnClearable(colMeta)) {
    return `Column **${colMeta.title || columnIndex1Based}** is not user-editable here (system or formula column).`;
  }
  const targetRow = (sheet.rows || []).find((r) => Number(r.rowNumber) === Number(rowNumber));
  if (!targetRow) {
    return `I could not find **row ${rowNumber}** on this sheet.`;
  }
  const existing = findCellOnRow(targetRow, colMeta.id);
  const value = pickCellValueForBackgroundUpdate(colMeta, existing);
  const priorFormat = existing?.format;
  const newFormat = mergeSmartsheetFormatDescriptorBold(priorFormat, wantBold);
  await smartsheetRequest(`/sheets/${getActiveSmartsheetSheetId()}/rows`, {
    method: 'PUT',
    body: [{ id: targetRow.id, cells: [{ columnId: colMeta.id, value, format: newFormat }] }],
  });
  invalidateSmartsheetColumnCache();
  const title = String(colMeta.title || `Column${columnIndex1Based}`).trim();
  const state = wantBold ? '**bold**' : 'default weight (not bold)';
  return `Done — set ${state} for **${title}** (column ${columnIndex1Based}), row **${rowNumber}**.`;
}

function collectSmartsheetKeepRowNumbersFromClearMessage(lower) {
  const set = new Set();
  const re = /(?:except(?:\s+for)?|excluding|but\s+keep|leave)\s+row\s*(\d+)/gi;
  let m;
  while ((m = re.exec(lower)) !== null) {
    const n = Number(m[1]);
    if (Number.isFinite(n) && n >= 1) set.add(n);
  }
  return set;
}

/**
 * Clear cells in one column only (optionally keep some row numbers, optional Yes/No picklist + default No on kept rows).
 * Does not match whole-sheet clears like "empty every cell" with no column.
 */
function parseSmartsheetClearColumnCellsRequest(text) {
  const lower = String(text || '').toLowerCase();
  if (!/\b(empty|clear|wipe|erase)\b/i.test(lower)) return null;
  const colM = lower.match(/\bcolumn\s*(\d+)\b/i);
  if (!colM) return null;
  const col = Number(colM[1]);
  if (col < 1) return null;

  const scoped =
    /\b(empty|clear|wipe|erase)\b[\s\S]{0,120}\b(all|every)\b[\s\S]{0,80}\bcells?\b[\s\S]{0,120}\bcolumn\s*\d+/i.test(
      lower
    ) ||
    /\b(empty|clear|wipe|erase)\b[\s\S]{0,60}\bcells?\b[\s\S]{0,100}\bcolumn\s*\d+/i.test(lower) ||
    /\b(empty|clear|wipe|erase)\b[\s\S]{0,100}\bcolumn\s*\d+\b[\s\S]{0,120}\b(except|excluding|leave|but\s+keep)/i.test(
      lower
    ) ||
    /\b(empty|clear|wipe|erase)\b\s+column\s*\d+\b/i.test(lower);

  if (!scoped) return null;

  const keepRows = collectSmartsheetKeepRowNumbersFromClearMessage(lower);
  const wantPicklist =
    /\b(yes\s*\/\s*no|yes-no)\b/i.test(lower) ||
    (/\b(drop[\s-]*downs?|dropdowns?|picklists?)\b/i.test(lower) && /\byes\b/.test(lower) && /\bno\b/.test(lower));
  const defaultNo =
    /\bdefault\b[\s\S]{0,40}\b(no|n)\b/i.test(lower) ||
    /\bdefaults?\s+to\s+(no|n)\b/i.test(lower) ||
    /\bdefault\s+be\s+["']?no["']?\b/i.test(lower);

  return { col, keepRows, wantPicklist, defaultNo };
}

/** Natural language: "create yes/no dropdown in column 4 …" (picklist is per-column; optional row + default No). */
function parseSmartsheetYesNoPicklistColumnRequest(text) {
  const lower = String(text || '').toLowerCase();
  if (!/\b(create|add|make|set|convert|change|turn|put)\b/i.test(lower)) return null;
  const colM = lower.match(/\bcolumn\s*(\d+)\b/i);
  if (!colM) return null;
  const col = Number(colM[1]);
  if (col < 1) return null;
  const hasDropdownWord = /\b(drop[\s-]*downs?|dropdowns?|picklists?)\b/i.test(lower);
  const hasYesNoSlash = /\byes\s*\/\s*no\b/i.test(lower) || /\byes-no\b/i.test(lower);
  const hasYesAndNo = /\byes\b/.test(lower) && /\bno\b/.test(lower);
  if (!hasDropdownWord && !hasYesNoSlash && !hasYesAndNo) return null;
  const wantsDefaultNo =
    /\bdefault\b[\s\S]{0,40}\b(no|n)\b/i.test(lower) ||
    /\bdefaults?\s+to\s+(no|n)\b/i.test(lower) ||
    /\bdefault\s+be\s+["']?no["']?\b/i.test(lower);
  let defaultNoRow = null;
  if (wantsDefaultNo) {
    const rowM = lower.match(/\brow\s*(\d+)\b/i);
    if (rowM) {
      const rn = Number(rowM[1]);
      if (Number.isFinite(rn) && rn >= 1) defaultNoRow = rn;
    }
  }
  return { col, defaultNoRow };
}

/** True if Smartsheet column metadata is already a Yes/No style PICKLIST. */
function smartsheetColumnIsYesNoPicklist(col) {
  if (!col || String(col.type || '') !== 'PICKLIST') return false;
  const opts = col.options;
  if (!Array.isArray(opts) || opts.length < 2) return false;
  const norm = opts.map((o) => String(o).trim().toLowerCase()).filter(Boolean);
  return norm.includes('yes') && norm.includes('no');
}

async function executeSmartsheetColumnPicklistYesNo(columnIndex1Based) {
  const sheetId = getActiveSmartsheetSheetId();
  if (!sheetId) {
    throw new Error('SMARTSHEET_SHEET_ID is not configured (set SMARTSHEET_SHEET_ID or SMARTSHEET_SHEET_MAP).');
  }
  const sheet = await smartsheetRequest(`/sheets/${sheetId}`);
  const sheetName = String(sheet.name || '').trim() || '(unnamed sheet)';
  const columns = sheet.columns || [];
  if (columnIndex1Based < 1 || columnIndex1Based > columns.length) {
    return `This sheet has ${columns.length} column(s); column ${columnIndex1Based} is out of range.`;
  }
  const col = columns[columnIndex1Based - 1];
  if (col.primary === true || Number(col.primary) === 1) {
    return 'Smartsheet does not allow converting the **primary column** to a picklist via API. Pick a different column index.';
  }
  const ctype = String(col.type || '');
  if (
    ctype === 'PREDECESSOR' ||
    ctype === 'AUTO_NUMBER' ||
    ctype.includes('SYSTEM') ||
    ctype === 'ABSTRACT_DATETIME'
  ) {
    return `Column **${col.title || columnIndex1Based}** is type \`${ctype}\` and cannot be switched to a Yes/No picklist here.`;
  }
  const title = String(col.title || `Column${columnIndex1Based}`).trim();
  const limitation =
    '**Smartsheet cannot put a Yes/No dropdown on only one cell** — a dropdown is a **column** setting, so **every row** in that column shares the same Yes/No list (not Excel-style validation on one cell).';

  if (smartsheetColumnIsYesNoPicklist(col)) {
    return (
      `${limitation}\n\n` +
      `**${title}** (column ${columnIndex1Based}) on **"${sheetName}"** (sheet id \`${sheetId}\`) **was already** a Yes/No \`PICKLIST\`. No column-type API call was sent.\n\n` +
      '_If this is the wrong sheet, say **in Exact Sheet Name sheet,** or **in Exact Sheet Name,** so it matches `SMARTSHEET_SHEET_MAP`._'
    );
  }

  const idx = Number.isFinite(Number(col.index)) ? Number(col.index) : columnIndex1Based - 1;
  await smartsheetRequest(`/sheets/${sheetId}/columns/${col.id}`, {
    method: 'PUT',
    body: {
      title,
      index: idx,
      type: 'PICKLIST',
      options: ['Yes', 'No'],
    },
  });
  invalidateSmartsheetColumnCache();

  const sheet2 = await smartsheetRequest(`/sheets/${sheetId}`);
  const col2 = (sheet2.columns || [])[columnIndex1Based - 1];
  if (!smartsheetColumnIsYesNoPicklist(col2)) {
    const got = col2 ? String(col2.type || '') : 'missing column';
    return (
      `${limitation}\n\n` +
      `The bot called Smartsheet to update **${title}** (column ${columnIndex1Based}) on **"${sheetName}"** (id \`${sheetId}\`), but **after** the call the column was **not** a Yes/No picklist (Smartsheet now reports type \`${got}\`). **Do not trust an old “Done” reply** — please check this sheet in the UI or retry.\n\n` +
      `_Column type before the call was \`${ctype}\`._`
    );
  }

  const afterTitle = String(col2.title || title).trim();
  return (
    `${limitation}\n\n` +
    `**Done** — **${afterTitle}** (column ${columnIndex1Based}) on **"${sheetName}"** (id \`${sheetId}\`) is confirmed as a Yes/No \`PICKLIST\` after an API re-check.\n\n` +
    '_A row number in your message only selects which cell to set to **No** afterward (when you asked for a default), not which rows get the dropdown._\n' +
    '_Smartsheet may convert or clear existing values in that column when the type changes._'
  );
}

async function executeSmartsheetClearColumnCells(spec) {
  const columnIndex1Based = spec.col;
  const keep = spec.keepRows instanceof Set ? spec.keepRows : new Set(Array.isArray(spec.keepRows) ? spec.keepRows : []);
  const wantPicklist = Boolean(spec.wantPicklist);
  const defaultNo = Boolean(spec.defaultNo);
  const sheetId = getActiveSmartsheetSheetId();
  if (!sheetId) {
    throw new Error('SMARTSHEET_SHEET_ID is not configured (set SMARTSHEET_SHEET_ID or SMARTSHEET_SHEET_MAP).');
  }

  let { rows, columns } = await getAllSmartsheetRowsAndColumns();
  if (columnIndex1Based < 1 || columnIndex1Based > columns.length) {
    return `This sheet has ${columns.length} column(s); column ${columnIndex1Based} is out of range.`;
  }

  if (wantPicklist) {
    const pickMsg = await executeSmartsheetColumnPicklistYesNo(columnIndex1Based);
    const pickOk =
      /\*\*Done\*\*/.test(pickMsg) ||
      /\*\*was already\*\*/.test(pickMsg) ||
      /No column-type API call was sent/.test(pickMsg);
    if (!pickOk) {
      return pickMsg;
    }
    ({ rows, columns } = await getAllSmartsheetRowsAndColumns());
  }

  const col = columns[columnIndex1Based - 1];
  if (!col || col.id == null) {
    return 'Could not resolve that column after loading the sheet.';
  }
  if (!isSmartsheetColumnClearable(col)) {
    return `Column **${col.title || columnIndex1Based}** is not user-editable here (system or formula column).`;
  }

  const payloads = [];
  let clearedOtherRows = 0;
  for (const row of rows) {
    const rn = Number(row.rowNumber);
    const existing = findCellOnRow(row, col.id);
    let value;
    if (keep.has(rn)) {
      if (wantPicklist && defaultNo) {
        value = 'No';
      } else {
        value = pickCellValueForBackgroundUpdate(col, existing);
      }
    } else {
      value = pickCellValueForBackgroundUpdate(col, null);
      clearedOtherRows += 1;
    }
    payloads.push({ id: row.id, cells: [{ columnId: col.id, value }] });
  }

  for (let i = 0; i < payloads.length; i += SMARTSHEET_ROW_PUT_BATCH) {
    const batch = payloads.slice(i, i + SMARTSHEET_ROW_PUT_BATCH);
    await smartsheetRequest(`/sheets/${sheetId}/rows`, { method: 'PUT', body: batch });
  }
  invalidateSmartsheetColumnCache();

  const title = String(col.title || `Column${columnIndex1Based}`).trim();
  const keptSorted = Array.from(keep).sort((a, b) => a - b);
  let msg =
    `Done — **${title}** (column ${columnIndex1Based}): cleared **${clearedOtherRows}** cell(s) in that column only (other columns untouched).`;
  if (keptSorted.length > 0) {
    msg += ` Row number(s) **${keptSorted.join(', ')}** were kept`;
    msg +=
      wantPicklist && defaultNo
        ? ' and set to **No**.'
        : wantPicklist
          ? ' (values preserved where possible).'
          : ' (values preserved).';
  }
  if (wantPicklist) {
    msg +=
      ' The column is a Yes/No `PICKLIST` for **every row** in Smartsheet; only values were scoped as you asked.';
  }
  return msg;
}

async function findRowIdByMatch(columnsByTitle, matchColumn, matchValue) {
  if (!matchColumn || !matchValue) return { rowId: null, ambiguous: false };
  const col = columnsByTitle.get(String(matchColumn).toLowerCase());
  if (!col) return { rowId: null, ambiguous: false };
  const sheet = await getSmartsheetSheet();
  const target = String(matchValue).toLowerCase().trim();
  const matches = [];
  for (const row of sheet.rows || []) {
    for (const cell of row.cells || []) {
      if (cell.columnId !== col.id) continue;
      const raw = cell.displayValue ?? cell.value;
      if (raw == null) continue;
      if (String(raw).toLowerCase().trim() === target) matches.push(row.id);
    }
  }
  if (matches.length === 1) return { rowId: matches[0], ambiguous: false };
  if (matches.length > 1) return { rowId: null, ambiguous: true };
  return { rowId: null, ambiguous: false };
}

async function handlePlainEnglishSmartsheet(userText) {
  try {
    const columnsByTitle = await getSmartsheetColumns();
    const scaledCopy = parseSmartsheetCellCopyScaledFromCellRequest(userText);
    if (scaledCopy) {
      try {
        return await executeSmartsheetCellCopyScaledFromCell(scaledCopy);
      } catch (e) {
        return `Smartsheet could not set that cell (${String(e?.message || e).slice(0, 900)}).`;
      }
    }
    let numDelta = parseSmartsheetCellNumericDeltaRequest(userText, null);
    if (!numDelta && mightNeedNamedColumnForNumericDelta(userText)) {
      const sNum = await getSmartsheetSheetForCellValues();
      numDelta = parseSmartsheetCellNumericDeltaRequest(userText, sNum.columns || []);
    }
    if (numDelta) {
      try {
        return await executeSmartsheetCellNumericDelta(numDelta.col, numDelta.row, numDelta.delta);
      } catch (e) {
        return `Smartsheet could not update that cell (${String(e?.message || e).slice(0, 900)}).`;
      }
    }
    const readCoords = parseSmartsheetColumnRowRead(userText);
    if (readCoords && looksLikeSmartsheetColumnRowReadQuestion(userText)) {
      return await readSmartsheetCellAtCoordinates(readCoords.col, readCoords.row);
    }
    const boldReq = parseSmartsheetCellBoldRequest(userText);
    if (boldReq) {
      try {
        return await executeSmartsheetCellBoldAtCoordinates(boldReq.col, boldReq.row, boldReq.bold);
      } catch (e) {
        return `Smartsheet could not set bold on that cell (${String(e?.message || e).slice(0, 900)}).`;
      }
    }
    const colClearSpec = parseSmartsheetClearColumnCellsRequest(userText);
    if (colClearSpec) {
      try {
        return await executeSmartsheetClearColumnCells(colClearSpec);
      } catch (e) {
        return `Smartsheet could not clear/update that column (${String(e?.message || e).slice(0, 900)}).`;
      }
    }
    const pickReq = parseSmartsheetYesNoPicklistColumnRequest(userText);
    if (pickReq) {
      try {
        let msg = await executeSmartsheetColumnPicklistYesNo(pickReq.col);
        if (pickReq.defaultNoRow != null) {
          const hardFail =
            /out of range|primary column|cannot be switched|was \*\*not\*\* a Yes\/No|cannot put/i.test(msg);
          const allowCell =
            !hardFail &&
            (/\*\*Done\*\*/.test(msg) ||
              /\*\*was already\*\*/.test(msg) ||
              /No column-type API call was sent/.test(msg));
          if (allowCell) {
            try {
              const cellMsg = await setSmartsheetCellValueAtColumnRow(
                pickReq.col,
                pickReq.defaultNoRow,
                'No'
              );
              msg = `${msg}\n\n${cellMsg}`;
            } catch (cellErr) {
              msg = `${msg}\n\nCould not set default **No** on row ${pickReq.defaultNoRow}: ${String(
                cellErr?.message || cellErr
              ).slice(0, 500)}`;
            }
          }
        }
        return msg;
      } catch (e) {
        return `Smartsheet could not update that column (${String(e?.message || e).slice(0, 900)}). If the column already has incompatible formatting or formulas, change it in the Smartsheet UI or pick another column.`;
      }
    }
    const containsReplaceReply = await executeContainsReplaceRequest(userText, columnsByTitle);
    if (containsReplaceReply) return containsReplaceReply;
    const implicitStatusReply = await executeImplicitJobFinishedUpdate(userText, columnsByTitle);
    if (implicitStatusReply) return implicitStatusReply;
    const changeJobLocReply = await executeChangeValueForJobLocation(userText, columnsByTitle);
    if (changeJobLocReply) return changeJobLocReply;
    if (looksLikeClearWholeSheetRequest(userText)) {
      return await executeClearEntireSheet();
    }
    const colorSwap = parseBackgroundColorChangeRequest(userText);
    if (colorSwap) return await executeResetBackgroundColor(columnsByTitle, colorSwap.sourceHex, colorSwap.targetHex);
    const intent = await interpretNaturalLanguageSmartsheetIntent(userText, columnsByTitle);
    const range = parseRowNumberRange(userText);
    const heuristicColor = inferBackgroundColorFromText(userText);
    const heuristicColumn = inferColumnTitleFromText(userText, columnsByTitle);
    const canHeuristicFormat = Boolean(range && heuristicColor && heuristicColumn);

    async function applyHeuristicFormat() {
      const mapped = await getRowIdsByRowNumbers(range.start, range.end);
      if (!mapped.ok) {
        return `Could not find Smartsheet row number ${mapped.missingRowNumber} in this sheet (maybe fewer rows than requested).`;
      }
      const { paletteHex } = await applyColumnBackgroundForRows(
        mapped.ids,
        heuristicColumn,
        columnsByTitle,
        heuristicColor
      );
      return `Done — set background ${describeAppliedSmartsheetBackground(heuristicColor, paletteHex)} for "${heuristicColumn}" on row numbers ${range.start}-${range.end}.`;
    }

    if (!intent || intent.action === 'none') {
      const conf = Number(intent?.confidence);
      if (canHeuristicFormat) return applyHeuristicFormat();
      if (looksLikeSheetTopic(userText)) {
        const titles = Array.from(columnsByTitle.values())
          .map((c) => c.title)
          .join(', ');
        return (
          'I can update this Smartsheet via API, but I need one concrete instruction.\n' +
          `Columns: ${titles}\n` +
          'Examples:\n' +
          '- Read one cell: "What is in column 3 row 1?" (column numbers are 1-based from the left)\n' +
          '- Bold one cell: "Make column 1 row 7 bold" or `smartsheet bold 1 7` (`smartsheet unbold 1 7` to remove)\n' +
          '- Yes/No dropdown on a column: "Create a yes/no drop-down in column 4" or `smartsheet picklist 4`\n' +
          '- Clear one column only: "Clear all cells in column 4 except row 1" (optional: add yes/no + default No on kept rows)\n' +
          '- Add row: "Add row: Primary Column = X | Column2 = Y"\n' +
          '- Update by row id: "Update row 1234567890: Column2 = Done"\n' +
          '- Green background rows 1-5 in Column2: "Set Column2 background green for rows 1-5"\n' +
          '- Bump a number in one cell: "Increase column 1 row 10 by 1" or "Increase Primary Column row 10 by 2"\n' +
          '- Copy value × multiplier: "Have column 2 row 10 equal to column 1 row 10 * 2"\n' +
          '- Clear the whole sheet (editable cells only): "Empty every cell and start over" or `smartsheet clear all`\n' +
          '- Or explicit: `smartsheet format Column2 1 5 #00FF00`'
        );
      }
      if (conf && conf < 0.5) return null;
      return null;
    }
    if (intent.clarification) {
      if (canHeuristicFormat) return applyHeuristicFormat();
      return intent.clarification;
    }
    if (intent.action === 'clear_sheet') {
      return await executeClearEntireSheet();
    }
    if (intent.action === 'read_cell') {
      const r = intent.read || {};
      let colIdx = Number(r.column_index);
      const rowNum = Number(r.row_number);
      if ((!Number.isFinite(colIdx) || colIdx < 1) && r.column_title) {
        const sheet = await getSmartsheetSheetForCellValues();
        const cols = sheet.columns || [];
        const t = String(r.column_title).toLowerCase().trim();
        const idx = cols.findIndex((c) => String(c.title || '').toLowerCase().trim() === t);
        if (idx >= 0) colIdx = idx + 1;
      }
      if (Number.isFinite(colIdx) && colIdx >= 1 && Number.isFinite(rowNum) && rowNum >= 1) {
        return await readSmartsheetCellAtCoordinates(colIdx, rowNum);
      }
      return 'Say which cell to read like **column 3 row 1** (column numbers are 1-based from the left), or name the exact column title and row number.';
    }
    const pairs = Object.entries(intent.fields || {}).map(([key, value]) => ({
      key,
      value: String(value),
    }));
    const { cells, missingColumns } = toSmartsheetCells(pairs, columnsByTitle);
    if (missingColumns.length > 0) {
      return `I could not find these column titles: ${missingColumns.join(', ')}. Use exact titles from the sheet: ${Array.from(
        columnsByTitle.values()
      )
        .map((c) => c.title)
        .join(', ')}`;
    }
    if (intent.action === 'add') {
      if (cells.length === 0) return 'What should I put in the new row? Name the columns and values.';
      const result = await smartsheetRequest(`/sheets/${getActiveSmartsheetSheetId()}/rows`, {
        method: 'POST',
        body: [{ toBottom: true, cells }],
      });
      const rowId = result.result?.[0]?.id || 'created';
      return formatConfirmation('add', rowId, pairs);
    }
    if (intent.action === 'update') {
      if (cells.length === 0) return 'Which fields should I change, and to what values?';
      let rowId = intent.row_id ? Number(intent.row_id) : null;
      if (!rowId) {
        const found = await findRowIdByMatch(columnsByTitle, intent.match_column, intent.match_value);
        if (found.ambiguous) {
          return `I found more than one row where "${intent.match_column}" is "${intent.match_value}". Which row (row id), or give more specific match details?`;
        }
        rowId = found.rowId;
      }
      if (!rowId) {
        return 'I could not tell which row to update. Say the Smartsheet row id, or say e.g. "update the row where Task is Client onboarding".';
      }
      await smartsheetRequest(`/sheets/${getActiveSmartsheetSheetId()}/rows`, {
        method: 'PUT',
        body: [{ id: rowId, cells }],
      });
      return formatConfirmation('update', rowId, pairs);
    }
    if (intent.action === 'format') {
      const fmt = intent.format || {};
      const colTitle = fmt.column || heuristicColumn;
      const rowStart = Number.isFinite(Number(fmt.row_start)) ? Number(fmt.row_start) : range?.start;
      const rowEnd = Number.isFinite(Number(fmt.row_end)) ? Number(fmt.row_end) : range?.end;
      const color = normalizeColorToHex(
        fmt.background_color || inferBackgroundColorFromText(userText) || heuristicColor || ''
      );
      if (!colTitle || !Number.isFinite(rowStart) || !Number.isFinite(rowEnd) || !color) {
        if (canHeuristicFormat) return applyHeuristicFormat();
        return 'I could not parse that formatting request. Try: "Set Column2 background green for rows 1-5".';
      }
      const mapped = await getRowIdsByRowNumbers(rowStart, rowEnd);
      if (!mapped.ok) {
        return `Could not find Smartsheet row number ${mapped.missingRowNumber} in this sheet (maybe fewer rows than requested).`;
      }
      const { paletteHex } = await applyColumnBackgroundForRows(mapped.ids, colTitle, columnsByTitle, color);
      return `Done — set background ${describeAppliedSmartsheetBackground(color, paletteHex)} for "${colTitle}" on row numbers ${rowStart}-${rowEnd}.`;
    }
    return null;
  } catch (err) {
    return `Smartsheet error: ${err?.message || err}`;
  }
}

app.get('/health', (_req, res) => {
  res.status(200)
    .type('application/json')
    .send(
      JSON.stringify({
        ok: true,
        build: AGENT_CODE_BUILD,
        smartsheetMcpAgent: String(process.env.SMARTSHEET_USE_OFFICIAL_MCP || '').trim() === '1',
      })
    );
});

app.get('/slack/events', (_req, res) => {
  res.status(200).type('application/json').send(
    JSON.stringify({
      ok: true,
      build: AGENT_CODE_BUILD,
      note: 'Slack Event Subscriptions use HTTP POST to this same path, not GET. If you see this in a browser, the URL is reachable but Slack must POST JSON here.',
    })
  );
});

app.get('/autodesk/connect', (req, res) => {
  if (!buildingConnectedConfigured()) {
    return res.status(500).type('text/plain').send(
      'Missing AUTODESK_CLIENT_ID/AUTODESK_CLIENT_SECRET. Configure .env and restart.'
    );
  }
  const authUrl = getAutodeskAuthorizeUrl(req);
  return res.redirect(authUrl);
});

app.get('/autodesk/callback', async (req, res) => {
  try {
    const state = String(req.query?.state || '');
    const code = String(req.query?.code || '');
    const error = String(req.query?.error || '');
    if (error) {
      return res
        .status(400)
        .type('text/plain')
        .send(`Autodesk authorization failed: ${error}. You can retry by opening /autodesk/connect.`);
    }
    if (!isValidAutodeskOauthState(state)) {
      return res
        .status(400)
        .type('text/plain')
        .send('Invalid or expired OAuth state. Retry by opening /autodesk/connect and approving again.');
    }
    if (!code) {
      return res.status(400).type('text/plain').send('Missing authorization code from Autodesk callback.');
    }
    const redirectUri = getAutodeskRedirectUri(req);
    await exchangeAutodeskAuthorizationCode(code, redirectUri);
    return res
      .status(200)
      .type('text/plain')
      .send('Autodesk authorization complete. Return to Slack and run `bc whoami` or `bc projects 5`.');
  } catch (err) {
    return res.status(500).type('text/plain').send(`OAuth callback failed: ${String(err?.message || err)}`);
  }
});

app.post('/slack/events', async (req, res) => {
  console.log('[slack/events] incoming POST', new Date().toISOString(), {
    bodyType: String(req.body?.type || ''),
    innerType: String(req.body?.event?.type || ''),
  });
  if (req.body?.type === 'url_verification') {
    return res.json({ challenge: req.body.challenge });
  }
  const verified = verifySlack(req);
  if (!verified) {
    console.warn('[slack/events] verify FAILED → 403:', slackSignatureRejectReason(req));
    return res.sendStatus(403);
  }
  const envelopeType = req.body?.type;
  const event = req.body?.event;
  if (envelopeType !== 'event_callback' || !event || typeof event !== 'object') {
    console.warn('[slack/events] ignored envelope:', envelopeType, 'hasEvent=', Boolean(event));
    return res.sendStatus(200);
  }
  console.log('[slack/events]', envelopeType, 'inner=', event.type, 'ch=', event.channel, 'sub=', event.subtype || '');
  console.log('Slack event received:', envelopeType, event.type);
  if (event.subtype === 'bot_message') {
    console.log('[slack/events] skip: bot_message');
    return res.sendStatus(200);
  }
  if (slackBotUserId && event.user && event.user === slackBotUserId) {
    console.log('[slack/events] skip: event.user matches bot user id', slackBotUserId);
    return res.sendStatus(200);
  }
  if (event.type !== 'message' && event.type !== 'app_mention') {
    console.log('[slack/events] skip: inner type not message/app_mention:', event.type);
    return res.sendStatus(200);
  }
  if (!event.channel) {
    console.warn('[slack/events] skip: missing event.channel');
    return res.sendStatus(200);
  }
  const ignoredMessageSubtypes = new Set([
    'message_changed',
    'message_deleted',
    'channel_join',
    'channel_leave',
    'channel_topic',
    'channel_purpose',
    'channel_name',
    'channel_archive',
    'channel_unarchive',
    'pinned_item',
    'unpinned_item',
  ]);
  if (event.type === 'message' && event.subtype && ignoredMessageSubtypes.has(event.subtype)) {
    console.log('[slack/events] skip: ignored message subtype', event.subtype);
    return res.sendStatus(200);
  }
  res.sendStatus(200);
  try {
    const replyThreadTs = slackReplyThreadTs(event);
    const audioFile = (event.files || []).find(isAudioSlackFile);
    let userText = slackUserMessagePlainText(event);

    if (audioFile) {
      const privateUrl = audioFile.url_private_download || audioFile.url_private;
      if (!privateUrl) {
        await postSlackMessage({
          channel: event.channel,
          text: 'I found an audio file but could not access its download URL.',
          thread_ts: replyThreadTs,
        });
        return;
      }
      const slowHint = startSlackSlowWorkHintTimer({
        channel: event.channel,
        thread_ts: replyThreadTs,
        text: '_Transcribing audio — that usually takes more than a few seconds…_',
      });
      let audioBuffer;
      let transcript;
      try {
        audioBuffer = await downloadSlackFile(privateUrl);
        transcript = await transcribeAudioBuffer(
          audioBuffer,
          audioFile.name || 'audio-message.m4a',
          audioFile.mimetype
        );
      } finally {
        slowHint.cancel();
      }
      userText = transcript.trim();
      if (!userText) {
        await postSlackMessage({
          channel: event.channel,
          text: 'I could not transcribe that audio. Please try a clearer recording.',
          thread_ts: replyThreadTs,
        });
        return;
      }
    }

    if (!userText) {
      await postSlackMessage({
        channel: event.channel,
        text:
          event.type === 'app_mention'
            ? 'I did not see any text after the mention. What should I do?'
            : 'I did not get any plain text from that message (sometimes Slack only sends formatting or attachments). Try again with a short written instruction, or @ mention me with your request.',
        thread_ts: replyThreadTs,
      });
      return;
    }

    /** One short “started” post per user message; final outcome is always a separate postSlackMessage. */
    let slackWorkAckPosted = false;
    async function postSlackWorkStartedIfNeeded() {
      if (slackWorkAckPosted) return;
      slackWorkAckPosted = true;
      await postSlackMessage({
        channel: event.channel,
        text: '_Working on that — I’ll post another message when the result is ready._',
        thread_ts: replyThreadTs,
      });
    }

    if (shouldHandleAsProposalPipeline(userText, { channel: event.channel, threadTs: replyThreadTs })) {
      await postSlackWorkStartedIfNeeded();
      await handleProposalPipeline(userText, {
        slackClient: slack,
        channel: event.channel,
        threadTs: replyThreadTs,
        anthropicClient: client,
      });
      return;
    }

    const healthReply = getBasicHealthReply(userText);
    if (healthReply) {
      await postSlackMessage({
        channel: event.channel,
        text: healthReply,
        thread_ts: replyThreadTs,
      });
      return;
    }

    let threadPriorBlock = '';
    if (event.channel && replyThreadTs && event.ts && !event._local_post_to_channel) {
      threadPriorBlock = await fetchSlackThreadPriorTranscript({
        channel: event.channel,
        threadRootTs: replyThreadTs,
        currentMessageTs: event.ts,
      });
    }
    const combinedRouting = threadPriorBlock ? `${threadPriorBlock}${userText}` : userText;

    const threadTsEarly = replyThreadTs;
    const publicBidsHint = startSlackSlowWorkHintTimer({
      channel: event.channel,
      thread_ts: threadTsEarly,
      text: '_Fetching Public Bid Tracker — large pages can take a bit…_',
    });
    let publicBidsReply;
    try {
      const trimmedPbt = userText.trim();
      const pbtTok = trimmedPbt.split(/\s+/);
      if (
        /^(publicbids|pbt)$/i.test(pbtTok[0] || '') &&
        String(pbtTok[1] || '').toLowerCase() === 'search' &&
        pbtTok.length > 2
      ) {
        await postSlackWorkStartedIfNeeded();
      }
      publicBidsReply = await require('./public-bid-tracker').handlePublicBidsSlackCommand(userText);
    } catch (pbtErr) {
      console.error('Public Bid Tracker command failed:', pbtErr?.message || pbtErr);
      publicBidsReply = `*Public Bid Tracker*\nSomething went wrong while fetching or parsing the page: ${String(
        pbtErr?.message || pbtErr
      ).slice(0, 800)}`;
    } finally {
      publicBidsHint.cancel();
    }
    if (publicBidsReply) {
      await postSlackMessage({
        channel: event.channel,
        text: publicBidsReply,
        thread_ts: threadTsEarly,
      });
      return;
    }

    const sheetSel = resolveSheetSelection(userText, event.channel);
    const defaultSid = String(process.env.SMARTSHEET_SHEET_ID || '').trim();
    let sheetId = String(sheetSel.sheetId || defaultSid || '');
    const knownSheetNames = [...getSmartsheetNameToIdMap().keys()].join(', ');

    if (sheetSel.unknownName) {
      await postSlackMessage({
        channel: event.channel,
        text: [
          `I don't recognize the sheet name "${sheetSel.label}".`,
          knownSheetNames
            ? `Configured names: ${knownSheetNames}.`
            : 'No names are configured yet — add SMARTSHEET_SHEET_MAP to your .env (JSON: friendly name → Smartsheet sheet id).',
          'Example: SMARTSHEET_SHEET_MAP={"Job Sim":"1234567890123456"}',
          'I did not run that change on your default sheet, so nothing was updated.',
        ].join(' '),
        thread_ts: replyThreadTs,
      });
      return;
    } else if (sheetSel.cleared) {
      await postSlackMessage({
        channel: event.channel,
        text: 'OK — this channel is back on the default Smartsheet (SMARTSHEET_SHEET_ID).',
        thread_ts: replyThreadTs,
      });
    } else if (
      sheetSel.switched &&
      sheetSel.label &&
      !sheetSel.cleared &&
      !sheetSel.unknownName &&
      !hasInlineSheetIntro(userText) &&
      !hasSheetOperationIntent(userText)
    ) {
      await postSlackMessage({
        channel: event.channel,
        text: `OK — this channel will use the "${sheetSel.label}" Smartsheet until you switch again.`,
        thread_ts: replyThreadTs,
      });
    }

    const multiMapNoDefault =
      getSmartsheetNameToIdMap().size > 1 && !defaultSid && !slackChannelActiveSheetId.has(event.channel);
    if (multiMapNoDefault && !sheetId && userText.trim().toLowerCase().startsWith('smartsheet')) {
      await postSlackMessage({
        channel: event.channel,
        text: `Set SMARTSHEET_SHEET_ID to a default sheet, or say which sheet first (e.g. "Switch to Job Sim sheet"). Configured names: ${knownSheetNames}.`,
        thread_ts: replyThreadTs,
      });
      return;
    }

    if (!sheetId && getSmartsheetNameToIdMap().size === 1) {
      sheetId = Array.from(getSmartsheetNameToIdMap().values())[0];
    }

    const sheetThreadTs = replyThreadTs;
    const sheetSlowHint = startSlackSlowWorkHintTimer({
      channel: event.channel,
      thread_ts: sheetThreadTs,
    });
    try {
      await smartsheetSheetContext.run({ sheetId }, async () => {
      const bcReply = await handleBuildingConnectedCommand(userText);
      if (bcReply) {
        await postSlackMessage({
          channel: event.channel,
          text: bcReply,
          thread_ts: replyThreadTs,
        });
        return;
      }
      const smTrim = userText.trim();
      if (/^smartsheet\s+(clear|format|add|update|picklist|bold|unbold)\b/i.test(smTrim)) {
        await postSlackWorkStartedIfNeeded();
      }
      const smartsheetReply = await handleSmartsheetCommand(userText);
      if (smartsheetReply) {
        await postSlackMessage({
          channel: event.channel,
          text: smartsheetReply,
          thread_ts: replyThreadTs,
        });
        return;
      }
      if (smartsheetConfigured() && !skipSmartsheetNaturalLanguageForMessage(combinedRouting)) {
        await postSlackWorkStartedIfNeeded();
        let sheetReply;
        if (String(process.env.SMARTSHEET_USE_OFFICIAL_MCP || '').trim() === '1') {
          const mcpAgent = await import('./smartsheet-mcp-agent.mjs');
          sheetReply = await mcpAgent.handleSmartsheetPlainEnglishViaOfficialMcp(userText, {
            threadPriorBlock,
            sheetId,
            anthropicClient: client,
          });
        } else {
          sheetReply = await handlePlainEnglishSmartsheet(userText);
        }
        if (sheetReply) {
          await postSlackMessage({
            channel: event.channel,
            text: sheetReply,
            thread_ts: replyThreadTs,
          });
          return;
        }
        if (looksLikeSheetTopic(userText)) {
          const onlySwitchContext =
            sheetSel.switched &&
            !hasInlineSheetIntro(userText) &&
            !hasSheetOperationIntent(userText);
          if (!onlySwitchContext) {
            await postSlackMessage({
              channel: event.channel,
              text:
                'I could not safely map that to a Smartsheet change yet. Try an explicit command like `smartsheet help` or `smartsheet format Column2 1 5 #00FF00`.',
              thread_ts: replyThreadTs,
            });
            return;
          }
        }
      }
      await postSlackWorkStartedIfNeeded();
      const bcDiscoverReply = await handleBuildingConnectedDiscoveryAnswer(userText, { threadPriorBlock });
      if (bcDiscoverReply) {
        await postSlackMessage({
          channel: event.channel,
          text: bcDiscoverReply,
          thread_ts: replyThreadTs,
        });
        return;
      }
      let safeReply;
      try {
        safeReply = await generateAssistantReply(userText, { threadPriorBlock });
      } catch (genErr) {
        console.error('generateAssistantReply failed:', genErr?.message || genErr);
        safeReply = `I could not get a model reply just then (${String(genErr?.message || genErr).slice(0, 400)}). Please try again.`;
      }
      let out = String(safeReply || '').trim() || 'I did not get any reply text back. Please try again.';
      out = replaceFalseBcConnectDemand(combinedRouting, out);
      await postSlackMessage({
        channel: event.channel,
        text: out.slice(0, 39000),
        thread_ts: replyThreadTs,
      });
    });
    } finally {
      sheetSlowHint.cancel();
    }
  } catch (err) {
    console.error('Failed to handle Slack message:', err?.message || err, {
      channel: event?.channel,
      textPreview: String(event?.text || '').slice(0, 200),
    });
    try {
      await postSlackMessage({
        channel: event?.channel,
        text: `I hit an internal error while handling that message (${String(err?.message || err).slice(0, 500)}). Please try again in a moment.`,
        thread_ts: slackReplyThreadTs(event),
      });
    } catch (postErr) {
      console.error('Failed to send Slack fallback message:', postErr?.message || postErr);
    }
  }
});

app.listen(3000, () => {
  loadAutodeskTokenCacheFromDisk();
  const missing = REQUIRED_ENV_VARS.filter((name) => !process.env[name]);
  if (missing.length > 0) {
    console.error(`Missing required env vars: ${missing.join(', ')}`);
  } else {
    console.log('All required env vars are loaded');
  }
  if (!process.env.OPENAI_API_KEY) {
    console.warn('OPENAI_API_KEY is not set. Audio transcription is disabled.');
  }
  if (!buildingConnectedConfigured()) {
    console.warn(
      'BuildingConnected integration is disabled. Set AUTODESK_CLIENT_ID and AUTODESK_CLIENT_SECRET.'
    );
  } else if (!process.env.AUTODESK_REDIRECT_URI && !process.env.PUBLIC_BASE_URL) {
    console.warn(
      'Set AUTODESK_REDIRECT_URI (or PUBLIC_BASE_URL) so Autodesk can redirect to /autodesk/callback during bc connect.'
    );
  }
  if (String(process.env.BC_WEB_FALLBACK_ENABLED || '').trim() === '1') {
    console.warn(
      'BC_WEB_FALLBACK_ENABLED is set in .env but BuildingConnected web/UI fallback is disabled in this build (API-only); you can remove that variable.'
    );
  }
  const smartsheetReady =
    Boolean(process.env.SMARTSHEET_API_TOKEN) &&
    (Boolean(process.env.SMARTSHEET_SHEET_ID) || getSmartsheetNameToIdMap().size > 0);
  if (!smartsheetReady) {
    console.warn(
      'Smartsheet integration is disabled. Set SMARTSHEET_API_TOKEN and SMARTSHEET_SHEET_ID (or SMARTSHEET_SHEET_MAP with a default SMARTSHEET_SHEET_ID).'
    );
  }
  if (String(process.env.SMARTSHEET_USE_OFFICIAL_MCP || '').trim() === '1') {
    console.log(
      'SMARTSHEET_USE_OFFICIAL_MCP=1 — plain-English Smartsheet edits use hosted Smartsheet MCP + Claude tool loop (README-MCP.md).'
    );
  }
  console.log('Bot is running on port 3000');
  console.log(
    `Bot build: ${AGENT_CODE_BUILD} | executeResetBackgroundColor=${typeof executeResetBackgroundColor}`
  );
  console.log(
    'Slack inbound (same as before): Slack must POST to this process. For local Mac use `npm run startup` (scripts/startup.sh: starts ./ngrok if needed, prints https://…/slack/events) or `npm run start:all`. Plain `npm run start` = replies work only for `npm run slack:ping-local`, not for messages typed in Slack.'
  );
  slack.auth
    .test()
    .then((r) => {
      if (r.user_id) slackBotUserId = String(r.user_id).trim();
      console.log('[slack] auth.test ok bot_user_id=', slackBotUserId, 'team=', r.team);
    })
    .catch((e) => console.warn('[slack] auth.test failed (self-message filter may be incomplete):', e?.message || e));
  (async () => {
    const urls = ['https://slack.com/api/api.test', 'https://www.slack.com'];
    for (const url of urls) {
      try {
        const response = await fetch(url, { signal: fetchTimeoutSignal(8000) });
        console.log(`Outbound Slack reachability (${url}):`, response.status);
        return;
      } catch (err) {
        console.error(`Outbound Slack reachability FAILED (${url}):`, err?.message || err);
      }
    }
  })();
});

// Keep a lightweight heartbeat so this process stays alive in local terminal sessions.
setInterval(() => {
  // Intentionally minimal; useful only to keep the event loop active.
}, 60_000);