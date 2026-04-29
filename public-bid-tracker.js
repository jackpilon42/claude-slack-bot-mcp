'use strict';

const DEFAULT_PBT_URL = 'https://publicbidtracker.com/california/open-bids/';
const DEFAULT_PORTAL = 'https://caleprocure.ca.gov/';

function getPublicBidTrackerUrl() {
  const u = String(process.env.PUBLIC_BID_TRACKER_URL || '').trim();
  return u || DEFAULT_PBT_URL;
}

function htmlToPlain(html) {
  let t = String(html || '')
    .replace(/<script[\s\S]*?<\/script>/gi, '\n')
    .replace(/<style[\s\S]*?<\/style>/gi, '\n')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/(tr|div|p|h[1-6]|li)\s*>/gi, '\n')
    .replace(/<[^>]+>/g, ' ');
  t = t.replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&');
  t = t.replace(/&#(\d+);/g, (_, n) => String.fromCharCode(Number(n)));
  t = t.replace(/[ \t\r\f\v]+/g, ' ').replace(/\n\s*\n/g, '\n');
  return t.trim();
}

function tryParseNextDataBids(html) {
  const m = html.match(/<script id="__NEXT_DATA__" type="application\/json">([\s\S]*?)<\/script>/);
  if (!m) return null;
  let data;
  try {
    data = JSON.parse(m[1]);
  } catch {
    return null;
  }
  const found = findBidLikeArray(data, 0);
  if (!found?.length) return null;
  return found.map((row) => ({
    id: String(row.bidNumber || row.id || row.solicitationId || row.number || '').trim() || 'unknown',
    agency: String(row.agency || row.organization || row.department || row.issuingAgency || '').trim(),
    description: String(row.description || row.title || row.fullDescription || row.name || '').trim(),
    portalUrl: String(row.portalUrl || row.sourceUrl || row.url || DEFAULT_PORTAL).trim() || DEFAULT_PORTAL,
  }));
}

function findBidLikeArray(obj, depth) {
  if (depth > 18 || obj == null) return null;
  if (Array.isArray(obj) && obj.length > 2 && typeof obj[0] === 'object') {
    const keys = Object.keys(obj[0]).join(' ').toLowerCase();
    if (
      (keys.includes('description') || keys.includes('title')) &&
      (keys.includes('bid') || keys.includes('solicit') || keys.includes('organization'))
    ) {
      return obj;
    }
  }
  if (typeof obj !== 'object') return null;
  for (const v of Object.values(obj)) {
    const hit = findBidLikeArray(v, depth + 1);
    if (hit?.length) return hit;
  }
  return null;
}

function parseBidsFromPlainText(text) {
  const chunks = text.split(/(?=Bid\s*\/\s*Solicitation\s*#)/i);
  const out = [];
  for (const chunk of chunks) {
    if (!/Bid\s*\/\s*Solicitation\s*#/i.test(chunk)) continue;
    const idM = chunk.match(/Bid\s*\/\s*Solicitation\s*#\s*([A-Za-z0-9\-#.]+)/i);
    const agM = chunk.match(/Issuing\s*Agency\s*(.+?)\s+Full\s*Description/is);
    const fdM = chunk.match(/Full\s*Description\s*(.+?)\s+Procurement\s*Method/is);
    const urlM = chunk.match(/https:\/\/caleprocure\.ca\.gov[^)\s"'<>]*/) || chunk.match(/https:\/\/[^\s)"'<>]+/);
    if (idM && fdM) {
      out.push({
        id: idM[1].trim(),
        agency: agM ? agM[1].replace(/\s+/g, ' ').trim() : '',
        description: fdM[1].replace(/\s+/g, ' ').trim(),
        portalUrl: urlM ? urlM[0].replace(/[,;.]$/, '') : DEFAULT_PORTAL,
      });
    }
  }
  return out;
}

function parseMarkdownPipeRows(text) {
  const out = [];
  const rowRe = /\|\s*([^|]+?)\s*\|\s*([^|]+?)\s*\|\s*([^|]+?)\s*\|\s*Posted\s*\|/g;
  let pr;
  while ((pr = rowRe.exec(text)) !== null) {
    const id = pr[1].trim();
    if (/^bid #$/i.test(id) || /^organization$/i.test(pr[2].trim())) continue;
    if (/^description$/i.test(pr[3].trim())) continue;
    out.push({
      id,
      agency: pr[2].trim(),
      description: pr[3].trim(),
      portalUrl: DEFAULT_PORTAL,
    });
  }
  return out;
}

function dedupeBids(rows) {
  const seen = new Set();
  const out = [];
  for (const r of rows) {
    const k = `${r.id}|${r.description}`.slice(0, 400);
    if (seen.has(k)) continue;
    seen.add(k);
    out.push(r);
  }
  return out;
}

async function fetchCaliforniaBidRows() {
  const url = getPublicBidTrackerUrl();
  const res = await fetch(url, {
    headers: {
      'user-agent': 'TransformSimSlackBot/1.0 (internal; contact: procurement)',
      accept: 'text/html,application/xhtml+xml',
    },
    signal: AbortSignal.timeout(Number(process.env.PUBLIC_BID_TRACKER_TIMEOUT_MS || 60_000)),
  });
  if (!res.ok) throw new Error(`HTTP ${res.status} fetching Public Bid Tracker`);
  const html = await res.text();
  let rows = tryParseNextDataBids(html);
  const plain = htmlToPlain(html);
  if (!rows?.length) {
    const a = parseBidsFromPlainText(plain);
    const b = parseMarkdownPipeRows(plain);
    rows = dedupeBids([...a, ...b]);
  }
  if (!rows?.length) {
    throw new Error(
      'Could not parse any bids from the page (layout may have changed). Try updating public-bid-tracker.js parsers.'
    );
  }
  return dedupeBids(rows);
}

function filterByKeywords(rows, needlesLower) {
  if (!needlesLower.length) return rows;
  return rows.filter((r) => {
    const blob = `${r.id} ${r.agency} ${r.description}`.toLowerCase();
    return needlesLower.every((n) => blob.includes(n));
  });
}

function formatSlackReply(rows, needlesLower, sourceUrl) {
  const max = Math.min(Number(process.env.PUBLIC_BIDS_MAX_RESULTS || 20) || 20, 40);
  let msg =
    `*Public Bid Tracker (California)*\n` +
    `Source: ${sourceUrl}\n` +
    `Keywords (all must match title/agency/description): ${needlesLower.join(', ')}\n` +
    `_Mirrors public listings; confirm details and documents on Cal eProcure / the issuing portal._\n\n`;
  if (!rows.length) {
    msg += '_No rows matched your keywords._ Try broader terms or fewer keywords (`publicbids search roofing`).';
    return msg;
  }
  msg += `Showing **${Math.min(rows.length, max)}** of **${rows.length}** match(es):\n\n`;
  const slice = rows.slice(0, max);
  for (const r of slice) {
    msg += `• **${r.id}** — ${r.description}\n  _${r.agency}_\n  <${r.portalUrl}|Cal eProcure / portal>\n\n`;
  }
  if (rows.length > max) msg += `_…and ${rows.length - max} more._\n`;
  return msg.trimEnd();
}

async function runPublicBidSearch(needlesLower) {
  const sourceUrl = getPublicBidTrackerUrl();
  const all = await fetchCaliforniaBidRows();
  const filtered = needlesLower.length ? filterByKeywords(all, needlesLower) : all.slice(0, 30);
  return formatSlackReply(filtered, needlesLower, sourceUrl);
}

async function handlePublicBidsSlackCommand(userText) {
  const trimmed = String(userText || '').trim();
  if (!/^publicbids\b|^pbt\b/i.test(trimmed)) return null;
  const parts = trimmed.split(/\s+/);
  const cmd = (parts[0] || '').toLowerCase();
  if (cmd !== 'publicbids' && cmd !== 'pbt') return null;
  const action = (parts[1] || '').toLowerCase();
  if (!action || action === 'help') {
    return (
      '*Public Bid Tracker*\n' +
      '- `publicbids search roofing modesto` — keywords (all must match)\n' +
      '- `pbt search stanislaus roof` — same (alias)\n' +
      '- Optional `.env`: `PUBLIC_BID_TRACKER_URL` (default: California open bids page)\n' +
      '_Data is scraped from the public HTML page; layout changes can break parsing._'
    );
  }
  if (action !== 'search') {
    return 'Unknown subcommand. Use `publicbids help`.';
  }
  const tokens = parts.slice(2).map((p) => p.toLowerCase()).filter(Boolean);
  if (!tokens.length) {
    return 'Usage: `publicbids search roofing modesto` — provide one or more keywords.';
  }
  try {
    return await runPublicBidSearch(tokens);
  } catch (e) {
    return `*Public Bid Tracker*\nCould not fetch or parse the page: ${String(e?.message || e).slice(0, 900)}`;
  }
}

module.exports = {
  getPublicBidTrackerUrl,
  fetchCaliforniaBidRows,
  filterByKeywords,
  handlePublicBidsSlackCommand,
  htmlToPlain,
  parseBidsFromPlainText,
};
