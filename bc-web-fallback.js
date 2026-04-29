'use strict';

const fs = require('fs');
const path = require('path');

/** When true, Playwright / UI scraping is never run; env cannot re-enable (product policy). */
const BC_WEB_FALLBACK_DISABLED = true;

function getBcWebStorageStatePath() {
  const e = String(process.env.BC_WEB_STORAGE_STATE_PATH || '').trim();
  if (e) return path.isAbsolute(e) ? e : path.join(__dirname, e);
  return path.join(__dirname, 'playwright', '.bc-storage-state.json');
}

function bcWebFallbackEnabledFlag() {
  if (BC_WEB_FALLBACK_DISABLED) return false;
  return String(process.env.BC_WEB_FALLBACK_ENABLED || '').trim() === '1';
}

function bcWebFallbackReady() {
  return bcWebFallbackEnabledFlag() && fs.existsSync(getBcWebStorageStatePath());
}

function parseFallbackUrls() {
  const raw = String(process.env.BC_WEB_FALLBACK_URLS || '').trim();
  if (raw) {
    return raw
      .split(/[\s,]+/)
      .map((u) => u.trim())
      .filter(Boolean);
  }
  return [
    'https://app.buildingconnected.com/bid-board',
    'https://app.buildingconnected.com/opportunities',
    'https://app.buildingconnected.com/',
  ];
}

function formatWebStatus() {
  if (BC_WEB_FALLBACK_DISABLED) {
    return (
      '*BuildingConnected web fallback*\n' +
      '**Disabled** in this bot build (API-only). `BC_WEB_FALLBACK_ENABLED` and related env vars have no effect. ' +
      'You can remove them from `.env` to avoid confusion.'
    );
  }
  const enabled = bcWebFallbackEnabledFlag();
  const p = getBcWebStorageStatePath();
  const exists = fs.existsSync(p);
  let play = 'not checked';
  try {
    require.resolve('playwright');
    play = 'playwright package present';
  } catch {
    play = 'playwright missing — run `npm install` then `npx playwright install chromium`';
  }
  return (
    `*BuildingConnected web fallback*\n` +
    `- BC_WEB_FALLBACK_ENABLED: ${enabled ? '1 (on)' : 'unset/off'}\n` +
    `- Session file: \`${p}\` — ${exists ? 'found' : 'missing (run save-session script)'}\n` +
    `- ${play}\n` +
    `- Optional: \`BC_WEB_START_URL=\` first page to open; \`BC_WEB_FALLBACK_URLS=\` comma URLs to scan\n` +
    `_Web mode matches visible text lines; it does not compute road miles._`
  );
}

async function pageLooksLikeLogin(page) {
  const url = (page.url() || '').toLowerCase();
  if (url.includes('signin') || url.includes('/login') || url.includes('/authorize')) return true;
  let title = '';
  try {
    title = await page.title();
  } catch {
    /* ignore */
  }
  if (/sign in|log in/i.test(title)) return true;
  let t = '';
  try {
    t = await page.evaluate(() => (document.body && document.body.innerText) || '');
  } catch {
    /* ignore */
  }
  const low = t.slice(0, 6000).toLowerCase();
  if (low.includes('sign in') && (low.includes('autodesk') || low.includes('buildingconnected'))) return true;
  return false;
}

/**
 * @param {string[]} needlesLower keywords, all must appear in a line (substring match)
 * @returns {Promise<{ ok: boolean, text?: string, error?: string }>}
 */
async function runBcWebKeywordSearch(needlesLower) {
  if (BC_WEB_FALLBACK_DISABLED) {
    return { ok: false, error: 'Web/UI fallback is disabled in this bot (BuildingConnected API only).' };
  }
  if (!needlesLower.length) {
    return { ok: false, error: 'No keywords to search.' };
  }
  let playwright;
  try {
    playwright = require('playwright');
  } catch {
    return {
      ok: false,
      error: 'Playwright not installed. On the bot host: `npm install` then `npx playwright install chromium`.',
    };
  }
  const storage = getBcWebStorageStatePath();
  if (!fs.existsSync(storage)) {
    return {
      ok: false,
      error: `No saved browser session. Run: \`node scripts/bc-web-save-session.js\` (saves \`${storage}\`).`,
    };
  }
  const headless = String(process.env.BC_WEB_HEADFUL || '').trim() !== '1';
  const browser = await playwright.chromium.launch({ headless });
  try {
    const context = await browser.newContext({ storageState: storage });
    const page = await context.newPage();
    const start = String(process.env.BC_WEB_START_URL || '').trim();
    const urls = [...new Set([...(start ? [start] : []), ...parseFallbackUrls()])];
    const chunks = [];
    for (const u of urls) {
      try {
        await page.goto(u, { waitUntil: 'domcontentloaded', timeout: 60_000 });
        await page.waitForLoadState('networkidle', { timeout: 20_000 }).catch(() => {});
        if (await pageLooksLikeLogin(page)) {
          return {
            ok: false,
            error:
              'Browser landed on a sign-in page — saved session expired or is invalid. Re-run `node scripts/bc-web-save-session.js`.',
          };
        }
        const body = await page.evaluate(() => (document.body && document.body.innerText) || '');
        if (body && body.length > 80) chunks.push(body);
      } catch (e) {
        chunks.push(`[${u}] ${String(e?.message || e)}`);
      }
    }
    const blob = chunks.join('\n\n---\n\n');
    const lines = blob
      .split('\n')
      .map((l) => l.trim())
      .filter((l) => l.length > 4 && l.length < 600);
    const matched = [];
    for (const line of lines) {
      const low = line.toLowerCase();
      if (needlesLower.every((n) => low.includes(n))) matched.push(line);
    }
    const uniq = [...new Set(matched)];
    const maxShow = Math.min(Number(process.env.BC_WEB_MAX_LINES || 18) || 18, 40);
    let text =
      `*BuildingConnected web fallback (Playwright)*\n` +
      `Keywords (all must appear in the same line of visible text): ${needlesLower.join(', ')}\n` +
      `_Not GIS distance — use BC filters in the app for exact mileage._\n\n`;
    if (!uniq.length) {
      text +=
        '_No matching lines found._ Try `BC_WEB_START_URL` pointing at the list page you use (Bid Board / opportunities), then re-save the session after navigating there once.';
      return { ok: true, text };
    }
    text += `Found **${uniq.length}** line(s). Showing up to ${maxShow}:\n\n`;
    text += uniq
      .slice(0, maxShow)
      .map((l) => `• ${l}`)
      .join('\n');
    if (uniq.length > maxShow) text += `\n…_and ${uniq.length - maxShow} more._`;
    return { ok: true, text };
  } finally {
    await browser.close().catch(() => {});
  }
}

module.exports = {
  getBcWebStorageStatePath,
  bcWebFallbackEnabledFlag,
  bcWebFallbackReady,
  formatWebStatus,
  runBcWebKeywordSearch,
};
