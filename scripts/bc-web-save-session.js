#!/usr/bin/env node
/**
 * Legacy: login capture for BuildingConnected Playwright fallback.
 * The bot build disables web fallback (API-only); this script is only useful if you re-enable it in code.
 *
 * Requires: npm install playwright && npx playwright install chromium
 */
'use strict';

const fs = require('fs');
const path = require('path');
const readline = require('readline');

const outDir = path.join(__dirname, '..', 'playwright');
const outFile =
  String(process.env.BC_WEB_STORAGE_STATE_PATH || '').trim() ||
  path.join(outDir, '.bc-storage-state.json');
const target = path.isAbsolute(outFile) ? outFile : path.join(__dirname, '..', outFile);

async function main() {
  let playwright;
  try {
    playwright = require('playwright');
  } catch {
    console.error('Install Playwright first: npm install && npx playwright install chromium');
    process.exit(1);
  }
  fs.mkdirSync(path.dirname(target), { recursive: true });
  console.log('Launching browser. Log in to BuildingConnected (Autodesk) in the window.');
  console.log('When you see your normal logged-in home or Bid Board, return here and press Enter.\n');
  const browser = await playwright.chromium.launch({ headless: false });
  const context = await browser.newContext();
  const page = await context.newPage();
  const start = String(process.env.BC_WEB_SAVE_SESSION_URL || 'https://app.buildingconnected.com/').trim();
  await page.goto(start, { waitUntil: 'domcontentloaded', timeout: 120_000 });
  const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
  await new Promise((resolve) => rl.question('Press Enter after you are fully logged in… ', resolve));
  rl.close();
  await context.storageState({ path: target });
  await browser.close();
  console.log(`Saved session to: ${target}`);
  console.log('Note: current bot build has web fallback disabled in code; this file is unused until that changes.');
  if (!String(process.env.BC_WEB_STORAGE_STATE_PATH || '').trim()) {
    console.log('(Optional) BC_WEB_STORAGE_STATE_PATH if you use a non-default path.');
  }
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
