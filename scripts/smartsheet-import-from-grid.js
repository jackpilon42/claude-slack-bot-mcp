#!/usr/bin/env node
/**
 * Plain-English intent (what this script does):
 * 1. Create a new Smartsheet named like the template.
 * 2. Add one column per Excel column (first column = Primary Column).
 * 3. Add each Excel row as a Smartsheet row, copying text/numbers/dates into the matching columns.
 *
 * Requires: SMARTSHEET_API_TOKEN in .env (same as the Slack bot).
 * Optional: SMARTSHEET_WORKSPACE_ID — if set, sheet is created in that workspace; otherwise Home.
 *
 * Usage:
 *   node scripts/smartsheet-import-from-grid.js [path/to/grid.json]
 */
const fs = require('fs');
const path = require('path');
require('dotenv').config({ path: path.join(__dirname, '..', '.env') });

const BASE = 'https://api.smartsheet.com/2.0';

async function api(method, pth, body) {
  const token = String(process.env.SMARTSHEET_API_TOKEN || '').trim();
  if (!token) throw new Error('SMARTSHEET_API_TOKEN missing in .env');
  const res = await fetch(`${BASE}${pth}`, {
    method,
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json',
    },
    body: body != null ? JSON.stringify(body) : undefined,
  });
  const text = await res.text();
  let json;
  try {
    json = JSON.parse(text);
  } catch {
    json = { raw: text };
  }
  if (!res.ok) {
    const err = new Error(`Smartsheet ${res.status}: ${text.slice(0, 2000)}`);
    err.status = res.status;
    err.body = json;
    throw err;
  }
  return json;
}

function columnDefs(count) {
  const cols = [];
  for (let i = 0; i < count; i += 1) {
    if (i === 0) {
      cols.push({ title: 'Primary Column', primary: true, type: 'TEXT_NUMBER' });
    } else {
      cols.push({ title: `Field ${i + 1}`, type: 'TEXT_NUMBER' });
    }
  }
  return cols;
}

function rowPayload(columnIds, values) {
  const cells = [];
  for (let i = 0; i < columnIds.length; i += 1) {
    const v = values[i];
    if (v === '' || v == null) continue;
    const s = typeof v === 'number' ? String(v) : String(v);
    cells.push({ columnId: columnIds[i], value: s.length > 4000 ? s.slice(0, 4000) : s });
  }
  return { toBottom: true, cells };
}

async function main() {
  const jsonPath = path.resolve(process.argv[2] || path.join(__dirname, '_project_timeline_grid.json'));
  const raw = JSON.parse(fs.readFileSync(jsonPath, 'utf8'));
  const { name, rows: grid } = raw;
  if (!grid?.length) throw new Error('grid.rows missing');
  const width = grid[0].length;
  const ws = String(process.env.SMARTSHEET_WORKSPACE_ID || '').trim();

  const createPath = ws ? `/workspaces/${ws}/sheets` : '/sheets';
  const sheetBody = { name: name || 'Imported grid', columns: columnDefs(width) };

  const created = await api('POST', createPath, sheetBody);
  const sheet = created.result || created;
  const sheetId = sheet.id;
  const columns = sheet.columns || [];
  const columnIds = columns.map((c) => c.id);

  const BATCH = 25;
  for (let i = 0; i < grid.length; i += BATCH) {
    const chunk = grid.slice(i, i + BATCH).map((values) => rowPayload(columnIds, values));
    await api('POST', `/sheets/${sheetId}/rows`, chunk);
  }

  const url = `https://app.smartsheet.com/sheets/${sheetId}`;
  console.log(JSON.stringify({ sheetId, name: sheet.name, url, rowsImported: grid.length, columns: width }, null, 2));
}

main().catch((e) => {
  console.error(e.message || e);
  process.exit(1);
});
