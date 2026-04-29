# Transform Sim — Smartsheet MCP branch

This directory is a **copy** of `claude-slack-bot` with an **optional** integration path: employee-style plain English for Smartsheet can be handled by **Claude + Smartsheet’s hosted MCP** instead of the hand-maintained natural-language parsers in `index.js`.

Your original project at `../claude-slack-bot` is **not modified** by this work.

## When to use this copy

- You want the model to choose from **many** Smartsheet operations via official MCP tools.
- You are okay with **LLM + tool-call** behavior (variance, latency, token cost) vs. deterministic regex handlers.

## When to stay on the original project

- You need predictable, minimal behavior and no dependency on Smartsheet’s MCP service.
- You prefer the existing fast paths (`smartsheet …` commands, parsers).

## Prerequisites (Smartsheet)

Smartsheet documents MCP access requirements (plan, token, region). See:

- [Install the Smartsheet MCP server](https://developers.smartsheet.com/api/smartsheet/ai-integration/install-the-smartsheet-mcp-server)

Use the correct **region URL** if not US (`https://mcp.smartsheet.com`).

## Configuration

1. Copy `.env` from your working bot (or copy `.env.example` → `.env` and fill values).
2. Set:

```bash
SMARTSHEET_USE_OFFICIAL_MCP=1
# Optional if not US:
# SMARTSHEET_MCP_URL=https://mcp.smartsheet.eu
```

3. Keep `SMARTSHEET_API_TOKEN` and sheet routing (`SMARTSHEET_SHEET_ID` / `SMARTSHEET_SHEET_MAP`) as today — the MCP agent receives the **active sheet id** from the same Slack per-channel context as the original bot.

4. Install dependencies and run:

```bash
cd claude-slack-bot-mcp
npm install
npm run startup
```

5. `/health` includes `"smartsheetMcpAgent": true` when MCP mode is on.

## How it works

- If `SMARTSHEET_USE_OFFICIAL_MCP=1`, after `smartsheet …` commands, **plain-English Smartsheet** messages are sent to `smartsheet-mcp-agent.mjs`.
- That module connects to **`SMARTSHEET_MCP_URL`** with `Authorization: Bearer SMARTSHEET_API_TOKEN`, loads MCP tools, and runs a **Claude tool loop** until the model responds with text (or hits a tool round limit).
- Explicit `smartsheet …` slash-style commands still use the **existing REST** helpers in `index.js` (unchanged behavior).

## Reverting

- Point Slack Event Subscriptions / process manager back at **`claude-slack-bot`** and run `npm run startup` there, **or**
- In this repo, unset `SMARTSHEET_USE_OFFICIAL_MCP` (or set it to `0`) and restart — parsers behave like the original copy at fork time.

## Files only in this branch

- `README-MCP.md` — this document  
- `smartsheet-mcp-agent.mjs` — MCP client + Claude tool loop  
- `package.json` — adds `@modelcontextprotocol/sdk`, package name `claude-slack-bot-mcp`  
- Small edits in `index.js` (feature flag, build id, `/health` payload)

---

## Notes from prior Transform Sim / Smartsheet work (REST era → keep in mind for MCP)

These came out of long-running Slack + Smartsheet debugging before this branch; they are useful whether you use MCP or parsers.

### Sheet targeting in Slack

- The server resolves **`SMARTSHEET_SHEET_ID`**, per-channel memory, and **`SMARTSHEET_SHEET_MAP`** before calling the MCP agent. The agent receives the resulting **`sheetId`**.
- Users often say **`in My Sheet, …`** (comma, no word “sheet”). The bot matches that against **`SMARTSHEET_SHEET_MAP`** when the name is unambiguous—ensure friendly names in the map match how people speak.

### Smartsheet product semantics (avoid wrong expectations)

- **PICKLIST / Yes–No** is a **column** setting: every row shares the same allowed values. It is **not** Excel-style validation on one cell. If the user names a row, they may want that **cell value** set to `Yes` or `No` after the column type change—do that with tools when it makes sense.
- **Reads** (“what is in column 3 row 1?”) must come from **API/tool data**, not from guessing (e.g. confusing column *title* with cell *value*).

### Hybrid routing in this repo

- **`smartsheet …` commands** (explicit REST helpers in `index.js`) always run **without** MCP, even when `SMARTSHEET_USE_OFFICIAL_MCP=1`.
- **Plain English** Smartsheet text uses **MCP + tool loop** when the flag is on, otherwise **`handlePlainEnglishSmartsheet`** (regex + intent JSON + special cases like column-scoped clear / picklist helpers).

If a phrase was carefully handled in **`handlePlainEnglishSmartsheet`** (e.g. scoped “clear all cells in column N except row …”) and MCP tools cannot express it, **temporarily set `SMARTSHEET_USE_OFFICIAL_MCP=0`** for that workflow or extend the REST path—don’t assume MCP alone matches parser coverage.

### Optional env (MCP agent tuning)

| Variable | Purpose |
|----------|---------|
| `SMARTSHEET_MCP_URL` | Hosted MCP base URL (default US). |
| `SMARTSHEET_MCP_CLAUDE_MODEL` | Model id for the tool loop (default `claude-sonnet-4-20250514`). |
| `SMARTSHEET_MCP_MAX_TOOL_ROUNDS` | Cap on tool rounds per Slack message (default `12`, clamped 4–24). |
| `SMARTSHEET_MCP_ANTHROPIC_TIMEOUT_MS` | Per–`messages.create` timeout (default `120000`, clamped 30s–600s). |

### Slack delivery (not MCP-specific, but recurring)

- **`@bot ping`** uses **`app_mention`**. Plain messages in a channel need **`message.channels`** (and the bot in the channel) or Slack will never POST them.
- If users get **no reply**, check server logs for **`[slack/events]`** and verify the public **Request URL** matches the running process.

### Further reading (external)

- Smartsheet hosted MCP install / plan / region: [Install the Smartsheet MCP server](https://developers.smartsheet.com/api/smartsheet/ai-integration/install-the-smartsheet-mcp-server)  
- Column types (why dropdowns are column-wide): [Smartsheet column types](https://help.smartsheet.com/articles/2480016-column-types)  
- Community example of many Smartsheet tools + MCP surface (not maintained here): [Abdeltoto/smartsheet-controller](https://github.com/Abdeltoto/smartsheet-controller)
