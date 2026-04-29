# BuildingConnected integration — handoff for Cursor

This file summarizes what Jack and Cursor worked through for **Autodesk BuildingConnected** on the **`claude-slack-bot`** Slack bot. Paste or `@`-attach this in a new Cursor chat so the agent can continue from the same assumptions.

**Repo path (Jack’s machine):** `/Users/jackpilon/claude-slack-bot`  
**Primary Cursor agent transcript (this machine):**  
`~/.cursor/projects/1776101275580/agent-transcripts/e04e658e-65b0-4acb-9f05-09428dd62930/e04e658e-65b0-4acb-9f05-09428dd62930.jsonl`  
*(Your friend will not have this path unless they use the same Cursor project id; this handoff file is the portable substitute.)*  
**Daily hook archive (prompt/response summaries):** `~/.cursor/conversation-archive/` — relevant days include `2026-04-20.jsonl` and `2026-04-21.jsonl` (timestamps may cross UTC midnight).

---

## 1. Product context

- **Bot:** Node/Express Slack app (`index.js`) with Anthropic, optional OpenAI, **Smartsheet**, and **BuildingConnected**.
- **Goal:** Use the **official BuildingConnected v2 API** first; if `GET /projects` (or discovery that depends on it) fails, optionally use a **Playwright “web fallback”** that scans visible text on BC web pages using a **saved browser session** (not “Claude clicking” — the **server process** runs Chromium).

---

## 2. Timeline (what happened)

### April 16–17, 2026 (broader bot work)

From the conversation archive summary: Smartsheet plain-English behavior, multi-sheet switching, ngrok/Slack wiring, and ops debugging. That is the same codebase but **not** the BC-specific thread.

### April 20, 2026 — BuildingConnected API (primary path)

- Integrated **APS OAuth** for BuildingConnected: **3-legged** authorization code flow (not 2-legged client-credentials for project listing).
- **Endpoints in use:**
  - Base: `https://developer.api.autodesk.com/construction/buildingconnected/v2`
  - OAuth token: `https://developer.api.autodesk.com/authentication/v2/token`
  - Authorize: `https://developer.api.autodesk.com/authentication/v2/authorize`
- **Working after OAuth:** `GET /users/me` returns **200** with a normal user payload.
- **Blocked:** `GET /projects` returns **403** with body containing **`BC_PRO_SUBSCRIPTION_REQUIRED`** (before any client-side filtering). This is treated as an **Autodesk entitlement / subscription / office-context** signal, **not** a “wrong OAuth grant type” issue once `/users/me` works on the same token.
- **Support email drafted** (full ticket + shorter form version) asking Autodesk to confirm what controls access to `GET /projects` vs the Pro subscription the org believes it has.

### April 21–22, 2026 — Autodesk reply and research

- **Salesforce / support case 25814570** reply (`.msg` shared in chat): **routing only** — ticket was on the “wrong” support system; **ADN** called out for direct API support; community paths (Stack Overflow, forums, APS docs) referenced. **No** resolution of `BC_PRO_SUBSCRIPTION_REQUIRED` in that email.
- **Community research:** Strongest public match is Stack Overflow *“Getting Forbidden Error connecting to Building Connected Pro APIs”* — same `detail` string on a BC v2 route with 3-legged tokens; **no solved answer**; Autodesk staff comment suggests checking **which endpoints pass/fail** and whether the **OAuth user’s office** actually has BC Pro, with private follow-up to **`aps.help@autodesk.com`**.
- **Clarified:** **ADN (paid developer network)** is mainly about **support channel**, not the same thing as **BuildingConnected Pro** licensing. Having Pro in the product sense does not automatically disprove a 403 from the API’s entitlement check.

### April 22, 2026 — Web fallback (backup path)

- User asked to **prioritize backup** while API entitlement is sorted: “roofing within 50 miles of Modesto” style questions.
- **Important limitation:** Neither the API path nor the web fallback performs **GIS / road mileage**. Discovery uses **keyword extraction** (trade terms, city names, etc.) and requires **all** keywords to match either **project JSON** (API) or a **single line of visible text** (web). “50 miles” is **not** computed; messaging in the bot is explicit about that.

---

## 3. What was implemented (code map)

| Piece | Role |
|--------|------|
| `index.js` | APS OAuth routes (`/autodesk/connect`, `/autodesk/callback`), token load/save (`.autodesk-tokens.json`), `buildingConnectedRequest`, Slack **`bc`** commands, natural-language **discovery** for BC-style questions, **Smartsheet skip** when message looks like BC/geo discovery, **post-process** to stop the model from wrongly demanding `bc connect` when tokens exist, **API-first search** with optional web fallback append on failure. |
| `bc-web-fallback.js` | If `BC_WEB_FALLBACK_ENABLED=1` and storage state exists: Playwright Chromium loads saved session, visits configurable URLs, scans `innerText` lines for **all** keywords. `bc web status`, `bc web search`. |
| `scripts/bc-web-save-session.js` | Interactive one-time login; saves `playwright/.bc-storage-state.json` (or `BC_WEB_STORAGE_STATE_PATH`). |
| `.env.example` | Documents `AUTODESK_*`, optional `PUBLIC_BASE_URL` / `AUTODESK_REDIRECT_URI`, and `BC_WEB_*` vars. |
| `.gitignore` | Ignores token file and Playwright storage state. |

**Slack commands (BuildingConnected):**

- `bc help` — list commands  
- `bc connect` — start OAuth (see ngrok / `PUBLIC_BASE_URL` notes in help text when URL is localhost)  
- `bc auth status`  
- `bc projects [limit]`  
- `bc search <keywords…>` — all keywords must appear in project JSON (API path)  
- `bc project <id>` / `bc whoami`  
- `bc web status` / `bc web search <keywords…>` — web fallback only  

Natural-language messages that mention BuildingConnected or look like **geo + bid discovery** can trigger **server-side** `handleBuildingConnectedDiscoveryAnswer` (API first, then web block if configured and API failed).

---

## 4. Environment variables (no secrets here)

Copy from **`.env.example`**. Required for BC API:

- `AUTODESK_CLIENT_ID`
- `AUTODESK_CLIENT_SECRET`
- Optional: `AUTODESK_REDIRECT_URI` or use **`PUBLIC_BASE_URL`** so redirect is `https://…/autodesk/callback` (must match Autodesk app callback URLs exactly).
- `AUTODESK_SCOPES` — default in code/docs is `data:read` unless you change it.

Tokens persist to **`.autodesk-tokens.json`** (gitignored), path overridable with `AUTODESK_TOKEN_STORE_PATH`.

**Web fallback:**

- `BC_WEB_FALLBACK_ENABLED=1`
- Optional: `BC_WEB_STORAGE_STATE_PATH`, `BC_WEB_START_URL`, `BC_WEB_FALLBACK_URLS`, `BC_WEB_HEADFUL=1` for debugging  
- Run `npm install` and **`npx playwright install chromium`** on the host.

---

## 5. OAuth / ops notes that burned time

- **`bc connect` in Slack** prints a URL. If it shows **`localhost`**, only the machine running the bot can open it. For Slack on a phone or another PC, use **ngrok** (or similar), set **`PUBLIC_BASE_URL`** (and matching **Autodesk callback URL**), restart, then reconnect.
- **`GET /projects` vs 2-legged:** Project listing is **user (3-legged)** territory; 2-legged is the wrong tool for `/projects` per public Q&A — but that is **separate** from `BC_PRO_SUBSCRIPTION_REQUIRED` when 3-legged `/users/me` already works.

---

## 6. Suggested next steps for your friend’s Cursor

1. Open the repo, read **`index.js`** (search `BuildingConnected`, `bc `, `autodesk`), **`bc-web-fallback.js`**, **`scripts/bc-web-save-session.js`**, **`.env.example`**.
2. Reproduce: `bc connect` → `bc whoami` → `bc projects`. If **403** + `BC_PRO_SUBSCRIPTION_REQUIRED`, confirm **which Autodesk identity** completed OAuth and **BC admin/office** assignment; consider **`aps.help@autodesk.com`** with endpoint matrix (from SO thread pattern).
3. If API stays blocked and org **accepts** UI automation risk: enable **`BC_WEB_FALLBACK_ENABLED`**, run save-session script on the **same host as the bot**, verify **`bc web status`**, test **`bc web search roofing modesto`** and natural-language discovery.

---

## 7. Compliance

Treat Playwright fallback as **RPA**: verify against **Autodesk / company policy** before relying on it in production.

---

*Generated as a portable summary from Cursor transcripts and `~/.cursor/conversation-archive/` entries around 2026-04-20–22. Adjust paths if the repo lives somewhere else on your friend’s machine.*
