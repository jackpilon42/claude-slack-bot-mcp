#!/usr/bin/env bash
set -euo pipefail

PROJECT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
ENV_FILE="${PROJECT_DIR}/.env"
NGROK_API="http://127.0.0.1:4040/api/tunnels"

cd "${PROJECT_DIR}"

PASS=0
FAIL=0
WARN=0

ok() {
  PASS=$((PASS + 1))
  echo "[OK]   $1"
}

bad() {
  FAIL=$((FAIL + 1))
  echo "[FAIL] $1"
}

warn() {
  WARN=$((WARN + 1))
  echo "[WARN] $1"
}

is_loopback_proxy() {
  local value="${1:-}"
  [[ "${value}" == http://127.0.0.1:* || "${value}" == https://127.0.0.1:* || "${value}" == socks5://127.0.0.1:* ]]
}

echo "Claude Slack Bot Status"
echo "Project: ${PROJECT_DIR}"
echo ""

if [[ -f "${ENV_FILE}" ]]; then
  ok ".env file found"
  set -a
  # shellcheck disable=SC1090
  source "${ENV_FILE}"
  set +a
else
  bad ".env file missing at ${ENV_FILE}"
  echo ""
  echo "Summary: ${PASS} OK, ${WARN} WARN, ${FAIL} FAIL"
  exit 1
fi

REQUIRED_VARS=(
  "SLACK_BOT_TOKEN"
  "SLACK_SIGNING_SECRET"
  "ANTHROPIC_API_KEY"
  "SMARTSHEET_API_TOKEN"
  "SMARTSHEET_SHEET_ID"
)

for name in "${REQUIRED_VARS[@]}"; do
  if [[ -n "${!name:-}" ]]; then
    ok "${name} is set"
  else
    bad "${name} is missing"
  fi
done

if [[ -n "${OPENAI_API_KEY:-}" ]]; then
  ok "OPENAI_API_KEY is set (audio enabled)"
else
  warn "OPENAI_API_KEY is not set (audio transcription disabled)"
fi

# NOTE: Do not use `$(curl ... || echo 000)` — on failure curl often prints 000 for %{http_code},
# then echo runs too, producing "000000" and a false OK.
BOT_HTTP_CODE="$(curl -sS --max-time 3 -o /dev/null -w "%{http_code}" "http://127.0.0.1:3000/health" 2>/dev/null)" || BOT_HTTP_CODE="000"
BOT_HTTP_CODE="$(printf '%s' "${BOT_HTTP_CODE}" | tr -d '\r\n[:space:]')"
if [[ ! "${BOT_HTTP_CODE}" =~ ^[0-9]{3}$ ]] || [[ "${BOT_HTTP_CODE}" == "000" ]]; then
  bad "Bot is not healthy on http://127.0.0.1:3000/health (got HTTP ${BOT_HTTP_CODE:-none}; start the bot: node index.js or ./scripts/startup.sh)"
else
  ok "Bot responds on localhost:3000/health (HTTP ${BOT_HTTP_CODE})"
fi

EVENT_HTTP_CODE="$(curl -sS --max-time 3 -o /dev/null -w "%{http_code}" -X POST http://127.0.0.1:3000/slack/events -H "Content-Type: application/json" -d '{"type":"url_verification","challenge":"status-check"}' 2>/dev/null)" || EVENT_HTTP_CODE="000"
EVENT_HTTP_CODE="$(printf '%s' "${EVENT_HTTP_CODE}" | tr -d '\r\n[:space:]')"
if [[ "${EVENT_HTTP_CODE}" == "200" ]]; then
  ok "/slack/events responds locally (HTTP 200, url_verification)"
elif [[ "${EVENT_HTTP_CODE}" == "403" ]]; then
  bad "/slack/events returned 403 (Slack signature rejected). url_verification should bypass signing — restart the bot with the latest index.js."
else
  bad "/slack/events local check failed (HTTP ${EVENT_HTTP_CODE:-none}; bot not listening or wrong process on :3000)"
fi

TUNNEL_JSON="$(curl -s "${NGROK_API}" 2>/dev/null || true)"
if [[ -z "${TUNNEL_JSON}" ]]; then
  bad "ngrok API not reachable at ${NGROK_API}"
else
  TUNNEL_URL="$(python3 - <<'PY' "${TUNNEL_JSON}"
import json,sys
raw = sys.argv[1]
try:
    data = json.loads(raw)
except Exception:
    print("")
    raise SystemExit
for t in data.get("tunnels", []):
    if t.get("proto") == "https":
        print(t.get("public_url", ""))
        raise SystemExit
print("")
PY
)"
  if [[ -n "${TUNNEL_URL}" ]]; then
    ok "ngrok HTTPS tunnel is live: ${TUNNEL_URL}"
    echo "      Slack Request URL should be: ${TUNNEL_URL}/slack/events"
  else
    bad "ngrok API reachable, but no HTTPS tunnel found"
  fi
fi

SLACK_REACH_TMP="$(mktemp)"
SLACK_REACH_CODE="$(curl -sS -o /dev/null -w "%{http_code}" https://slack.com/api/api.test 2>"${SLACK_REACH_TMP}" || true)"
if [[ "${SLACK_REACH_CODE}" == "200" ]]; then
  rm -f "${SLACK_REACH_TMP}"
  ok "Outbound reachability to slack.com is healthy"
else
  CURL_ERR_MSG="$(tr -d '\n' < "${SLACK_REACH_TMP}")"
  rm -f "${SLACK_REACH_TMP}"
  if is_loopback_proxy "${HTTP_PROXY:-}" || is_loopback_proxy "${HTTPS_PROXY:-}" || is_loopback_proxy "${ALL_PROXY:-}" || is_loopback_proxy "${http_proxy:-}" || is_loopback_proxy "${https_proxy:-}" || is_loopback_proxy "${all_proxy:-}"; then
    DIRECT_TMP="$(mktemp)"
    DIRECT_CODE="$(env -u HTTP_PROXY -u HTTPS_PROXY -u ALL_PROXY -u http_proxy -u https_proxy -u all_proxy -u SOCKS_PROXY -u SOCKS5_PROXY -u socks_proxy -u socks5_proxy -u GIT_HTTP_PROXY -u GIT_HTTPS_PROXY curl -sS -o /dev/null -w "%{http_code}" https://slack.com/api/api.test 2>"${DIRECT_TMP}" || true)"
    DIRECT_ERR="$(tr -d '\n' < "${DIRECT_TMP}")"
    rm -f "${DIRECT_TMP}"
    if [[ "${DIRECT_CODE}" == "200" ]]; then
      warn "Slack reachability only works without local proxy env. Current proxy appears to block outbound Slack (${CURL_ERR_MSG:-HTTP ${SLACK_REACH_CODE:-000}})."
    elif [[ -n "${CURL_ERR_MSG}" ]]; then
      bad "Outbound reachability to slack.com failed (HTTP ${SLACK_REACH_CODE:-000}; ${CURL_ERR_MSG}). Direct no-proxy test also failed (HTTP ${DIRECT_CODE:-000}; ${DIRECT_ERR:-no curl error text})."
    else
      bad "Outbound reachability to slack.com failed (HTTP ${SLACK_REACH_CODE:-000}). Direct no-proxy test also failed (HTTP ${DIRECT_CODE:-000}; ${DIRECT_ERR:-no curl error text})."
    fi
  elif [[ -n "${CURL_ERR_MSG}" ]]; then
    bad "Outbound reachability to slack.com failed (HTTP ${SLACK_REACH_CODE:-000}; ${CURL_ERR_MSG})"
  else
    bad "Outbound reachability to slack.com failed (HTTP ${SLACK_REACH_CODE:-000})"
  fi
fi

echo ""
echo "Summary: ${PASS} OK, ${WARN} WARN, ${FAIL} FAIL"
if [[ ${FAIL} -gt 0 ]]; then
  exit 1
fi
