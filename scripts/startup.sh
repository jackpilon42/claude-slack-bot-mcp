#!/usr/bin/env bash
set -euo pipefail

PROJECT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
ENV_FILE="${PROJECT_DIR}/.env"
NGROK_BIN="${PROJECT_DIR}/ngrok"
if [[ ! -x "${NGROK_BIN}" ]] && command -v ngrok >/dev/null 2>&1; then
  NGROK_BIN="$(command -v ngrok)"
fi
NGROK_API="http://127.0.0.1:4040/api/tunnels"
NGROK_LOG="/tmp/claude-slack-bot-ngrok.log"

cd "${PROJECT_DIR}"

if [[ ! -f "${ENV_FILE}" ]]; then
  echo "Missing .env file at ${ENV_FILE}"
  exit 1
fi

set -a
# shellcheck disable=SC1090
source "${ENV_FILE}"
set +a

REQUIRED_VARS=(
  "SLACK_BOT_TOKEN"
  "SLACK_SIGNING_SECRET"
  "ANTHROPIC_API_KEY"
  "SMARTSHEET_API_TOKEN"
  "SMARTSHEET_SHEET_ID"
)

disable_loopback_proxy_env() {
  local value="${1:-}"
  [[ "${value}" == http://127.0.0.1:* || "${value}" == https://127.0.0.1:* || "${value}" == socks5://127.0.0.1:* ]]
}

if disable_loopback_proxy_env "${HTTP_PROXY:-}" || disable_loopback_proxy_env "${HTTPS_PROXY:-}" || disable_loopback_proxy_env "${ALL_PROXY:-}" || disable_loopback_proxy_env "${http_proxy:-}" || disable_loopback_proxy_env "${https_proxy:-}" || disable_loopback_proxy_env "${all_proxy:-}"; then
  echo "Detected local loopback proxy vars; unsetting proxy env for bot startup."
  unset HTTP_PROXY HTTPS_PROXY ALL_PROXY http_proxy https_proxy all_proxy
  unset SOCKS_PROXY SOCKS5_PROXY socks_proxy socks5_proxy
  unset GIT_HTTP_PROXY GIT_HTTPS_PROXY
fi

MISSING=()
for name in "${REQUIRED_VARS[@]}"; do
  if [[ -z "${!name:-}" ]]; then
    MISSING+=("${name}")
  fi
done

if [[ ${#MISSING[@]} -gt 0 ]]; then
  echo "Missing required .env values: ${MISSING[*]}"
  echo "Open ${ENV_FILE} and fill them in."
  exit 1
fi

if [[ ! -x "${NGROK_BIN}" ]]; then
  echo "ngrok not found. Install:"
  echo "  1) Download a binary named 'ngrok' into ${PROJECT_DIR}/ngrok and chmod +x it, OR install ngrok on PATH."
  echo "  2) Run once: ngrok config add-authtoken <token> (from https://dashboard.ngrok.com/)"
  exit 1
fi

get_tunnel_url() {
  curl -s --max-time 2 "${NGROK_API}" | python3 - <<'PY'
import json, sys
try:
    data = json.load(sys.stdin)
except Exception:
    print("")
    raise SystemExit
for t in data.get("tunnels", []):
    if t.get("proto") == "https":
        print(t.get("public_url", ""))
        raise SystemExit
print("")
PY
}

wait_for_tunnel_url() {
  local attempts="${1:-10}"
  local delay_seconds="${2:-2}"
  local url=""
  local i=1
  while [[ "${i}" -le "${attempts}" ]]; do
    url="$(get_tunnel_url || true)"
    if [[ -n "${url}" ]]; then
      echo "${url}"
      return 0
    fi
    sleep "${delay_seconds}"
    i=$((i + 1))
  done
  echo ""
}

TUNNEL_URL="$(get_tunnel_url || true)"

if [[ -z "${TUNNEL_URL}" ]]; then
  echo "Starting ngrok for port 3000 (${NGROK_BIN})..."
  : >"${NGROK_LOG}"
  nohup "${NGROK_BIN}" http 3000 >>"${NGROK_LOG}" 2>&1 &
  # Poll the local API often so we notice the tunnel quickly; cap total wait ~12s.
  TUNNEL_URL="$(wait_for_tunnel_url 24 0.5 || true)"
fi

if [[ -n "${TUNNEL_URL}" ]]; then
  echo ""
  echo "Slack Request URL:"
  echo "${TUNNEL_URL}/slack/events"
  echo ""
else
  echo ""
  echo "ERROR: ngrok tunnel URL not detected (Slack cannot reach this Mac until this works)."
  echo "Log file: ${NGROK_LOG}"
  echo "--- last 40 lines of ngrok log ---"
  tail -n 40 "${NGROK_LOG}" 2>/dev/null || true
  echo "--- end ---"
  echo "Common fixes: run \`ngrok config add-authtoken …\`, free port 4040, or run in foreground to see errors:"
  echo "  ${NGROK_BIN} http 3000"
  echo ""
fi

if command -v lsof >/dev/null 2>&1; then
  if lsof -iTCP:3000 -sTCP:LISTEN -P -n >/dev/null 2>&1; then
    echo ""
    echo "ERROR: Port 3000 is already in use (often an older copy of this bot)."
    echo "Quit that process, then run this script again. Details:"
    lsof -iTCP:3000 -sTCP:LISTEN -P -n || true
    echo ""
    exit 1
  fi
fi

echo "Starting bot..."
exec node index.js
