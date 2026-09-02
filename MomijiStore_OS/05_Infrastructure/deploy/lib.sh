#!/usr/bin/env bash
# =============================================================
#  lib.sh — Deploy.sh / Rollback.sh の共通処理
# =============================================================
#  Git(GitHub main)が唯一の正本。NASは実行環境。
#  PROJECT_CHARTER §6 の運用に従う。
# =============================================================

# --- 設定 ----------------------------------------------------
NAS_HOST="${NAS_HOST:-momiji-nas}"
NAS_DIR="${NAS_DIR:-/volume1/MomijiStore}"
NAS_URL="${NAS_URL:-http://192.168.0.8:8000}"
COMPOSE_FILE="momiji-stack.yml"
SERVICE="momiji-mcp"          # ここ以外のサービスは操作しない
DB_SERVICE="momiji-postgres"  # 状態確認のみ。停止・再作成はしない
SNAPSHOT_DIR="${NAS_DIR}/_deploy_backup"
DEPLOYED_COMMIT_FILE="${NAS_DIR}/.deployed_commit"

# リポジトリ内のパス → NAS上の配置先
REPO_SUBDIR="MomijiStore_OS/05_Infrastructure"
DEPLOY_ITEMS=("mcp-server" "docker/${COMPOSE_FILE}")

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_ROOT="$(cd "${SCRIPT_DIR}/../../.." && pwd)"
LOG_DIR="${SCRIPT_DIR}/logs"

# --- ログ ----------------------------------------------------
init_log() {
    local prefix="$1"
    mkdir -p "$LOG_DIR"
    LOG_FILE="${LOG_DIR}/${prefix}_$(date +%Y%m%d_%H%M%S).log"
    # 標準出力・標準エラーの両方をログへ複製する
    exec > >(tee -a "$LOG_FILE") 2>&1
    echo "ログ: $LOG_FILE"
}

log()      { echo "[$(date +%H:%M:%S)] $*"; }
log_step() { echo; echo "━━━ $* ━━━"; }
log_ok()   { echo "  ✅ $*"; }
log_err()  { echo "  ❌ $*" >&2; }

# --- 異常終了 ------------------------------------------------
#  途中で失敗したら必ずここを通す。黙って続行させない。
fail() {
    log_err "$*"
    echo
    echo "════════════════════════════════════════════"
    echo " FAILED — 処理を中止しました"
    echo "════════════════════════════════════════════"
    echo " ログ: ${LOG_FILE:-未作成}"
    if [[ -n "${SNAPSHOT_ID:-}" ]]; then
        echo
        echo " 復旧するには:"
        echo "   ${SCRIPT_DIR}/Rollback.sh ${SNAPSHOT_ID}"
    fi
    exit 1
}

# --- 安全装置 ------------------------------------------------
#  スクリプト自身に破壊的コマンドが混入していないか毎回検査する。
#  PostgreSQLのボリュームを消す事故を構造的に防ぐ。
self_check() {
    local found
    found=$(grep -nE 'down\s+.*-v|down\s+--volumes|volume\s+rm' \
        "${SCRIPT_DIR}"/*.sh | grep -v 'grep -nE' || true)
    if [[ -n "$found" ]]; then
        log_err "破壊的コマンドを検出したため実行を中止します:"
        echo "$found" >&2
        exit 1
    fi
}

# --- SSH -----------------------------------------------------
nas() { ssh -o BatchMode=yes -o ConnectTimeout=10 "$NAS_HOST" "$@"; }

require_ssh() {
    nas true 2>/dev/null || fail "NASへSSH接続できません (host=${NAS_HOST})。~/.ssh/config と鍵を確認してください。"
    log_ok "SSH接続 OK (${NAS_HOST})"
}

require_docker() {
    # sudoなしでdockerを使えることが前提(usermod -aG docker 済み)
    nas 'docker ps >/dev/null 2>&1' \
        || fail "NAS上でdockerを実行できません。'sudo usermod -aG docker momiji-admin' 実施後、SSHを再接続してください。"
    log_ok "docker 実行権限 OK"
}

# --- トークン取得 --------------------------------------------
#  ⚠️ 取得した値はログにも画面にも出さない。
get_token() {
    local raw
    raw=$(nas "grep -m1 -E '^MOMIJI_API_KEYS=|^MOMIJI_API_KEY=' ${NAS_DIR}/.env 2>/dev/null | cut -d= -f2-") \
        || fail ".env からトークンを取得できません"
    # カンマ区切りの先頭1本を使う
    raw="${raw%%,*}"
    raw="$(echo "$raw" | tr -d '[:space:]')"
    [[ -n "$raw" ]] || fail ".env に MOMIJI_API_KEYS が設定されていません"
    printf '%s' "$raw"
}

# --- 動作確認 ------------------------------------------------
check_health() {
    local tries="${1:-30}" i code
    for ((i = 1; i <= tries; i++)); do
        code=$(curl -s -o /dev/null -w '%{http_code}' -m 5 "${NAS_URL}/health" || true)
        if [[ "$code" == "200" ]]; then
            log_ok "/health 200 (${i}回目で応答)"
            return 0
        fi
        sleep 2
    done
    fail "/health が ${tries} 回試行しても 200 になりません (最後の応答: ${code:-なし})"
}

check_postgres() {
    local status
    status=$(nas "docker inspect -f '{{.State.Health.Status}}' ${DB_SERVICE} 2>/dev/null" || echo "unknown")
    [[ "$status" == "healthy" ]] \
        || fail "${DB_SERVICE} が healthy ではありません (status=${status})"
    log_ok "${DB_SERVICE} healthy を維持"
}

#  search_products を実際に呼んで結果を確認する。
#  トークンはヘッダーで渡し、コマンドラインにもログにも残さない。
check_search_products() {
    local token session result
    token="$(get_token)"

    session=$(curl -s -D - -o /dev/null -m 10 -X POST "${NAS_URL}/mcp" \
        -H "Authorization: Bearer ${token}" \
        -H 'Content-Type: application/json' \
        -H 'Accept: application/json, text/event-stream' \
        -d '{"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"2025-06-18","capabilities":{},"clientInfo":{"name":"deploy","version":"1"}}}' \
        | tr -d '\r' | awk 'tolower($1)=="mcp-session-id:"{print $2}')
    [[ -n "$session" ]] || fail "MCP initialize に失敗しました(認証エラーの可能性)"

    curl -s -m 10 -X POST "${NAS_URL}/mcp" \
        -H "Authorization: Bearer ${token}" \
        -H 'Content-Type: application/json' \
        -H 'Accept: application/json, text/event-stream' \
        -H "mcp-session-id: ${session}" \
        -d '{"jsonrpc":"2.0","method":"notifications/initialized"}' >/dev/null

    result=$(curl -s -m 15 -X POST "${NAS_URL}/mcp" \
        -H "Authorization: Bearer ${token}" \
        -H 'Content-Type: application/json' \
        -H 'Accept: application/json, text/event-stream' \
        -H "mcp-session-id: ${session}" \
        -d '{"jsonrpc":"2.0","id":2,"method":"tools/call","params":{"name":"search_products","arguments":{"internal_id":"P000021"}}}' \
        | sed -n 's/^data: //p')

    if [[ "$result" != *'"内部管理ID"'* ]] || [[ "$result" == *'"error"'* ]]; then
        log_err "search_products の応答: ${result:-なし}"
        fail "search_products が想定どおりに応答しません"
    fi
    SEARCH_RESULT="P000021 を1件取得"
    log_ok "search_products OK (${SEARCH_RESULT})"
}

fmt_elapsed() {
    local sec=$(( $(date +%s) - START_TIME ))
    printf '%dm%02ds' $((sec / 60)) $((sec % 60))
}
