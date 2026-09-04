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
#  UGOSでは momiji-admin のホーム(/home/momiji-admin)が存在せず、
#  作成もできない(root でも Operation not permitted)。
#  docker CLI は起動時に HOME を作ろうとして失敗するため、
#  書き込み可能な場所を HOME として与える。
NAS_HOME="${NAS_HOME:-/tmp/momiji-deploy-home}"

nas() {
    ssh -o BatchMode=yes -o ConnectTimeout=10 "$NAS_HOST" \
        "export HOME='${NAS_HOME}'; mkdir -p \"\$HOME\" 2>/dev/null; $*"
}

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

# --- 転送 ----------------------------------------------------
#  rsync は使えない。UGOSのrsyncは独自改変版で、一般ユーザーだと
#    "cannot set euid as root" → "invalid path" で失敗する(UGOS仕様)。
#  そのため tar over SSH を使う。追加依存はなく、確実に動く。
#  COPYFILE_DISABLE=1 は macOS が ._* (AppleDouble) を混ぜるのを防ぐ。
#
#  ※ 差分削除は行わない(上書きと追加のみ)。Gitから削除されたファイルが
#     NASに残る場合は、スナップショットを確認のうえ手動で整理する。

sync_dir() {
    local src="$1" dst="$2"
    [[ -d "$src" ]] || fail "転送元がありません: $src"
    COPYFILE_DISABLE=1 tar -C "$src" -czf - . \
        | nas "mkdir -p '$dst' && tar -xzf - -C '$dst'" \
        || fail "ディレクトリの転送に失敗しました: $src → $dst"
}

sync_file() {
    local src="$1" dst_dir="$2"
    [[ -f "$src" ]] || fail "転送元がありません: $src"
    COPYFILE_DISABLE=1 tar -C "$(dirname "$src")" -czf - "$(basename "$src")" \
        | nas "tar -xzf - -C '$dst_dir'" \
        || fail "ファイルの転送に失敗しました: $src → $dst_dir"
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

#  MCPセッションを開いてセッションIDを返す。
#  トークンはヘッダーで渡し、コマンドラインにもログにも残さない。
mcp_open_session() {
    local token="$1" headers session code
    headers=$(curl -s -D - -o /dev/null -m 10 -X POST "${NAS_URL}/mcp" \
        -H "Authorization: Bearer ${token}" \
        -H 'Content-Type: application/json' \
        -H 'Accept: application/json, text/event-stream' \
        -d '{"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"2025-06-18","capabilities":{},"clientInfo":{"name":"deploy","version":"1"}}}' \
        | tr -d '\r')
    session=$(echo "$headers" | awk 'tolower($1)=="mcp-session-id:"{print $2}')
    code=$(echo "$headers" | awk 'NR==1{print $2}')

    if [[ -z "$session" ]]; then
        # 原因を推測しない。実際のHTTPステータスとサーバーログを提示する。
        log_err "MCP initialize が失敗しました (HTTP ${code:-不明})"
        log_err "サーバーログ(直近10行):"
        nas 'docker logs --tail 10 momiji-mcp 2>&1' | sed 's/^/      /' >&2 || true
        fail "MCP initialize に失敗しました"
    fi

    curl -s -m 10 -X POST "${NAS_URL}/mcp" \
        -H "Authorization: Bearer ${token}" \
        -H 'Content-Type: application/json' \
        -H 'Accept: application/json, text/event-stream' \
        -H "mcp-session-id: ${session}" \
        -d '{"jsonrpc":"2.0","method":"notifications/initialized"}' >/dev/null

    printf '%s' "$session"
}

#  開いているセッションでJSON-RPCを1回呼ぶ。応答本文(data:行)を返す。
mcp_call() {
    local token="$1" session="$2" payload="$3"
    curl -s -m 15 -X POST "${NAS_URL}/mcp" \
        -H "Authorization: Bearer ${token}" \
        -H 'Content-Type: application/json' \
        -H 'Accept: application/json, text/event-stream' \
        -H "mcp-session-id: ${session}" \
        -d "$payload" \
        | sed -n 's/^data: //p'
}

#  search_products を実際に呼んで結果を確認する。
check_search_products() {
    local token session result
    token="$(get_token)"
    session="$(mcp_open_session "$token")"

    result=$(mcp_call "$token" "$session" \
        '{"jsonrpc":"2.0","id":2,"method":"tools/call","params":{"name":"search_products","arguments":{"internal_id":"P000021"}}}')

    # tools/call の結果はJSON文字列として入れ子になり、引用符がエスケープされる。
    # そのため引用符を含まない部分文字列で判定する。
    if [[ "$result" != *'内部管理ID'* ]] || [[ "$result" == *'"isError":true'* ]]; then
        log_err "search_products の応答: ${result:-なし}"
        fail "search_products が想定どおりに応答しません"
    fi
    local count
    count=$(printf '%s' "$result" | sed -n 's/.*\\"count\\": \([0-9]*\).*/\1/p' | head -1)
    SEARCH_RESULT="P000021 を ${count:-1} 件取得"
    log_ok "search_products OK (${SEARCH_RESULT})"
}

#  期待するツールが実際に登録されているかを確認する。
#  「デプロイできた」と「ツールが使える」は別。必ず列挙して確かめる。
#  ※ 記録系は呼び出さない(確認のためにダミーの判断を残さない)。
EXPECTED_TOOLS=(
    health_check
    search_products
    record_decision
    record_outcome
    search_decisions
    list_pending_reviews
)

check_tools() {
    local token session result name
    local missing=()
    token="$(get_token)"
    session="$(mcp_open_session "$token")"

    result=$(mcp_call "$token" "$session" '{"jsonrpc":"2.0","id":3,"method":"tools/list"}')

    for name in "${EXPECTED_TOOLS[@]}"; do
        [[ "$result" == *"\"${name}\""* ]] || missing+=("$name")
    done

    if (( ${#missing[@]} > 0 )); then
        log_err "登録されていないツール: ${missing[*]}"
        log_err "tools/list の応答: ${result:-なし}"
        fail "MCPツールの登録を確認できません"
    fi

    TOOLS_RESULT="${#EXPECTED_TOOLS[@]}種すべて登録済"
    log_ok "MCPツール OK (${TOOLS_RESULT})"
}

#  Intelligence Layer が実際にDBへ届くかを確認する。
#
#  ツールが登録されていることと、それが動くことは別。
#  2026-09-05、DBパスワードの不整合が「最初にツールを呼んだ時」まで
#  表面化しなかった。Migrate.sh は docker exec のローカルソケット
#  (trust認証)を通るためパスワードを検証できず、素通りしていた。
#  デプロイのたびに実際の接続経路を通す。
#
#  ※ 読み取りのみ。確認のために判断記録を作らない(記録は追記専用で消せない)。
check_intelligence() {
    local token session result
    token="$(get_token)"
    session="$(mcp_open_session "$token")"

    result=$(mcp_call "$token" "$session" \
        '{"jsonrpc":"2.0","id":4,"method":"tools/call","params":{"name":"list_pending_reviews","arguments":{"only_due":false}}}')

    if [[ "$result" == *'"isError":true'* ]] || [[ "$result" == *'error'* ]] \
       || [[ "$result" != *'matched'* ]]; then
        log_err "list_pending_reviews の応答: ${result:-なし}"
        log_err "サーバーログ(直近10行):"
        nas 'docker logs --tail 10 momiji-mcp 2>&1' | sed 's/^/      /' >&2 || true
        fail "Intelligence Layer がDBへ到達できません"
    fi

    local pending
    pending=$(printf '%s' "$result" | sed -n 's/.*\\"matched\\": \([0-9]*\).*/\1/p' | head -1)
    INTEL_RESULT="DB接続OK / 結果未記録 ${pending:-?} 件"
    log_ok "Intelligence Layer OK (${INTEL_RESULT})"
}

fmt_elapsed() {
    local sec=$(( $(date +%s) - START_TIME ))
    printf '%dm%02ds' $((sec / 60)) $((sec % 60))
}
