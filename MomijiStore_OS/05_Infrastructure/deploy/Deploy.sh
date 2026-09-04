#!/usr/bin/env bash
# =============================================================
#  Deploy.sh — GitHub main を NAS へデプロイする
# =============================================================
#  GitHub の origin/main が唯一の正本。
#  ローカルの作業ツリーは一切参照しない(未コミットの変更が
#  NASへ紛れ込む事故を防ぐため、git archive origin/main を使う)。
#
#  使い方:
#    ./Deploy.sh          … Current/Target を表示して確認後に実行
#    ./Deploy.sh --yes    … 確認を省略(自動実行用)
#
#  触れないもの:
#    momiji-postgres / momiji_pg_data / .env / docker-compose.yaml
#    data/products(商品マスターは SyncData.sh で別途扱う)
#    **DBスキーマ**(Migrate.sh の領分)
#
#  ⚠️ スキーマ変更はここでは行わない。コードの入れ替えとDBの構造変更を
#     同じ操作にすると、どちらが原因で壊れたか切り分けられなくなる。
#     Intelligence Layer のような新機能は Migrate.sh → Deploy.sh の順で行う。
# =============================================================
set -euo pipefail

source "$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)/lib.sh"

START_TIME=$(date +%s)
ASSUME_YES=false
[[ "${1:-}" == "--yes" || "${1:-}" == "-y" ]] && ASSUME_YES=true

self_check
init_log "deploy"

echo "════════════════════════════════════════════"
echo " MomijiStore OS — Deploy"
echo "════════════════════════════════════════════"

# --- 1. 事前チェック -----------------------------------------
log_step "1/9 事前チェック"
require_ssh
require_docker
nas "test -f ${NAS_DIR}/.env" || fail "${NAS_DIR}/.env がありません"
log_ok "NAS上の .env を確認"

# --- 2. GitHubから最新を取得 ---------------------------------
log_step "2/9 GitHub から取得"
cd "$REPO_ROOT"
git fetch --quiet origin main || fail "git fetch に失敗しました"
TARGET_COMMIT=$(git rev-parse origin/main)
TARGET_SHORT=$(git rev-parse --short origin/main)
TARGET_SUBJECT=$(git log -1 --format=%s origin/main)
log_ok "origin/main を取得"

CURRENT_COMMIT=$(nas "cat ${DEPLOYED_COMMIT_FILE} 2>/dev/null" || true)
CURRENT_COMMIT="${CURRENT_COMMIT:-unknown}"
if [[ "$CURRENT_COMMIT" != "unknown" ]]; then
    CURRENT_SHORT="${CURRENT_COMMIT:0:7}"
    CURRENT_SUBJECT=$(git log -1 --format=%s "$CURRENT_COMMIT" 2>/dev/null || echo "(不明なコミット)")
else
    CURRENT_SHORT="unknown"
    CURRENT_SUBJECT="(初回デプロイ、または記録なし)"
fi

# --- 3. 差分の提示と確認 -------------------------------------
echo
echo "  Current Commit : ${CURRENT_SHORT}  ${CURRENT_SUBJECT}"
echo "  Target  Commit : ${TARGET_SHORT}  ${TARGET_SUBJECT}"
echo

if [[ "$CURRENT_COMMIT" == "$TARGET_COMMIT" ]]; then
    echo "  ※ 同一コミットです。再デプロイは冪等なので安全に実行できます。"
    echo
fi

if [[ "$ASSUME_YES" == false ]]; then
    if [[ -t 0 ]]; then
        read -r -p "  このコミットをデプロイしますか? [y/N] " answer
        [[ "$answer" =~ ^[Yy]$ ]] || { echo "  中止しました。"; exit 0; }
    else
        fail "対話できない環境です。確認を省略する場合は --yes を付けてください。"
    fi
fi

# --- 4. スナップショット取得(ロールバック用) ----------------
log_step "3/9 スナップショット取得"
SNAPSHOT_ID=$(date +%Y%m%d_%H%M%S)
SNAP_PATH="${SNAPSHOT_DIR}/${SNAPSHOT_ID}"
nas "mkdir -p '${SNAP_PATH}' \
     && cp -a '${NAS_DIR}/${COMPOSE_FILE}' '${SNAP_PATH}/' 2>/dev/null || true; \
     cp -a '${NAS_DIR}/mcp-server' '${SNAP_PATH}/' 2>/dev/null || true; \
     printf '%s' '${CURRENT_COMMIT}' > '${SNAP_PATH}/.commit'" \
    || fail "スナップショットの作成に失敗しました"
log_ok "スナップショット: ${SNAPSHOT_ID}"

# --- 5. origin/main の内容を転送 -----------------------------
#  作業ツリーではなく git archive を使う。これによりローカルの
#  未コミット変更は構造上NASへ入らない。
log_step "4/9 転送 (origin/main → NAS)"
STAGE=$(mktemp -d)
trap 'rm -rf "$STAGE"' EXIT

for item in "${DEPLOY_ITEMS[@]}"; do
    git archive origin/main "${REPO_SUBDIR}/${item}" \
        | tar -x -C "$STAGE" || fail "git archive に失敗しました (${item})"
done

sync_dir "${STAGE}/${REPO_SUBDIR}/mcp-server" "${NAS_DIR}/mcp-server"
sync_file "${STAGE}/${REPO_SUBDIR}/docker/${COMPOSE_FILE}" "${NAS_DIR}"
log_ok "転送完了"

# --- 6. ビルド -----------------------------------------------
log_step "5/9 ビルド"
nas "cd '${NAS_DIR}' && docker compose -f '${COMPOSE_FILE}' build ${SERVICE}" \
    || fail "docker build に失敗しました"
log_ok "ビルド完了"

# --- 7. 起動(momiji-mcp のみ。DBには触れない) ---------------
log_step "6/9 起動"
nas "cd '${NAS_DIR}' && docker compose -f '${COMPOSE_FILE}' up -d ${SERVICE}" \
    || fail "コンテナの起動に失敗しました"
log_ok "${SERVICE} を起動"

# --- 8. 動作確認 ---------------------------------------------
log_step "7/9 ヘルスチェック"
check_health 30
check_postgres

log_step "8/9 search_products 疎通確認"
check_search_products

log_step "9/9 MCPツールの登録確認"
check_tools

# --- 9. デプロイ済みコミットを記録 ---------------------------
nas "printf '%s' '${TARGET_COMMIT}' > '${DEPLOYED_COMMIT_FILE}'" \
    || fail "デプロイ済みコミットの記録に失敗しました"

ELAPSED=$(fmt_elapsed)
echo
echo "════════════════════════════════════════════"
echo " SUCCESS"
echo "════════════════════════════════════════════"
echo "  Commit          : ${TARGET_SHORT}  ${TARGET_SUBJECT}"
echo "  Health          : /health 200 / ${DB_SERVICE} healthy"
echo "  search_products : ${SEARCH_RESULT}"
echo "  MCP Tools       : ${TOOLS_RESULT}"
echo "  Elapsed Time    : ${ELAPSED}"
echo "  Snapshot        : ${SNAPSHOT_ID}  (Rollback.sh ${SNAPSHOT_ID})"
echo "  Log             : ${LOG_FILE}"
echo "════════════════════════════════════════════"
