#!/usr/bin/env bash
# =============================================================
#  Rollback.sh — Deploy前のスナップショットへ戻す
# =============================================================
#  使い方:
#    ./Rollback.sh --list          … スナップショット一覧
#    ./Rollback.sh                 … 直近のスナップショットへ戻す
#    ./Rollback.sh 20260903_153000 … 指定時点へ戻す
#    ./Rollback.sh <id> --yes      … 確認を省略
#
#  触れないもの:
#    momiji-postgres / momiji_pg_data / .env / docker-compose.yaml
#  古いスナップショットは自動削除しない(CHARTER §5)。
# =============================================================
set -euo pipefail

source "$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)/lib.sh"

START_TIME=$(date +%s)
TARGET_ID=""
ASSUME_YES=false
for arg in "$@"; do
    case "$arg" in
        --list) LIST_ONLY=true ;;
        --yes|-y) ASSUME_YES=true ;;
        *) TARGET_ID="$arg" ;;
    esac
done

self_check

# --- 一覧表示 ------------------------------------------------
if [[ "${LIST_ONLY:-false}" == true ]]; then
    require_ssh
    echo "利用可能なスナップショット (${SNAPSHOT_DIR}):"
    nas "for d in ${SNAPSHOT_DIR}/*/; do
            [ -d \"\$d\" ] || continue
            id=\$(basename \"\$d\")
            c=\$(cat \"\$d/.commit\" 2>/dev/null | cut -c1-7)
            printf '  %s  commit=%s\n' \"\$id\" \"\${c:-unknown}\"
         done" 2>/dev/null || echo "  (スナップショットがありません)"
    exit 0
fi

init_log "rollback"

echo "════════════════════════════════════════════"
echo " MomijiStore OS — Rollback"
echo "════════════════════════════════════════════"

log_step "1/5 事前チェック"
require_ssh
require_docker

# --- 対象スナップショットの決定 ------------------------------
if [[ -z "$TARGET_ID" ]]; then
    TARGET_ID=$(nas "ls -1 ${SNAPSHOT_DIR} 2>/dev/null | sort | tail -1" || true)
    [[ -n "$TARGET_ID" ]] || fail "スナップショットが1つもありません"
    log "対象を自動選択: ${TARGET_ID}(直近)"
fi

SNAP_PATH="${SNAPSHOT_DIR}/${TARGET_ID}"
nas "test -d '${SNAP_PATH}'" || fail "スナップショットが見つかりません: ${TARGET_ID}"

SNAP_COMMIT=$(nas "cat '${SNAP_PATH}/.commit' 2>/dev/null" || echo "unknown")
CURRENT_COMMIT=$(nas "cat ${DEPLOYED_COMMIT_FILE} 2>/dev/null" || echo "unknown")

echo
echo "  Current Commit : ${CURRENT_COMMIT:0:7}"
echo "  Rollback To    : ${SNAP_COMMIT:0:7}  (snapshot ${TARGET_ID})"
echo

if [[ "$ASSUME_YES" == false ]]; then
    if [[ -t 0 ]]; then
        read -r -p "  この状態へ戻しますか? [y/N] " answer
        [[ "$answer" =~ ^[Yy]$ ]] || { echo "  中止しました。"; exit 0; }
    else
        fail "対話できない環境です。確認を省略する場合は --yes を付けてください。"
    fi
fi

# --- 復元 ----------------------------------------------------
#  戻す前に現状も退避する。ロールバック自体をやり直せるようにするため。
log_step "2/5 現状を退避"
PRE_ID="pre_rollback_$(date +%Y%m%d_%H%M%S)"
nas "mkdir -p '${SNAPSHOT_DIR}/${PRE_ID}' \
     && cp -a '${NAS_DIR}/${COMPOSE_FILE}' '${SNAPSHOT_DIR}/${PRE_ID}/' 2>/dev/null || true; \
     cp -a '${NAS_DIR}/mcp-server' '${SNAPSHOT_DIR}/${PRE_ID}/' 2>/dev/null || true; \
     printf '%s' '${CURRENT_COMMIT}' > '${SNAPSHOT_DIR}/${PRE_ID}/.commit'"
log_ok "退避: ${PRE_ID}"

log_step "3/5 復元"
nas "cp -a '${SNAP_PATH}/${COMPOSE_FILE}' '${NAS_DIR}/${COMPOSE_FILE}' \
     && rm -rf '${NAS_DIR}/mcp-server.rollback_tmp' \
     && cp -a '${SNAP_PATH}/mcp-server' '${NAS_DIR}/mcp-server.rollback_tmp' \
     && rm -rf '${NAS_DIR}/mcp-server' \
     && mv '${NAS_DIR}/mcp-server.rollback_tmp' '${NAS_DIR}/mcp-server'" \
    || fail "ファイルの復元に失敗しました"
log_ok "ファイルを復元"

log_step "4/5 再ビルド・起動"
nas "cd '${NAS_DIR}' && docker compose -f '${COMPOSE_FILE}' up -d --build ${SERVICE}" \
    || fail "再ビルド・起動に失敗しました"
log_ok "${SERVICE} を再起動"

log_step "5/5 動作確認"
check_health 30
check_postgres
check_search_products

nas "printf '%s' '${SNAP_COMMIT}' > '${DEPLOYED_COMMIT_FILE}'"

ELAPSED=$(fmt_elapsed)
echo
echo "════════════════════════════════════════════"
echo " ROLLBACK SUCCESS"
echo "════════════════════════════════════════════"
echo "  Restored To     : ${SNAP_COMMIT:0:7}  (snapshot ${TARGET_ID})"
echo "  Health          : /health 200 / ${DB_SERVICE} healthy"
echo "  search_products : ${SEARCH_RESULT}"
echo "  Elapsed Time    : ${ELAPSED}"
echo "  やり直す場合     : Rollback.sh ${PRE_ID}"
echo "  Log             : ${LOG_FILE}"
echo "════════════════════════════════════════════"
