#!/usr/bin/env bash
# =============================================================
#  run_product_core_tests.sh — Product Core の DDL・制約テストを
#  NAS 上の「隔離された使い捨て PostgreSQL」で実行する
# =============================================================
#  使い方:  ./run_product_core_tests.sh
#
#  必須条件(2026-09-26 承認)と、このスクリプトでの担保:
#    本番 DB へ接続しない          … ネットワーク none で起動(どこにも接続できない)。本番コンテナには exec しない
#    本番 DB のボリュームを使わない … データ領域は tmpfs(メモリ)。ボリューム・バインドは付けない
#    本番の環境ファイル・認証情報を使わない … 読まない・渡さない。lib.sh も読み込まない(本番設定の読み出し関数を持つため)
#    テスト専用の一時 DB・ユーザー・パスワード … 実行ごとに乱数で生成。パスワードは引数に乗せず標準入力で渡す
#    外部ポートを公開しない        … ポート指定をしない
#    終了後にコンテナを削除        … trap で必ず削除し、消えたことを確認する
#    本番と明確に異なる名前        … aikos-pctest-<日時>-<乱数>。本番名と一致したら止める
#  起動後にも docker inspect で上の条件を検査し、1つでも満たさなければ即座に削除して止める。
#
#  前提: NAS に SSH 鍵で入れること(~/.ssh/config の momiji-nas)、sudo なしで docker を使えること。
# =============================================================
set -euo pipefail

NAS_HOST="${NAS_HOST:-momiji-nas}"
NAS_HOME="/tmp/momiji-deploy-home"          # UGOS では HOME が作れないため(lib.sh と同じ理由)
IMAGE="postgres:16.14-alpine"               # 本番と同じイメージ
PROD_CONTAINERS=("momiji-postgres" "momiji-mcp")
RUN_ID="$(date +%Y%m%d%H%M%S)-${RANDOM}"
NAME="aikos-pctest-${RUN_ID}"
TEST_DB="pctest"
TEST_USER="pctest_admin"

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
INFRA_DIR="$(cd "${SCRIPT_DIR}/../.." && pwd)"
MIG_DIR="${INFRA_DIR}/db/migrations"
TEST_DIR="${SCRIPT_DIR}/product_core"
LOG_DIR="${SCRIPT_DIR}/logs"
[[ -d "$LOG_DIR" ]] || mkdir "$LOG_DIR"
LOG_FILE="${LOG_DIR}/product_core_test_${RUN_ID}.log"
exec > >(tee -a "$LOG_FILE") 2>&1

log()  { echo "[$(date +%H:%M:%S)] $*"; }
ok()   { echo "  ✅ $*"; }
ng()   { echo "  ❌ $*" >&2; }
fail() { ng "$*"; echo " FAILED — ログ: ${LOG_FILE}"; exit 1; }

# ---- 安全装置 0: スクリプト自身に禁止した操作が紛れ込んでいないか ----------------
#  パターンは実行時に組み立てる(このファイル自身の定義行に一致しないように)
self_check() {
    local pats=( "$(printf '%s' '\.e' 'nv')" "$(printf '%s' 'momiji_pg' '_data')" "$(printf '%s' 'docker ' 'compose')"
                 "$(printf '%s' 'docker-' 'compose')" "$(printf '%s' 'env-' 'file')" "$(printf '%s' '--vol' 'ume')"
                 "$(printf '%s' '--mo' 'unt')" "$(printf '%s' '--pub' 'lish')" "$(printf '%s' ' -' 'v ')" "$(printf '%s' ' -' 'p ')"
                 "$(printf '%s' 'MOMIJI_' 'DB_')" "$(printf '%s' 'lib' '.sh\"')" )
    local p hit
    for p in "${pats[@]}"; do
        hit=$(grep -nE -- "$p" "${BASH_SOURCE[0]}" | grep -v 'printf' || true)
        [[ -z "$hit" ]] || fail "スクリプトに禁止した操作が含まれています: ${hit}"
    done
    ok "スクリプトの自己検査(禁止した操作なし)"
}

# ---- 安全装置 1: 対象コンテナ名 --------------------------------------------------
guard_name() {
    [[ "$NAME" =~ ^aikos-pctest-[0-9]{14}-[0-9]+$ ]] || fail "テストコンテナ名の形式が不正: ${NAME}"
    local p
    for p in "${PROD_CONTAINERS[@]}"; do
        [[ "$NAME" != "$p" ]] || fail "本番コンテナ名と一致したため中止: ${NAME}"
    done
}

nas() {
    ssh -o BatchMode=yes -o ConnectTimeout=10 "$NAS_HOST" "export HOME='${NAS_HOME}'; mkdir \"\$HOME\" 2>/dev/null; $*"
}

# テストコンテナ以外には exec しない
test_exec() {
    guard_name
    nas "docker exec -i '${NAME}' $*"
}

CREATED=false
cleanup() {
    local rc=$?
    if [[ "$CREATED" == true ]]; then
        guard_name
        log "テストコンテナを削除: ${NAME}"
        nas "docker rm -f '${NAME}' >/dev/null 2>&1" || true
        if [[ -z "$(nas "docker ps -aq --filter 'name=^/${NAME}\$'")" ]]; then
            ok "テストコンテナの削除を確認(残存なし)"
        else
            ng "テストコンテナが残っています: ${NAME}(手動で確認してください)"
            rc=1
        fi
    fi
    exit $rc
}
trap cleanup EXIT

echo "════════════════════════════════════════════"
echo " Product Core 制約テスト(使い捨て PostgreSQL)"
echo "════════════════════════════════════════════"
echo " ログ: ${LOG_FILE}"

# ---- 1. 事前確認 ----------------------------------------------------------------
log "1/6 事前確認"
self_check
guard_name
ok "テストコンテナ名: ${NAME}"
for f in 001_intelligence_layer.sql 002_product_core.sql 003_product_core_roles.sql; do
    [[ -f "${MIG_DIR}/${f}" ]] || fail "マイグレーションがありません: ${f}"
done
# Migrate.sh と同じ破壊的 SQL の検査を、テスト前にも通しておく
found=$(grep -inE '(drop[[:space:]]+(table|schema|database|column))|truncate[[:space:]]|delete[[:space:]]+from' "${MIG_DIR}"/*.sql || true)
[[ -z "$found" ]] || fail "マイグレーションに破壊的 SQL: ${found}"
ok "マイグレーションに破壊的 SQL なし(Migrate.sh と同じ検査)"
nas true 2>/dev/null || fail "NAS へ SSH 接続できません(${NAS_HOST})。UGOS で SSH を有効にし、鍵認証を確認してください"
ok "SSH 接続 OK"
nas 'docker ps >/dev/null 2>&1' || fail "NAS で docker を実行できません"
nas "docker image inspect '${IMAGE}' >/dev/null 2>&1" || fail "NAS に ${IMAGE} がありません(自動で取得はしません)"
ok "イメージ ${IMAGE} が NAS 上にある"
PROD_STARTED_BEFORE=$(nas "docker inspect --format '{{.State.StartedAt}}' momiji-postgres 2>/dev/null" || echo "なし")
log "本番 DB コンテナの起動時刻(読み取りのみ): ${PROD_STARTED_BEFORE}"

# ---- 2. 使い捨てコンテナの起動 ------------------------------------------------------
log "2/6 使い捨てコンテナの起動"
TEST_PW="$(openssl rand -hex 24)"
# パスワードは標準入力で渡す(コマンドライン・ログに残さない)
printf '%s' "$TEST_PW" | nas "PCTEST_PW=\$(cat); docker run -d --rm --name '${NAME}' --network none \
    --tmpfs /var/lib/postgresql/data:rw,size=512m --memory 768m \
    --label aikos.purpose=product-core-test --label aikos.run='${RUN_ID}' \
    -e POSTGRES_USER='${TEST_USER}' -e POSTGRES_DB='${TEST_DB}' -e POSTGRES_PASSWORD=\"\$PCTEST_PW\" \
    '${IMAGE}' >/dev/null" || fail "テストコンテナを起動できません"
unset TEST_PW
CREATED=true
ok "起動: ${NAME}"

# ---- 3. 隔離の検査(1つでも満たさなければ止める → trap で削除) -------------------------
log "3/6 隔離の検査"
INSPECT=$(nas "docker inspect --format '{{.Name}}|{{.HostConfig.NetworkMode}}|{{len .Mounts}}|{{json .HostConfig.PortBindings}}|{{.Config.Image}}|{{index .Config.Labels \"aikos.purpose\"}}' '${NAME}'")
IFS='|' read -r i_name i_net i_mounts i_ports i_image i_label <<< "$INSPECT"
[[ "$i_name" == "/${NAME}" ]]                        || fail "コンテナ名が一致しません: ${i_name}"
[[ "$i_net" == "none" ]]                             || fail "ネットワークが none ではありません: ${i_net}"
[[ "$i_mounts" == "0" ]]                             || fail "ボリューム / バインドが付いています: ${i_mounts}"
[[ "$i_ports" == "{}" || "$i_ports" == "null" ]]     || fail "ポートが公開されています: ${i_ports}"
[[ "$i_image" == "$IMAGE" ]]                         || fail "イメージが違います: ${i_image}"
[[ "$i_label" == "product-core-test" ]]              || fail "テスト用ラベルがありません"
ENV_HITS=$(nas "docker inspect --format '{{range .Config.Env}}{{println .}}{{end}}' '${NAME}'" | grep -c '^MOMIJI' || true)
[[ "$ENV_HITS" == "0" ]]                             || fail "本番の環境変数が渡っています"
ok "ネットワーク none / ボリューム 0 / ポート公開なし / 本番の環境変数なし / イメージ ${IMAGE}"

for i in $(seq 1 60); do
    if nas "docker logs '${NAME}' 2>&1 | grep -q 'init process complete'" \
       && test_exec "pg_isready -q -U '${TEST_USER}' -d '${TEST_DB}'" >/dev/null 2>&1; then
        break
    fi
    [[ $i -lt 60 ]] || fail "テスト DB が起動しません"
    sleep 1
done
ok "テスト DB 起動完了"

PSQL="psql -X -q --set ON_ERROR_STOP=1 -U '${TEST_USER}' -d '${TEST_DB}'"

# ---- 4. マイグレーション適用(2回: 再適用でも壊れないこと) ----------------------------
log "4/6 マイグレーション適用(001 → 002 → 003、Migrate.sh と同じく1トランザクション)"
cat "${MIG_DIR}/001_intelligence_layer.sql" | test_exec "${PSQL} --single-transaction" || fail "001 の適用に失敗"
cat "${MIG_DIR}/002_product_core.sql" "${MIG_DIR}/003_product_core_roles.sql" | test_exec "${PSQL} --single-transaction" \
    || fail "002 / 003 の適用に失敗"
ok "初回適用"
cat "${MIG_DIR}/002_product_core.sql" "${MIG_DIR}/003_product_core_roles.sql" | test_exec "${PSQL} --single-transaction" \
    || fail "002 / 003 の再適用に失敗(冪等でない)"
ok "再適用(冪等性)"

# ---- 5. 制約テスト -----------------------------------------------------------------
log "5/6 制約テスト"
set +e
cat "${TEST_DIR}/00_harness.sql" "${TEST_DIR}/10_core.sql" "${TEST_DIR}/20_cost.sql" \
    "${TEST_DIR}/30_identifier_roles.sql" "${TEST_DIR}/50_revision.sql" \
    | test_exec "psql -X -q --set ON_ERROR_STOP=1 -U '${TEST_USER}' -d '${TEST_DB}'" >/dev/null
TEST_RC=$?
# AI のプロセスと同じ条件(スーパーユーザーではない pc_ingest メンバー)の別セッション
if [[ $TEST_RC -eq 0 ]]; then
    cat "${TEST_DIR}/40_ai_session.sql" \
        | test_exec "psql -X -q --set ON_ERROR_STOP=1 -U pctest_ai_login -d '${TEST_DB}'" >/dev/null
    TEST_RC=$?
fi
cat "${TEST_DIR}/90_report.sql" | test_exec "psql -X --set ON_ERROR_STOP=1 -U '${TEST_USER}' -d '${TEST_DB}'" || TEST_RC=1
set -e

# ---- 6. 本番への影響が無いことの確認(読み取りのみ) --------------------------------------
log "6/6 本番への影響確認"
PROD_STARTED_AFTER=$(nas "docker inspect --format '{{.State.StartedAt}}' momiji-postgres 2>/dev/null" || echo "なし")
if [[ "$PROD_STARTED_BEFORE" == "$PROD_STARTED_AFTER" ]]; then
    ok "本番 DB コンテナは再起動・変更されていない(起動時刻 ${PROD_STARTED_AFTER})"
else
    ng "本番 DB コンテナの起動時刻が変わりました: ${PROD_STARTED_BEFORE} → ${PROD_STARTED_AFTER}"
    TEST_RC=1
fi

if [[ $TEST_RC -eq 0 ]]; then
    echo "════════════════════════════════════════════"
    echo " PASS — 全テスト成功"
    echo "════════════════════════════════════════════"
else
    echo "════════════════════════════════════════════"
    echo " FAIL — ログを確認してください: ${LOG_FILE}"
    echo "════════════════════════════════════════════"
fi
exit $TEST_RC
