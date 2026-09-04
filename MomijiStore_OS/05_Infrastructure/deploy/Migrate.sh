#!/usr/bin/env bash
# =============================================================
#  Migrate.sh — DBスキーマを GitHub main の定義に合わせる
# =============================================================
#  GitHub の origin/main が唯一の正本。Deploy.sh と同じ原則で、
#  ローカルの作業ツリーは一切参照しない(git archive を使う)。
#
#  使い方:
#    ./Migrate.sh          … 未適用のマイグレーションを表示して確認後に適用
#    ./Migrate.sh --yes    … 確認を省略(自動実行用)
#    ./Migrate.sh --status … 適用状況の確認のみ(何も変更しない)
#
#  触れないもの:
#    momiji_pg_data(ボリューム)/ .env / momiji-mcp / data/products
#    既存の業務データ。この仕組みで DROP / TRUNCATE / DELETE は行えない。
#
#  【なぜ psql をコンテナ内で動かすか】
#  momiji-postgres は ports を開けていない(外部公開しない設計)。
#  docker exec 経由なら Unix ソケット接続になり、パスワードを
#  コマンドラインにもログにも出さずに済む。
# =============================================================
set -euo pipefail

source "$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)/lib.sh"

START_TIME=$(date +%s)
MODE="apply"
case "${1:-}" in
    --yes | -y)   ASSUME_YES=true ;;
    --status)     MODE="status"; ASSUME_YES=true ;;
    "")           ASSUME_YES=false ;;
    *)            echo "不明な引数: $1" >&2; exit 1 ;;
esac

MIGRATIONS_SUBDIR="${REPO_SUBDIR}/db/migrations"

self_check
init_log "migrate"

echo "════════════════════════════════════════════"
echo " MomijiStore OS — DB Migrate"
echo "════════════════════════════════════════════"

# --- psql 実行 -----------------------------------------------
#  SQLは標準入力で渡す。コマンドラインに乗せないことで、
#  ps や履歴・ログに内容が残らない。
#  -X: ~/.psqlrc を読まない / -q: 冗長出力を抑える
#  ON_ERROR_STOP=1: 途中でエラーが出たら必ず止める(黙って進めない)
psql_pipe() {
    local extra="${1:-}"
    nas "docker exec -i ${DB_SERVICE} sh -c 'psql -X -q -v ON_ERROR_STOP=1 ${extra} -U \"\$POSTGRES_USER\" -d \"\$POSTGRES_DB\"'"
}

# 値を1つ取り出す(タプルのみ・区切りなし)
psql_value() {
    printf '%s\n' "$1" | psql_pipe "-t -A"
}

# --- 破壊的SQLの検査 -----------------------------------------
#  lib.sh の self_check と同じ考え方。スクリプトだけでなく
#  流し込むSQLも検査する。データを消すマイグレーションは
#  この仕組みでは流せない(必要なら人が手順を設計して実施する)。
sql_safety_check() {
    local dir="$1" found
    found=$(grep -inE '(drop[[:space:]]+(table|schema|database|column))|truncate[[:space:]]|delete[[:space:]]+from' \
        "$dir"/*.sql 2>/dev/null || true)
    if [[ -n "$found" ]]; then
        log_err "マイグレーションに破壊的SQLを検出したため中止します:"
        echo "$found" >&2
        echo >&2
        echo "  データを削除する変更は Migrate.sh では実行しません。" >&2
        echo "  影響範囲と復元方法を提示し、承認を得たうえで個別に実施してください。" >&2
        exit 1
    fi
    log_ok "破壊的SQLなし"
}

# --- 1. 事前チェック -----------------------------------------
log_step "1/5 事前チェック"
require_ssh
require_docker
check_postgres

# --- 2. GitHub から定義を取得 --------------------------------
log_step "2/5 GitHub からマイグレーションを取得"
cd "$REPO_ROOT"
git fetch --quiet origin main || fail "git fetch に失敗しました"
TARGET_SHORT=$(git rev-parse --short origin/main)

STAGE=$(mktemp -d)
trap 'rm -rf "$STAGE"' EXIT
git archive origin/main "${MIGRATIONS_SUBDIR}" | tar -x -C "$STAGE" \
    || fail "git archive に失敗しました (${MIGRATIONS_SUBDIR})"

MIG_DIR="${STAGE}/${MIGRATIONS_SUBDIR}"
[[ -d "$MIG_DIR" ]] || fail "マイグレーションディレクトリがありません: ${MIGRATIONS_SUBDIR}"

shopt -s nullglob
MIGRATIONS=("$MIG_DIR"/*.sql)
shopt -u nullglob
(( ${#MIGRATIONS[@]} > 0 )) || fail "マイグレーションファイルが1つもありません"
log_ok "origin/main (${TARGET_SHORT}) から ${#MIGRATIONS[@]} 件を取得"

sql_safety_check "$MIG_DIR"

# --- 3. 適用履歴テーブルを用意 -------------------------------
#  「何をいつ適用したか」が残らない変更は認めない(BL-7)。
#
#  ※ --status は何も変更しない。テーブルが無ければ「履歴なし」として扱い、
#     作成はしない(確認のつもりで実行したらDBが変わっていた、を起こさない)。
log_step "3/5 適用履歴の確認"
HISTORY_EXISTS=$(psql_value "SELECT to_regclass('public.schema_migrations') IS NOT NULL;") \
    || fail "適用履歴テーブルの有無を確認できません"

if [[ "$HISTORY_EXISTS" != "t" ]]; then
    if [[ "$MODE" == "status" ]]; then
        log "適用履歴テーブルは未作成です(--status のため作成しません)"
    else
        printf '%s\n' "
CREATE TABLE public.schema_migrations (
    filename    text        PRIMARY KEY,
    checksum    text        NOT NULL,
    applied_at  timestamptz NOT NULL DEFAULT now(),
    applied_by  text
);
COMMENT ON TABLE public.schema_migrations IS 'Migrate.sh の適用履歴。checksum で適用後の改変を検出する。';
" | psql_pipe || fail "適用履歴テーブルを用意できません"
        log_ok "適用履歴テーブルを作成しました"
    fi
    APPLIED_LIST=""
else
    APPLIED_LIST=$(psql_value "SELECT filename || ' ' || checksum FROM public.schema_migrations;") \
        || fail "適用履歴を取得できません"
    log_ok "適用履歴テーブルを確認"
fi

# --- 4. 差分の判定 -------------------------------------------
log_step "4/5 差分の判定"
PENDING=()
declare -a PENDING_SUM=()
for file in "${MIGRATIONS[@]}"; do
    name=$(basename "$file")
    sum=$(shasum -a 256 "$file" | awk '{print $1}')
    recorded=$(printf '%s\n' "$APPLIED_LIST" | awk -v n="$name" '$1==n {print $2}')

    if [[ -z "$recorded" ]]; then
        echo "  [未適用] ${name}"
        PENDING+=("$file")
        PENDING_SUM+=("$sum")
    elif [[ "$recorded" != "$sum" ]]; then
        # 適用済みファイルの中身が変わっている = 何が入っているか分からない状態。
        # 黙って再適用も無視もしない。必ず人が判断する。
        log_err "適用済みのマイグレーションが変更されています: ${name}"
        echo "    適用時: ${recorded}" >&2
        echo "    現在  : ${sum}" >&2
        echo >&2
        echo "  適用済みのファイルは変更しないでください。" >&2
        echo "  変更が必要な場合は、新しい番号のファイルを追加します。" >&2
        exit 1
    else
        echo "  [適用済] ${name}"
    fi
done

if [[ "$MODE" == "status" ]]; then
    echo
    echo "  未適用: ${#PENDING[@]} 件 (--status のため適用しません)"
    echo "  ログ: ${LOG_FILE}"
    exit 0
fi

if (( ${#PENDING[@]} == 0 )); then
    echo
    log_ok "未適用のマイグレーションはありません"
else
    echo
    if [[ "$ASSUME_YES" == false ]]; then
        if [[ -t 0 ]]; then
            read -r -p "  ${#PENDING[@]} 件を適用しますか? [y/N] " answer
            [[ "$answer" =~ ^[Yy]$ ]] || { echo "  中止しました。"; exit 0; }
        else
            fail "対話できない環境です。確認を省略する場合は --yes を付けてください。"
        fi
    fi

    for i in "${!PENDING[@]}"; do
        file="${PENDING[$i]}"
        sum="${PENDING_SUM[$i]}"
        name=$(basename "$file")
        log "適用中: ${name}"

        # SQL本体と履歴への記録を1つのトランザクションで流す(-1)。
        # 途中で失敗すれば両方とも入らない。「適用されたのに履歴が無い」
        # 状態を構造的に作らない。
        {
            cat "$file"
            printf '\nINSERT INTO public.schema_migrations (filename, checksum, applied_by) VALUES (%s, %s, %s);\n' \
                "'${name}'" "'${sum}'" "'$(whoami)@$(hostname -s)'"
        } | psql_pipe "-1" || fail "マイグレーションの適用に失敗しました: ${name}"

        log_ok "適用完了: ${name}"
    done
fi

# --- 5. 構造の確認 -------------------------------------------
#  「適用できた」ではなく「意図した構造になっている」ことを確認する。
log_step "5/5 構造の確認"

check_object() {
    local label="$1" query="$2" actual
    actual=$(psql_value "$query") || fail "確認クエリに失敗しました: ${label}"
    [[ "$actual" == "t" ]] || fail "${label} が存在しません"
    log_ok "${label}"
}

check_object "intelligence スキーマ" \
    "SELECT EXISTS(SELECT 1 FROM pg_namespace WHERE nspname='intelligence');"
check_object "intelligence.decisions" \
    "SELECT to_regclass('intelligence.decisions') IS NOT NULL;"
check_object "intelligence.outcomes" \
    "SELECT to_regclass('intelligence.outcomes') IS NOT NULL;"
check_object "intelligence.pending_reviews (view)" \
    "SELECT to_regclass('intelligence.pending_reviews') IS NOT NULL;"
check_object "追記専用トリガー (decisions/outcomes)" \
    "SELECT count(*)=2 FROM pg_trigger WHERE tgname IN ('decisions_append_only','outcomes_append_only') AND NOT tgisinternal;"
check_object "判断者を人に限る制約" \
    "SELECT EXISTS(SELECT 1 FROM pg_constraint WHERE conname='decisions_decided_by_is_human');"

# 追記専用が実際に効くかを、記録を残さずに確認する。
# INSERT → UPDATE が拒否されることを確認 → 例外でブロック全体を巻き戻す。
# ※ bigserial の採番は巻き戻らないため decision_id に欠番が出る(想定どおり)。
log "追記専用の動作確認(記録は残しません)"
SELFTEST=$(printf '%s\n' "
DO \$\$
DECLARE
    v_id bigint;
    v_denied boolean := false;
BEGIN
    INSERT INTO intelligence.decisions
        (decision_type, action, action_kind, reason, decided_by, source)
    VALUES
        ('DEC-SYS-02', 'Migrate.sh の自己診断', 'unchanged',
         '追記専用トリガーが機能しているかを確認するため', 'migrate-selftest', 'selftest')
    RETURNING decision_id INTO v_id;

    BEGIN
        UPDATE intelligence.decisions SET note = 'must fail' WHERE decision_id = v_id;
    EXCEPTION WHEN others THEN
        v_denied := true;
    END;

    IF NOT v_denied THEN
        RAISE EXCEPTION 'APPEND_ONLY_NOT_ENFORCED';
    END IF;

    -- ここまで来れば正常。診断用の行を残さないため、意図的に巻き戻す。
    RAISE EXCEPTION 'SELFTEST_ROLLBACK';
EXCEPTION WHEN others THEN
    IF SQLERRM = 'SELFTEST_ROLLBACK' THEN
        RAISE NOTICE 'APPEND_ONLY_OK';
    ELSE
        RAISE;
    END IF;
END
\$\$;
" | psql_pipe 2>&1) || fail "追記専用の動作確認に失敗しました: ${SELFTEST}"

if [[ "$SELFTEST" == *"APPEND_ONLY_OK"* ]]; then
    log_ok "追記専用トリガーが機能しています(UPDATE は拒否されました)"
else
    fail "追記専用の動作を確認できませんでした: ${SELFTEST:-出力なし}"
fi

DECISION_COUNT=$(psql_value "SELECT count(*) FROM intelligence.decisions;")
OUTCOME_COUNT=$(psql_value "SELECT count(*) FROM intelligence.outcomes;")
PENDING_COUNT=$(psql_value "SELECT count(*) FROM intelligence.pending_reviews;")

ELAPSED=$(fmt_elapsed)
echo
echo "════════════════════════════════════════════"
echo " SUCCESS"
echo "════════════════════════════════════════════"
echo "  Commit          : ${TARGET_SHORT}"
echo "  適用            : ${#PENDING[@]} 件"
echo "  判断の記録      : ${DECISION_COUNT} 件"
echo "  結果の記録      : ${OUTCOME_COUNT} 件"
echo "  結果未記録      : ${PENDING_COUNT} 件"
echo "  Elapsed Time    : ${ELAPSED}"
echo "  Log             : ${LOG_FILE}"
echo "════════════════════════════════════════════"
