"""intelligence — Intelligence Layer(知識層)の記録と参照

Architecture.md「6. Intelligence Layer」の実装。
判断(何を・なぜ)と結果(どうなったか)を PostgreSQL へ追記し、参照する。

【この層だけが書き込みを許される理由】
MCPサーバーは Excel に対しては今後も読み取り専用である。
ここで書き込むのは intelligence スキーマの2表だけで、これらは
Excelに対応物を持たない「DBで生まれるデータ」である。
正本が二重化しないため BL-5 に抵触しない。

【追記専用】
UPDATE / DELETE は行わない。DB側のトリガーでも拒否される。
訂正は supersedes に旧IDを入れた新しい行で行う(BL-7)。

スキーマ定義: 05_Infrastructure/db/migrations/001_intelligence_layer.sql
"""

from __future__ import annotations

import logging
import os
import re
from datetime import date, datetime
from decimal import Decimal
from typing import Any

logger = logging.getLogger("momiji-mcp.intelligence")

# --- 語彙 -------------------------------------------------------------------
#  DBのCHECK制約と同じ値をここにも置く。DBだけに任せるとエラーが
#  「制約違反」としか返らず、利用者が何を指定すべきか分からないため。
ACTION_KINDS = ("changed", "unchanged", "rejected")
ASSESSMENTS = ("success", "partial", "failure", "unclear")

# Decision_Catalog.md の Decision ID(Architecture.md Naming Convention)
DECISION_TYPE_RE = re.compile(r"^DEC-[A-Z]{3}-\d{2}$")
# Business Logic 番号(MomijiStore_OS_Logical_Design_v1.0.md)
BUSINESS_LOGIC_RE = re.compile(r"^BL-\d{1,3}$")
# 判断者にAIを指定させない(AI Constitution 第1条)。DBのCHECKと同じ判定。
AI_PREFIX_RE = re.compile(r"^ai[:_-]", re.IGNORECASE)

DEFAULT_LIMIT = int(os.getenv("MOMIJI_INTEL_DEFAULT_LIMIT", "50"))
MAX_LIMIT = int(os.getenv("MOMIJI_INTEL_MAX_LIMIT", "500"))

# 接続待ちの上限。NASのDBが落ちている時に呼び出しが固まらないようにする。
CONNECT_TIMEOUT_SEC = int(os.getenv("MOMIJI_DB_CONNECT_TIMEOUT", "5"))


class IntelligenceError(Exception):
    """利用者へそのまま返せる内容のエラー。"""


# --- 入力検証 ---------------------------------------------------------------
#  DBへ渡す前に落とす。理由や判断者が空のまま貯まると、記録はあるのに
#  学習に使えないという最悪の状態になる。


def _required_text(value: Any, field: str) -> str:
    if value is None or not str(value).strip():
        raise IntelligenceError(f"{field} は必須です(空文字は不可)。")
    return str(value).strip()


def _optional_text(value: Any) -> str | None:
    if value is None:
        return None
    text = str(value).strip()
    return text or None


def _validate_decision_type(value: Any) -> str:
    text = _required_text(value, "decision_type")
    if not DECISION_TYPE_RE.match(text):
        raise IntelligenceError(
            f"decision_type の形式が不正です: 「{text}」。"
            "DEC-<領域3文字>-<連番2桁> の形で、Decision_Catalog.md に"
            "登録済みのIDを指定してください(例: DEC-SAL-01)。"
            "該当する判断の型が無い場合は、先にカタログへ追加します。"
        )
    return text


def _validate_choice(value: Any, allowed: tuple[str, ...], field: str) -> str:
    text = _required_text(value, field)
    if text not in allowed:
        raise IntelligenceError(
            f"{field} は {' / '.join(allowed)} のいずれかです(指定値: 「{text}」)。"
        )
    return text


def _validate_decided_by(value: Any) -> str:
    text = _required_text(value, "decided_by")
    if AI_PREFIX_RE.match(text):
        raise IntelligenceError(
            "decided_by にAIを指定することはできません。"
            "決めるのは人です(FOUNDATION.md AI Constitution 第1条)。"
            "AIが提案した場合は proposed_by に 'ai:claude' のように記録し、"
            "decided_by には決めた人の名前を入れてください。"
        )
    return text


def _validate_business_logic(value: Any) -> list[str] | None:
    if value is None:
        return None
    if isinstance(value, str):
        items = [v.strip() for v in value.split(",")]
    elif isinstance(value, (list, tuple)):
        items = [str(v).strip() for v in value]
    else:
        raise IntelligenceError(
            "business_logic は 'BL-3,BL-4' のような文字列か、リストで指定してください。"
        )
    items = [i for i in items if i]
    for item in items:
        if not BUSINESS_LOGIC_RE.match(item):
            raise IntelligenceError(
                f"business_logic の形式が不正です: 「{item}」。BL-3 のように指定してください。"
            )
    return items or None


def _validate_date(value: Any, field: str) -> date | None:
    if value is None or (isinstance(value, str) and not value.strip()):
        return None
    if isinstance(value, date) and not isinstance(value, datetime):
        return value
    try:
        return date.fromisoformat(str(value).strip())
    except ValueError as exc:
        raise IntelligenceError(
            f"{field} は YYYY-MM-DD 形式で指定してください(指定値: 「{value}」)。"
        ) from exc


def _validate_timestamp(value: Any, field: str) -> datetime | None:
    if value is None or (isinstance(value, str) and not value.strip()):
        return None
    if isinstance(value, datetime):
        return value
    text = str(value).strip()
    try:
        # 日付だけの指定も許す(その日の 00:00 として扱う)
        if len(text) == 10:
            return datetime.fromisoformat(text)
        return datetime.fromisoformat(text.replace("Z", "+00:00"))
    except ValueError as exc:
        raise IntelligenceError(
            f"{field} は YYYY-MM-DD または ISO8601 形式で指定してください"
            f"(指定値: 「{value}」)。"
        ) from exc


def _validate_number(value: Any, field: str) -> float | None:
    if value is None or (isinstance(value, str) and not value.strip()):
        return None
    try:
        return float(value)
    except (TypeError, ValueError) as exc:
        raise IntelligenceError(f"{field} は数値で指定してください(指定値: 「{value}」)。") from exc


def _effective_limit(limit: Any) -> int:
    try:
        value = int(limit)
    except (TypeError, ValueError):
        value = DEFAULT_LIMIT
    return max(1, min(value, MAX_LIMIT))


# --- 応答の整形 -------------------------------------------------------------
def _jsonable(value: Any) -> Any:
    """MCPの応答はJSONになるため、DB由来の型を素の型へ落とす。"""
    if isinstance(value, Decimal):
        return float(value)
    if isinstance(value, datetime):
        return value.isoformat()
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, (list, tuple)):
        return [_jsonable(v) for v in value]
    if isinstance(value, dict):
        return {k: _jsonable(v) for k, v in value.items()}
    return value


def _rows(cursor) -> list[dict[str, Any]]:
    columns = [c.name for c in cursor.description]
    return [_jsonable(dict(zip(columns, row))) for row in cursor.fetchall()]


# --- 接続 -------------------------------------------------------------------
#  psycopg は import を遅延させる。コンテナには入っているが(Dockerfileの
#  build時チェックで担保)、検証スクリプトを psycopg 無しの環境で
#  動かせるようにしておく。
def _conninfo() -> dict[str, Any]:
    password = os.environ.get("MOMIJI_DB_PASSWORD")
    if not password:
        raise IntelligenceError(
            "MOMIJI_DB_PASSWORD が設定されていません。DBへ接続できません。"
        )
    return {
        "host": os.environ.get("MOMIJI_DB_HOST", "momiji-postgres"),
        "port": int(os.environ.get("MOMIJI_DB_PORT", "5432")),
        "dbname": os.environ.get("MOMIJI_DB_NAME", "momiji"),
        "user": os.environ.get("MOMIJI_DB_USER", "momiji"),
        "password": password,
        "connect_timeout": CONNECT_TIMEOUT_SEC,
    }


def _connect():
    """1呼び出し1接続。判断の記録は1日数件で、プールは要らない。

    プール(psycopg_pool)は依存が増えるため入れない。
    件数が増えて実測で問題になった時に、Decision Log へ記録して導入する。
    """
    try:
        import psycopg
    except ImportError as exc:  # pragma: no cover — コンテナには必ず入っている
        raise IntelligenceError(
            "psycopg が入っていません。requirements.lock を確認してください。"
        ) from exc

    try:
        return psycopg.connect(**_conninfo())
    except Exception as exc:
        # 例外を握り潰さない。実際の型とメッセージを必ず残す。
        # 接続情報にパスワードが含まれるため、host/port/dbname だけを出す。
        info = _conninfo()
        logger.exception("DBへ接続できません host=%s port=%s db=%s",
                         info["host"], info["port"], info["dbname"])
        raise IntelligenceError(
            f"DBへ接続できません({type(exc).__name__}: {exc})。"
            f" 接続先={info['host']}:{info['port']}/{info['dbname']}"
        ) from exc


# --- 記録 -------------------------------------------------------------------
def record_decision(
    decision_type: str,
    action: str,
    action_kind: str,
    reason: str,
    decided_by: str,
    subject_type: str | None = None,
    subject_id: str | None = None,
    subject_label: str | None = None,
    alternatives: str | None = None,
    expected: str | None = None,
    expected_metric: str | None = None,
    expected_value: float | None = None,
    review_due: str | None = None,
    decided_at: str | None = None,
    proposed_by: str | None = None,
    business_logic: Any = None,
    supersedes: int | None = None,
    note: str | None = None,
) -> dict[str, Any]:
    """判断を1件記録する。返り値に decision_id を含む。"""
    values = {
        "decision_type": _validate_decision_type(decision_type),
        "action": _required_text(action, "action"),
        "action_kind": _validate_choice(action_kind, ACTION_KINDS, "action_kind"),
        "reason": _required_text(reason, "reason"),
        "decided_by": _validate_decided_by(decided_by),
        "subject_type": _optional_text(subject_type),
        "subject_id": _optional_text(subject_id),
        "subject_label": _optional_text(subject_label),
        "alternatives": _optional_text(alternatives),
        "expected": _optional_text(expected),
        "expected_metric": _optional_text(expected_metric),
        "expected_value": _validate_number(expected_value, "expected_value"),
        "review_due": _validate_date(review_due, "review_due"),
        "proposed_by": _optional_text(proposed_by),
        "business_logic": _validate_business_logic(business_logic),
        "supersedes": int(supersedes) if supersedes is not None else None,
        "note": _optional_text(note),
    }
    decided = _validate_timestamp(decided_at, "decided_at")

    columns = list(values.keys())
    params: list[Any] = [values[c] for c in columns]
    if decided is not None:
        columns.append("decided_at")
        params.append(decided)

    placeholders = ", ".join(["%s"] * len(columns))
    sql = (
        f"INSERT INTO intelligence.decisions ({', '.join(columns)}) "
        f"VALUES ({placeholders}) "
        "RETURNING decision_id, decided_at, review_due"
    )

    with _connect() as conn, conn.cursor() as cur:
        cur.execute(sql, params)
        row = _rows(cur)[0]

    logger.info(
        "判断を記録 decision_id=%s type=%s kind=%s subject=%s",
        row["decision_id"], values["decision_type"],
        values["action_kind"], values["subject_id"],
    )
    result = {
        "recorded": True,
        "decision_id": row["decision_id"],
        "decided_at": row["decided_at"],
        "review_due": row["review_due"],
    }
    if row["review_due"] is None:
        # 期限が無い判断は pending_reviews に出ても放置されやすい。
        # 「やりっぱなし禁止」を守るため、記録時点で気付けるようにする。
        result["warning"] = (
            "review_due が未設定です。結果をいつ確認するかを決めないと、"
            "この判断は学習素材になりません。"
        )
    return result


def record_outcome(
    decision_id: int,
    assessment: str,
    summary: str,
    measured_by: str,
    metric: str | None = None,
    actual_value: float | None = None,
    period_start: str | None = None,
    period_end: str | None = None,
    learning: str | None = None,
    measured_at: str | None = None,
    note: str | None = None,
) -> dict[str, Any]:
    """判断の結果を1件記録する。同じ判断に複数回記録してよい。"""
    try:
        target_id = int(decision_id)
    except (TypeError, ValueError) as exc:
        raise IntelligenceError("decision_id は数値で指定してください。") from exc

    values = {
        "decision_id": target_id,
        "assessment": _validate_choice(assessment, ASSESSMENTS, "assessment"),
        "summary": _required_text(summary, "summary"),
        "measured_by": _required_text(measured_by, "measured_by"),
        "metric": _optional_text(metric),
        "actual_value": _validate_number(actual_value, "actual_value"),
        "period_start": _validate_date(period_start, "period_start"),
        "period_end": _validate_date(period_end, "period_end"),
        "learning": _optional_text(learning),
        "note": _optional_text(note),
    }
    measured = _validate_timestamp(measured_at, "measured_at")

    columns = list(values.keys())
    params: list[Any] = [values[c] for c in columns]
    if measured is not None:
        columns.append("measured_at")
        params.append(measured)

    placeholders = ", ".join(["%s"] * len(columns))
    sql = (
        f"INSERT INTO intelligence.outcomes ({', '.join(columns)}) "
        f"VALUES ({placeholders}) "
        "RETURNING outcome_id, measured_at"
    )

    with _connect() as conn, conn.cursor() as cur:
        # 存在しない判断に結果だけ付くと、後から辿れない記録になる。
        # 外部キーでも弾けるが、利用者に分かるメッセージで返したい。
        cur.execute(
            "SELECT decision_type, action, expected, expected_metric, expected_value"
            "  FROM intelligence.decisions WHERE decision_id = %s",
            (target_id,),
        )
        decision = _rows(cur)
        if not decision:
            raise IntelligenceError(
                f"decision_id={target_id} の判断が見つかりません。"
                "先に record_decision で判断を記録してください。"
            )
        cur.execute(sql, params)
        row = _rows(cur)[0]

    logger.info(
        "結果を記録 outcome_id=%s decision_id=%s assessment=%s",
        row["outcome_id"], target_id, values["assessment"],
    )
    return {
        "recorded": True,
        "outcome_id": row["outcome_id"],
        "decision_id": target_id,
        "measured_at": row["measured_at"],
        # 「何に対する結果か」をその場で返し、取り違えに気付けるようにする
        "decision": decision[0],
    }


# --- 参照 -------------------------------------------------------------------
def search_decisions(
    decision_type: str | None = None,
    subject_id: str | None = None,
    action_kind: str | None = None,
    decided_by: str | None = None,
    since: str | None = None,
    until: str | None = None,
    keyword: str | None = None,
    with_outcomes: bool = True,
    limit: int = DEFAULT_LIMIT,
) -> dict[str, Any]:
    """判断を検索する。条件はAND。結果(outcomes)も併せて返す。"""
    where: list[str] = []
    params: list[Any] = []

    if decision_type:
        where.append("d.decision_type = %s")
        params.append(_validate_decision_type(decision_type))
    if subject_id:
        where.append("d.subject_id = %s")
        params.append(str(subject_id).strip())
    if action_kind:
        where.append("d.action_kind = %s")
        params.append(_validate_choice(action_kind, ACTION_KINDS, "action_kind"))
    if decided_by:
        where.append("d.decided_by = %s")
        params.append(str(decided_by).strip())

    since_ts = _validate_timestamp(since, "since")
    if since_ts is not None:
        where.append("d.decided_at >= %s")
        params.append(since_ts)
    until_ts = _validate_timestamp(until, "until")
    if until_ts is not None:
        where.append("d.decided_at <= %s")
        params.append(until_ts)

    if keyword:
        # 施策・理由・対象名を横断する。過去の似た判断を探すための入口。
        where.append(
            "(d.action ILIKE %s OR d.reason ILIKE %s"
            " OR COALESCE(d.subject_label,'') ILIKE %s)"
        )
        pattern = f"%{str(keyword).strip()}%"
        params.extend([pattern, pattern, pattern])

    effective_limit = _effective_limit(limit)
    clause = f"WHERE {' AND '.join(where)}" if where else ""

    with _connect() as conn, conn.cursor() as cur:
        cur.execute(
            f"SELECT COUNT(*) AS matched FROM intelligence.decisions d {clause}",
            params,
        )
        matched = _rows(cur)[0]["matched"]

        cur.execute(
            "SELECT d.* FROM intelligence.decisions d "
            f"{clause} ORDER BY d.decided_at DESC, d.decision_id DESC LIMIT %s",
            [*params, effective_limit],
        )
        decisions = _rows(cur)

        if with_outcomes and decisions:
            ids = [d["decision_id"] for d in decisions]
            cur.execute(
                "SELECT * FROM intelligence.outcomes"
                " WHERE decision_id = ANY(%s) ORDER BY measured_at",
                (ids,),
            )
            by_decision: dict[int, list[dict[str, Any]]] = {}
            for outcome in _rows(cur):
                by_decision.setdefault(outcome["decision_id"], []).append(outcome)
            for item in decisions:
                item["outcomes"] = by_decision.get(item["decision_id"], [])

    logger.info(
        "search_decisions 該当=%d件 返却=%d件 条件数=%d",
        matched, len(decisions), len(where),
    )
    return {
        "results": decisions,
        "count": len(decisions),
        "matched": matched,
        "truncated": matched > len(decisions),
    }


def list_pending_reviews(
    as_of: str | None = None,
    only_due: bool = True,
    limit: int = DEFAULT_LIMIT,
) -> dict[str, Any]:
    """結果が未記録の判断を返す。「やりっぱなし」を可視化する。"""
    as_of_date = _validate_date(as_of, "as_of")
    where: list[str] = []
    params: list[Any] = []
    if only_due:
        # 期限切れに加え、期限が未設定のものも「放置されうる」ため含める。
        where.append("(review_due IS NULL OR review_due <= %s)")
        params.append(as_of_date or date.today())

    clause = f"WHERE {' AND '.join(where)}" if where else ""
    effective_limit = _effective_limit(limit)

    with _connect() as conn, conn.cursor() as cur:
        cur.execute(
            f"SELECT COUNT(*) AS matched FROM intelligence.pending_reviews {clause}",
            params,
        )
        matched = _rows(cur)[0]["matched"]
        cur.execute(
            f"SELECT * FROM intelligence.pending_reviews {clause}"
            " ORDER BY review_due NULLS LAST, decided_at LIMIT %s",
            [*params, effective_limit],
        )
        rows = _rows(cur)

    logger.info("list_pending_reviews 該当=%d件 返却=%d件", matched, len(rows))
    return {
        "results": rows,
        "count": len(rows),
        "matched": matched,
        "truncated": matched > len(rows),
        "note": (
            "結果が未記録の判断です。ここが溜まるほど、AIは自分の提案が"
            "正しかったかを学習できません(Architecture.md §6 やりっぱなし禁止)。"
        ),
    }
