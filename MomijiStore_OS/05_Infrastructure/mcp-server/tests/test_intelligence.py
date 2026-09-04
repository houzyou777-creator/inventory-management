"""intelligence モジュールの検証。

DBを立てずに動かせる範囲を確認する。
  - 入力検証(不正な記録がDBへ届く前に落ちるか)
  - 組み立てたSQLとパラメータ
  - DB由来の型がJSONにできる形へ落ちるか

DB本体の制約(CHECK・追記専用トリガー・ビュー)は
deploy/Migrate.sh の「5/5 構造の確認」がNAS上で検証する。
役割を分けているので、両方を通して初めて確認完了とする。

実行: python3 -m unittest discover -s tests -v
"""

from __future__ import annotations

import os
import sys
import unittest
from datetime import date, datetime
from decimal import Decimal

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

import intelligence  # noqa: E402
from intelligence import IntelligenceError  # noqa: E402


# --- DBの代わり -------------------------------------------------------------
class FakeColumn:
    def __init__(self, name: str) -> None:
        self.name = name


class FakeCursor:
    """execute された内容を記録し、あらかじめ渡した結果を返す。"""

    def __init__(self, results: list) -> None:
        self._results = list(results)
        self.executed: list[tuple] = []
        self.description = None
        self._rows: list = []

    def execute(self, sql, params=None):
        self.executed.append((" ".join(sql.split()), params))
        columns, rows = self._results.pop(0) if self._results else ([], [])
        self.description = [FakeColumn(c) for c in columns]
        self._rows = rows

    def fetchall(self):
        return self._rows

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        return False


class FakeConnection:
    def __init__(self, cursor: FakeCursor) -> None:
        self._cursor = cursor

    def cursor(self):
        return self._cursor

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        return False


class DBTestCase(unittest.TestCase):
    """_connect を差し替えて、SQLの組み立てだけを確認する土台。"""

    def use_results(self, results: list) -> FakeCursor:
        cursor = FakeCursor(results)
        self._original = intelligence._connect
        intelligence._connect = lambda: FakeConnection(cursor)
        self.addCleanup(lambda: setattr(intelligence, "_connect", self._original))
        return cursor


# --- 入力検証 ---------------------------------------------------------------
class TestValidation(unittest.TestCase):
    def test_decision_type_must_match_catalog_format(self):
        for bad in ["SAL-01", "DEC-SALES-01", "DEC-SAL-1", "dec-sal-01", ""]:
            with self.subTest(bad=bad):
                with self.assertRaises(IntelligenceError):
                    intelligence._validate_decision_type(bad)
        self.assertEqual(intelligence._validate_decision_type(" DEC-SAL-01 "), "DEC-SAL-01")

    def test_action_kind_is_limited_to_three_values(self):
        for good in intelligence.ACTION_KINDS:
            self.assertEqual(
                intelligence._validate_choice(good, intelligence.ACTION_KINDS, "action_kind"),
                good,
            )
        with self.assertRaises(IntelligenceError):
            intelligence._validate_choice("skipped", intelligence.ACTION_KINDS, "action_kind")

    def test_ai_cannot_be_the_decider(self):
        """AI Constitution 第1条を、記録の入口でも守る。"""
        for bad in ["ai:claude", "AI-claude", "ai_gpt", "AI:Claude"]:
            with self.subTest(bad=bad):
                with self.assertRaises(IntelligenceError) as ctx:
                    intelligence._validate_decided_by(bad)
                self.assertIn("決めるのは人", str(ctx.exception))
        # 人の名前は通る。'ai' で始まる人名を巻き込まないことも確認する。
        self.assertEqual(intelligence._validate_decided_by("蛯名"), "蛯名")
        self.assertEqual(intelligence._validate_decided_by("aiko"), "aiko")

    def test_reason_cannot_be_blank(self):
        with self.assertRaises(IntelligenceError):
            intelligence._required_text("   ", "reason")

    def test_business_logic_accepts_string_or_list(self):
        self.assertEqual(intelligence._validate_business_logic("BL-3,BL-11"), ["BL-3", "BL-11"])
        self.assertEqual(intelligence._validate_business_logic(["BL-4"]), ["BL-4"])
        self.assertIsNone(intelligence._validate_business_logic(None))
        self.assertIsNone(intelligence._validate_business_logic(""))
        with self.assertRaises(IntelligenceError):
            intelligence._validate_business_logic("BL3")

    def test_dates_must_be_iso(self):
        self.assertEqual(intelligence._validate_date("2026-09-30", "review_due"), date(2026, 9, 30))
        self.assertIsNone(intelligence._validate_date(None, "review_due"))
        with self.assertRaises(IntelligenceError):
            intelligence._validate_date("2026/09/30", "review_due")

    def test_limit_is_clamped(self):
        self.assertEqual(intelligence._effective_limit(0), 1)
        self.assertEqual(intelligence._effective_limit(10), 10)
        self.assertEqual(intelligence._effective_limit(99999), intelligence.MAX_LIMIT)
        self.assertEqual(intelligence._effective_limit("abc"), intelligence.DEFAULT_LIMIT)


class TestJsonable(unittest.TestCase):
    def test_db_types_become_json_safe(self):
        value = intelligence._jsonable(
            {
                "amount": Decimal("1980.50"),
                "at": datetime(2026, 9, 5, 7, 30),
                "due": date(2026, 9, 30),
                "bl": ["BL-3"],
            }
        )
        self.assertEqual(value["amount"], 1980.5)
        self.assertEqual(value["at"], "2026-09-05T07:30:00")
        self.assertEqual(value["due"], "2026-09-30")
        self.assertEqual(value["bl"], ["BL-3"])


# --- 記録 -------------------------------------------------------------------
VALID_DECISION = {
    "decision_type": "DEC-SAL-01",
    "action": "1,980円へ値下げした",
    "action_kind": "changed",
    "reason": "カート喪失。配送優位では維持できず、利益率15%を確保できる下限まで下げた",
    "decided_by": "蛯名",
}


class TestRecordDecision(DBTestCase):
    def test_insert_contains_required_columns(self):
        cursor = self.use_results(
            [(["decision_id", "decided_at", "review_due"], [(1, datetime(2026, 9, 5), date(2026, 10, 5))])]
        )
        result = intelligence.record_decision(
            **VALID_DECISION,
            subject_id="P000021",
            business_logic="BL-3,BL-10",
            review_due="2026-10-05",
            expected_value=15,
        )
        self.assertTrue(result["recorded"])
        self.assertEqual(result["decision_id"], 1)
        self.assertNotIn("warning", result)

        sql, params = cursor.executed[0]
        self.assertIn("INSERT INTO intelligence.decisions", sql)
        self.assertIn("RETURNING decision_id", sql)
        # プレースホルダの数と値の数が一致していること(ズレると別列へ入る)
        self.assertEqual(sql.count("%s"), len(params))
        self.assertIn("P000021", params)
        self.assertIn(["BL-3", "BL-10"], params)
        self.assertIn(date(2026, 10, 5), params)

    def test_missing_review_due_returns_warning(self):
        """期限の無い判断は放置されやすい。記録時点で気付けること。"""
        self.use_results([(["decision_id", "decided_at", "review_due"], [(2, datetime(2026, 9, 5), None)])])
        result = intelligence.record_decision(**VALID_DECISION)
        self.assertIn("warning", result)
        self.assertIn("学習素材", result["warning"])

    def test_unchanged_is_recordable(self):
        """BL-11「変更しない」も記録できなければならない。"""
        cursor = self.use_results(
            [(["decision_id", "decided_at", "review_due"], [(3, datetime(2026, 9, 5), date(2026, 10, 5))])]
        )
        intelligence.record_decision(
            decision_type="DEC-SAL-01",
            action="価格を据え置いた",
            action_kind="unchanged",
            reason="配送優位でカートを維持できているため、値下げの必要がない",
            decided_by="蛯名",
            review_due="2026-10-05",
        )
        _, params = cursor.executed[0]
        self.assertIn("unchanged", params)

    def test_invalid_input_never_reaches_db(self):
        cursor = self.use_results([])
        bad = dict(VALID_DECISION, decided_by="ai:claude")
        with self.assertRaises(IntelligenceError):
            intelligence.record_decision(**bad)
        self.assertEqual(cursor.executed, [], "検証前にDBへ問い合わせてはならない")

    def test_decided_at_is_appended_only_when_given(self):
        cursor = self.use_results(
            [(["decision_id", "decided_at", "review_due"], [(4, datetime(2026, 9, 1), None)])]
        )
        intelligence.record_decision(**VALID_DECISION, decided_at="2026-09-01")
        sql, params = cursor.executed[0]
        self.assertIn("decided_at", sql)
        self.assertEqual(sql.count("%s"), len(params))


class TestRecordOutcome(DBTestCase):
    def test_outcome_requires_existing_decision(self):
        cursor = self.use_results([([], [])])  # 判断が見つからない
        with self.assertRaises(IntelligenceError) as ctx:
            intelligence.record_outcome(
                decision_id=99, assessment="success", summary="売れた", measured_by="蛯名"
            )
        self.assertIn("見つかりません", str(ctx.exception))
        self.assertEqual(len(cursor.executed), 1, "存在確認だけで止まること")

    def test_outcome_insert_and_echo_back_decision(self):
        cursor = self.use_results(
            [
                (
                    ["decision_type", "action", "expected", "expected_metric", "expected_value"],
                    [("DEC-SAL-01", "1,980円へ値下げした", "カート復帰", "利益額", Decimal("15"))],
                ),
                (["outcome_id", "measured_at"], [(7, datetime(2026, 10, 5))]),
            ]
        )
        result = intelligence.record_outcome(
            decision_id=1,
            assessment="partial",
            summary="カートは戻ったが利益率は13%",
            measured_by="蛯名",
            metric="利益率",
            actual_value=13,
            period_start="2026-09-05",
            period_end="2026-10-05",
            learning="配送優位を先に確認していれば値下げ幅を抑えられた",
        )
        self.assertTrue(result["recorded"])
        self.assertEqual(result["outcome_id"], 7)
        # 何に対する結果かをその場で返し、取り違えに気付けること
        self.assertEqual(result["decision"]["decision_type"], "DEC-SAL-01")

        sql, params = cursor.executed[1]
        self.assertIn("INSERT INTO intelligence.outcomes", sql)
        self.assertEqual(sql.count("%s"), len(params))
        self.assertIn("partial", params)

    def test_assessment_is_validated(self):
        cursor = self.use_results([])
        with self.assertRaises(IntelligenceError):
            intelligence.record_outcome(
                decision_id=1, assessment="good", summary="s", measured_by="蛯名"
            )
        self.assertEqual(cursor.executed, [])


# --- 参照 -------------------------------------------------------------------
class TestSearchDecisions(DBTestCase):
    def test_no_conditions_returns_recent_records(self):
        cursor = self.use_results(
            [
                (["matched"], [(2,)]),
                (["decision_id"], [(1,), (2,)]),
                (["outcome_id", "decision_id"], [(9, 1)]),
            ]
        )
        result = intelligence.search_decisions()
        self.assertEqual(result["matched"], 2)
        self.assertEqual(result["count"], 2)
        self.assertFalse(result["truncated"])
        # 結果は該当する判断にだけ付く
        self.assertEqual(len(result["results"][0]["outcomes"]), 1)
        self.assertEqual(result["results"][1]["outcomes"], [])
        self.assertNotIn("WHERE", cursor.executed[0][0])

    def test_conditions_become_and_clauses(self):
        cursor = self.use_results(
            [(["matched"], [(0,)]), (["decision_id"], [])]
        )
        intelligence.search_decisions(
            decision_type="DEC-PUR-02",
            subject_id="P000021",
            action_kind="rejected",
            since="2026-09-01",
            keyword="値崩れ",
        )
        count_sql, count_params = cursor.executed[0]
        self.assertIn("WHERE", count_sql)
        self.assertEqual(count_sql.count("%s"), len(count_params))
        # keyword は3列を横断するため %s が3つ増える
        self.assertEqual(count_params.count("%値崩れ%"), 3)

        list_sql, list_params = cursor.executed[1]
        self.assertIn("ORDER BY d.decided_at DESC", list_sql)
        self.assertEqual(list_params[-1], intelligence.DEFAULT_LIMIT)

    def test_invalid_condition_is_rejected_before_query(self):
        cursor = self.use_results([])
        with self.assertRaises(IntelligenceError):
            intelligence.search_decisions(decision_type="DEC-XX-1")
        self.assertEqual(cursor.executed, [])

    def test_with_outcomes_false_skips_second_query(self):
        cursor = self.use_results(
            [(["matched"], [(1,)]), (["decision_id"], [(1,)])]
        )
        result = intelligence.search_decisions(with_outcomes=False)
        self.assertEqual(len(cursor.executed), 2)
        self.assertNotIn("outcomes", result["results"][0])


class TestPendingReviews(DBTestCase):
    def test_only_due_filters_by_date(self):
        cursor = self.use_results(
            [(["matched"], [(1,)]), (["decision_id", "review_due"], [(1, date(2026, 8, 1))])]
        )
        result = intelligence.list_pending_reviews(as_of="2026-09-05")
        self.assertEqual(result["matched"], 1)
        count_sql, count_params = cursor.executed[0]
        self.assertIn("review_due IS NULL OR review_due <= %s", count_sql)
        self.assertEqual(count_params, [date(2026, 9, 5)])

    def test_all_pending_when_only_due_false(self):
        cursor = self.use_results(
            [(["matched"], [(3,)]), (["decision_id"], [(1,), (2,), (3,)])]
        )
        intelligence.list_pending_reviews(only_due=False)
        self.assertNotIn("WHERE", cursor.executed[0][0])


if __name__ == "__main__":
    unittest.main(verbosity=2)
