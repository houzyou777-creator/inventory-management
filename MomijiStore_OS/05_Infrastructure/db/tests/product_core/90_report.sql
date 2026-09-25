-- =============================================================
--  90_report.sql — 結果の集計。1件でも FAIL があれば異常終了させる
-- =============================================================
\set ON_ERROR_STOP 1
\pset pager off

\echo '---- 区分別 ----'
SELECT area AS "区分", count(*) AS "件数", count(*) FILTER (WHERE ok) AS "PASS", count(*) FILTER (WHERE NOT ok) AS "FAIL"
  FROM pc_test.result GROUP BY area ORDER BY min(seq);

\echo '---- 種類別(正常系 LIVES / 異常系 THROWS / 値の確認 CHECK)----'
SELECT kind AS "種類", count(*) AS "件数", count(*) FILTER (WHERE ok) AS "PASS", count(*) FILTER (WHERE NOT ok) AS "FAIL"
  FROM pc_test.result GROUP BY kind ORDER BY kind;

\echo '---- FAIL 一覧 ----'
SELECT seq, area, name, kind, detail FROM pc_test.result WHERE NOT ok ORDER BY seq;

SELECT format('RESULT total=%s pass=%s fail=%s', count(*), count(*) FILTER (WHERE ok), count(*) FILTER (WHERE NOT ok)) AS summary
  FROM pc_test.result;

DO $$
DECLARE v_fail integer;
BEGIN
    SELECT count(*) INTO v_fail FROM pc_test.result WHERE NOT ok;
    IF v_fail > 0 THEN
        RAISE EXCEPTION 'FAIL が % 件あります', v_fail;
    END IF;
    IF (SELECT count(*) FROM pc_test.result) < 100 THEN
        RAISE EXCEPTION 'テスト件数が想定より少ない(途中で止まった可能性)';
    END IF;
END;
$$;
