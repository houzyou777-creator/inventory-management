-- =============================================================
--  00_harness.sql — Product Core 制約テストの土台(使い捨て DB 専用)
-- =============================================================
--  ⚠️ このファイルと同じフォルダのテストは、run_product_core_tests.sh が起動する
--     隔離された使い捨て PostgreSQL でだけ流す。本番 DB では絶対に流さない
--     (テスト用のロールと行を作るため)。db/migrations/ には置かない。
--
--  仕組み:
--    pc_test.lives(...)  … 例外が出なければ PASS(正常系)
--    pc_test.throws(...) … 例外が出れば PASS(異常系)。期待するメッセージ / SQLSTATE も照合できる
--    どちらもサブトランザクションで実行し、指定したロールに SET LOCAL ROLE して流す。
-- =============================================================
\set ON_ERROR_STOP 1

-- 本番ではないことの確認(テストスクリプトが作る DB 名だけを許す)
DO $$
BEGIN
    IF current_database() <> 'pctest' THEN
        RAISE EXCEPTION 'テストは使い捨て DB(pctest)でだけ実行できます。現在: %', current_database();
    END IF;
END;
$$;

CREATE SCHEMA pc_test;
CREATE TABLE pc_test.result (
    seq    bigserial PRIMARY KEY,
    area   text      NOT NULL,
    name   text      NOT NULL,
    kind   text      NOT NULL,   -- LIVES / THROWS / CHECK
    ok     boolean   NOT NULL,
    detail text
);
GRANT USAGE ON SCHEMA pc_test TO PUBLIC;
GRANT SELECT, INSERT ON pc_test.result TO PUBLIC;
GRANT USAGE ON SEQUENCE pc_test.result_seq_seq TO PUBLIC;

-- テスト用のログインしないロール(グループロールのメンバー)
CREATE ROLE pctest_ingest   NOLOGIN INHERIT IN ROLE pc_ingest;
CREATE ROLE pctest_reviewer NOLOGIN INHERIT IN ROLE pc_reviewer;
CREATE ROLE pctest_reader   NOLOGIN INHERIT IN ROLE pc_reader;
CREATE ROLE pctest_nobody   NOLOGIN;
-- AI のプロセスと同じ条件(スーパーユーザーではないログインユーザーが pc_ingest のメンバー)で
-- 別セッションとして接続して検証するためのユーザー。パスワードは持たない(ネットワークの無い
-- 使い捨てコンテナ内の Unix ソケット接続のみ)
CREATE ROLE pctest_ai_login LOGIN NOSUPERUSER NOCREATEROLE INHERIT IN ROLE pc_ingest;

CREATE FUNCTION pc_test.record(p_area text, p_name text, p_kind text, p_ok boolean, p_detail text)
RETURNS void LANGUAGE sql AS $$
    INSERT INTO pc_test.result (area, name, kind, ok, detail) VALUES (p_area, p_name, p_kind, p_ok, p_detail)
$$;

-- 正常系: 例外なく実行できること
CREATE FUNCTION pc_test.lives(p_area text, p_name text, p_sql text[], p_role text DEFAULT NULL)
RETURNS void LANGUAGE plpgsql AS $$
DECLARE v_stmt text; v_err text; v_state text;
BEGIN
    BEGIN
        IF p_role IS NOT NULL THEN EXECUTE format('SET LOCAL ROLE %I', p_role); END IF;
        FOREACH v_stmt IN ARRAY p_sql LOOP
            EXECUTE v_stmt;
        END LOOP;
        SET CONSTRAINTS ALL IMMEDIATE;
        SET CONSTRAINTS ALL DEFERRED;
        RESET ROLE;
    EXCEPTION WHEN OTHERS THEN
        GET STACKED DIAGNOSTICS v_err = MESSAGE_TEXT, v_state = RETURNED_SQLSTATE;
        PERFORM pc_test.record(p_area, p_name, 'LIVES', false, format('予期しない例外 [%s] %s', v_state, v_err));
        RETURN;
    END;
    PERFORM pc_test.record(p_area, p_name, 'LIVES', true, NULL);
END;
$$;

-- 異常系: 例外で拒否されること(p_expect は SQLSTATE か、メッセージに含まれる語)
CREATE FUNCTION pc_test.throws(p_area text, p_name text, p_sql text[], p_role text DEFAULT NULL,
                               p_expect text DEFAULT NULL)
RETURNS void LANGUAGE plpgsql AS $$
DECLARE v_stmt text; v_err text; v_state text;
BEGIN
    BEGIN
        IF p_role IS NOT NULL THEN EXECUTE format('SET LOCAL ROLE %I', p_role); END IF;
        FOREACH v_stmt IN ARRAY p_sql LOOP
            EXECUTE v_stmt;
        END LOOP;
        -- 遅延制約(COMMIT 時の検査)もここで発火させる
        SET CONSTRAINTS ALL IMMEDIATE;
        SET CONSTRAINTS ALL DEFERRED;
        RAISE EXCEPTION USING ERRCODE = 'P0T01', MESSAGE = '__NO_ERROR__';
    EXCEPTION WHEN OTHERS THEN
        GET STACKED DIAGNOSTICS v_err = MESSAGE_TEXT, v_state = RETURNED_SQLSTATE;
    END;
    IF v_state = 'P0T01' THEN
        PERFORM pc_test.record(p_area, p_name, 'THROWS', false, '拒否されるべき操作が通った');
    ELSIF p_expect IS NOT NULL AND v_state <> p_expect AND position(p_expect IN v_err) = 0 THEN
        PERFORM pc_test.record(p_area, p_name, 'THROWS', false,
                               format('拒否はされたが理由が違う [%s] %s(期待: %s)', v_state, v_err, p_expect));
    ELSE
        PERFORM pc_test.record(p_area, p_name, 'THROWS', true, format('[%s] %s', v_state, v_err));
    END IF;
END;
$$;

-- 値の確認
CREATE FUNCTION pc_test.check(p_area text, p_name text, p_cond boolean, p_detail text DEFAULT NULL)
RETURNS void LANGUAGE sql AS $$
    SELECT pc_test.record(p_area, p_name, 'CHECK', coalesce(p_cond, false), p_detail)
$$;

-- 単一文の短縮形
CREATE FUNCTION pc_test.lives(p_area text, p_name text, p_sql text, p_role text DEFAULT NULL)
RETURNS void LANGUAGE sql AS $$ SELECT pc_test.lives(p_area, p_name, ARRAY[p_sql], p_role) $$;
CREATE FUNCTION pc_test.throws(p_area text, p_name text, p_sql text, p_role text DEFAULT NULL, p_expect text DEFAULT NULL)
RETURNS void LANGUAGE sql AS $$ SELECT pc_test.throws(p_area, p_name, ARRAY[p_sql], p_role, p_expect) $$;
