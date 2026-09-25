-- =============================================================
--  30_identifier_roles.sql — Identifier(S6)と DB ロール(S5)
-- =============================================================
\set ON_ERROR_STOP 1

-- ------------------------------------------------------------
--  IDENTIFIER: コード登録簿(P3)とコード → 対象の候補(P1)
-- ------------------------------------------------------------
SELECT pc_test.lives('IDENTIFIER', '取込ロールがコードを登録できる',
    $q$INSERT INTO product_core.identifier (identifier_id, id_type, value, checkdigit_valid, first_seen_source, first_seen_ref, created_by) VALUES
       ('ID-10000001', 'JAN', '4900000000011', true, 'legacy_product_master', 'sha#1', 'ingest:t'),
       ('ID-10000002', 'JAN', '4900000000028', true, 'legacy_product_master', 'sha#2', 'ingest:t'),
       ('ID-10000003', 'JAN', '4900000000035', true, 'legacy_product_master', 'sha#3', 'ingest:t'),
       ('ID-10000004', 'JAN', '4900000000042', true, 'legacy_product_master', 'sha#4', 'ingest:t'),
       ('ID-10000005', 'JAN', '4900000000059', true, 'legacy_product_master', 'sha#5', 'ingest:t'),
       ('ID-10000006', 'JAN_PARTIAL4', '0011', NULL, 'set_sku', 'SET-0011-0028', 'ingest:t'),
       ('ID-10000007', 'ASIN', 'B0TESTA001', NULL, 'legacy_product_master', 'sha#7', 'ingest:t'),
       ('ID-10000008', 'GTIN14', '14900000000018', true, 'stocktake_scan', 'scan#8', 'ingest:t'),
       ('ID-10000009', 'JAN', '4900000000066', true, 'legacy_product_master', 'sha#9', 'ingest:t'),
       ('ID-10000010', 'TEMP_ID', 'T-000123', NULL, 'stocktake_scan', 'scan#10', 'ingest:t')$q$, 'pctest_ingest');
SELECT pc_test.throws('IDENTIFIER', 'JAN の形式を検査', $q$INSERT INTO product_core.identifier (id_type, value, first_seen_source, first_seen_ref, created_by)
    VALUES ('JAN', '12345', 's', 'r', 'ingest:t')$q$, 'pctest_ingest', '23514');
SELECT pc_test.throws('IDENTIFIER', '同じコードは1行(重複を拒否)', $q$INSERT INTO product_core.identifier (id_type, value, first_seen_source, first_seen_ref, created_by)
    VALUES ('JAN', '4900000000011', 's', 'r', 'ingest:t')$q$, 'pctest_ingest', '23505');
SELECT pc_test.throws('IDENTIFIER', 'P3: コードは変更できない', $q$UPDATE product_core.identifier SET value = '4900000000073' WHERE identifier_id = 'ID-10000001'$q$, NULL, '追記専用');
SELECT pc_test.throws('IDENTIFIER', 'P3: コードは削除できない', $q$DELETE FROM product_core.identifier WHERE identifier_id = 'ID-10000001'$q$, NULL, '追記専用');

SELECT pc_test.lives('IDENTIFIER', 'AI が紐付けの候補(PROPOSED)を作れる — 同じ JAN を単品とセットへ(共有 JAN)',
    $q$INSERT INTO product_core.identifier_link (link_id, identifier_id, entity_id, entity_type, link_basis, confidence, evidence, created_by) VALUES
       ('IL-10000001', 'ID-10000001', 'PP-100001', 'PHYSICAL_PRODUCT', 'LEGACY_MASTER', 'HIGH', '{}', 'rule:id/1.0'),
       ('IL-10000002', 'ID-10000002', 'PP-100002', 'PHYSICAL_PRODUCT', 'LEGACY_MASTER', 'HIGH', '{}', 'rule:id/1.0'),
       ('IL-10000003', 'ID-10000002', 'CP-100001', 'COMPOSITION', 'LEGACY_MASTER', 'MEDIUM', '{"shared_jan":true}', 'rule:id/1.0'),
       ('IL-10000004', 'ID-10000003', 'PP-100007', 'PHYSICAL_PRODUCT', 'LEGACY_MASTER', 'HIGH', '{}', 'rule:id/1.0'),
       ('IL-10000005', 'ID-10000004', 'PP-100003', 'PHYSICAL_PRODUCT', 'LEGACY_MASTER', 'HIGH', '{}', 'rule:id/1.0'),
       ('IL-10000006', 'ID-10000005', 'PP-100008', 'PHYSICAL_PRODUCT', 'LEGACY_MASTER', 'HIGH', '{}', 'rule:id/1.0'),
       ('IL-10000009', 'ID-10000009', 'PP-100002', 'PHYSICAL_PRODUCT', 'LEGACY_MASTER', 'HIGH', '{}', 'rule:id/1.0')$q$, 'pctest_ingest');
SELECT pc_test.throws('IDENTIFIER', 'S6: AI は自動確定の許可(AUTO_IF_UNIQUE)付きで提案できない',
    $q$INSERT INTO product_core.identifier_link (identifier_id, entity_id, entity_type, link_basis, scan_policy, confidence, evidence, created_by)
       VALUES ('ID-10000005', 'PP-100001', 'PHYSICAL_PRODUCT', 'LEGACY_MASTER', 'AUTO_IF_UNIQUE', 'HIGH', '{}', 'ai:x')$q$, 'pctest_ingest', '23514');
SELECT pc_test.throws('IDENTIFIER', 'S6: JAN 下4桁(部分識別子)はスキャンに使えない',
    $q$INSERT INTO product_core.identifier_link (identifier_id, entity_id, entity_type, link_basis, confidence, evidence, created_by)
       VALUES ('ID-10000006', 'PP-100001', 'PHYSICAL_PRODUCT', 'DERIVED_FROM_SKU', 'LOW', '{}', 'ai:x')$q$, 'pctest_ingest', 'NOT_FOR_SCAN');
SELECT pc_test.throws('IDENTIFIER', 'S6: JAN 欄の誤記(MISFILED)はスキャンに使えない',
    $q$INSERT INTO product_core.identifier_link (identifier_id, entity_id, entity_type, link_basis, confidence, evidence, created_by)
       VALUES ('ID-10000007', 'LS-100001', 'LISTING', 'MISFILED', 'LOW', '{}', 'ai:x')$q$, 'pctest_ingest', '23514');
SELECT pc_test.lives('IDENTIFIER', '部分識別子・誤記は NOT_FOR_SCAN なら記録できる',
    $q$INSERT INTO product_core.identifier_link (link_id, identifier_id, entity_id, entity_type, link_basis, scan_policy, confidence, evidence, created_by) VALUES
       ('IL-10000007', 'ID-10000006', 'PP-100001', 'PHYSICAL_PRODUCT', 'DERIVED_FROM_SKU', 'NOT_FOR_SCAN', 'LOW', '{}', 'ai:x'),
       ('IL-10000008', 'ID-10000007', 'LS-100001', 'LISTING', 'MISFILED', 'NOT_FOR_SCAN', 'LOW', '{}', 'ai:x')$q$, 'pctest_ingest');
SELECT pc_test.throws('IDENTIFIER', 'S5: AI は scan_policy を変えられない(列権限)',
    $q$UPDATE product_core.identifier_link SET scan_policy = 'AUTO_IF_UNIQUE' WHERE link_id = 'IL-10000001'$q$, 'pctest_ingest', '42501');
SELECT pc_test.throws('IDENTIFIER', 'C1: IDENTIFIER_APPROVE の無い人は承認できない',
    $q$UPDATE product_core.identifier_link SET record_status = 'ACTIVE', approved_by = 'human:op_rev', updated_by = 'human:op_rev' WHERE link_id = 'IL-10000001'$q$,
    'pctest_reviewer', 'IDENTIFIER_APPROVE');
SELECT pc_test.lives('IDENTIFIER', '人が承認し、自動確定の可否を決める', ARRAY[
    $q$UPDATE product_core.identifier_link SET record_status = 'ACTIVE', approved_by = 'human:op_admin', updated_by = 'human:op_admin', scan_policy = 'AUTO_IF_UNIQUE'
        WHERE link_id IN ('IL-10000001', 'IL-10000002', 'IL-10000003', 'IL-10000005')$q$,
    $q$UPDATE product_core.identifier_link SET record_status = 'ACTIVE', approved_by = 'human:op_admin', updated_by = 'human:op_admin'
        WHERE link_id = 'IL-10000004'$q$,
    $q$UPDATE product_core.identifier_link SET record_status = 'ACTIVE', approved_by = 'human:op_admin', updated_by = 'human:op_admin', scan_policy = 'AUTO_IF_UNIQUE'
        WHERE link_id = 'IL-10000009'$q$,
    $q$UPDATE product_core.identifier_link SET valid_to = '2026-01-31', updated_by = 'human:op_admin' WHERE link_id = 'IL-10000009'$q$], 'pctest_reviewer');
SELECT pc_test.throws('IDENTIFIER', 'C3: 紐付けの対象は変更できない',
    $q$UPDATE product_core.identifier_link SET entity_id = 'PP-100002' WHERE link_id = 'IL-10000001'$q$, NULL, '内容列');

-- スキャン解決(S6)
SELECT pc_test.check('RESOLVE', '一意・自動確定可・ACTIVE → UNIQUE_AUTO で確定',
    (SELECT resolution = 'UNIQUE_AUTO' AND confirmed_entity_id = 'PP-100001'
       FROM product_core.resolve_identifier('JAN', '4900000000011', '2026-09-26')));
SELECT pc_test.check('RESOLVE', 'S6: 単品とセットで共有の JAN → NEEDS_SELECTION(自動確定しない)',
    (SELECT count(*) = 2 AND bool_and(resolution = 'NEEDS_SELECTION') AND bool_and(confirmed_entity_id IS NULL)
       FROM product_core.resolve_identifier('JAN', '4900000000028', '2026-09-26')));
SELECT pc_test.check('RESOLVE', 'S6: 候補1件でも CONFIRM_ALWAYS → NEEDS_SELECTION',
    (SELECT resolution = 'NEEDS_SELECTION' AND confirmed_entity_id IS NULL
       FROM product_core.resolve_identifier('JAN', '4900000000035', '2026-09-26')));
SELECT pc_test.check('RESOLVE', 'S6: 対象が PROVISIONAL → NEEDS_SELECTION',
    (SELECT resolution = 'NEEDS_SELECTION' AND confirmed_entity_id IS NULL
       FROM product_core.resolve_identifier('JAN', '4900000000042', '2026-09-26')));
SELECT pc_test.check('RESOLVE', 'S6: 未承認(PROPOSED)の紐付けだけ → NO_ACTIVE_LINK',
    (SELECT resolution = 'NO_ACTIVE_LINK' AND confirmed_entity_id IS NULL
       FROM product_core.resolve_identifier('JAN', '4900000000059', '2026-09-26')));
SELECT pc_test.check('RESOLVE', 'S6: 誤記の ASIN はスキャンで使われない → NO_ACTIVE_LINK',
    (SELECT resolution = 'NO_ACTIVE_LINK' FROM product_core.resolve_identifier('ASIN', 'B0TESTA001', '2026-09-26')));
SELECT pc_test.check('RESOLVE', '未登録のケースコード(紐付け無し)→ NO_ACTIVE_LINK',
    (SELECT resolution = 'NO_ACTIVE_LINK' FROM product_core.resolve_identifier('GTIN14', '14900000000018', '2026-09-26')));
SELECT pc_test.check('RESOLVE', '登録の無いコード → NOT_FOUND',
    (SELECT resolution = 'NOT_FOUND' FROM product_core.resolve_identifier('JAN', '4999999999999', '2026-09-26')));
SELECT pc_test.check('RESOLVE', '有効期間内は確定、期間後(旧 JAN)は NO_ACTIVE_LINK',
    (SELECT resolution = 'UNIQUE_AUTO' FROM product_core.resolve_identifier('JAN', '4900000000066', '2026-01-15'))
    AND (SELECT resolution = 'NO_ACTIVE_LINK' FROM product_core.resolve_identifier('JAN', '4900000000066', '2026-09-26')));

-- ------------------------------------------------------------
--  ROLES: DB ロールの権限(S5)
-- ------------------------------------------------------------
SELECT pc_test.lives('ROLES', '参照ロールは読める', $q$SELECT count(*) FROM product_core.physical_product$q$, 'pctest_reader');
SELECT pc_test.throws('ROLES', '参照ロールは書けない', $q$INSERT INTO product_core.identifier (id_type, value, first_seen_source, first_seen_ref, created_by)
    VALUES ('JAN', '4900000000073', 's', 'r', 'ingest:t')$q$, 'pctest_reader', '42501');
SELECT pc_test.throws('ROLES', '権限の無いロールはスキーマに触れられない', $q$SELECT count(*) FROM product_core.physical_product$q$, 'pctest_nobody', '42501');
-- SET ROLE の可否はログイン中のユーザーで決まるため、切り替えの拒否は 40_ai_session.sql(AI と同じ条件の別セッション)で検証する
SELECT pc_test.check('ROLES', 'S5: 取込ロールは承認ロールのメンバーではない',
    NOT pg_has_role('pc_ingest', 'pc_reviewer', 'MEMBER') AND NOT pg_has_role('pctest_ai_login', 'pc_reviewer', 'MEMBER'));
SELECT pc_test.throws('ROLES', 'S5: 取込ロールは権限付与を書けない',
    $q$INSERT INTO product_core.operator_permission (operator_id, permission, granted_by) VALUES ('op_rev', 'COST_APPROVE', 'human:op_admin')$q$, 'pctest_ingest', '42501');
SELECT pc_test.throws('ROLES', 'S5: 取込ロールは正式原価(CP)を作れない',
    $q$INSERT INTO product_core.composition_cost_history (cp_id, set_cost_excl_tax, purchase_evidence, valid_from, source_observation_id, normalization, basis, assessment_id, approved_by)
       VALUES ('CP-100001', 1, '{}', '2030-01-01', 'CO-100000016', '{}', 'x', 'AS-100000004', 'human:op_cost')$q$, 'pctest_ingest', '42501');
SELECT pc_test.check('ROLES', 'S5: pc_ingest は承認・review・状態の列に書き込み権限を持たない',
    NOT has_column_privilege('pc_ingest', 'product_core.legacy_mapping', 'approved_by', 'INSERT')
    AND NOT has_column_privilege('pc_ingest', 'product_core.legacy_mapping', 'approved_by', 'UPDATE')
    AND NOT has_column_privilege('pc_ingest', 'product_core.legacy_mapping', 'record_status', 'INSERT')
    AND NOT has_column_privilege('pc_ingest', 'product_core.listing_reference_history', 'record_status', 'UPDATE')
    AND NOT has_column_privilege('pc_ingest', 'product_core.identifier_link', 'approved_by', 'INSERT')
    AND NOT has_column_privilege('pc_ingest', 'product_core.product_relationship', 'approved_by', 'INSERT')
    AND NOT has_column_privilege('pc_ingest', 'product_core.assessment', 'review_status', 'INSERT')
    AND NOT has_column_privilege('pc_ingest', 'product_core.assessment', 'reviewed_by', 'UPDATE')
    AND NOT has_column_privilege('pc_ingest', 'product_core.assessment', 'content_hash', 'INSERT')
    AND NOT has_column_privilege('pc_ingest', 'product_core.physical_product', 'status', 'INSERT')
    AND NOT has_column_privilege('pc_ingest', 'product_core.physical_product', 'activated_by', 'UPDATE')
    AND NOT has_column_privilege('pc_ingest', 'product_core.physical_product', 'tax_category', 'UPDATE')
    AND NOT has_column_privilege('pc_ingest', 'product_core.composition', 'activated_by', 'UPDATE'));
SELECT pc_test.check('ROLES', 'S5: pc_ingest は正式原価・operator・core_entity に書き込み権限を持たない',
    NOT has_table_privilege('pc_ingest', 'product_core.cost_history', 'INSERT')
    AND NOT has_table_privilege('pc_ingest', 'product_core.composition_cost_history', 'INSERT')
    AND NOT has_table_privilege('pc_ingest', 'product_core.operator', 'INSERT')
    AND NOT has_table_privilege('pc_ingest', 'product_core.operator_permission', 'INSERT')
    AND NOT has_table_privilege('pc_ingest', 'product_core.core_entity', 'INSERT'));
SELECT pc_test.check('ROLES', 'pc_reader はどの表にも書き込み権限を持たない',
    NOT EXISTS (SELECT 1 FROM information_schema.tables t
                 WHERE t.table_schema IN ('product_core', 'legacy_ingest') AND t.table_type = 'BASE TABLE'
                   AND (has_table_privilege('pc_reader', format('%I.%I', t.table_schema, t.table_name), 'INSERT')
                        OR has_table_privilege('pc_reader', format('%I.%I', t.table_schema, t.table_name), 'UPDATE')
                        OR has_table_privilege('pc_reader', format('%I.%I', t.table_schema, t.table_name), 'DELETE'))));
SELECT pc_test.check('ROLES', 'どのロールにも表の DELETE 権限が無い(構成品の下書きを除く)',
    NOT EXISTS (SELECT 1 FROM information_schema.tables t CROSS JOIN unnest(ARRAY['pc_ingest', 'pc_reviewer', 'pc_reader']) r
                 WHERE t.table_schema IN ('product_core', 'legacy_ingest') AND t.table_type = 'BASE TABLE'
                   AND t.table_name <> 'composition_component'
                   AND has_table_privilege(r, format('%I.%I', t.table_schema, t.table_name), 'DELETE')));
SELECT pc_test.check('ROLES', 'Product Core のロールはスーパーユーザーでもロール作成権限持ちでもない',
    NOT EXISTS (SELECT 1 FROM pg_roles WHERE rolname LIKE 'pc\_%' AND (rolsuper OR rolcreaterole OR rolcreatedb OR rolcanlogin)));
SELECT pc_test.check('SAFETY', 'D1: DB にファイル書き出し・外部接続の拡張が無い(書き戻し経路なし)',
    NOT EXISTS (SELECT 1 FROM pg_extension WHERE extname IN ('dblink', 'postgres_fdw', 'file_fdw', 'adminpack', 'plpython3u', 'plperlu')));
SELECT pc_test.check('SAFETY', 'D1: Product Core の関数にファイル書き出し・COPY が無い',
    NOT EXISTS (SELECT 1 FROM pg_proc p JOIN pg_namespace n ON n.oid = p.pronamespace
                 WHERE n.nspname IN ('product_core', 'legacy_ingest')
                   AND (p.prosrc ~* '\mcopy\M' OR p.prosrc ~* 'lo_export' OR p.prosrc ~* 'pg_file_write' OR p.prosrc ~* 'dblink')));
