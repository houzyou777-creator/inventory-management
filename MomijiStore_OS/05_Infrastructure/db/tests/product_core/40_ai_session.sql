-- =============================================================
--  40_ai_session.sql — AI のプロセスと同じ条件での検証(S5)
-- =============================================================
--  このファイルだけは、スーパーユーザーではなく pctest_ai_login(pc_ingest のメンバー)で
--  別セッションとして接続して流す。SET ROLE の可否はログイン中のユーザーで決まるため、
--  スーパーユーザーのまま SET LOCAL ROLE する他のテストでは、この点を検証できない。
\set ON_ERROR_STOP 1

SELECT pc_test.check('AI_SESSION', 'このセッションは AI 用のログインユーザーで、スーパーユーザーではない',
    session_user = 'pctest_ai_login' AND NOT (SELECT rolsuper FROM pg_roles WHERE rolname = session_user));
SELECT pc_test.throws('AI_SESSION', 'S5: AI は承認ロール(pc_reviewer)に切り替えられない', $q$SET ROLE pc_reviewer$q$, NULL, '42501');
SELECT pc_test.throws('AI_SESSION', 'S5: AI は承認用のログインロールに切り替えられない', $q$SET ROLE pctest_reviewer$q$, NULL, '42501');
SELECT pc_test.throws('AI_SESSION', 'S5: AI はスーパーユーザーに切り替えられない', $q$SET ROLE pctest_admin$q$, NULL, '42501');
SELECT pc_test.throws('AI_SESSION', 'S5: AI が human: を名乗って review を承認することはできない',
    $q$UPDATE product_core.assessment SET review_status = 'APPROVED', reviewed_by = 'human:op_admin' WHERE assessment_id = 'AS-100000003'$q$, NULL, '42501');
SELECT pc_test.throws('AI_SESSION', 'S5: AI が human: を名乗って対応付けを ACTIVE で作ることはできない',
    $q$INSERT INTO product_core.legacy_mapping (legacy_p, entity_id, entity_type, mapping_role, legacy_snapshot_ref, created_by, record_status, approved_by)
       VALUES ('P000077', 'PP-100002', 'PHYSICAL_PRODUCT', 'SOLE', 'x', 'human:op_admin', 'ACTIVE', 'human:op_admin')$q$, NULL, '42501');
SELECT pc_test.throws('AI_SESSION', 'S5: AI が human: を名乗って PP を ACTIVE 化することはできない',
    $q$UPDATE product_core.physical_product SET status = 'ACTIVE', activated_by = 'human:op_prod' WHERE pp_id = 'PP-100003'$q$, NULL, '42501');
SELECT pc_test.throws('AI_SESSION', 'S5: AI が human: を名乗って正式原価を作ることはできない',
    $q$INSERT INTO product_core.cost_history (pp_id, unit_cost_excl_tax, valid_from, source_observation_id, normalization, basis, assessment_id, approved_by)
       VALUES ('PP-100001', 1000, '2031-01-01', 'CO-100000001', '{}', 'x', 'AS-100000004', 'human:op_cost')$q$, NULL, '42501');
SELECT pc_test.throws('AI_SESSION', 'S5: AI が human: を名乗って権限を付与することはできない',
    $q$INSERT INTO product_core.operator_permission (operator_id, permission, granted_by) VALUES ('op_rev', 'OPERATOR_ADMIN', 'human:op_admin')$q$, NULL, '42501');
SELECT pc_test.lives('AI_SESSION', 'AI は判定(assessment)を UNREVIEWED で記録できる(正常系)',
    $q$INSERT INTO product_core.assessment (assessment_type, subject_key, verdict, confidence, reason, evidence, rule_version, assessed_by)
       VALUES ('LEGACY_MAPPING', 'P000077', 'NEEDS_REVIEW', 'LOW', 'AI セッションからの記録', '{}', 'legacy-map/1.0', 'ai:test/1')$q$);
SELECT pc_test.check('AI_SESSION', 'AI が記録した判定は UNREVIEWED のまま',
    (SELECT review_status = 'UNREVIEWED' FROM product_core.assessment WHERE subject_key = 'P000077' AND assessed_by = 'ai:test/1'));
SELECT pc_test.throws('AI_SESSION', '1.6: AI は review_batch(人の確認表の束)を作れない',
    $q$INSERT INTO product_core.review_batch (review_type, file_name, created_by, updated_by) VALUES ('LEGACY_LISTING', 'x.xlsx', 'human:op_rev', 'human:op_rev')$q$, NULL, '42501');
SELECT pc_test.throws('AI_SESSION', '1.6: AI は保留理由・review_batch_id(人の判断)を書けない',
    $q$UPDATE product_core.assessment SET review_status = 'ON_HOLD', reviewed_by = 'human:op_rev', review_note = 'x', review_reason_code = 'HOLD',
           review_channel = 'EXCEL', review_batch_id = 'RB-00000001' WHERE assessment_id = 'AS-100000003'$q$, NULL, '42501');
SELECT pc_test.throws('AI_SESSION', '1.4: AI は人の修正コード(HUMAN_CORRECTED)を登録できない',
    $q$INSERT INTO product_core.identifier (id_type, value, first_seen_source, first_seen_ref, created_by, origin) VALUES ('JAN', '4900000000998', 's', 'r', 'human:op_admin', 'HUMAN_CORRECTED')$q$, NULL, '42501');
SELECT pc_test.throws('AI_SESSION', '1.1: AI が human: を名乗って C番号の対応を ACTIVE で作ることはできない',
    $q$INSERT INTO product_core.legacy_listing_mapping (legacy_listing_id, listing_id, match_basis, legacy_snapshot_ref, created_by, record_status, approved_by)
       VALUES ('C000077', 'LS-100002', 'MANUAL', 'x', 'human:op_admin', 'ACTIVE', 'human:op_admin')$q$, NULL, '42501');
SELECT pc_test.throws('AI_SESSION', '1.3: AI は PP の ID を指定できない',
    $q$INSERT INTO product_core.physical_product (pp_id, name, origin_key, created_by) VALUES ('PP-100077', 'x', 'k:ai77', 'ai:x')$q$, NULL, '42501');
