-- =============================================================
--  50_revision.sql — Phase 1 DDL Revision の制約テスト(A〜F)
-- =============================================================
--  A. legacy_listing_mapping(C番号 → 出品)
--  B. PER_LEGACY_P(既存 P あたりの原価は単独で正式原価にならない)
--  C. DB 採番(取込ロールは ID を指定できない・RETURNING・再実行で重複しない)
--  D. GTIN12 / LEADING_ZERO_SUSPECT(12桁を元のまま・13桁は人の修正後だけ)
--  E. 取込の冪等性(ファイル SHA256 + レコード・ロールバック後の再実行・同じ業務キーの別版)
--  F. Human Review(review_batch_id・review_reason_code の追跡・AI は人の判断を書けない)
--  10〜30 のテストで作った仮データ(PP-1000xx・LS-10000x・operator)を前提にする。
\set ON_ERROR_STOP 1

-- ------------------------------------------------------------
--  A. legacy_listing_mapping
-- ------------------------------------------------------------
SELECT pc_test.lives('LEGACY_LISTING', 'AI が C番号 → 出品の候補(PROPOSED)を作れる(複数の C番号 → 同じ出品を含む)',
    $q$INSERT INTO product_core.legacy_listing_mapping (legacy_listing_id, listing_id, match_basis, observed_legacy_p, legacy_snapshot_ref, created_by, ingest_run_id) VALUES
       ('C000001', 'LS-100001', 'SAME_NATURAL_KEY', 'P000001', 'sha#c1', 'rule:llm/1.0', 'IR-20260926-000002-abcd'),
       ('C000011', 'LS-100001', 'DUPLICATE_LEGACY_ROW', 'P000001', 'sha#c11', 'rule:llm/1.0', 'IR-20260926-000002-abcd'),
       ('C000003', 'LS-100003', 'SAME_NATURAL_KEY', 'P000002', 'sha#c3', 'rule:llm/1.0', 'IR-20260926-000002-abcd')$q$, 'pctest_ingest');
SELECT pc_test.check('LEGACY_LISTING', 'ID は DB が採番(LLM- + 8桁)',
    (SELECT bool_and(llm_id ~ '^LLM-[0-9]{8}$') AND count(*) = 3 FROM product_core.legacy_listing_mapping));
SELECT pc_test.throws('LEGACY_LISTING', 'AI は ACTIVE の対応を作れない(列権限)',
    $q$INSERT INTO product_core.legacy_listing_mapping (legacy_listing_id, listing_id, match_basis, legacy_snapshot_ref, created_by, record_status, approved_by)
       VALUES ('C000099', 'LS-100002', 'MANUAL', 'x', 'ai:x', 'ACTIVE', 'human:op_admin')$q$, 'pctest_ingest', '42501');
SELECT pc_test.throws('LEGACY_LISTING', 'AI は ID(llm_id)を指定できない(列権限)',
    $q$INSERT INTO product_core.legacy_listing_mapping (llm_id, legacy_listing_id, listing_id, match_basis, legacy_snapshot_ref, created_by)
       VALUES ('LLM-99999999', 'C000099', 'LS-100002', 'MANUAL', 'x', 'ai:x')$q$, 'pctest_ingest', '42501');
SELECT pc_test.throws('LEGACY_LISTING', 'AI は候補を ACTIVE 化できない(列権限)',
    $q$UPDATE product_core.legacy_listing_mapping SET record_status = 'ACTIVE', approved_by = 'human:op_admin' WHERE legacy_listing_id = 'C000001'$q$,
    'pctest_ingest', '42501');
SELECT pc_test.throws('LEGACY_LISTING', '同じ C番号 → 同じ出品の候補は重ねて作られない(再取込で増えない)',
    $q$INSERT INTO product_core.legacy_listing_mapping (legacy_listing_id, listing_id, match_basis, legacy_snapshot_ref, created_by)
       VALUES ('C000001', 'LS-100001', 'SAME_NATURAL_KEY', 'sha#c1', 'rule:llm/1.0')$q$, 'pctest_ingest', 'legacy_listing_mapping_open_uq');
SELECT pc_test.throws('LEGACY_LISTING', 'C番号の形式を検査',
    $q$INSERT INTO product_core.legacy_listing_mapping (legacy_listing_id, listing_id, match_basis, legacy_snapshot_ref, created_by)
       VALUES ('LS-100001', 'LS-100001', 'MANUAL', 'x', 'ai:x')$q$, 'pctest_ingest', '23514');
SELECT pc_test.throws('LEGACY_LISTING', '存在しない出品へは対応付けできない(孤児を作らない)',
    $q$INSERT INTO product_core.legacy_listing_mapping (legacy_listing_id, listing_id, match_basis, legacy_snapshot_ref, created_by)
       VALUES ('C000098', 'LS-999999', 'MANUAL', 'x', 'ai:x')$q$, 'pctest_ingest', '23503');
SELECT pc_test.throws('LEGACY_LISTING', '権限(LISTING_REFERENCE_APPROVE)の無い人は ACTIVE 化できない',
    $q$UPDATE product_core.legacy_listing_mapping SET record_status = 'ACTIVE', approved_by = 'human:op_rev', updated_by = 'human:op_rev' WHERE legacy_listing_id = 'C000001'$q$,
    'pctest_reviewer', 'LISTING_REFERENCE_APPROVE');
SELECT pc_test.check('LEGACY_LISTING', '未承認の対応は正式処理用ビューに出ない',
    NOT EXISTS (SELECT 1 FROM product_core.v_active_legacy_listing_mapping));
SELECT pc_test.lives('LEGACY_LISTING', '権限者が ACTIVE 化できる',
    $q$UPDATE product_core.legacy_listing_mapping SET record_status = 'ACTIVE', approved_by = 'human:op_admin', updated_by = 'human:op_admin'
        WHERE record_status = 'PROPOSED'$q$, 'pctest_reviewer');
SELECT pc_test.check('LEGACY_LISTING', '複数の C番号 → 同じ出品(C000001・C000011 → LS-100001)を許可',
    (SELECT array_agg(legacy_listing_id ORDER BY legacy_listing_id) FROM product_core.v_active_legacy_listing_mapping WHERE listing_id = 'LS-100001')
    = ARRAY['C000001', 'C000011']);
SELECT pc_test.throws('LEGACY_LISTING', '1 C番号 → 1 出品: 同じ C番号に2つ目の ACTIVE は作れない',
    $q$INSERT INTO product_core.legacy_listing_mapping (legacy_listing_id, listing_id, match_basis, legacy_snapshot_ref, created_by, record_status, approved_by)
       VALUES ('C000001', 'LS-100002', 'MANUAL', 'x', 'human:op_admin', 'ACTIVE', 'human:op_admin')$q$, 'pctest_reviewer', 'legacy_listing_mapping_active_uq');
SELECT pc_test.lives('LEGACY_LISTING', '付け替え: 新しい候補 → 旧 ACTIVE を SUPERSEDED → 新を ACTIVE', ARRAY[
    $q$SET LOCAL ROLE pctest_ingest$q$,
    $q$INSERT INTO product_core.legacy_listing_mapping (legacy_listing_id, listing_id, match_basis, legacy_snapshot_ref, created_by)
       VALUES ('C000003', 'LS-100002', 'MANUAL', 'sha#c3-v2', 'ai:x')$q$,
    $q$SET LOCAL ROLE pctest_reviewer$q$,
    $q$UPDATE product_core.legacy_listing_mapping o SET record_status = 'SUPERSEDED', updated_by = 'human:op_admin',
           superseded_by = (SELECT n.llm_id FROM product_core.legacy_listing_mapping n WHERE n.legacy_listing_id = 'C000003' AND n.record_status = 'PROPOSED')
        WHERE o.legacy_listing_id = 'C000003' AND o.record_status = 'ACTIVE'$q$,
    $q$UPDATE product_core.legacy_listing_mapping SET record_status = 'ACTIVE', approved_by = 'human:op_admin', updated_by = 'human:op_admin'
        WHERE legacy_listing_id = 'C000003' AND record_status = 'PROPOSED'$q$]);
SELECT pc_test.check('LEGACY_LISTING', 'SUPERSEDED 後も旧い対応(C000003 → LS-100003)の履歴が残り、付け替え先を指す',
    (SELECT count(*) FROM product_core.legacy_listing_mapping WHERE legacy_listing_id = 'C000003') = 2
    AND EXISTS (SELECT 1 FROM product_core.legacy_listing_mapping o JOIN product_core.legacy_listing_mapping n ON n.llm_id = o.superseded_by
                 WHERE o.legacy_listing_id = 'C000003' AND o.listing_id = 'LS-100003' AND o.record_status = 'SUPERSEDED'
                   AND n.listing_id = 'LS-100002' AND n.record_status = 'ACTIVE'));
SELECT pc_test.check('LEGACY_LISTING', '旧 C番号から出品を逆引きできる(C000003 → 現在は LS-100002)',
    (SELECT listing_id FROM product_core.v_active_legacy_listing_mapping WHERE legacy_listing_id = 'C000003') = 'LS-100002');
SELECT pc_test.throws('LEGACY_LISTING', '別の C番号の行を付け替え先にできない',
    $q$UPDATE product_core.legacy_listing_mapping SET record_status = 'SUPERSEDED', updated_by = 'human:op_admin',
           superseded_by = (SELECT llm_id FROM product_core.legacy_listing_mapping WHERE legacy_listing_id = 'C000011')
        WHERE legacy_listing_id = 'C000001' AND record_status = 'ACTIVE'$q$, 'pctest_reviewer', '同じ C番号');
SELECT pc_test.throws('LEGACY_LISTING', '対応は削除できない', $q$DELETE FROM product_core.legacy_listing_mapping WHERE legacy_listing_id = 'C000001'$q$, NULL, '削除できません');
SELECT pc_test.throws('LEGACY_LISTING', '対応先の出品は削除できない(孤児にならない)',
    $q$DELETE FROM product_core.listing WHERE listing_id = 'LS-100001'$q$, NULL, '削除できません');
SELECT pc_test.throws('LEGACY_LISTING', '対応先(内容列)は変更できない',
    $q$UPDATE product_core.legacy_listing_mapping SET listing_id = 'LS-100002' WHERE legacy_listing_id = 'C000001'$q$, NULL, '内容列');

-- ------------------------------------------------------------
--  B. PER_LEGACY_P
-- ------------------------------------------------------------
SELECT pc_test.lives('PER_LEGACY_P', '既存 P あたりの原価(PER_LEGACY_P・単位 UNKNOWN)を元の値のまま保存できる',
    $q$INSERT INTO product_core.cost_observation (source, source_ref, source_reliability, observed_at, observed_amount, tax_inclusion, tax_inclusion_basis,
            tax_rate_basis, amount_unit, subject_legacy_p, source_unit_basis, source_record_key, source_record_hash, evidence, created_by, ingest_run_id)
       VALUES ('legacy_product_master', 'master#P000021', 'MEDIUM', '2026-09-15', 500, 'EXCLUDED', 'SOURCE_SPEC', 'NONE', 'UNKNOWN',
               'P000021', 'PER_LEGACY_P', 'legacy_p=P000021', repeat('b', 64), '{"column":"標準原価"}', 'ingest:t', 'IR-20260926-000002-abcd')$q$, 'pctest_ingest');
SELECT pc_test.check('PER_LEGACY_P', '保存した値と単位の根拠がそのまま残る',
    (SELECT observed_amount = 500 AND source_unit_basis = 'PER_LEGACY_P' AND amount_unit = 'UNKNOWN'
       FROM product_core.cost_observation WHERE source_ref = 'master#P000021'));
SELECT pc_test.throws('PER_LEGACY_P', 'PER_LEGACY_P を PP 単価(PER_PP)と名乗れない',
    $q$INSERT INTO product_core.cost_observation (source, source_ref, source_reliability, observed_at, observed_amount, tax_inclusion, tax_inclusion_basis,
            tax_rate_basis, amount_unit, subject_legacy_p, source_unit_basis, source_record_key, source_record_hash, evidence, created_by, ingest_run_id)
       VALUES ('legacy_product_master', 'm#x1', 'MEDIUM', now(), 500, 'EXCLUDED', 'SOURCE_SPEC', 'NONE', 'PER_PP',
               'P000022', 'PER_LEGACY_P', 'legacy_p=P000022', repeat('c', 64), '{}', 'ingest:t', 'IR-20260926-000002-abcd')$q$, 'pctest_ingest', '23514');
SELECT pc_test.throws('PER_LEGACY_P', '既存 P に付いた値を PER_PP の単位根拠で記録できない(P が単品か未確定)',
    $q$INSERT INTO product_core.cost_observation (source, source_ref, source_reliability, observed_at, observed_amount, tax_inclusion, tax_inclusion_basis,
            tax_rate_basis, amount_unit, subject_legacy_p, source_unit_basis, source_record_key, source_record_hash, evidence, created_by, ingest_run_id)
       VALUES ('legacy_product_master', 'm#x2', 'MEDIUM', now(), 500, 'EXCLUDED', 'SOURCE_SPEC', 'NONE', 'PER_PP',
               'P000022', 'PER_PP', 'legacy_p=P000022', repeat('d', 64), '{}', 'ingest:t', 'IR-20260926-000002-abcd')$q$, 'pctest_ingest', '23514');
SELECT pc_test.throws('PER_LEGACY_P', 'PER_LEGACY_P は既存 P を対象にした値だけ',
    $q$INSERT INTO product_core.cost_observation (source, source_ref, source_reliability, observed_at, observed_amount, tax_inclusion, tax_inclusion_basis,
            tax_rate_basis, amount_unit, subject_entity_id, subject_entity_type, source_unit_basis, source_record_key, source_record_hash, evidence, created_by, ingest_run_id)
       VALUES ('legacy_product_master', 'm#x3', 'MEDIUM', now(), 500, 'EXCLUDED', 'SOURCE_SPEC', 'NONE', 'UNKNOWN',
               'PP-100006', 'PHYSICAL_PRODUCT', 'PER_LEGACY_P', 'legacy_p=P000023', repeat('e', 64), '{}', 'ingest:t', 'IR-20260926-000002-abcd')$q$, 'pctest_ingest', '23514');

-- 判定の準備: LEGACY_MAPPING(P000021 → PP-100006)と UNIT_BASIS を AI が記録する(未 review)
SELECT pc_test.lives('PER_LEGACY_P', '(準備)AI が対応付けと単位の判定を記録する', ARRAY[
    $q$SELECT setval('product_core.as_seq', 100000300)$q$,
    $q$SET LOCAL ROLE pctest_ingest$q$,
    $q$INSERT INTO product_core.assessment (assessment_type, subject_key, verdict, proposed_entity_id, proposed_entity_type, confidence, reason, evidence, rule_version, assessed_by) VALUES
       ('LEGACY_MAPPING', 'P000021', 'PHYSICAL_PRODUCT_CANDIDATE', 'PP-100006', 'PHYSICAL_PRODUCT', 'HIGH', 'JAN一意・単品', '{}', 'legacy-map/1.0', 'rule:legacy-map/1.0')$q$,
    $q$INSERT INTO product_core.assessment (assessment_type, subject_key, verdict, confidence, reason, evidence, rule_version, assessed_by)
       SELECT 'UNIT_BASIS', observation_id, 'PER_PP', 'MEDIUM', '単品の P', '{}', 'unit/1.0', 'ai:unit/1.0'
         FROM product_core.cost_observation WHERE source_ref = 'master#P000021'$q$,
    $q$INSERT INTO product_core.assessment (assessment_type, subject_key, verdict, confidence, reason, evidence, rule_version, assessed_by)
       SELECT 'UNIT_BASIS', observation_id, 'PER_PP', 'MEDIUM', '単品の P(再判定)', '{}', 'unit/1.0', 'ai:unit/1.0'
         FROM product_core.cost_observation WHERE source_ref = 'master#P000021'$q$,
    $q$INSERT INTO product_core.assessment (assessment_type, subject_key, verdict, confidence, reason, evidence, rule_version, assessed_by)
       VALUES ('UNIT_BASIS', 'CO-100000001', 'PER_PP', 'MEDIUM', '別の観測値の判定', '{}', 'unit/1.0', 'ai:unit/1.0')$q$,
    $q$RESET ROLE$q$,
    $q$SELECT setval('product_core.as_seq', 100000400)$q$]);
SELECT pc_test.check('PER_LEGACY_P', '(準備)判定の ID が想定どおり(AS-100000301〜304)',
    (SELECT string_agg(assessment_id || '=' || assessment_type, ',' ORDER BY assessment_id) FROM product_core.assessment WHERE assessment_id BETWEEN 'AS-100000301' AND 'AS-100000399')
    = 'AS-100000301=LEGACY_MAPPING,AS-100000302=UNIT_BASIS,AS-100000303=UNIT_BASIS,AS-100000304=UNIT_BASIS');
SELECT pc_test.lives('PER_LEGACY_P', '(準備)別の観測値の UNIT_BASIS を人が承認(流用の検査用)',
    $q$UPDATE product_core.assessment SET review_status = 'APPROVED', reviewed_by = 'human:op_cost' WHERE assessment_id = 'AS-100000304'$q$, 'pctest_reviewer');

\set norm_lp '{"observed_amount":500,"tax_inclusion":"EXCLUDED","tax_inclusion_basis":"SOURCE_SPEC","tax_rate":null,"tax_inclusion_applied":"EXCLUDED","tax_rate_applied":null,"tax_rate_basis_applied":"NOT_APPLICABLE","unit_divisor":1,"formula":"500"}'
SELECT pc_test.throws('PER_LEGACY_P', '承認済みの legacy_mapping が無ければ正式原価にできない',
    format($q$INSERT INTO product_core.cost_history (pp_id, unit_cost_excl_tax, valid_from, valid_to, source_observation_id, normalization, basis, assessment_id, unit_basis_assessment_id, approved_by)
       SELECT 'PP-100006', 500, '2025-01-01', '2025-12-31', observation_id, %L, 'x', 'AS-100000004', 'AS-100000302', 'human:op_cost'
         FROM product_core.cost_observation WHERE source_ref = 'master#P000021'$q$, :'norm_lp'),
    'pctest_reviewer', '承認済みの対応付け');
SELECT pc_test.lives('PER_LEGACY_P', '(準備)人が対応付けの判定を承認し、P000021 → PP-100006 を ACTIVE 化', ARRAY[
    $q$UPDATE product_core.assessment SET review_status = 'APPROVED', reviewed_by = 'human:op_rev' WHERE assessment_id = 'AS-100000301'$q$,
    $q$INSERT INTO product_core.legacy_mapping (legacy_p, entity_id, entity_type, mapping_role, legacy_snapshot_ref, assessment_id, created_by, record_status, approved_by)
       VALUES ('P000021', 'PP-100006', 'PHYSICAL_PRODUCT', 'SOLE', 'sha#P000021', 'AS-100000301', 'human:op_admin', 'ACTIVE', 'human:op_admin')$q$], 'pctest_reviewer');
SELECT pc_test.throws('PER_LEGACY_P', '対応付けが承認済みでも、UNIT_BASIS の確定が無ければ正式原価にできない',
    format($q$INSERT INTO product_core.cost_history (pp_id, unit_cost_excl_tax, valid_from, valid_to, source_observation_id, normalization, basis, assessment_id, approved_by)
       SELECT 'PP-100006', 500, '2025-01-01', '2025-12-31', observation_id, %L, 'x', 'AS-100000004', 'human:op_cost'
         FROM product_core.cost_observation WHERE source_ref = 'master#P000021'$q$, :'norm_lp'),
    'pctest_reviewer', '単位の根拠が未確定');
SELECT pc_test.throws('PER_LEGACY_P', 'AI だけの UNIT_BASIS 判定(未 review)では正式原価にできない',
    format($q$INSERT INTO product_core.cost_history (pp_id, unit_cost_excl_tax, valid_from, valid_to, source_observation_id, normalization, basis, assessment_id, unit_basis_assessment_id, approved_by)
       SELECT 'PP-100006', 500, '2025-01-01', '2025-12-31', observation_id, %L, 'x', 'AS-100000004', 'AS-100000302', 'human:op_cost'
         FROM product_core.cost_observation WHERE source_ref = 'master#P000021'$q$, :'norm_lp'),
    'pctest_reviewer', '人が確定したもの');
SELECT pc_test.throws('PER_LEGACY_P', 'AI は UNIT_BASIS の review を書けない(列権限)',
    $q$UPDATE product_core.assessment SET review_status = 'APPROVED', reviewed_by = 'human:op_cost' WHERE assessment_id = 'AS-100000302'$q$, 'pctest_ingest', '42501');
SELECT pc_test.throws('PER_LEGACY_P', '別の観測値について人が確定した UNIT_BASIS は流用できない',
    format($q$INSERT INTO product_core.cost_history (pp_id, unit_cost_excl_tax, valid_from, valid_to, source_observation_id, normalization, basis, assessment_id, unit_basis_assessment_id, approved_by)
       SELECT 'PP-100006', 500, '2025-01-01', '2025-12-31', observation_id, %L, 'x', 'AS-100000004', 'AS-100000304', 'human:op_cost'
         FROM product_core.cost_observation WHERE source_ref = 'master#P000021'$q$, :'norm_lp'),
    'pctest_reviewer', '人が確定したもの');
SELECT pc_test.lives('PER_LEGACY_P', '人が「実はセット単位」と修正(CORRECTED → PER_COMPOSITION)',
    $q$UPDATE product_core.assessment SET review_status = 'CORRECTED', reviewed_by = 'human:op_cost', corrected_value = 'PER_COMPOSITION',
           review_note = '3個セットの価格だった' WHERE assessment_id = 'AS-100000302'$q$, 'pctest_reviewer');
SELECT pc_test.throws('PER_LEGACY_P', '人がセット単位と確定した値は PP の正式原価にできない(PER_LEGACY_P ≠ PP 単価)',
    format($q$INSERT INTO product_core.cost_history (pp_id, unit_cost_excl_tax, valid_from, valid_to, source_observation_id, normalization, basis, assessment_id, unit_basis_assessment_id, approved_by)
       SELECT 'PP-100006', 500, '2025-01-01', '2025-12-31', observation_id, %L, 'x', 'AS-100000004', 'AS-100000302', 'human:op_cost'
         FROM product_core.cost_observation WHERE source_ref = 'master#P000021'$q$, :'norm_lp'),
    'pctest_reviewer', 'PER_PP ではありません');
SELECT pc_test.lives('PER_LEGACY_P', '承認済みの対応付け + 人が PER_PP と確定 → 正式原価にできる', ARRAY[
    $q$UPDATE product_core.assessment SET review_status = 'APPROVED', reviewed_by = 'human:op_cost' WHERE assessment_id = 'AS-100000303'$q$,
    format($q$INSERT INTO product_core.cost_history (pp_id, unit_cost_excl_tax, valid_from, valid_to, source_observation_id, normalization, basis, assessment_id, unit_basis_assessment_id, approved_by)
       SELECT 'PP-100006', 500, '2025-01-01', '2025-12-31', observation_id, %L, '商品マスター(人が PP 単位と確定)', 'AS-100000004', 'AS-100000303', 'human:op_cost'
         FROM product_core.cost_observation WHERE source_ref = 'master#P000021'$q$, :'norm_lp')], 'pctest_reviewer');
SELECT pc_test.check('PER_LEGACY_P', '正式原価の作成後も観測値は PER_LEGACY_P のまま残る',
    (SELECT source_unit_basis = 'PER_LEGACY_P' AND observed_amount = 500 FROM product_core.cost_observation WHERE source_ref = 'master#P000021')
    AND EXISTS (SELECT 1 FROM product_core.cost_history WHERE pp_id = 'PP-100006' AND unit_basis_assessment_id = 'AS-100000303'));

-- ------------------------------------------------------------
--  C. DB 採番(1.3)
-- ------------------------------------------------------------
SELECT pc_test.check('DB_ID', '取込ロールは Product Core の ID 列に INSERT 権限を持たない(13列)',
    NOT EXISTS (SELECT 1 FROM (VALUES
        ('product_core.physical_product', 'pp_id'), ('product_core.composition', 'cp_id'), ('product_core.listing', 'listing_id'),
        ('product_core.listing_group', 'listing_group_id'), ('product_core.identifier', 'identifier_id'),
        ('product_core.identifier_link', 'link_id'), ('product_core.legacy_mapping', 'mapping_id'),
        ('product_core.product_relationship', 'relationship_id'), ('product_core.listing_reference_history', 'ref_id'),
        ('product_core.assessment', 'assessment_id'), ('product_core.cost_observation', 'observation_id'),
        ('product_core.composition_component', 'component_id'), ('product_core.legacy_listing_mapping', 'llm_id')) AS c(t, col)
      WHERE has_column_privilege('pc_ingest', c.t, c.col, 'INSERT')));
SELECT pc_test.throws('DB_ID', '取込ロールは CP の ID を指定できない',
    $q$INSERT INTO product_core.composition (cp_id, name, comp_type, origin_key, created_by) VALUES ('CP-100090', 'x', 'FIXED', 'k:cp90', 'ingest:t')$q$, 'pctest_ingest', '42501');
SELECT pc_test.throws('DB_ID', '取込ロールは出品の ID を指定できない',
    $q$INSERT INTO product_core.listing (listing_id, channel, seller_sku, source_type, source_ref, created_by) VALUES ('LS-100090', 'AMAZON', 'SKU-ID', 't', 't', 'ingest:t')$q$, 'pctest_ingest', '42501');
SELECT pc_test.throws('DB_ID', '取込ロールは判定の ID を指定できない',
    $q$INSERT INTO product_core.assessment (assessment_id, assessment_type, subject_key, verdict, confidence, reason, evidence, rule_version, assessed_by)
       VALUES ('AS-199999999', 'LEGACY_MAPPING', 'P000090', 'NEEDS_REVIEW', 'LOW', 'x', '{}', 'v', 'ai:x')$q$, 'pctest_ingest', '42501');
SELECT pc_test.throws('DB_ID', '取込ロールはコードの ID を指定できない',
    $q$INSERT INTO product_core.identifier (identifier_id, id_type, value, first_seen_source, first_seen_ref, created_by)
       VALUES ('ID-19999999', 'JAN', '4900000000905', 's', 'r', 'ingest:t')$q$, 'pctest_ingest', '42501');
SELECT pc_test.throws('DB_ID', '取込ロールは観測値の ID を指定できない',
    $q$INSERT INTO product_core.cost_observation (observation_id, source, source_ref, source_reliability, observed_at, observed_amount, tax_inclusion, tax_inclusion_basis,
            tax_rate_basis, amount_unit, subject_entity_id, subject_entity_type, source_unit_basis, source_record_key, source_record_hash, evidence, created_by, ingest_run_id)
       VALUES ('CO-199999999', 'sku_string', 'id#1', 'LOW', now(), 1, 'UNKNOWN', 'NONE', 'NONE', 'PER_PP', 'PP-100001', 'PHYSICAL_PRODUCT', 'PER_PP', 'id#1', repeat('f', 64), '{}', 'ingest:t', 'IR-20260926-000002-abcd')$q$,
    'pctest_ingest', '42501');
SELECT pc_test.lives('DB_ID', 'INSERT ... RETURNING で採番された ID を受け取れる',
    $q$WITH ins AS (INSERT INTO product_core.physical_product (name, origin_key, created_by, ingest_run_id)
                    VALUES ('採番確認', 'legacy:P000031', 'ingest:t', 'IR-20260926-000002-abcd') RETURNING pp_id)
       INSERT INTO pc_test.result (area, name, kind, ok, detail)
       SELECT 'DB_ID', 'RETURNING の値が PP- 形式で、登録簿にも同じ ID がある', 'CHECK',
              pp_id ~ '^PP-[0-9]{6,}$', pp_id FROM ins$q$, 'pctest_ingest');
SELECT pc_test.check('DB_ID', 'RETURNING で受け取った ID が実体・登録簿と一致',
    (SELECT p.pp_id = e.entity_id FROM product_core.physical_product p JOIN product_core.core_entity e ON e.entity_id = p.pp_id
      WHERE p.origin_key = 'legacy:P000031'));
SELECT pc_test.lives('DB_ID', '取込関数で PP・CP・出品を作れる(ID は DB 採番)', ARRAY[
    $q$SELECT * FROM product_core.import_physical_product('legacy:P000032', '取込関数の単品', 'GOODS', 'ingest:t', 'IR-20260926-000002-abcd')$q$,
    $q$SELECT * FROM product_core.import_composition('legacy:P000932', '取込関数のセット', 'FIXED', 'ingest:t', 'IR-20260926-000002-abcd')$q$,
    $q$SELECT * FROM product_core.import_listing('AMAZON', 'SKU-IMP-1', 'B0TESTIMP1', NULL, NULL, 'test', 'imp#1', 'ingest:t', 'IR-20260926-000002-abcd')$q$,
    $q$SELECT * FROM product_core.import_listing('RAKUTEN', NULL, NULL, 'r-imp-1', NULL, 'test', 'imp#2', 'ingest:t', 'IR-20260926-000002-abcd')$q$,
    $q$SELECT * FROM product_core.import_listing('RAKUTEN', NULL, NULL, 'r-imp-1', 'sku-1', 'test', 'imp#3', 'ingest:t', 'IR-20260926-000002-abcd')$q$], 'pctest_ingest');
SELECT pc_test.lives('DB_ID', '取込関数を再実行しても同じ ID を返し、作り直さない', ARRAY[
    $q$INSERT INTO pc_test.result (area, name, kind, ok, detail)
       SELECT 'DB_ID', '再実行: PP は created = false で同じ ID', 'CHECK', NOT r.created AND r.pp_id = p.pp_id, r.pp_id
         FROM product_core.import_physical_product('legacy:P000032', '名前が変わっても', 'GOODS', 'ingest:t2', 'IR-20260926-000002-abcd') r
         JOIN product_core.physical_product p ON p.origin_key = 'legacy:P000032'$q$,
    $q$INSERT INTO pc_test.result (area, name, kind, ok, detail)
       SELECT 'DB_ID', '再実行: CP は created = false', 'CHECK', NOT r.created, r.cp_id
         FROM product_core.import_composition('legacy:P000932', 'x', 'FIXED', 'ingest:t2', 'IR-20260926-000002-abcd') r$q$,
    $q$INSERT INTO pc_test.result (area, name, kind, ok, detail)
       SELECT 'DB_ID', '再実行: Amazon 出品は created = false', 'CHECK', NOT r.created, r.listing_id
         FROM product_core.import_listing('AMAZON', 'SKU-IMP-1', 'B0TESTIMP1', NULL, NULL, 'test', 'imp#1b', 'ingest:t2', 'IR-20260926-000002-abcd') r$q$,
    $q$INSERT INTO pc_test.result (area, name, kind, ok, detail)
       SELECT 'DB_ID', '再実行: 楽天(SKU 空)は created = false', 'CHECK', NOT r.created, r.listing_id
         FROM product_core.import_listing('RAKUTEN', NULL, NULL, 'r-imp-1', NULL, 'test', 'imp#2b', 'ingest:t2', 'IR-20260926-000002-abcd') r$q$], 'pctest_ingest');
SELECT pc_test.check('DB_ID', '再実行しても件数は増えず、既存の名前も上書きしない',
    (SELECT count(*) = 1 AND min(name) = '取込関数の単品' FROM product_core.physical_product WHERE origin_key = 'legacy:P000032')
    AND (SELECT count(*) FROM product_core.composition WHERE origin_key = 'legacy:P000932') = 1
    AND (SELECT count(*) FROM product_core.listing WHERE seller_sku = 'SKU-IMP-1') = 1
    AND (SELECT count(*) FROM product_core.listing WHERE rakuten_item_id = 'r-imp-1') = 2);
SELECT pc_test.check('DB_ID', 'DB 採番は既存 ID と衝突しない(PP・CP・出品・登録簿の ID がすべて一意)',
    (SELECT count(*) = count(DISTINCT entity_id) FROM product_core.core_entity)
    AND (SELECT count(*) FROM product_core.core_entity e
          WHERE NOT EXISTS (SELECT 1 FROM product_core.physical_product WHERE pp_id = e.entity_id)
            AND NOT EXISTS (SELECT 1 FROM product_core.composition WHERE cp_id = e.entity_id)
            AND NOT EXISTS (SELECT 1 FROM product_core.listing WHERE listing_id = e.entity_id)) = 0);
SELECT pc_test.throws('DB_ID', '途中失敗: 取込関数の後でエラー → 全体が取り消される',
    ARRAY[$q$SELECT * FROM product_core.import_physical_product('legacy:P000033', '途中失敗', 'GOODS', 'ingest:t', 'IR-20260926-000002-abcd')$q$,
          $q$SELECT 1 / 0$q$], 'pctest_ingest', '22012');
SELECT pc_test.check('DB_ID', '途中失敗の後、PP も登録簿も残っていない(孤児なし)',
    NOT EXISTS (SELECT 1 FROM product_core.physical_product WHERE origin_key = 'legacy:P000033')
    AND (SELECT count(*) FROM product_core.core_entity WHERE entity_type = 'PHYSICAL_PRODUCT')
        = (SELECT count(*) FROM product_core.physical_product));
SELECT pc_test.lives('DB_ID', '途中失敗の後の再実行 → 1件だけ作られ、さらに再実行しても増えない', ARRAY[
    $q$SELECT * FROM product_core.import_physical_product('legacy:P000033', '途中失敗', 'GOODS', 'ingest:t', 'IR-20260926-000002-abcd')$q$,
    $q$SELECT * FROM product_core.import_physical_product('legacy:P000033', '途中失敗', 'GOODS', 'ingest:t', 'IR-20260926-000002-abcd')$q$], 'pctest_ingest');
SELECT pc_test.check('DB_ID', '再実行後の件数は1件',
    (SELECT count(*) FROM product_core.physical_product WHERE origin_key = 'legacy:P000033') = 1);

-- ------------------------------------------------------------
--  D. GTIN12 / LEADING_ZERO_SUSPECT(1.4)
-- ------------------------------------------------------------
SELECT pc_test.lives('GTIN12', '12桁のコードを元の値のまま(GTIN12)保存できる',
    $q$INSERT INTO product_core.identifier (id_type, value, checkdigit_valid, first_seen_source, first_seen_ref, created_by, ingest_run_id)
       VALUES ('GTIN12', '012345678905', true, 'legacy_product_master', 'sha#g12', 'ingest:t', 'IR-20260926-000002-abcd')$q$, 'pctest_ingest');
SELECT pc_test.check('GTIN12', '保存した値は12桁のまま・OBSERVED',
    (SELECT value = '012345678905' AND length(value) = 12 AND origin = 'OBSERVED' FROM product_core.identifier WHERE id_type = 'GTIN12'));
SELECT pc_test.throws('GTIN12', '12桁を JAN として登録できない(JAN は 8桁 / 13桁)',
    $q$INSERT INTO product_core.identifier (id_type, value, first_seen_source, first_seen_ref, created_by) VALUES ('JAN', '012345678905', 's', 'r', 'ingest:t')$q$, 'pctest_ingest', '23514');
SELECT pc_test.check('GTIN12', '勝手に13桁へ変換していない(0 を補った JAN が存在しない)',
    NOT EXISTS (SELECT 1 FROM product_core.identifier WHERE value = '0012345678905'));
SELECT pc_test.throws('GTIN12', 'AI は 0 を補った13桁を取込元の値として登録できない(人の review が必要)',
    $q$INSERT INTO product_core.identifier (id_type, value, first_seen_source, first_seen_ref, created_by) VALUES ('JAN', '0012345678905', 'ai_guess', 'r', 'ai:x')$q$,
    'pctest_ingest', '先頭0を補った値');
SELECT pc_test.lives('GTIN12', 'AI は LEADING_ZERO_SUSPECT を提案(判定として記録)できる', ARRAY[
    $q$SELECT setval('product_core.as_seq', 100000500)$q$,
    $q$SET LOCAL ROLE pctest_ingest$q$,
    $q$INSERT INTO product_core.assessment (assessment_type, subject_key, verdict, confidence, reason, evidence, rule_version, assessed_by)
       VALUES ('IDENTIFIER_LINK', 'GTIN12:012345678905', 'LEADING_ZERO_SUSPECT', 'MEDIUM', '12桁・国内商品のため先頭0欠落の疑い',
               '{"original":"012345678905","candidate":"0012345678905"}', 'gtin/1.0', 'rule:gtin/1.0')$q$,
    $q$RESET ROLE$q$,
    $q$SELECT setval('product_core.as_seq', 100000600)$q$]);
SELECT pc_test.throws('GTIN12', 'AI は正式な13桁(HUMAN_CORRECTED)を作れない(列権限)',
    $q$INSERT INTO product_core.identifier (id_type, value, first_seen_source, first_seen_ref, created_by, origin, derived_from_identifier_id, correction_assessment_id)
       SELECT 'JAN', '0012345678905', 'review', 'AS-100000501', 'ai:x', 'HUMAN_CORRECTED', identifier_id, 'AS-100000501'
         FROM product_core.identifier WHERE id_type = 'GTIN12' AND value = '012345678905'$q$, 'pctest_ingest', '42501');
SELECT pc_test.throws('GTIN12', '人の review(CORRECTED)の前は、人でも13桁を作れない',
    $q$INSERT INTO product_core.identifier (id_type, value, first_seen_source, first_seen_ref, created_by, origin, derived_from_identifier_id, correction_assessment_id)
       SELECT 'JAN', '0012345678905', 'review', 'AS-100000501', 'human:op_admin', 'HUMAN_CORRECTED', identifier_id, 'AS-100000501'
         FROM product_core.identifier WHERE id_type = 'GTIN12' AND value = '012345678905'$q$, 'pctest_reviewer', 'CORRECTED');
SELECT pc_test.throws('GTIN12', 'AI は判定を CORRECTED にできない(列権限)',
    $q$UPDATE product_core.assessment SET review_status = 'CORRECTED', reviewed_by = 'human:op_rev', corrected_value = '0012345678905', review_note = 'x'
        WHERE assessment_id = 'AS-100000501'$q$, 'pctest_ingest', '42501');
SELECT pc_test.lives('GTIN12', '人が review で13桁へ修正(CORRECTED)',
    $q$UPDATE product_core.assessment SET review_status = 'CORRECTED', reviewed_by = 'human:op_rev', corrected_value = '0012345678905',
           review_note = '現物の印字が13桁(先頭0)' WHERE assessment_id = 'AS-100000501'$q$, 'pctest_reviewer');
SELECT pc_test.throws('GTIN12', 'IDENTIFIER_APPROVE の無い人は修正コードを登録できない',
    $q$INSERT INTO product_core.identifier (id_type, value, first_seen_source, first_seen_ref, created_by, origin, derived_from_identifier_id, correction_assessment_id)
       SELECT 'JAN', '0012345678905', 'review', 'AS-100000501', 'human:op_rev', 'HUMAN_CORRECTED', identifier_id, 'AS-100000501'
         FROM product_core.identifier WHERE id_type = 'GTIN12' AND value = '012345678905'$q$, 'pctest_reviewer', 'IDENTIFIER_APPROVE');
SELECT pc_test.throws('GTIN12', '人が確定した値と違う13桁は登録できない',
    $q$INSERT INTO product_core.identifier (id_type, value, first_seen_source, first_seen_ref, created_by, origin, derived_from_identifier_id, correction_assessment_id)
       SELECT 'JAN', '4912345678905', 'review', 'AS-100000501', 'human:op_admin', 'HUMAN_CORRECTED', identifier_id, 'AS-100000501'
         FROM product_core.identifier WHERE id_type = 'GTIN12' AND value = '012345678905'$q$, 'pctest_reviewer', '先頭0を補った');
SELECT pc_test.lives('GTIN12', '人の修正の後、権限者だけが修正コード(13桁)を登録できる',
    $q$INSERT INTO product_core.identifier (id_type, value, checkdigit_valid, first_seen_source, first_seen_ref, created_by, origin, derived_from_identifier_id, correction_assessment_id)
       SELECT 'JAN', '0012345678905', true, 'review', 'AS-100000501', 'human:op_admin', 'HUMAN_CORRECTED', identifier_id, 'AS-100000501'
         FROM product_core.identifier WHERE id_type = 'GTIN12' AND value = '012345678905'$q$, 'pctest_reviewer');
SELECT pc_test.check('GTIN12', '元の12桁は残り、13桁は派生元と根拠の判定を指す',
    EXISTS (SELECT 1 FROM product_core.identifier n JOIN product_core.identifier o ON o.identifier_id = n.derived_from_identifier_id
             WHERE n.value = '0012345678905' AND n.origin = 'HUMAN_CORRECTED' AND n.correction_assessment_id = 'AS-100000501'
               AND o.id_type = 'GTIN12' AND o.value = '012345678905' AND o.origin = 'OBSERVED'));
SELECT pc_test.throws('GTIN12', '元の12桁は変更できない(追記専用)',
    $q$UPDATE product_core.identifier SET value = '112345678905' WHERE id_type = 'GTIN12'$q$, NULL, '追記専用');
SELECT pc_test.throws('GTIN12', '元の12桁は削除できない(追記専用)',
    $q$DELETE FROM product_core.identifier WHERE id_type = 'GTIN12'$q$, NULL, '追記専用');

-- ------------------------------------------------------------
--  E. 取込の冪等性(1.5)
-- ------------------------------------------------------------
SELECT pc_test.lives('IDEMPOTENCY', 'ファイル(SHA256)とレコードを登録し、PP と対応表を作る', ARRAY[
    $q$SELECT * FROM legacy_ingest.import_source_file_register('IR-20260926-000002-abcd', 'legacy_product_master', 'master_v1.xlsx', repeat('1', 64), '2026-09-26 10:00+09', 3)$q$,
    $q$SELECT * FROM legacy_ingest.import_source_record_register(
           (SELECT file_id FROM legacy_ingest.import_source_file WHERE file_sha256 = repeat('1', 64)), 'IR-20260926-000002-abcd',
           'legacy_product_master', 'legacy_p=P000041', repeat('a', 64), '{"name":"冪等テスト品","cost":700}')$q$,
    $q$INSERT INTO legacy_ingest.import_record_map (record_id, target_table, target_id)
       SELECT r.record_id, 'physical_product', p.pp_id
         FROM legacy_ingest.import_source_record r,
              product_core.import_physical_product('legacy:P000041', '冪等テスト品', 'GOODS', 'ingest:t', 'IR-20260926-000002-abcd') p
        WHERE r.source_record_key = 'legacy_p=P000041' AND r.source_record_hash = repeat('a', 64)
       ON CONFLICT DO NOTHING$q$], 'pctest_ingest');
SELECT pc_test.lives('IDEMPOTENCY', '同じファイル SHA256 + 同じレコードを再投入(全手順をやり直す)', ARRAY[
    $q$INSERT INTO pc_test.result (area, name, kind, ok, detail)
       SELECT 'IDEMPOTENCY', '再投入: ファイルは created = false', 'CHECK', NOT created, file_id::text
         FROM legacy_ingest.import_source_file_register('IR-20260926-000002-abcd', 'legacy_product_master', 'master_v1_copy.xlsx', repeat('1', 64), NULL, 3)$q$,
    $q$INSERT INTO pc_test.result (area, name, kind, ok, detail)
       SELECT 'IDEMPOTENCY', '再投入: レコードは created = false・食い違いなし', 'CHECK', NOT created AND NOT is_conflict, record_id::text
         FROM legacy_ingest.import_source_record_register(
              (SELECT file_id FROM legacy_ingest.import_source_file WHERE file_sha256 = repeat('1', 64)), 'IR-20260926-000002-abcd',
              'legacy_product_master', 'legacy_p=P000041', repeat('a', 64), '{"name":"冪等テスト品","cost":700}')$q$,
    $q$INSERT INTO legacy_ingest.import_record_map (record_id, target_table, target_id)
       SELECT r.record_id, 'physical_product', p.pp_id
         FROM legacy_ingest.import_source_record r,
              product_core.import_physical_product('legacy:P000041', '冪等テスト品', 'GOODS', 'ingest:t', 'IR-20260926-000002-abcd') p
        WHERE r.source_record_key = 'legacy_p=P000041' AND r.source_record_hash = repeat('a', 64)
       ON CONFLICT DO NOTHING$q$], 'pctest_ingest');
SELECT pc_test.check('IDEMPOTENCY', '再投入しても ファイル1・レコード1・PP 1・対応表1 のまま',
    (SELECT count(*) FROM legacy_ingest.import_source_file WHERE file_sha256 = repeat('1', 64)) = 1
    AND (SELECT count(*) FROM legacy_ingest.import_source_record WHERE source_record_key = 'legacy_p=P000041') = 1
    AND (SELECT count(*) FROM product_core.physical_product WHERE origin_key = 'legacy:P000041') = 1
    AND (SELECT count(*) FROM legacy_ingest.import_record_map m JOIN legacy_ingest.import_source_record r USING (record_id)
          WHERE r.source_record_key = 'legacy_p=P000041') = 1);
SELECT pc_test.throws('IDEMPOTENCY', '途中失敗: ファイル・レコード・PP を作った後でエラー → 取り消される',
    ARRAY[$q$SELECT * FROM legacy_ingest.import_source_file_register('IR-20260926-000002-abcd', 'legacy_product_master', 'master_v2.xlsx', repeat('2', 64), NULL, 1)$q$,
          $q$SELECT * FROM legacy_ingest.import_source_record_register(
                 (SELECT file_id FROM legacy_ingest.import_source_file WHERE file_sha256 = repeat('2', 64)), 'IR-20260926-000002-abcd',
                 'legacy_product_master', 'legacy_p=P000042', repeat('b', 64), '{"name":"ロールバック品"}')$q$,
          $q$SELECT * FROM product_core.import_physical_product('legacy:P000042', 'ロールバック品', 'GOODS', 'ingest:t', 'IR-20260926-000002-abcd')$q$,
          $q$SELECT 1 / 0$q$], 'pctest_ingest', '22012');
SELECT pc_test.check('IDEMPOTENCY', 'ロールバック後はファイル・レコード・PP・登録簿のどれも残らない',
    NOT EXISTS (SELECT 1 FROM legacy_ingest.import_source_file WHERE file_sha256 = repeat('2', 64))
    AND NOT EXISTS (SELECT 1 FROM legacy_ingest.import_source_record WHERE source_record_key = 'legacy_p=P000042')
    AND NOT EXISTS (SELECT 1 FROM product_core.physical_product WHERE origin_key = 'legacy:P000042')
    AND (SELECT count(*) FROM product_core.core_entity WHERE entity_type = 'PHYSICAL_PRODUCT') = (SELECT count(*) FROM product_core.physical_product));
SELECT pc_test.lives('IDEMPOTENCY', 'ロールバック後の再実行(2回)', ARRAY[
    $q$SELECT * FROM legacy_ingest.import_source_file_register('IR-20260926-000002-abcd', 'legacy_product_master', 'master_v2.xlsx', repeat('2', 64), NULL, 1)$q$,
    $q$SELECT * FROM legacy_ingest.import_source_record_register(
           (SELECT file_id FROM legacy_ingest.import_source_file WHERE file_sha256 = repeat('2', 64)), 'IR-20260926-000002-abcd',
           'legacy_product_master', 'legacy_p=P000042', repeat('b', 64), '{"name":"ロールバック品"}')$q$,
    $q$SELECT * FROM product_core.import_physical_product('legacy:P000042', 'ロールバック品', 'GOODS', 'ingest:t', 'IR-20260926-000002-abcd')$q$,
    $q$SELECT * FROM legacy_ingest.import_source_file_register('IR-20260926-000002-abcd', 'legacy_product_master', 'master_v2.xlsx', repeat('2', 64), NULL, 1)$q$,
    $q$SELECT * FROM product_core.import_physical_product('legacy:P000042', 'ロールバック品', 'GOODS', 'ingest:t', 'IR-20260926-000002-abcd')$q$], 'pctest_ingest');
SELECT pc_test.check('IDEMPOTENCY', '再実行後は ファイル1・レコード1・PP 1 で整合',
    (SELECT count(*) FROM legacy_ingest.import_source_file WHERE file_sha256 = repeat('2', 64)) = 1
    AND (SELECT count(*) FROM legacy_ingest.import_source_record WHERE source_record_key = 'legacy_p=P000042') = 1
    AND (SELECT count(*) FROM product_core.physical_product WHERE origin_key = 'legacy:P000042') = 1);
SELECT pc_test.lives('IDEMPOTENCY', '別のファイルで同じ業務キー・違う内容が来た',
    $q$INSERT INTO pc_test.result (area, name, kind, ok, detail)
       SELECT 'IDEMPOTENCY', '別の版は新しいレコードとして追記され、食い違い(is_conflict)が返る', 'CHECK', created AND is_conflict, record_id::text
         FROM legacy_ingest.import_source_record_register(
              (SELECT file_id FROM legacy_ingest.import_source_file_register('IR-20260926-000002-abcd', 'legacy_product_master', 'master_v3.xlsx', repeat('3', 64), NULL, 1)),
              'IR-20260926-000002-abcd', 'legacy_product_master', 'legacy_p=P000041', repeat('c', 64), '{"name":"冪等テスト品(改名)","cost":900}')$q$,
    'pctest_ingest');
SELECT pc_test.check('IDEMPOTENCY', '黙って上書きしない: 最初の版の内容はそのまま残る',
    (SELECT raw ->> 'cost' FROM legacy_ingest.import_source_record WHERE source_record_key = 'legacy_p=P000041' AND source_record_hash = repeat('a', 64)) = '700');
SELECT pc_test.check('IDEMPOTENCY', '食い違いは確認用ビュー(v_source_key_conflicts)に出る',
    (SELECT versions FROM legacy_ingest.v_source_key_conflicts WHERE source_record_key = 'legacy_p=P000041') = 2);
SELECT pc_test.lives('IDEMPOTENCY', '別の版から PP を取り込もうとしても既存 PP を返す(名前を上書きしない)',
    $q$SELECT * FROM product_core.import_physical_product('legacy:P000041', '冪等テスト品(改名)', 'GOODS', 'ingest:t', 'IR-20260926-000002-abcd')$q$, 'pctest_ingest');
SELECT pc_test.check('IDEMPOTENCY', 'PP は1件のまま・名前は最初のまま',
    (SELECT count(*) = 1 AND min(name) = '冪等テスト品' FROM product_core.physical_product WHERE origin_key = 'legacy:P000041'));
SELECT pc_test.lives('IDEMPOTENCY', '観測値: 同じ業務キーで違う内容は別の行として追記(上書きしない)', ARRAY[
    $q$INSERT INTO product_core.cost_observation (source, source_ref, source_reliability, observed_at, observed_amount, tax_inclusion, tax_inclusion_basis,
            tax_rate_basis, amount_unit, subject_legacy_p, source_unit_basis, source_record_key, source_record_hash, evidence, created_by, ingest_run_id)
       VALUES ('legacy_product_master', 'master_v1.xlsx#P000041', 'MEDIUM', '2026-09-26', 700, 'UNKNOWN', 'NONE', 'NONE', 'UNKNOWN',
               'P000041', 'PER_LEGACY_P', 'legacy_p=P000041', repeat('a', 64), '{}', 'ingest:t', 'IR-20260926-000002-abcd')$q$,
    $q$INSERT INTO product_core.cost_observation (source, source_ref, source_reliability, observed_at, observed_amount, tax_inclusion, tax_inclusion_basis,
            tax_rate_basis, amount_unit, subject_legacy_p, source_unit_basis, source_record_key, source_record_hash, evidence, created_by, ingest_run_id)
       VALUES ('legacy_product_master', 'master_v3.xlsx#P000041', 'MEDIUM', '2026-09-27', 900, 'UNKNOWN', 'NONE', 'NONE', 'UNKNOWN',
               'P000041', 'PER_LEGACY_P', 'legacy_p=P000041', repeat('c', 64), '{}', 'ingest:t', 'IR-20260926-000002-abcd')$q$,
    $q$INSERT INTO product_core.cost_observation (source, source_ref, source_reliability, observed_at, observed_amount, tax_inclusion, tax_inclusion_basis,
            tax_rate_basis, amount_unit, subject_legacy_p, source_unit_basis, source_record_key, source_record_hash, evidence, created_by, ingest_run_id)
       VALUES ('legacy_product_master', 'master_v1_copy.xlsx#P000041', 'MEDIUM', '2026-09-28', 700, 'UNKNOWN', 'NONE', 'NONE', 'UNKNOWN',
               'P000041', 'PER_LEGACY_P', 'legacy_p=P000041', repeat('a', 64), '{}', 'ingest:t', 'IR-20260926-000002-abcd')
       ON CONFLICT DO NOTHING$q$], 'pctest_ingest');
SELECT pc_test.check('IDEMPOTENCY', '観測値は2版(700・900)で、同じ版の再投入は増えず、食い違いがビューに出る',
    (SELECT array_agg(observed_amount::int ORDER BY observed_amount) FROM product_core.cost_observation WHERE source_record_key = 'legacy_p=P000041') = ARRAY[700, 900]
    AND (SELECT versions FROM product_core.v_cost_observation_key_conflicts WHERE source_record_key = 'legacy_p=P000041') = 2);
SELECT pc_test.throws('IDEMPOTENCY', '取込元レコードは変更できない(追記専用)',
    $q$UPDATE legacy_ingest.import_source_record SET raw = '{}' WHERE source_record_key = 'legacy_p=P000041'$q$, NULL, '追記専用');
SELECT pc_test.throws('IDEMPOTENCY', '取込元ファイルの記録は削除できない(追記専用)',
    $q$DELETE FROM legacy_ingest.import_source_file WHERE file_sha256 = repeat('1', 64)$q$, NULL, '追記専用');
SELECT pc_test.throws('IDEMPOTENCY', 'ファイル SHA256 の形式を検査',
    $q$SELECT * FROM legacy_ingest.import_source_file_register('IR-20260926-000002-abcd', 'legacy_product_master', 'x.xlsx', 'not-a-hash', NULL, 1)$q$, 'pctest_ingest', '23514');

-- ------------------------------------------------------------
--  F. Human Review(1.6)
-- ------------------------------------------------------------
SELECT pc_test.throws('HUMAN_REVIEW', 'AI(取込ロール)は review_batch を作れない',
    $q$INSERT INTO product_core.review_batch (review_type, file_name, created_by, updated_by) VALUES ('LEGACY_LISTING', 'x.xlsx', 'human:op_rev', 'human:op_rev')$q$,
    'pctest_ingest', '42501');
SELECT pc_test.throws('HUMAN_REVIEW', 'REVIEW 権限の無い人の名義では review_batch を作れない',
    $q$INSERT INTO product_core.review_batch (review_type, file_name, created_by, updated_by) VALUES ('LEGACY_LISTING', 'x.xlsx', 'human:ghost', 'human:ghost')$q$,
    'pctest_reviewer', 'REVIEW');
-- 直前の拒否された挿入でもシーケンスは進む(番号は欠番になる)ため、他の準備と同じく setval で ID を固定する
SELECT pc_test.lives('HUMAN_REVIEW', '人が確認表の束(RB-00000002)を作る', ARRAY[
    $q$SELECT setval('product_core.rb_seq', 1)$q$,
    $q$SET LOCAL ROLE pctest_reviewer$q$,
    $q$INSERT INTO product_core.review_batch (review_type, file_name, file_sha256, created_by, updated_by)
       VALUES ('LEGACY_LISTING', 'review_legacy_listing_RB-00000002.xlsx', repeat('e', 64), 'human:op_rev', 'human:op_rev')$q$,
    $q$RESET ROLE$q$,
    $q$SELECT setval('product_core.rb_seq', 100)$q$]);
SELECT pc_test.check('HUMAN_REVIEW', 'review_batch は PENDING で作られ、ID は DB 採番',
    (SELECT status = 'PENDING' FROM product_core.review_batch WHERE review_batch_id = 'RB-00000002'));
SELECT pc_test.lives('HUMAN_REVIEW', '(準備)AI が C番号の対応判定を記録', ARRAY[
    $q$SELECT setval('product_core.as_seq', 100000700)$q$,
    $q$SET LOCAL ROLE pctest_ingest$q$,
    $q$INSERT INTO product_core.assessment (assessment_type, subject_key, verdict, confidence, reason, evidence, rule_version, assessed_by) VALUES
       ('LEGACY_LISTING', 'C000021', 'SAME_NATURAL_KEY', 'HIGH', 'SKU 一致', '{}', 'llm/1.0', 'rule:llm/1.0'),
       ('LEGACY_LISTING', 'C000022', 'AMBIGUOUS', 'LOW', '同じ SKU が2行', '{}', 'llm/1.0', 'rule:llm/1.0')$q$,
    $q$RESET ROLE$q$,
    $q$SELECT setval('product_core.as_seq', 100000800)$q$]);
SELECT pc_test.throws('HUMAN_REVIEW', 'AI は判定を作るときに保留理由・review_batch を書けない(列権限)',
    $q$INSERT INTO product_core.assessment (assessment_type, subject_key, verdict, confidence, reason, evidence, rule_version, assessed_by, review_reason_code, review_batch_id)
       VALUES ('LEGACY_LISTING', 'C000023', 'AMBIGUOUS', 'LOW', 'x', '{}', 'v', 'ai:x', 'HOLD', 'RB-00000002')$q$, 'pctest_ingest', '42501');
SELECT pc_test.throws('HUMAN_REVIEW', 'AI は review_batch_id・保留理由を後から書けない(列権限)',
    $q$UPDATE product_core.assessment SET review_batch_id = 'RB-00000002', review_reason_code = 'HOLD' WHERE assessment_id = 'AS-100000702'$q$, 'pctest_ingest', '42501');
SELECT pc_test.throws('HUMAN_REVIEW', '保留(ON_HOLD)には理由コードが必要',
    $q$UPDATE product_core.assessment SET review_status = 'ON_HOLD', reviewed_by = 'human:op_rev', review_note = '判断保留', review_channel = 'EXCEL',
           review_batch_id = 'RB-00000002' WHERE assessment_id = 'AS-100000702'$q$, 'pctest_reviewer', '23514');
SELECT pc_test.throws('HUMAN_REVIEW', 'NEEDS_MORE_INFO には不足している情報の指定が必要',
    $q$UPDATE product_core.assessment SET review_status = 'ON_HOLD', reviewed_by = 'human:op_rev', review_note = '資料待ち', review_channel = 'EXCEL',
           review_batch_id = 'RB-00000002', review_reason_code = 'NEEDS_MORE_INFO' WHERE assessment_id = 'AS-100000702'$q$, 'pctest_reviewer', '23514');
SELECT pc_test.throws('HUMAN_REVIEW', 'Excel での review には review_batch_id が必要',
    $q$UPDATE product_core.assessment SET review_status = 'APPROVED', reviewed_by = 'human:op_rev', review_channel = 'EXCEL' WHERE assessment_id = 'AS-100000701'$q$,
    'pctest_reviewer', '23514');
SELECT pc_test.throws('HUMAN_REVIEW', '存在しない review_batch は指定できない',
    $q$UPDATE product_core.assessment SET review_status = 'APPROVED', reviewed_by = 'human:op_rev', review_channel = 'EXCEL', review_batch_id = 'RB-99999999'
        WHERE assessment_id = 'AS-100000701'$q$, 'pctest_reviewer', '23503');
SELECT pc_test.lives('HUMAN_REVIEW', '人が束の中で承認・保留(NEEDS_MORE_INFO / PURCHASE_RECORD)を記録', ARRAY[
    $q$UPDATE product_core.assessment SET review_status = 'APPROVED', reviewed_by = 'human:op_rev', review_channel = 'EXCEL', review_batch_id = 'RB-00000002',
           review_ref = 'row#2' WHERE assessment_id = 'AS-100000701'$q$,
    $q$UPDATE product_core.assessment SET review_status = 'ON_HOLD', reviewed_by = 'human:op_rev', review_note = '仕入記録を確認したい', review_channel = 'EXCEL',
           review_batch_id = 'RB-00000002', review_ref = 'row#3', review_reason_code = 'NEEDS_MORE_INFO', review_info_needed = 'PURCHASE_RECORD'
        WHERE assessment_id = 'AS-100000702'$q$], 'pctest_reviewer');
SELECT pc_test.check('HUMAN_REVIEW', 'review_batch_id で束の判断を追跡できる(誰が・何を・なぜ保留)',
    (SELECT string_agg(assessment_id || ':' || review_status || ':' || coalesce(review_reason_code, '-') || ':' || coalesce(review_info_needed, '-') || ':' || reviewed_by,
                       ',' ORDER BY assessment_id)
       FROM product_core.assessment WHERE review_batch_id = 'RB-00000002')
    = 'AS-100000701:APPROVED:-:-:human:op_rev,AS-100000702:ON_HOLD:NEEDS_MORE_INFO:PURCHASE_RECORD:human:op_rev');
SELECT pc_test.lives('HUMAN_REVIEW', '保留 → 承認しても、保留理由は記録として残る',
    $q$UPDATE product_core.assessment SET review_status = 'APPROVED', reviewed_by = 'human:op_rev', review_note = '仕入記録で確認済み'
        WHERE assessment_id = 'AS-100000702'$q$, 'pctest_reviewer');
SELECT pc_test.check('HUMAN_REVIEW', '承認後も review_reason_code = NEEDS_MORE_INFO が残る',
    (SELECT review_status = 'APPROVED' AND review_reason_code = 'NEEDS_MORE_INFO' FROM product_core.assessment WHERE assessment_id = 'AS-100000702'));
SELECT pc_test.check('HUMAN_REVIEW', 'review 欄の記入で AI 部分の content_hash は変わらない(review 欄は hash の対象外)',
    (SELECT content_hash = encode(sha256(convert_to((to_jsonb(a) - product_core.assessment_review_columns() - 'content_hash')::text, 'UTF8')), 'hex')
       FROM product_core.assessment a WHERE assessment_id = 'AS-100000702'));
SELECT pc_test.throws('HUMAN_REVIEW', 'review_batch は PENDING から直接 IMPORTED にできない',
    $q$UPDATE product_core.review_batch SET status = 'IMPORTED', reviewed_file_sha256 = repeat('f', 64), updated_by = 'human:op_rev' WHERE review_batch_id = 'RB-00000002'$q$,
    'pctest_reviewer', '変更できません');
SELECT pc_test.throws('HUMAN_REVIEW', 'Excel の束は記入済みファイルの SHA256 なしに REVIEWED にできない',
    $q$UPDATE product_core.review_batch SET status = 'REVIEWED', updated_by = 'human:op_rev' WHERE review_batch_id = 'RB-00000002'$q$, 'pctest_reviewer', '23514');
SELECT pc_test.lives('HUMAN_REVIEW', 'PENDING → REVIEWED → IMPORTED', ARRAY[
    $q$UPDATE product_core.review_batch SET status = 'REVIEWED', reviewed_file_sha256 = repeat('f', 64), updated_by = 'human:op_rev' WHERE review_batch_id = 'RB-00000002'$q$,
    $q$UPDATE product_core.review_batch SET status = 'IMPORTED', updated_by = 'human:op_rev' WHERE review_batch_id = 'RB-00000002'$q$], 'pctest_reviewer');
SELECT pc_test.throws('HUMAN_REVIEW', '取込済み(IMPORTED)の束は変更できない',
    $q$UPDATE product_core.review_batch SET note = 'x', updated_by = 'human:op_rev' WHERE review_batch_id = 'RB-00000002'$q$, 'pctest_reviewer', '取込済み');
SELECT pc_test.throws('HUMAN_REVIEW', 'review_batch は削除できない', $q$DELETE FROM product_core.review_batch WHERE review_batch_id = 'RB-00000002'$q$, NULL, '削除できません');
SELECT pc_test.throws('HUMAN_REVIEW', 'AI(取込ロール)は review_batch の状態を変えられない(列権限)',
    $q$UPDATE product_core.review_batch SET status = 'REVIEWED' WHERE review_batch_id = 'RB-00000001'$q$, 'pctest_ingest', '42501');
