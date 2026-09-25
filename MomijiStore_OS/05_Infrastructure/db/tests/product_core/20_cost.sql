-- =============================================================
--  20_cost.sql — 観測値・正式原価・税抜標準化・原価差異警告・NUMERIC 精度
-- =============================================================
\set ON_ERROR_STOP 1

-- ------------------------------------------------------------
--  OBSERVATION: 原価の観測値(P3・D12)
-- ------------------------------------------------------------
SELECT pc_test.lives('OBSERVATION', '取込ロールが税込・税抜・不明の観測値を元の値のまま記録できる',
    $q$INSERT INTO product_core.cost_observation (observation_id, source, source_ref, source_document_id, source_reliability, observed_at,
            observed_amount, tax_inclusion, tax_inclusion_basis, tax_rate, tax_rate_basis, amount_unit,
            subject_entity_id, subject_entity_type, subject_legacy_p, evidence, created_by, ingest_run_id) VALUES
       ('CO-100000001', 'invoice', 'inv#1', 'DOC-20260926-000001', 'HIGH', '2026-09-15', 1000, 'EXCLUDED', 'DOCUMENT_STATED', 0.10, 'DOCUMENT_STATED', 'PER_PP', 'PP-100001', 'PHYSICAL_PRODUCT', NULL, '{"line_unit_price":"1,000","tax_label":"税抜","issuer":"問屋A"}', 'ingest:t', 'IR-20260926-000002-abcd'),
       ('CO-100000002', 'legacy_product_master', 'master#2', NULL, 'MEDIUM', '2026-09-15', 1100, 'INCLUDED', 'OPERATIONAL_CONVENTION', NULL, 'NONE', 'PER_PP', NULL, NULL, 'P000001', '{"column":"標準原価"}', 'ingest:t', 'IR-20260926-000002-abcd'),
       ('CO-100000003', 'sku_string', 'sku#3', NULL, 'LOW', '2026-09-15', 1234, 'UNKNOWN', 'NONE', NULL, 'NONE', 'PER_PP', 'PP-100006', 'PHYSICAL_PRODUCT', NULL, '{"sku":"X@1234"}', 'ingest:t', 'IR-20260926-000002-abcd'),
       ('CO-100000004', 'delivery_note', 'dn#4', 'DOC-20260926-000002', 'HIGH', '2026-09-15', 1100, 'INCLUDED', 'DOCUMENT_STATED', 0.10, 'DOCUMENT_STATED', 'PER_PP', 'PP-100002', 'PHYSICAL_PRODUCT', NULL, '{"tax_label":"税込"}', 'ingest:t', 'IR-20260926-000002-abcd'),
       ('CO-100000005', 'invoice', 'inv#5', 'DOC-20260926-000003', 'HIGH', '2026-09-20', 1060, 'EXCLUDED', 'DOCUMENT_STATED', 0.10, 'DOCUMENT_STATED', 'PER_PP', 'PP-100001', 'PHYSICAL_PRODUCT', NULL, '{}', 'ingest:t', 'IR-20260926-000002-abcd'),
       ('CO-100000006', 'invoice', 'inv#6', 'DOC-20260926-000004', 'HIGH', '2026-09-20', 1050, 'EXCLUDED', 'DOCUMENT_STATED', 0.10, 'DOCUMENT_STATED', 'PER_PP', 'PP-100001', 'PHYSICAL_PRODUCT', NULL, '{}', 'ingest:t', 'IR-20260926-000002-abcd'),
       ('CO-100000007', 'invoice', 'inv#7', 'DOC-20260926-000005', 'HIGH', '2026-09-20', 1049, 'EXCLUDED', 'DOCUMENT_STATED', 0.10, 'DOCUMENT_STATED', 'PER_PP', 'PP-100001', 'PHYSICAL_PRODUCT', NULL, '{}', 'ingest:t', 'IR-20260926-000002-abcd'),
       ('CO-100000008', 'invoice', 'inv#8', 'DOC-20260926-000006', 'HIGH', '2026-09-20', 1284, 'EXCLUDED', 'DOCUMENT_STATED', 0.10, 'DOCUMENT_STATED', 'PER_PP', 'PP-100001', 'PHYSICAL_PRODUCT', NULL, '{}', 'ingest:t', 'IR-20260926-000002-abcd'),
       ('CO-100000009', 'invoice', 'inv#9', 'DOC-20260926-000007', 'HIGH', '2026-09-20', 1224, 'EXCLUDED', 'DOCUMENT_STATED', 0.10, 'DOCUMENT_STATED', 'PER_PP', 'PP-100001', 'PHYSICAL_PRODUCT', NULL, '{}', 'ingest:t', 'IR-20260926-000002-abcd'),
       ('CO-100000010', 'invoice', 'inv#10', 'DOC-20260926-000008', 'HIGH', '2026-09-20', 1100, 'EXCLUDED', 'DOCUMENT_STATED', 0.10, 'DOCUMENT_STATED', 'PER_PP', 'PP-100001', 'PHYSICAL_PRODUCT', NULL, '{}', 'ingest:t', 'IR-20260926-000002-abcd'),
       ('CO-100000011', 'legacy_product_master', 'master#11', NULL, 'MEDIUM', '2026-09-15', 1080, 'INCLUDED', 'OPERATIONAL_CONVENTION', NULL, 'NONE', 'PER_PP', 'PP-100003', 'PHYSICAL_PRODUCT', NULL, '{}', 'ingest:t', 'IR-20260926-000002-abcd'),
       ('CO-100000012', 'invoice', 'inv#12', 'DOC-20260926-000009', 'HIGH', '2026-09-20', 1000, 'EXCLUDED', 'DOCUMENT_STATED', 0.10, 'DOCUMENT_STATED', 'PER_LISTING_UNIT', 'LS-100001', 'LISTING', NULL, '{}', 'ingest:t', 'IR-20260926-000002-abcd'),
       ('CO-100000013', 'invoice', 'inv#13', 'DOC-20260926-000010', 'HIGH', '2026-09-20', 0, 'EXCLUDED', 'DOCUMENT_STATED', 0.10, 'DOCUMENT_STATED', 'PER_PP', 'PP-100001', 'PHYSICAL_PRODUCT', NULL, '{}', 'ingest:t', 'IR-20260926-000002-abcd'),
       ('CO-100000014', 'maker_document', 'mk#14', 'DOC-20260926-000011', 'HIGH', '2026-09-20', 0.12, 'EXCLUDED', 'DOCUMENT_STATED', 0.10, 'DOCUMENT_STATED', 'PER_PP', 'PP-100007', 'PHYSICAL_PRODUCT', NULL, '{"note":"100枚で0.12円"}', 'ingest:t', 'IR-20260926-000002-abcd'),
       ('CO-100000015', 'price_list', 'pl#15', 'DOC-20260926-000012', 'HIGH', '2026-09-20', 99999999.99, 'EXCLUDED', 'DOCUMENT_STATED', 0.10, 'DOCUMENT_STATED', 'PER_PP', 'PP-100008', 'PHYSICAL_PRODUCT', NULL, '{}', 'ingest:t', 'IR-20260926-000002-abcd'),
       ('CO-100000016', 'invoice', 'inv#16', 'DOC-20260926-000013', 'HIGH', '2026-09-20', 1800, 'EXCLUDED', 'DOCUMENT_STATED', 0.10, 'DOCUMENT_STATED', 'PER_COMPOSITION', 'CP-100001', 'COMPOSITION', NULL, '{}', 'ingest:t', 'IR-20260926-000002-abcd'),
       ('CO-100000017', 'delivery_note', 'dn#17', 'DOC-20260926-000014', 'HIGH', '2026-09-20', 1080, 'INCLUDED', 'DOCUMENT_STATED', 0.08, 'DOCUMENT_STATED', 'PER_PP', 'PP-100002', 'PHYSICAL_PRODUCT', NULL, '{"tax_label":"税込(軽減8%)"}', 'ingest:t', 'IR-20260926-000002-abcd'),
       ('CO-100000018', 'invoice', 'inv#18', 'DOC-20260926-000015', 'HIGH', '2026-09-20', 1980, 'INCLUDED', 'DOCUMENT_STATED', NULL, 'NONE', 'PER_COMPOSITION', 'CP-100001', 'COMPOSITION', NULL, '{}', 'ingest:t', 'IR-20260926-000002-abcd')$q$,
    'pctest_ingest');
SELECT pc_test.check('OBSERVATION', '元の金額をそのまま保存(1100 は 1100 のまま)',
    (SELECT observed_amount = 1100 AND tax_inclusion = 'INCLUDED' FROM product_core.cost_observation WHERE observation_id = 'CO-100000002'));
SELECT pc_test.throws('OBSERVATION', '書類の観測値は DocID 必須',
    $q$INSERT INTO product_core.cost_observation (source, source_ref, source_reliability, observed_at, observed_amount, tax_inclusion, tax_inclusion_basis,
            tax_rate_basis, amount_unit, subject_entity_id, subject_entity_type, evidence, created_by, ingest_run_id)
       VALUES ('invoice', 'x#1', 'HIGH', now(), 1, 'EXCLUDED', 'DOCUMENT_STATED', 'NONE', 'PER_PP', 'PP-100001', 'PHYSICAL_PRODUCT', '{}', 'ingest:t', 'IR-20260926-000002-abcd')$q$,
    'pctest_ingest', '23514');
SELECT pc_test.throws('OBSERVATION', '税区分 UNKNOWN に根拠を付けられない',
    $q$INSERT INTO product_core.cost_observation (source, source_ref, source_reliability, observed_at, observed_amount, tax_inclusion, tax_inclusion_basis,
            tax_rate_basis, amount_unit, subject_entity_id, subject_entity_type, evidence, created_by, ingest_run_id)
       VALUES ('sku_string', 'x#2', 'LOW', now(), 1, 'UNKNOWN', 'DOCUMENT_STATED', 'NONE', 'PER_PP', 'PP-100001', 'PHYSICAL_PRODUCT', '{}', 'ingest:t', 'IR-20260926-000002-abcd')$q$,
    'pctest_ingest', '23514');
SELECT pc_test.throws('OBSERVATION', '税率には根拠が必要',
    $q$INSERT INTO product_core.cost_observation (source, source_ref, source_reliability, observed_at, observed_amount, tax_inclusion, tax_inclusion_basis,
            tax_rate, tax_rate_basis, amount_unit, subject_entity_id, subject_entity_type, evidence, created_by, ingest_run_id)
       VALUES ('sku_string', 'x#3', 'LOW', now(), 1, 'INCLUDED', 'OPERATIONAL_CONVENTION', 0.10, 'NONE', 'PER_PP', 'PP-100001', 'PHYSICAL_PRODUCT', '{}', 'ingest:t', 'IR-20260926-000002-abcd')$q$,
    'pctest_ingest', '23514');
SELECT pc_test.throws('OBSERVATION', '税率は 0 以上 1 未満',
    $q$INSERT INTO product_core.cost_observation (source, source_ref, source_reliability, observed_at, observed_amount, tax_inclusion, tax_inclusion_basis,
            tax_rate, tax_rate_basis, amount_unit, subject_entity_id, subject_entity_type, evidence, created_by, ingest_run_id)
       VALUES ('supplier_csv', 'x#4', 'LOW', now(), 1, 'INCLUDED', 'SOURCE_SPEC', 1.5, 'SOURCE_SPEC', 'PER_PP', 'PP-100001', 'PHYSICAL_PRODUCT', '{}', 'ingest:t', 'IR-20260926-000002-abcd')$q$,
    'pctest_ingest');
SELECT pc_test.throws('OBSERVATION', '対象は1つだけ(実体と legacy P の両方は不可)',
    $q$INSERT INTO product_core.cost_observation (source, source_ref, source_reliability, observed_at, observed_amount, tax_inclusion, tax_inclusion_basis,
            tax_rate_basis, amount_unit, subject_entity_id, subject_entity_type, subject_legacy_p, evidence, created_by, ingest_run_id)
       VALUES ('sku_string', 'x#5', 'LOW', now(), 1, 'UNKNOWN', 'NONE', 'NONE', 'PER_PP', 'PP-100001', 'PHYSICAL_PRODUCT', 'P000001', '{}', 'ingest:t', 'IR-20260926-000002-abcd')$q$,
    'pctest_ingest', '23514');
SELECT pc_test.throws('OBSERVATION', 'P3: 観測値の金額は変更できない',
    $q$UPDATE product_core.cost_observation SET observed_amount = 1 WHERE observation_id = 'CO-100000001'$q$, NULL, '追記専用');
SELECT pc_test.throws('OBSERVATION', 'P3: 観測値の evidence は変更できない',
    $q$UPDATE product_core.cost_observation SET evidence = '{}' WHERE observation_id = 'CO-100000001'$q$, NULL, '追記専用');
SELECT pc_test.throws('OBSERVATION', 'P3: 観測値は削除できない',
    $q$DELETE FROM product_core.cost_observation WHERE observation_id = 'CO-100000001'$q$, NULL, '追記専用');
SELECT pc_test.throws('OBSERVATION', '同じ資料の同じ行は二重に取り込まれない(冪等性)',
    $q$INSERT INTO product_core.cost_observation (source, source_ref, source_document_id, source_reliability, observed_at, observed_amount, tax_inclusion,
            tax_inclusion_basis, tax_rate, tax_rate_basis, amount_unit, subject_entity_id, subject_entity_type, evidence, created_by, ingest_run_id)
       VALUES ('invoice', 'inv#1', 'DOC-20260926-000001', 'HIGH', now(), 1000, 'EXCLUDED', 'DOCUMENT_STATED', 0.10, 'DOCUMENT_STATED', 'PER_PP', 'PP-100001', 'PHYSICAL_PRODUCT', '{}', 'ingest:t', 'IR-20260926-000002-abcd')$q$,
    'pctest_ingest', '23505');

-- ------------------------------------------------------------
--  COST: 正式原価(C4・C8・C10・C11・S3・S5・S7・S8)
-- ------------------------------------------------------------
--  正規化の記録(EXCLUDED 1000 → 1000)
\set norm_excl '{"observed_amount":1000,"tax_inclusion":"EXCLUDED","tax_inclusion_basis":"DOCUMENT_STATED","tax_rate":0.10,"tax_inclusion_applied":"EXCLUDED","tax_rate_applied":null,"tax_rate_basis_applied":"NOT_APPLICABLE","unit_divisor":1,"formula":"1000"}'

SELECT pc_test.throws('COST', 'S5: 取込ロール(AI)は正式原価を作れない',
    format($q$INSERT INTO product_core.cost_history (pp_id, unit_cost_excl_tax, valid_from, source_observation_id, normalization, basis, assessment_id, approved_by)
       VALUES ('PP-100001', 1000, '2026-09-01', 'CO-100000001', %L, 'x', 'AS-100000004', 'human:op_cost')$q$, :'norm_excl'),
    'pctest_ingest', '42501');
SELECT pc_test.throws('COST', 'C1: COST_APPROVE の無い人は正式原価を作れない',
    format($q$INSERT INTO product_core.cost_history (pp_id, unit_cost_excl_tax, valid_from, source_observation_id, normalization, basis, assessment_id, approved_by)
       VALUES ('PP-100001', 1000, '2026-09-01', 'CO-100000001', %L, 'x', 'AS-100000004', 'human:op_rev')$q$, :'norm_excl'),
    'pctest_reviewer', 'COST_APPROVE');
SELECT pc_test.throws('COST', 'S5: 承認者に AI を書けない',
    format($q$INSERT INTO product_core.cost_history (pp_id, unit_cost_excl_tax, valid_from, source_observation_id, normalization, basis, assessment_id, approved_by)
       VALUES ('PP-100001', 1000, '2026-09-01', 'CO-100000001', %L, 'x', 'AS-100000004', 'ai:gpt')$q$, :'norm_excl'),
    'pctest_reviewer');
SELECT pc_test.throws('COST', 'S3: PROVISIONAL の PP には正式原価を作れない',
    format($q$INSERT INTO product_core.cost_history (pp_id, unit_cost_excl_tax, valid_from, source_observation_id, normalization, basis, assessment_id, approved_by)
       VALUES ('PP-100003', 1000, '2026-09-01', 'CO-100000001', %L, 'x', 'AS-100000004', 'human:op_cost')$q$, :'norm_excl'),
    'pctest_reviewer', 'PROVISIONAL');
SELECT pc_test.throws('COST', 'S3: RETIRED の PP には正式原価を作れない',
    format($q$INSERT INTO product_core.cost_history (pp_id, unit_cost_excl_tax, valid_from, source_observation_id, normalization, basis, assessment_id, approved_by)
       VALUES ('PP-100004', 1000, '2026-09-01', 'CO-100000001', %L, 'x', 'AS-100000004', 'human:op_cost')$q$, :'norm_excl'),
    'pctest_reviewer', 'RETIRED');
SELECT pc_test.throws('COST', 'C11: 人の review の無い判定を根拠にできない',
    format($q$INSERT INTO product_core.cost_history (pp_id, unit_cost_excl_tax, valid_from, source_observation_id, normalization, basis, assessment_id, approved_by)
       VALUES ('PP-100001', 1000, '2026-09-01', 'CO-100000001', %L, 'x', 'AS-100000003', 'human:op_cost')$q$, :'norm_excl'),
    'pctest_reviewer', '承認されていません');
SELECT pc_test.throws('COST', '他の商品の観測値は流用できない',
    format($q$INSERT INTO product_core.cost_history (pp_id, unit_cost_excl_tax, valid_from, source_observation_id, normalization, basis, assessment_id, approved_by)
       VALUES ('PP-100002', 1000, '2026-09-01', 'CO-100000001', %L, 'x', 'AS-100000004', 'human:op_cost')$q$, :'norm_excl'),
    'pctest_reviewer', 'のものではありません');
SELECT pc_test.throws('COST', 'S8: 換算記録の元金額が観測値と違えば拒否',
    format($q$INSERT INTO product_core.cost_history (pp_id, unit_cost_excl_tax, valid_from, source_observation_id, normalization, basis, assessment_id, approved_by)
       VALUES ('PP-100001', 1000, '2026-09-01', 'CO-100000001', %L, 'x', 'AS-100000004', 'human:op_cost')$q$,
       jsonb_set(:'norm_excl'::jsonb, '{observed_amount}', '999')),
    'pctest_reviewer', '一致しません');
SELECT pc_test.throws('COST', 'S8: 換算記録の税区分が観測値と違えば拒否',
    format($q$INSERT INTO product_core.cost_history (pp_id, unit_cost_excl_tax, valid_from, source_observation_id, normalization, basis, assessment_id, approved_by)
       VALUES ('PP-100001', 1000, '2026-09-01', 'CO-100000001', %L, 'x', 'AS-100000004', 'human:op_cost')$q$,
       jsonb_set(:'norm_excl'::jsonb, '{tax_inclusion}', '"INCLUDED"')),
    'pctest_reviewer', '一致しません');
SELECT pc_test.throws('COST', 'S8: 換算記録に必須キーが無ければ拒否',
    format($q$INSERT INTO product_core.cost_history (pp_id, unit_cost_excl_tax, valid_from, source_observation_id, normalization, basis, assessment_id, approved_by)
       VALUES ('PP-100001', 1000, '2026-09-01', 'CO-100000001', %L, 'x', 'AS-100000004', 'human:op_cost')$q$,
       (:'norm_excl'::jsonb - 'tax_rate')),
    'pctest_reviewer', '23514');
SELECT pc_test.throws('COST', '換算結果と一致しない正式原価を拒否',
    format($q$INSERT INTO product_core.cost_history (pp_id, unit_cost_excl_tax, valid_from, source_observation_id, normalization, basis, assessment_id, approved_by)
       VALUES ('PP-100001', 999, '2026-09-01', 'CO-100000001', %L, 'x', 'AS-100000004', 'human:op_cost')$q$, :'norm_excl'),
    'pctest_reviewer', '一致しません');
SELECT pc_test.lives('COST', '証拠の確かな税抜の観測値から正式原価を作れる',
    format($q$INSERT INTO product_core.cost_history (cost_id, pp_id, unit_cost_excl_tax, valid_from, source_observation_id, normalization, basis, assessment_id, approved_by)
       VALUES ('CH-10000001', 'PP-100001', 1000, '2026-09-01', 'CO-100000001', %L, '請求書 税抜', 'AS-100000004', 'human:op_cost')$q$, :'norm_excl'),
    'pctest_reviewer');
SELECT pc_test.check('COST', 'S8: 正式原価の作成後も観測値の元金額・税区分・evidence が残る',
    (SELECT observed_amount = 1000 AND tax_inclusion = 'EXCLUDED' AND evidence ->> 'issuer' = '問屋A'
       FROM product_core.cost_observation WHERE observation_id = 'CO-100000001')
    AND (SELECT normalization ->> 'observed_amount' = '1000' FROM product_core.cost_history WHERE cost_id = 'CH-10000001'));
SELECT pc_test.throws('COST', 'S8: 正式原価が参照する観測値は削除できない',
    $q$DELETE FROM product_core.cost_observation WHERE observation_id = 'CO-100000001'$q$);
SELECT pc_test.lives('COST', '税込(記載税率10%)の書類 1100 → 税抜 1000',
    $q$INSERT INTO product_core.cost_history (pp_id, unit_cost_excl_tax, valid_from, source_observation_id, normalization, basis, assessment_id, approved_by)
       VALUES ('PP-100002', 1000, '2026-09-01', 'CO-100000004',
               '{"observed_amount":1100,"tax_inclusion":"INCLUDED","tax_inclusion_basis":"DOCUMENT_STATED","tax_rate":0.10,"tax_inclusion_applied":"INCLUDED","tax_rate_applied":0.10,"tax_rate_basis_applied":"OBSERVED","unit_divisor":1,"formula":"1100/1.10"}',
               '納品書 税込', 'AS-100000004', 'human:op_cost')$q$, 'pctest_reviewer');
SELECT pc_test.throws('COST', 'C8: 同じ PP で ACTIVE の期間が重なる正式原価を拒否(EXCLUDE)',
    format($q$INSERT INTO product_core.cost_history (pp_id, unit_cost_excl_tax, valid_from, source_observation_id, normalization, basis, assessment_id, approved_by)
       VALUES ('PP-100001', 1000, '2026-10-01', 'CO-100000001', %L, 'x', 'AS-100000004', 'human:op_cost')$q$, :'norm_excl'),
    'pctest_reviewer', '23P01');
SELECT pc_test.lives('COST', '期間を閉じてから次の期間の正式原価を作れる', ARRAY[
    $q$UPDATE product_core.cost_history SET valid_to = '2026-09-30', updated_by = 'human:op_cost' WHERE cost_id = 'CH-10000001'$q$,
    $q$INSERT INTO product_core.cost_history (cost_id, pp_id, unit_cost_excl_tax, valid_from, source_observation_id, normalization, basis, assessment_id, approved_by)
       VALUES ('CH-10000002', 'PP-100001', 1060, '2026-10-01', 'CO-100000005',
               '{"observed_amount":1060,"tax_inclusion":"EXCLUDED","tax_inclusion_basis":"DOCUMENT_STATED","tax_rate":0.10,"tax_inclusion_applied":"EXCLUDED","tax_rate_applied":null,"tax_rate_basis_applied":"NOT_APPLICABLE","unit_divisor":1,"formula":"1060"}',
               '請求書 税抜', 'AS-100000004', 'human:op_cost')$q$], 'pctest_reviewer');
SELECT pc_test.throws('COST', 'C4: valid_to は一度だけ設定できる',
    $q$UPDATE product_core.cost_history SET valid_to = '2026-09-29', updated_by = 'human:op_cost' WHERE cost_id = 'CH-10000001'$q$, 'pctest_reviewer', '一度だけ');
SELECT pc_test.throws('COST', 'C4: 正式原価の金額は変更できない',
    $q$UPDATE product_core.cost_history SET unit_cost_excl_tax = 1 WHERE cost_id = 'CH-10000001'$q$, NULL, '変更できません');
SELECT pc_test.throws('COST', 'C4: 正式原価は削除できない', $q$DELETE FROM product_core.cost_history WHERE cost_id = 'CH-10000001'$q$, NULL, '削除できません');
SELECT pc_test.lives('COST', 'C4: 誤登録は VOID で取り消す(行は残る)',
    $q$UPDATE product_core.cost_history SET record_status = 'VOID', updated_by = 'human:op_cost' WHERE cost_id = 'CH-10000002'$q$, 'pctest_reviewer');
SELECT pc_test.throws('COST', 'VOID から戻せない',
    $q$UPDATE product_core.cost_history SET record_status = 'ACTIVE', updated_by = 'human:op_cost' WHERE cost_id = 'CH-10000002'$q$, 'pctest_reviewer');

-- S7: 税区分が不明・運用慣行だけの観測値
\set norm_unknown '{"observed_amount":1234,"tax_inclusion":"UNKNOWN","tax_inclusion_basis":"NONE","tax_rate":null,"tax_inclusion_applied":"INCLUDED","tax_rate_applied":0.10,"tax_rate_basis_applied":"HUMAN_CONFIRMED","unit_divisor":1,"formula":"1234/1.10"}'
SELECT pc_test.throws('COST', 'S7: 税区分 UNKNOWN の観測値は、人の確定なしに正式原価にできない',
    format($q$INSERT INTO product_core.cost_history (pp_id, unit_cost_excl_tax, valid_from, source_observation_id, normalization, basis, assessment_id, approved_by)
       VALUES ('PP-100006', 1121.818182, '2026-09-01', 'CO-100000003', %L, 'x', 'AS-100000004', 'human:op_cost')$q$, :'norm_unknown'),
    'pctest_reviewer', '確定していません');
SELECT pc_test.throws('COST', 'S7: 未 review の TAX_BASIS 判定では足りない',
    format($q$INSERT INTO product_core.cost_history (pp_id, unit_cost_excl_tax, valid_from, source_observation_id, normalization, basis, assessment_id, tax_basis_assessment_id, approved_by)
       VALUES ('PP-100006', 1121.818182, '2026-09-01', 'CO-100000003', %L, 'x', 'AS-100000004', 'AS-100000005', 'human:op_cost')$q$, :'norm_unknown'),
    'pctest_reviewer', '人が確定したもの');
SELECT pc_test.lives('COST', '人が review で税区分を「税込」に修正(CORRECTED)',
    $q$UPDATE product_core.assessment SET review_status = 'CORRECTED', reviewed_by = 'human:op_cost', corrected_value = 'INCLUDED',
           review_note = '仕入先に税込と確認' WHERE assessment_id = 'AS-100000005'$q$, 'pctest_reviewer');
SELECT pc_test.throws('COST', 'S7: 人が確定した税区分と違う税区分は適用できない',
    format($q$INSERT INTO product_core.cost_history (pp_id, unit_cost_excl_tax, valid_from, source_observation_id, normalization, basis, assessment_id, tax_basis_assessment_id, approved_by)
       VALUES ('PP-100006', 1234, '2026-09-01', 'CO-100000003', %L, 'x', 'AS-100000004', 'AS-100000005', 'human:op_cost')$q$,
       jsonb_set(jsonb_set(:'norm_unknown'::jsonb, '{tax_inclusion_applied}', '"EXCLUDED"'), '{tax_rate_basis_applied}', '"NOT_APPLICABLE"')),
    'pctest_reviewer', '一致しません');
SELECT pc_test.lives('COST', 'S7: 人の確定後は、確定した税区分で正式原価にできる(1234 → 1121.818182)',
    format($q$INSERT INTO product_core.cost_history (pp_id, unit_cost_excl_tax, valid_from, source_observation_id, normalization, basis, assessment_id, tax_basis_assessment_id, approved_by)
       VALUES ('PP-100006', 1121.818182, '2026-09-01', 'CO-100000003', %L, '人が税込と確定', 'AS-100000004', 'AS-100000005', 'human:op_cost')$q$, :'norm_unknown'),
    'pctest_reviewer');
\set norm_conv '{"observed_amount":1100,"tax_inclusion":"INCLUDED","tax_inclusion_basis":"OPERATIONAL_CONVENTION","tax_rate":null,"tax_inclusion_applied":"INCLUDED","tax_rate_applied":0.10,"tax_rate_basis_applied":"PP_TAX_CATEGORY","unit_divisor":1,"formula":"1100/1.10"}'
SELECT pc_test.throws('COST', 'S7: 運用慣行(原則税込)だけの観測値は、人の確定なしに正式原価にできない',
    format($q$INSERT INTO product_core.cost_history (pp_id, unit_cost_excl_tax, valid_from, source_observation_id, normalization, basis, assessment_id, approved_by)
       VALUES ('PP-100001', 1000, '2027-01-01', 'CO-100000002', %L, 'x', 'AS-100000004', 'human:op_cost')$q$, :'norm_conv'),
    'pctest_reviewer', '確定していません');
SELECT pc_test.lives('COST', '人が「税込」を承認すれば、legacy_mapping 経由の商品マスター原価を使える', ARRAY[
    $q$UPDATE product_core.assessment SET review_status = 'APPROVED', reviewed_by = 'human:op_cost' WHERE assessment_id = 'AS-100000006'$q$,
    format($q$INSERT INTO product_core.cost_history (pp_id, unit_cost_excl_tax, valid_from, source_observation_id, normalization, basis, assessment_id, tax_basis_assessment_id, approved_by)
       VALUES ('PP-100001', 1000, '2027-01-01', 'CO-100000002', %L, '商品マスター(人が税込と確定)', 'AS-100000004', 'AS-100000006', 'human:op_cost')$q$, :'norm_conv')],
    'pctest_reviewer');

-- ------------------------------------------------------------
--  COMPOSITION_COST: CP の正式原価(直接仕入)と A → B の優先順位
-- ------------------------------------------------------------
SELECT pc_test.lives('COMPOSITION_COST', '直接仕入のセット原価(税抜 1800)を正式原価にできる',
    $q$INSERT INTO product_core.composition_cost_history (cp_id, set_cost_excl_tax, purchase_evidence, valid_from, source_observation_id, normalization, basis, assessment_id, approved_by)
       VALUES ('CP-100001', 1800, '{"supplier":"問屋A","invoice_no":"123"}', '2026-09-01', 'CO-100000016',
               '{"observed_amount":1800,"tax_inclusion":"EXCLUDED","tax_inclusion_basis":"DOCUMENT_STATED","tax_rate":0.10,"tax_inclusion_applied":"EXCLUDED","tax_rate_applied":null,"tax_rate_basis_applied":"NOT_APPLICABLE","unit_divisor":1,"formula":"1800"}',
               'セットとして直接仕入', 'AS-100000004', 'human:op_cost')$q$, 'pctest_reviewer');
SELECT pc_test.throws('COMPOSITION_COST', 'CP には PP の税区分からの税率を使えない',
    $q$INSERT INTO product_core.composition_cost_history (cp_id, set_cost_excl_tax, purchase_evidence, valid_from, source_observation_id, normalization, basis, assessment_id, approved_by)
       VALUES ('CP-100001', 1800, '{}', '2027-01-01', 'CO-100000018',
               '{"observed_amount":1980,"tax_inclusion":"INCLUDED","tax_inclusion_basis":"DOCUMENT_STATED","tax_rate":null,"tax_inclusion_applied":"INCLUDED","tax_rate_applied":0.10,"tax_rate_basis_applied":"PP_TAX_CATEGORY","unit_divisor":1,"formula":"1980/1.10"}',
               'x', 'AS-100000004', 'human:op_cost')$q$, 'pctest_reviewer', 'PP の税区分');
SELECT pc_test.throws('COMPOSITION_COST', 'S3: PROVISIONAL の CP には正式原価を作れない',
    $q$INSERT INTO product_core.composition_cost_history (cp_id, set_cost_excl_tax, purchase_evidence, valid_from, source_observation_id, normalization, basis, assessment_id, approved_by)
       VALUES ('CP-100003', 1800, '{}', '2026-09-01', 'CO-100000016',
               '{"observed_amount":1800,"tax_inclusion":"EXCLUDED","tax_inclusion_basis":"DOCUMENT_STATED","tax_rate":0.10,"tax_inclusion_applied":"EXCLUDED","tax_rate_applied":null,"tax_rate_basis_applied":"NOT_APPLICABLE","unit_divisor":1,"formula":"1800"}',
               'x', 'AS-100000004', 'human:op_cost')$q$, 'pctest_reviewer', 'PROVISIONAL');
SELECT pc_test.check('COMPOSITION_COST', 'A 優先: 直接仕入の承認原価 1800 を使い、構成品計算 3000 との差を WARN',
    (SELECT method = 'A' AND cost = 1800 AND component_cost = 3000 AND variance_flag = 'WARN'
       FROM product_core.effective_composition_cost('CP-100001', '2026-09-15')));
SELECT pc_test.check('COMPOSITION_COST', 'A 期間中で構成品の原価が無い日は、A を使い差異は NO_BASELINE',
    (SELECT method = 'A' AND cost = 1800 AND component_cost IS NULL AND variance_flag = 'NO_BASELINE'
       FROM product_core.effective_composition_cost('CP-100001', '2026-10-15')));
SELECT pc_test.lives('COMPOSITION_COST', '直接仕入原価の期間を閉じる',
    $q$UPDATE product_core.composition_cost_history SET valid_to = '2026-12-31', updated_by = 'human:op_cost' WHERE cp_id = 'CP-100001'$q$, 'pctest_reviewer');
SELECT pc_test.check('COMPOSITION_COST', 'A の期間外は B: 構成品の承認原価 × 数量(1000 × 3)',
    (SELECT method = 'B' AND cost = 3000 FROM product_core.effective_composition_cost('CP-100001', '2027-02-01')));
SELECT pc_test.check('COMPOSITION_COST', 'A も B も無い日は算出不可(0 で埋めない)',
    (SELECT method = 'UNAVAILABLE' AND cost IS NULL FROM product_core.effective_composition_cost('CP-100001', '2026-08-01')));

-- ------------------------------------------------------------
--  NUMERIC: NUMERIC(18,6) の境界値(U2)
-- ------------------------------------------------------------
SELECT pc_test.check('NUMERIC', '通常原価: 税込 1100(10%)→ 税抜 1000 ちょうど', round(1100::numeric(18,6) / 1.10, 6) = 1000);
SELECT pc_test.check('NUMERIC', '軽減税率: 税込 1080(8%)→ 税抜 1000 ちょうど', round(1080::numeric(18,6) / 1.08, 6) = 1000);
SELECT pc_test.check('NUMERIC', '軽減税率の観測値を関数で標準化 → 1000',
    (SELECT excl_tax_amount = 1000 AND certainty = 'CONFIRMED' FROM product_core.normalize_observation('CO-100000017')));
SELECT pc_test.check('NUMERIC', '税率 0.10 / 0.08 は NUMERIC(7,6) で正確に保存', 0.10::numeric(7,6) = 0.1 AND 0.08::numeric(7,6) = 0.08);
SELECT pc_test.check('NUMERIC', '途中で丸めない: (1100 / 1.1) × 3 = 3000 ちょうど', (1100::numeric(18,6) / 1.1) * 3 = 3000);
SELECT pc_test.check('NUMERIC', '税込→税抜の割り切れない換算: 1000 / 1.1 は保存時に 909.090909',
    round(1000::numeric / 1.1, 6) = 909.090909);
SELECT pc_test.check('NUMERIC', '税込に戻して円表示すると 1000 に戻る', round(909.090909 * 1.1, 0) = 1000);
SELECT pc_test.check('NUMERIC', '非常に小さい単価 0.000001 は正確に保存', 0.000001::numeric(18,6) = 0.000001);
SELECT pc_test.check('NUMERIC', '【境界】0.0000004 は保存時に 0 へ丸まる(7桁目以下は保存しない)', 0.0000004::numeric(18,6) = 0);
SELECT pc_test.check('NUMERIC', '【境界】0.0000005 は保存時に 0.000001 へ丸まる(四捨五入)', 0.0000005::numeric(18,6) = 0.000001);
SELECT pc_test.lives('NUMERIC', '非常に小さい単価を正式原価に(0.12円 ÷ 100枚 = 0.0012)',
    $q$INSERT INTO product_core.cost_history (cost_id, pp_id, unit_cost_excl_tax, valid_from, source_observation_id, normalization, basis, assessment_id, approved_by)
       VALUES ('CH-10000010', 'PP-100007', 0.0012, '2026-09-01', 'CO-100000014',
               '{"observed_amount":0.12,"tax_inclusion":"EXCLUDED","tax_inclusion_basis":"DOCUMENT_STATED","tax_rate":0.10,"tax_inclusion_applied":"EXCLUDED","tax_rate_applied":null,"tax_rate_basis_applied":"NOT_APPLICABLE","unit_divisor":100,"formula":"0.12/100"}',
               'メーカー資料(100枚単位)', 'AS-100000004', 'human:op_cost')$q$, 'pctest_reviewer');
SELECT pc_test.check('NUMERIC', '小さい単価が正確に保存されている', (SELECT unit_cost_excl_tax = 0.0012 FROM product_core.cost_history WHERE cost_id = 'CH-10000010'));
SELECT pc_test.lives('NUMERIC', '大きな原価(99,999,999.99)を正式原価に',
    $q$INSERT INTO product_core.cost_history (cost_id, pp_id, unit_cost_excl_tax, valid_from, source_observation_id, normalization, basis, assessment_id, approved_by)
       VALUES ('CH-10000011', 'PP-100008', 99999999.99, '2026-09-01', 'CO-100000015',
               '{"observed_amount":99999999.99,"tax_inclusion":"EXCLUDED","tax_inclusion_basis":"DOCUMENT_STATED","tax_rate":0.10,"tax_inclusion_applied":"EXCLUDED","tax_rate_applied":null,"tax_rate_basis_applied":"NOT_APPLICABLE","unit_divisor":1,"formula":"99999999.99"}',
               '価格表', 'AS-100000004', 'human:op_cost')$q$, 'pctest_reviewer');
SELECT pc_test.check('NUMERIC', 'Composition 原価合計: 0.0012 + 99,999,999.99 を途中で丸めずに合計',
    (SELECT method = 'B' AND cost = 99999999.9912 FROM product_core.effective_composition_cost('CP-100005', '2026-09-15')));
SELECT pc_test.check('NUMERIC', '大きなセット原価: 12,000,000 × 999 = 11,988,000,000 が保存できる',
    (12000000::numeric(18,6) * 999)::numeric(18,6) = 11988000000);
SELECT pc_test.check('NUMERIC', 'NUMERIC(18,6) の上限 999,999,999,999.999999 が保存できる',
    999999999999.999999::numeric(18,6) = 999999999999.999999);
SELECT pc_test.throws('NUMERIC', '上限を超える金額は黙って丸めず拒否(1兆円)', $q$SELECT 1000000000000::numeric(18,6)$q$, NULL, '22003');
SELECT pc_test.check('NUMERIC', '100個の税抜換算の合計誤差は保存精度の範囲内(0.00005 以下)で、円表示は一致',
    abs(100 * round(1 / 1.1, 6) - round(100 / 1.1, 6)) <= 0.00005
    AND round(100 * round(1 / 1.1, 6), 0) = round(100 / 1.1, 0));

-- ------------------------------------------------------------
--  WARN: 原価差異警告の2段階(CONFIRMED_WARN / REFERENCE_WARN)
-- ------------------------------------------------------------
SELECT pc_test.check('WARN', '書類の税抜 → 標準化の確度は CONFIRMED',
    (SELECT certainty = 'CONFIRMED' AND excl_tax_amount = 1000 FROM product_core.normalize_observation('CO-100000001')));
SELECT pc_test.check('WARN', '書類の税込(税率明記)→ CONFIRMED で 1000',
    (SELECT certainty = 'CONFIRMED' AND excl_tax_amount = 1000 FROM product_core.normalize_observation('CO-100000004')));
SELECT pc_test.check('WARN', '商品マスター(原則税込・運用慣行)+ PP の税区分 → REFERENCE で 1000',
    (SELECT certainty = 'REFERENCE' AND excl_tax_amount = 1000 FROM product_core.normalize_observation('CO-100000002', 'PP-100001')));
SELECT pc_test.check('WARN', '商品マスター 1100(税込)と納品書 1000(税抜)は税抜 1000 同士で比較され、差は無い(OK)',
    (SELECT verdict = 'OK' AND diff = 0 FROM product_core.cost_variance('CO-100000002', 'CO-100000001', 'PP-100001')));
SELECT pc_test.check('WARN', '推定を含む比較で 50円以上 かつ 5%以上 → REFERENCE_WARN(CONFIRMED と区別)',
    (SELECT verdict = 'REFERENCE_WARN' FROM product_core.cost_variance('CO-100000010', 'CO-100000002', 'PP-100001')));
SELECT pc_test.check('WARN', '双方が証拠で確認済み、差 60円・6% → CONFIRMED_WARN',
    (SELECT verdict = 'CONFIRMED_WARN' FROM product_core.cost_variance('CO-100000005', 'CO-100000001')));
SELECT pc_test.check('WARN', '境界: 差 50円・5.0% ちょうど → CONFIRMED_WARN(以上を含む)',
    (SELECT verdict = 'CONFIRMED_WARN' FROM product_core.cost_variance('CO-100000006', 'CO-100000001')));
SELECT pc_test.check('WARN', '境界: 差 49円 → OK(金額条件を満たさない)',
    (SELECT verdict = 'OK' FROM product_core.cost_variance('CO-100000007', 'CO-100000001')));
SELECT pc_test.check('WARN', '境界: 差 60円でも 4.9% → OK(率の条件を満たさない)',
    (SELECT verdict = 'OK' FROM product_core.cost_variance('CO-100000008', 'CO-100000009')));
SELECT pc_test.check('WARN', '税区分 UNKNOWN を含む比較は正式な警告を出さない(TAX_BASIS_UNKNOWN)',
    (SELECT verdict = 'TAX_BASIS_UNKNOWN' AND diff IS NULL FROM product_core.cost_variance('CO-100000003', 'CO-100000001')));
SELECT pc_test.check('WARN', '税率が決まらない税込の観測値は警告を出さない(TAX_RATE_UNKNOWN)',
    (SELECT verdict = 'TAX_RATE_UNKNOWN' FROM product_core.cost_variance('CO-100000011', 'CO-100000001', 'PP-100003')));
SELECT pc_test.check('WARN', '単位が違う観測値同士は比較しない(UNIT_UNKNOWN)',
    (SELECT verdict = 'UNIT_UNKNOWN' FROM product_core.cost_variance('CO-100000012', 'CO-100000001')));
SELECT pc_test.check('WARN', '基準値 0 は比較しない(NO_BASELINE)',
    (SELECT verdict = 'NO_BASELINE' FROM product_core.cost_variance('CO-100000001', 'CO-100000013')));
SELECT pc_test.lives('WARN', '判定記録に CONFIRMED_WARN / REFERENCE_WARN を残せる',
    $q$INSERT INTO product_core.assessment (assessment_type, subject_key, verdict, confidence, reason, evidence, rule_version, assessed_by) VALUES
       ('COST_VARIANCE', 'CO-100000005', 'CONFIRMED_WARN', 'HIGH', '差60円・6%', '{}', 'cost-var/1.0', 'rule:cost-var/1.0'),
       ('COST_VARIANCE', 'CO-100000010', 'REFERENCE_WARN', 'MEDIUM', '推定を含む比較', '{}', 'cost-var/1.0', 'rule:cost-var/1.0')$q$, 'pctest_ingest');
