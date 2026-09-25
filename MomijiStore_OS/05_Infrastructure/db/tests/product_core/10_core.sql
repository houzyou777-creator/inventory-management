-- =============================================================
--  10_core.sql — 構造・操作者・実体・出品・識別子・判定・対応付け・参照の制約テスト
-- =============================================================
--  すべて仮のデータ(実データは投入しない)。ID は 100000 番台を明示して使う。
\set ON_ERROR_STOP 1

-- ------------------------------------------------------------
--  STRUCT: 構造(btree_gist・テーブル・数値型)
-- ------------------------------------------------------------
SELECT pc_test.check('STRUCT', 'btree_gist 拡張が有効',
    EXISTS (SELECT 1 FROM pg_extension WHERE extname = 'btree_gist'));
SELECT pc_test.check('STRUCT', 'Phase 1 の論理13表 + 構造用3表がそろう',
    (SELECT count(*) FROM information_schema.tables WHERE table_schema = 'product_core' AND table_type = 'BASE TABLE'
        AND table_name IN ('physical_product', 'identifier', 'identifier_link', 'legacy_mapping', 'product_relationship',
                           'composition', 'composition_component', 'listing', 'listing_group',
                           'listing_reference_history', 'cost_history', 'composition_cost_history',
                           'cost_observation', 'assessment', 'core_entity', 'operator', 'operator_permission')) = 17);
SELECT pc_test.check('STRUCT', '取込ステージング4表がそろう',
    (SELECT count(*) FROM information_schema.tables WHERE table_schema = 'legacy_ingest'
        AND table_name IN ('ingest_run', 'legacy_product_snapshot', 'legacy_listing_snapshot', 'external_file_snapshot')) = 4);
SELECT pc_test.check('STRUCT', 'FLOAT 系の列が1つも無い',
    NOT EXISTS (SELECT 1 FROM information_schema.columns WHERE table_schema IN ('product_core', 'legacy_ingest')
                 AND data_type IN ('real', 'double precision')));
SELECT pc_test.check('STRUCT', '金額列は NUMERIC(18,6)',
    (SELECT count(*) FROM information_schema.columns WHERE table_schema = 'product_core'
        AND column_name IN ('observed_amount', 'unit_cost_excl_tax', 'set_cost_excl_tax')
        AND data_type = 'numeric' AND numeric_precision = 18 AND numeric_scale = 6) = 3);
SELECT pc_test.check('STRUCT', '税率列は NUMERIC(7,6)',
    EXISTS (SELECT 1 FROM information_schema.columns WHERE table_schema = 'product_core' AND table_name = 'cost_observation'
             AND column_name = 'tax_rate' AND data_type = 'numeric' AND numeric_precision = 7 AND numeric_scale = 6));
SELECT pc_test.check('STRUCT', 'public / intelligence スキーマに product_core の表を作っていない',
    NOT EXISTS (SELECT 1 FROM information_schema.tables WHERE table_schema IN ('public', 'intelligence')
                 AND table_name IN ('physical_product', 'composition', 'listing', 'cost_history')));

-- ------------------------------------------------------------
--  OPERATOR: 承認者と権限(C1・D15)
-- ------------------------------------------------------------
SELECT pc_test.throws('OPERATOR', '取込ロールは operator を登録できない',
    $q$INSERT INTO product_core.operator (operator_id, display_name, created_by) VALUES ('op_admin', 'x', 'human:op_admin')$q$,
    'pctest_ingest', '42501');
SELECT pc_test.throws('OPERATOR', '最初の operator を他人の名義で登録できない',
    $q$INSERT INTO product_core.operator (operator_id, display_name, created_by) VALUES ('op_admin', '管理者', 'human:someone')$q$,
    'pctest_reviewer', '本人の名義');
SELECT pc_test.lives('OPERATOR', '最初の管理者を本人名義で登録できる(ブートストラップ)',
    $q$INSERT INTO product_core.operator (operator_id, display_name, created_by) VALUES ('op_admin', '管理者', 'human:op_admin')$q$,
    'pctest_reviewer');
SELECT pc_test.lives('OPERATOR', '最初の OPERATOR_ADMIN を本人に付与できる(ブートストラップ)',
    $q$INSERT INTO product_core.operator_permission (operator_id, permission, granted_by) VALUES ('op_admin', 'OPERATOR_ADMIN', 'human:op_admin')$q$,
    'pctest_reviewer');
SELECT pc_test.lives('OPERATOR', '管理者が他の権限と operator を登録できる', ARRAY[
    $q$INSERT INTO product_core.operator_permission (operator_id, permission, granted_by)
       SELECT 'op_admin', p, 'human:op_admin' FROM unnest(ARRAY['PRODUCT_APPROVE','LEGACY_MAPPING_APPROVE','RELATIONSHIP_APPROVE',
              'IDENTIFIER_APPROVE','LISTING_REFERENCE_APPROVE','COST_APPROVE','REVIEW']) p$q$,
    $q$INSERT INTO product_core.operator (operator_id, display_name, created_by) VALUES
       ('op_rev', '確認担当', 'human:op_admin'), ('op_cost', '原価担当', 'human:op_admin'),
       ('op_prod', '商品担当', 'human:op_admin'), ('op_gone', '退職者', 'human:op_admin')$q$,
    $q$INSERT INTO product_core.operator_permission (operator_id, permission, granted_by) VALUES
       ('op_rev', 'REVIEW', 'human:op_admin'),
       ('op_cost', 'REVIEW', 'human:op_admin'), ('op_cost', 'COST_APPROVE', 'human:op_admin'),
       ('op_prod', 'REVIEW', 'human:op_admin'), ('op_prod', 'PRODUCT_APPROVE', 'human:op_admin'),
       ('op_gone', 'REVIEW', 'human:op_admin'), ('op_gone', 'PRODUCT_APPROVE', 'human:op_admin'),
       ('op_gone', 'LEGACY_MAPPING_APPROVE', 'human:op_admin'), ('op_gone', 'COST_APPROVE', 'human:op_admin')$q$,
    $q$UPDATE product_core.operator SET status = 'INACTIVE', updated_by = 'human:op_admin' WHERE operator_id = 'op_gone'$q$],
    'pctest_reviewer');
SELECT pc_test.throws('OPERATOR', '管理者以外は operator を登録できない',
    $q$INSERT INTO product_core.operator (operator_id, display_name, created_by) VALUES ('op_x', 'x', 'human:op_rev')$q$,
    'pctest_reviewer', 'OPERATOR_ADMIN');
SELECT pc_test.throws('OPERATOR', '管理者がいる状態で自己付与(ブートストラップの悪用)はできない',
    $q$INSERT INTO product_core.operator_permission (operator_id, permission, granted_by) VALUES ('op_rev', 'OPERATOR_ADMIN', 'human:op_rev')$q$,
    'pctest_reviewer', 'OPERATOR_ADMIN');
SELECT pc_test.throws('OPERATOR', '権限の付与記録は削除できない',
    $q$DELETE FROM product_core.operator_permission WHERE operator_id = 'op_rev'$q$, NULL, '削除できません');
SELECT pc_test.throws('OPERATOR', 'operator は削除できない',
    $q$DELETE FROM product_core.operator WHERE operator_id = 'op_gone'$q$, NULL, '削除できません');
SELECT pc_test.lives('OPERATOR', '管理者は権限を取り消せる', ARRAY[
    $q$INSERT INTO product_core.operator_permission (operator_id, permission, granted_by) VALUES ('op_rev', 'RELATIONSHIP_APPROVE', 'human:op_admin')$q$,
    $q$UPDATE product_core.operator_permission SET revoked_by = 'human:op_admin'
        WHERE operator_id = 'op_rev' AND permission = 'RELATIONSHIP_APPROVE' AND revoked_at IS NULL$q$], 'pctest_reviewer');
SELECT pc_test.throws('OPERATOR', '取り消し済みの付与は再び変更できない',
    $q$UPDATE product_core.operator_permission SET revoked_by = 'human:op_admin'
        WHERE operator_id = 'op_rev' AND permission = 'RELATIONSHIP_APPROVE'$q$, 'pctest_reviewer', '取り消し');
SELECT pc_test.check('OPERATOR', '取り消した権限は無効', NOT product_core.has_permission('human:op_rev', 'RELATIONSHIP_APPROVE'));
SELECT pc_test.check('OPERATOR', 'INACTIVE の operator は権限が無効', NOT product_core.has_permission('human:op_gone', 'PRODUCT_APPROVE'));

-- ------------------------------------------------------------
--  INGEST: 取込ステージング
-- ------------------------------------------------------------
SELECT pc_test.lives('INGEST', '取込の開始とスナップショットの追加', ARRAY[
    $q$INSERT INTO legacy_ingest.ingest_run (ingest_run_id, actor) VALUES ('IR-20260926-000000-abcd', 'human:op_admin')$q$,
    $q$INSERT INTO legacy_ingest.legacy_product_snapshot (ingest_run_id, row_no, legacy_p, name, raw)
       VALUES ('IR-20260926-000000-abcd', 2, 'P000001', 'テスト単品', '{}')$q$], 'pctest_ingest');
SELECT pc_test.throws('INGEST', 'スナップショットは変更できない(追記専用)',
    $q$UPDATE legacy_ingest.legacy_product_snapshot SET name = 'x'$q$, NULL, '追記専用');
SELECT pc_test.throws('INGEST', 'スナップショットは削除できない(追記専用)',
    $q$DELETE FROM legacy_ingest.legacy_product_snapshot$q$, NULL, '追記専用');
SELECT pc_test.throws('INGEST', 'ingest_run の実行者は人でなければならない',
    $q$INSERT INTO legacy_ingest.ingest_run (ingest_run_id, actor) VALUES ('IR-20260926-000001-abcd', 'ai:gpt')$q$, 'pctest_ingest', '23514');
SELECT pc_test.lives('INGEST', '取込の完了', $q$UPDATE legacy_ingest.ingest_run SET status = 'SUCCEEDED', finished_at = now()
    WHERE ingest_run_id = 'IR-20260926-000000-abcd'$q$, 'pctest_ingest');
SELECT pc_test.throws('INGEST', '完了した取込は変更できない', $q$UPDATE legacy_ingest.ingest_run SET note = 'x'
    WHERE ingest_run_id = 'IR-20260926-000000-abcd'$q$, 'pctest_ingest', '完了');
SELECT pc_test.lives('INGEST', '以降のテスト用の取込を開始',
    $q$INSERT INTO legacy_ingest.ingest_run (ingest_run_id, actor) VALUES ('IR-20260926-000002-abcd', 'human:op_admin')$q$, 'pctest_ingest');

-- ------------------------------------------------------------
--  ENTITY: 物理商品(P5・C6・C12・C13・S1)
-- ------------------------------------------------------------
SELECT pc_test.lives('ENTITY', '取込ロールが PROVISIONAL の PP を作れる',
    $q$INSERT INTO product_core.physical_product (pp_id, name, origin_key, created_by) VALUES
       ('PP-100001', '単品A', 'legacy:P000001', 'ingest:test/1'), ('PP-100002', '単品B', 'legacy:P000002', 'ingest:test/1'),
       ('PP-100003', '未確定C', 'legacy:P000003', 'ingest:test/1'), ('PP-100004', '退役予定D', 'legacy:P000004', 'ingest:test/1'),
       ('PP-100006', '単品F', 'legacy:P000006', 'ingest:test/1'), ('PP-100007', '単品G', 'legacy:P000007', 'ingest:test/1'),
       ('PP-100008', '単品H', 'legacy:P000008', 'ingest:test/1'), ('PP-100009', '退役予定I', 'legacy:P000009', 'ingest:test/1')$q$, 'pctest_ingest');
SELECT pc_test.lives('ENTITY', '既定の採番(PP- + 6桁)', $q$INSERT INTO product_core.physical_product (name, origin_key, created_by)
    VALUES ('採番テスト', 'test:default-id', 'ingest:test/1')$q$, 'pctest_ingest');
SELECT pc_test.check('ENTITY', '採番された ID の形式', (SELECT pp_id ~ '^PP-[0-9]{6}$' FROM product_core.physical_product WHERE origin_key = 'test:default-id'));
SELECT pc_test.check('ENTITY', 'PP を作ると core_entity に同じ種類で登録される',
    EXISTS (SELECT 1 FROM product_core.core_entity WHERE entity_id = 'PP-100001' AND entity_type = 'PHYSICAL_PRODUCT'));
SELECT pc_test.throws('ENTITY', '取込ロールは status を指定できない(列権限)',
    $q$INSERT INTO product_core.physical_product (pp_id, name, origin_key, created_by, status) VALUES ('PP-100090', 'x', 'k:90', 'ingest:t', 'ACTIVE')$q$,
    'pctest_ingest', '42501');
SELECT pc_test.throws('ENTITY', 'ACTIVE の PP を直接作れない(トリガー)',
    $q$INSERT INTO product_core.physical_product (pp_id, name, origin_key, created_by, status, activated_by, activated_at)
       VALUES ('PP-100091', 'x', 'k:91', 'ingest:t', 'ACTIVE', 'human:op_prod', now())$q$, NULL, 'PROVISIONAL');
SELECT pc_test.throws('ENTITY', 'C13: 既存 P 番号を主キーにできない',
    $q$INSERT INTO product_core.physical_product (pp_id, name, origin_key, created_by) VALUES ('P000123', 'x', 'k:92', 'ingest:t')$q$, 'pctest_ingest', '23514');
SELECT pc_test.throws('ENTITY', 'S1: PP に別の entity_type を付けられない',
    $q$INSERT INTO product_core.physical_product (pp_id, entity_type, name, origin_key, created_by) VALUES ('PP-100093', 'COMPOSITION', 'x', 'k:93', 'ingest:t')$q$);
SELECT pc_test.throws('ENTITY', 'S1: 取込ロールは core_entity に直接登録できない',
    $q$INSERT INTO product_core.core_entity (entity_id, entity_type, created_by) VALUES ('PP-100094', 'PHYSICAL_PRODUCT', 'ingest:t')$q$, 'pctest_ingest', '42501');
SELECT pc_test.throws('ENTITY', 'S1: core_entity の ID 接頭辞と種類の食い違いを拒否',
    $q$INSERT INTO product_core.core_entity (entity_id, entity_type, created_by) VALUES ('PP-100095', 'COMPOSITION', 'ingest:t')$q$, NULL, '23514');
SELECT pc_test.throws('ENTITY', 'S1: 実体の無い core_entity は COMMIT 時に拒否(遅延制約)',
    $q$INSERT INTO product_core.core_entity (entity_id, entity_type, created_by) VALUES ('PP-100096', 'PHYSICAL_PRODUCT', 'ingest:t')$q$, NULL, '対応する実体');
SELECT pc_test.throws('ENTITY', 'C12: 列挙値の外(kind)を拒否',
    $q$INSERT INTO product_core.physical_product (pp_id, name, kind, origin_key, created_by) VALUES ('PP-100097', 'x', 'FOOD', 'k:97', 'ingest:t')$q$, 'pctest_ingest', '23514');
SELECT pc_test.throws('ENTITY', 'origin_key の重複を拒否(冪等性)',
    $q$INSERT INTO product_core.physical_product (pp_id, name, origin_key, created_by) VALUES ('PP-100098', 'x', 'legacy:P000001', 'ingest:t')$q$, 'pctest_ingest', '23505');

SELECT pc_test.throws('ENTITY', 'S5: 取込ロールは ACTIVE 化できない(列権限)',
    $q$UPDATE product_core.physical_product SET status = 'ACTIVE', activated_by = 'human:op_prod' WHERE pp_id = 'PP-100001'$q$, 'pctest_ingest', '42501');
SELECT pc_test.throws('ENTITY', 'S5: 承認者に AI を書けない',
    $q$UPDATE product_core.physical_product SET status = 'ACTIVE', activated_by = 'ai:gpt', updated_by = 'human:op_prod' WHERE pp_id = 'PP-100001'$q$, 'pctest_reviewer');
SELECT pc_test.throws('ENTITY', 'S5: 存在しない operator では承認できない',
    $q$UPDATE product_core.physical_product SET status = 'ACTIVE', activated_by = 'human:ghost', updated_by = 'human:ghost' WHERE pp_id = 'PP-100001'$q$, 'pctest_reviewer', 'PRODUCT_APPROVE');
SELECT pc_test.throws('ENTITY', 'S5: INACTIVE の operator では承認できない',
    $q$UPDATE product_core.physical_product SET status = 'ACTIVE', activated_by = 'human:op_gone', updated_by = 'human:op_gone' WHERE pp_id = 'PP-100001'$q$, 'pctest_reviewer', 'PRODUCT_APPROVE');
SELECT pc_test.throws('ENTITY', 'C1: 権限の無い operator では承認できない',
    $q$UPDATE product_core.physical_product SET status = 'ACTIVE', activated_by = 'human:op_rev', updated_by = 'human:op_rev' WHERE pp_id = 'PP-100001'$q$, 'pctest_reviewer', 'PRODUCT_APPROVE');
SELECT pc_test.lives('ENTITY', 'PRODUCT_APPROVE を持つ人が ACTIVE 化・税区分設定できる',
    $q$UPDATE product_core.physical_product SET status = 'ACTIVE', activated_by = 'human:op_prod', updated_by = 'human:op_prod', tax_category = 'STANDARD'
        WHERE pp_id IN ('PP-100001', 'PP-100002', 'PP-100004', 'PP-100006', 'PP-100007', 'PP-100008', 'PP-100009')$q$, 'pctest_reviewer');
SELECT pc_test.check('ENTITY', 'C2: 承認日時は承認者と同時に記録される',
    (SELECT bool_and(activated_at IS NOT NULL) FROM product_core.physical_product WHERE status = 'ACTIVE'));
SELECT pc_test.throws('ENTITY', '取込ロールは税区分を変えられない(列権限)',
    $q$UPDATE product_core.physical_product SET tax_category = 'REDUCED' WHERE pp_id = 'PP-100003'$q$, 'pctest_ingest', '42501');
SELECT pc_test.lives('ENTITY', '取込ロールは名称を更新できる',
    $q$UPDATE product_core.physical_product SET name = '単品A(改)', updated_by = 'ingest:test/2' WHERE pp_id = 'PP-100001'$q$, 'pctest_ingest');
SELECT pc_test.throws('ENTITY', 'C6: ACTIVE → PROVISIONAL へ戻せない',
    $q$UPDATE product_core.physical_product SET status = 'PROVISIONAL', updated_by = 'human:op_prod' WHERE pp_id = 'PP-100002'$q$, 'pctest_reviewer', '変更できません');
SELECT pc_test.lives('ENTITY', 'ACTIVE → RETIRED(退役)',
    $q$UPDATE product_core.physical_product SET status = 'RETIRED', retired_by = 'human:op_prod', retire_reason = 'テスト', updated_by = 'human:op_prod'
        WHERE pp_id = 'PP-100004'$q$, 'pctest_reviewer');
SELECT pc_test.throws('ENTITY', 'C6: RETIRED → ACTIVE へ戻せない',
    $q$UPDATE product_core.physical_product SET status = 'ACTIVE', activated_by = 'human:op_prod', updated_by = 'human:op_prod' WHERE pp_id = 'PP-100004'$q$, 'pctest_reviewer');
SELECT pc_test.throws('ENTITY', 'C6: PP は削除できない', $q$DELETE FROM product_core.physical_product WHERE pp_id = 'PP-100003'$q$, NULL, '削除できません');
SELECT pc_test.throws('ENTITY', 'ID・origin_key は変更できない',
    $q$UPDATE product_core.physical_product SET origin_key = 'x' WHERE pp_id = 'PP-100003'$q$, NULL, '変更できません');

-- ------------------------------------------------------------
--  COMPOSITION: 販売構成(C9)
-- ------------------------------------------------------------
SELECT pc_test.lives('COMPOSITION', '取込ロールが CP と構成品の下書きを作れる', ARRAY[
    $q$INSERT INTO product_core.composition (cp_id, name, comp_type, origin_key, created_by) VALUES
       ('CP-100001', 'A 3個セット', 'FIXED', 'legacy:P000901', 'ingest:t'),
       ('CP-100002', '福袋(構成未確認)', 'UNCLASSIFIED', 'legacy:P000902', 'ingest:t'),
       ('CP-100003', '未確定品入りセット', 'FIXED', 'legacy:P000903', 'ingest:t'),
       ('CP-100004', '構成品なしセット', 'FIXED', 'legacy:P000904', 'ingest:t'),
       ('CP-100005', '本体+詰替', 'FIXED', 'legacy:P000905', 'ingest:t')$q$,
    $q$INSERT INTO product_core.composition_component (cp_id, pp_id, quantity, evidence, created_by) VALUES
       ('CP-100001', 'PP-100001', 3, '{"from":"商品名 3個セット"}', 'ingest:t'),
       ('CP-100003', 'PP-100003', 2, '{}', 'ingest:t'),
       ('CP-100005', 'PP-100007', 1, '{}', 'ingest:t'), ('CP-100005', 'PP-100008', 1, '{}', 'ingest:t')$q$], 'pctest_ingest');
SELECT pc_test.throws('COMPOSITION', 'C9: UNCLASSIFIED(福袋)は構成品を持てない',
    $q$INSERT INTO product_core.composition_component (cp_id, pp_id, quantity, evidence, created_by) VALUES ('CP-100002', 'PP-100001', 1, '{}', 'ingest:t')$q$, 'pctest_ingest', 'FIXED');
SELECT pc_test.throws('COMPOSITION', 'C9: 構成品が ACTIVE でない CP は ACTIVE にできない',
    $q$UPDATE product_core.composition SET status = 'ACTIVE', activated_by = 'human:op_prod', updated_by = 'human:op_prod' WHERE cp_id = 'CP-100003'$q$, 'pctest_reviewer', 'ACTIVE でない');
SELECT pc_test.throws('COMPOSITION', 'C9: 構成品の無い CP は ACTIVE にできない',
    $q$UPDATE product_core.composition SET status = 'ACTIVE', activated_by = 'human:op_prod', updated_by = 'human:op_prod' WHERE cp_id = 'CP-100004'$q$, 'pctest_reviewer', '構成品が無い');
SELECT pc_test.throws('COMPOSITION', 'C9: UNCLASSIFIED の CP は ACTIVE にできない',
    $q$UPDATE product_core.composition SET status = 'ACTIVE', activated_by = 'human:op_prod', updated_by = 'human:op_prod' WHERE cp_id = 'CP-100002'$q$, 'pctest_reviewer', 'FIXED');
SELECT pc_test.lives('COMPOSITION', 'C9: 条件を満たす CP を人が ACTIVE 化できる',
    $q$UPDATE product_core.composition SET status = 'ACTIVE', activated_by = 'human:op_prod', updated_by = 'human:op_prod' WHERE cp_id IN ('CP-100001', 'CP-100005')$q$, 'pctest_reviewer');
SELECT pc_test.throws('COMPOSITION', 'C9: ACTIVE 後は構成品を追加できない',
    $q$INSERT INTO product_core.composition_component (cp_id, pp_id, quantity, evidence, created_by) VALUES ('CP-100001', 'PP-100002', 1, '{}', 'ingest:t')$q$, 'pctest_ingest', 'ACTIVE');
SELECT pc_test.throws('COMPOSITION', 'C9: ACTIVE 後は構成品の数量を変えられない',
    $q$UPDATE product_core.composition_component SET quantity = 4 WHERE cp_id = 'CP-100001'$q$, 'pctest_ingest', 'ACTIVE');
SELECT pc_test.throws('COMPOSITION', 'C9: ACTIVE 後は構成品を外せない',
    $q$DELETE FROM product_core.composition_component WHERE cp_id = 'CP-100001'$q$, 'pctest_ingest', 'ACTIVE');
SELECT pc_test.lives('COMPOSITION', 'PROVISIONAL の間は構成品を外せる',
    $q$DELETE FROM product_core.composition_component WHERE cp_id = 'CP-100003'$q$, 'pctest_ingest');
SELECT pc_test.throws('COMPOSITION', 'AI は SELECTION(可変構成)に分類できない',
    $q$UPDATE product_core.composition SET comp_type = 'SELECTION', updated_by = 'ingest:t' WHERE cp_id = 'CP-100002'$q$, 'pctest_ingest');
SELECT pc_test.lives('COMPOSITION', '人は SELECTION に分類できる',
    $q$UPDATE product_core.composition SET comp_type = 'SELECTION', updated_by = 'human:op_prod' WHERE cp_id = 'CP-100002'$q$, 'pctest_reviewer');
SELECT pc_test.throws('COMPOSITION', 'SELECTION は Phase 1 では ACTIVE にできない',
    $q$UPDATE product_core.composition SET status = 'ACTIVE', activated_by = 'human:op_prod', updated_by = 'human:op_prod' WHERE cp_id = 'CP-100002'$q$, 'pctest_reviewer', 'FIXED');

-- ------------------------------------------------------------
--  LISTING: 販売口(P4)
-- ------------------------------------------------------------
SELECT pc_test.lives('LISTING', '取込ロールが出品を写せる',
    $q$INSERT INTO product_core.listing (listing_id, channel, asin, seller_sku, fulfillment, legacy_listing_id, source_type, source_ref, created_by) VALUES
       ('LS-100001', 'AMAZON', 'B0TESTA001', 'SKU-A', 'SELF', 'C000001', 'test', 'test#1', 'ingest:t'),
       ('LS-100003', 'AMAZON', 'B0TESTA001', 'SKU-A-FBA', 'FBA', 'C000003', 'test', 'test#3', 'ingest:t')$q$, 'pctest_ingest');
SELECT pc_test.lives('LISTING', '楽天の出品(SKU 空)を写せる',
    $q$INSERT INTO product_core.listing (listing_id, channel, rakuten_item_id, source_type, source_ref, created_by) VALUES
       ('LS-100002', 'RAKUTEN', 'r-item-1', 'test', 'test#2', 'ingest:t')$q$, 'pctest_ingest');
SELECT pc_test.check('LISTING', 'ASIN は一意にしない(FBA と自己発送で同じ ASIN)',
    (SELECT count(*) FROM product_core.listing WHERE asin = 'B0TESTA001') = 2);
SELECT pc_test.throws('LISTING', '楽天の同じ管理番号 + 空 SKU の重複を拒否(NULLS NOT DISTINCT)',
    $q$INSERT INTO product_core.listing (listing_id, channel, rakuten_item_id, source_type, source_ref, created_by) VALUES
       ('LS-100009', 'RAKUTEN', 'r-item-1', 'test', 'test#9', 'ingest:t')$q$, 'pctest_ingest', '23505');
SELECT pc_test.throws('LISTING', 'Amazon の SKU 重複を拒否',
    $q$INSERT INTO product_core.listing (listing_id, channel, seller_sku, source_type, source_ref, created_by) VALUES
       ('LS-100010', 'AMAZON', 'SKU-A', 'test', 'test#10', 'ingest:t')$q$, 'pctest_ingest', '23505');
SELECT pc_test.throws('LISTING', 'Amazon の出品に楽天の管理番号を持たせない',
    $q$INSERT INTO product_core.listing (listing_id, channel, rakuten_item_id, source_type, source_ref, created_by) VALUES
       ('LS-100011', 'AMAZON', 'r-x', 'test', 'test#11', 'ingest:t')$q$, 'pctest_ingest', '23514');
SELECT pc_test.lives('LISTING', 'チャネル状態・最終確認日時は更新できる',
    $q$UPDATE product_core.listing SET channel_status = 'ACTIVE', last_seen_at = now(), updated_by = 'ingest:t' WHERE listing_id = 'LS-100001'$q$, 'pctest_ingest');
SELECT pc_test.throws('LISTING', 'P4: SKU(自然キー)は変更できない(取込ロール)',
    $q$UPDATE product_core.listing SET seller_sku = 'SKU-Z' WHERE listing_id = 'LS-100001'$q$, 'pctest_ingest', '42501');
SELECT pc_test.throws('LISTING', 'P4: SKU(自然キー)は変更できない(トリガー)',
    $q$UPDATE product_core.listing SET seller_sku = 'SKU-Z' WHERE listing_id = 'LS-100001'$q$, NULL, '自然キー');
SELECT pc_test.throws('LISTING', 'P4: 出品は削除できない', $q$DELETE FROM product_core.listing WHERE listing_id = 'LS-100001'$q$, NULL, '削除できません');

-- ------------------------------------------------------------
--  ASSESSMENT: AI 判定と人の review(D11・C5・S5)
-- ------------------------------------------------------------
SELECT pc_test.lives('ASSESSMENT', '取込ロール(AI)が判定を記録できる',
    $q$INSERT INTO product_core.assessment (assessment_id, assessment_type, subject_key, verdict, proposed_entity_id, proposed_entity_type,
            confidence, reason, evidence, rule_version, assessed_by) VALUES
       ('AS-100000001', 'LEGACY_MAPPING', 'P000001', 'PHYSICAL_PRODUCT_CANDIDATE', 'PP-100001', 'PHYSICAL_PRODUCT', 'HIGH', 'JAN一意・単品', '{}', 'legacy-map/1.0', 'rule:legacy-map/1.0'),
       ('AS-100000002', 'LEGACY_MAPPING', 'P000002', 'PHYSICAL_PRODUCT_CANDIDATE', 'PP-100002', 'PHYSICAL_PRODUCT', 'HIGH', 'JAN一意・単品', '{}', 'legacy-map/1.0', 'rule:legacy-map/1.0'),
       ('AS-100000003', 'LISTING_REFERENCE', 'B0TESTA001', 'CONSISTENT_HIGH', 'PP-100002', 'PHYSICAL_PRODUCT', 'HIGH', '商品名と入数が一致', '{}', 'asin-check/1.0', 'ai:test/1'),
       ('AS-100000004', 'COST_VARIANCE', 'CO-100000001', 'OK', NULL, NULL, 'HIGH', '差なし', '{}', 'cost-var/1.0', 'rule:cost-var/1.0'),
       ('AS-100000005', 'TAX_BASIS', 'CO-100000003', 'UNKNOWN', NULL, NULL, 'LOW', '税区分の記載なし', '{}', 'tax/1.0', 'rule:tax/1.0'),
       ('AS-100000006', 'TAX_BASIS', 'CO-100000002', 'INCLUDED', NULL, NULL, 'MEDIUM', '運用上は税込', '{}', 'tax/1.0', 'rule:tax/1.0'),
       ('AS-100000007', 'LEGACY_MAPPING', 'P000004', 'PHYSICAL_PRODUCT_CANDIDATE', 'PP-100004', 'PHYSICAL_PRODUCT', 'HIGH', 'x', '{}', 'legacy-map/1.0', 'rule:legacy-map/1.0')$q$,
    'pctest_ingest');
SELECT pc_test.check('ASSESSMENT', 'content_hash が自動計算される(sha256)',
    (SELECT bool_and(content_hash ~ '^[0-9a-f]{64}$') FROM product_core.assessment));
SELECT pc_test.check('ASSESSMENT', '「整合性高」でも review は UNREVIEWED のまま',
    (SELECT review_status = 'UNREVIEWED' FROM product_core.assessment WHERE assessment_id = 'AS-100000003'));
SELECT pc_test.throws('ASSESSMENT', 'S5: 取込ロールは review 欄を書けない(挿入)',
    $q$INSERT INTO product_core.assessment (assessment_type, subject_key, verdict, confidence, reason, evidence, rule_version, assessed_by, review_status, reviewed_by)
       VALUES ('LEGACY_MAPPING', 'P000009', 'NEEDS_REVIEW', 'LOW', 'x', '{}', 'v', 'ai:x', 'APPROVED', 'human:op_rev')$q$, 'pctest_ingest', '42501');
SELECT pc_test.throws('ASSESSMENT', 'S5: 取込ロールは review 欄を書けない(更新)',
    $q$UPDATE product_core.assessment SET review_status = 'APPROVED', reviewed_by = 'human:op_rev' WHERE assessment_id = 'AS-100000003'$q$, 'pctest_ingest', '42501');
SELECT pc_test.throws('ASSESSMENT', '承認済みの判定を最初から作れない(トリガー)',
    $q$INSERT INTO product_core.assessment (assessment_type, subject_key, verdict, confidence, reason, evidence, rule_version, assessed_by, review_status, reviewed_by, reviewed_at)
       VALUES ('LEGACY_MAPPING', 'P000009', 'NEEDS_REVIEW', 'LOW', 'x', '{}', 'v', 'ai:x', 'APPROVED', 'human:op_rev', now())$q$, NULL, 'UNREVIEWED');
SELECT pc_test.throws('ASSESSMENT', '人は assessment を作らない(assessed_by は ai:/rule: のみ)',
    $q$INSERT INTO product_core.assessment (assessment_type, subject_key, verdict, confidence, reason, evidence, rule_version, assessed_by)
       VALUES ('LEGACY_MAPPING', 'P000009', 'NEEDS_REVIEW', 'LOW', 'x', '{}', 'v', 'human:op_rev')$q$, 'pctest_ingest', '23514');
SELECT pc_test.throws('ASSESSMENT', 'C12: 種類に合わない verdict を拒否',
    $q$INSERT INTO product_core.assessment (assessment_type, subject_key, verdict, confidence, reason, evidence, rule_version, assessed_by)
       VALUES ('COST_VARIANCE', 'x', 'WARN', 'LOW', 'x', '{}', 'v', 'ai:x')$q$, 'pctest_ingest', 'assessment_verdict_by_type');
SELECT pc_test.throws('ASSESSMENT', 'S5: review 者に AI を書けない',
    $q$UPDATE product_core.assessment SET review_status = 'APPROVED', reviewed_by = 'ai:gpt' WHERE assessment_id = 'AS-100000001'$q$, 'pctest_reviewer');
SELECT pc_test.throws('ASSESSMENT', 'S5: 存在しない人の名義で review できない',
    $q$UPDATE product_core.assessment SET review_status = 'APPROVED', reviewed_by = 'human:ghost' WHERE assessment_id = 'AS-100000001'$q$, 'pctest_reviewer', 'REVIEW');
SELECT pc_test.throws('ASSESSMENT', '保留には理由が必要',
    $q$UPDATE product_core.assessment SET review_status = 'ON_HOLD', reviewed_by = 'human:op_rev' WHERE assessment_id = 'AS-100000002'$q$, 'pctest_reviewer', '23514');
SELECT pc_test.lives('ASSESSMENT', '保留 → 承認', ARRAY[
    $q$UPDATE product_core.assessment SET review_status = 'ON_HOLD', reviewed_by = 'human:op_rev', review_note = '資料待ち', review_channel = 'EXCEL', review_ref = 'sha#2' WHERE assessment_id = 'AS-100000002'$q$,
    $q$UPDATE product_core.assessment SET review_status = 'APPROVED', reviewed_by = 'human:op_rev' WHERE assessment_id = 'AS-100000002'$q$], 'pctest_reviewer');
SELECT pc_test.lives('ASSESSMENT', '人が承認できる',
    $q$UPDATE product_core.assessment SET review_status = 'APPROVED', reviewed_by = 'human:op_rev', review_channel = 'EXCEL', review_ref = 'sha#1'
        WHERE assessment_id IN ('AS-100000001', 'AS-100000004', 'AS-100000007')$q$, 'pctest_reviewer');
SELECT pc_test.throws('ASSESSMENT', '確定した review はやり直せない',
    $q$UPDATE product_core.assessment SET review_status = 'REJECTED', reviewed_by = 'human:op_rev', review_note = 'x' WHERE assessment_id = 'AS-100000001'$q$, 'pctest_reviewer', '確定済み');
SELECT pc_test.throws('ASSESSMENT', 'C5: AI 部分(verdict)は変更できない',
    $q$UPDATE product_core.assessment SET verdict = 'NEEDS_REVIEW' WHERE assessment_id = 'AS-100000003'$q$, NULL, 'AI 部分');
SELECT pc_test.throws('ASSESSMENT', 'C5: assessment は削除できない', $q$DELETE FROM product_core.assessment WHERE assessment_id = 'AS-100000003'$q$, NULL, '削除できません');

-- ------------------------------------------------------------
--  LEGACY_MAPPING: 既存 P の対応付け(P1・C3・C7・C8・C11・S4)
-- ------------------------------------------------------------
SELECT pc_test.lives('LEGACY_MAPPING', 'AI が PROPOSED の対応付けを作れる',
    $q$INSERT INTO product_core.legacy_mapping (mapping_id, legacy_p, entity_id, entity_type, mapping_role, legacy_snapshot_ref, assessment_id, created_by) VALUES
       ('LM-10000001', 'P000001', 'PP-100001', 'PHYSICAL_PRODUCT', 'SOLE', 'sha#row2', 'AS-100000001', 'rule:legacy-map/1.0'),
       ('LM-10000002', 'P000002', 'PP-100002', 'PHYSICAL_PRODUCT', 'SOLE', 'sha#row3', 'AS-100000003', 'rule:legacy-map/1.0')$q$, 'pctest_ingest');
SELECT pc_test.throws('LEGACY_MAPPING', 'S5: AI は ACTIVE の対応付けを作れない(列権限)',
    $q$INSERT INTO product_core.legacy_mapping (legacy_p, entity_id, entity_type, mapping_role, legacy_snapshot_ref, created_by, record_status, approved_by)
       VALUES ('P000005', 'PP-100001', 'PHYSICAL_PRODUCT', 'SOLE', 'x', 'ai:x', 'ACTIVE', 'human:op_admin')$q$, 'pctest_ingest', '42501');
SELECT pc_test.throws('LEGACY_MAPPING', 'C7: 種類の食い違う参照(PP を COMPOSITION として)を拒否',
    $q$INSERT INTO product_core.legacy_mapping (legacy_p, entity_id, entity_type, mapping_role, legacy_snapshot_ref, created_by)
       VALUES ('P000005', 'PP-100001', 'COMPOSITION', 'SOLE', 'x', 'ai:x')$q$, 'pctest_ingest', '23503');
SELECT pc_test.throws('LEGACY_MAPPING', 'C13: legacy P の形式を検査',
    $q$INSERT INTO product_core.legacy_mapping (legacy_p, entity_id, entity_type, mapping_role, legacy_snapshot_ref, created_by)
       VALUES ('PP-100001', 'PP-100001', 'PHYSICAL_PRODUCT', 'SOLE', 'x', 'ai:x')$q$, 'pctest_ingest', '23514');
SELECT pc_test.throws('LEGACY_MAPPING', 'C11: 人の review が無い判定を根拠に ACTIVE 化できない',
    $q$UPDATE product_core.legacy_mapping SET record_status = 'ACTIVE', approved_by = 'human:op_admin', updated_by = 'human:op_admin' WHERE mapping_id = 'LM-10000002'$q$,
    'pctest_reviewer', '承認されていません');
SELECT pc_test.throws('LEGACY_MAPPING', 'C1: 権限の無い人は ACTIVE 化できない',
    $q$UPDATE product_core.legacy_mapping SET record_status = 'ACTIVE', approved_by = 'human:op_rev', updated_by = 'human:op_rev' WHERE mapping_id = 'LM-10000001'$q$,
    'pctest_reviewer', 'LEGACY_MAPPING_APPROVE');
SELECT pc_test.check('LEGACY_MAPPING', 'S4: 未承認の対応付けは正式処理用ビューに出ない',
    NOT EXISTS (SELECT 1 FROM product_core.v_active_legacy_mapping WHERE legacy_p IN ('P000001', 'P000002')));
SELECT pc_test.lives('LEGACY_MAPPING', '人の review 済み判定を根拠に、権限者が ACTIVE 化できる',
    $q$UPDATE product_core.legacy_mapping SET record_status = 'ACTIVE', approved_by = 'human:op_admin', updated_by = 'human:op_admin' WHERE mapping_id = 'LM-10000001'$q$, 'pctest_reviewer');
SELECT pc_test.check('LEGACY_MAPPING', 'ACTIVE 化した対応付けだけがビューに出る',
    (SELECT count(*) FROM product_core.v_active_legacy_mapping WHERE legacy_p = 'P000001') = 1);
SELECT pc_test.throws('LEGACY_MAPPING', 'C8: 同じ legacy P に ACTIVE は1行だけ',
    $q$INSERT INTO product_core.legacy_mapping (legacy_p, entity_id, entity_type, mapping_role, legacy_snapshot_ref, created_by, record_status, approved_by, assessment_id)
       VALUES ('P000001', 'PP-100002', 'PHYSICAL_PRODUCT', 'SOLE', 'x', 'human:op_admin', 'ACTIVE', 'human:op_admin', 'AS-100000001')$q$, 'pctest_reviewer', '23505');
SELECT pc_test.throws('LEGACY_MAPPING', 'C3: 内容列(対象)は変更できない(トリガー)',
    $q$UPDATE product_core.legacy_mapping SET entity_id = 'PP-100002' WHERE mapping_id = 'LM-10000001'$q$, NULL, '内容列');
SELECT pc_test.throws('LEGACY_MAPPING', 'C3: 内容列(対象)は変更できない(列権限)',
    $q$UPDATE product_core.legacy_mapping SET entity_id = 'PP-100002' WHERE mapping_id = 'LM-10000001'$q$, 'pctest_reviewer', '42501');
SELECT pc_test.throws('LEGACY_MAPPING', 'P1: 対応付けは削除できない', $q$DELETE FROM product_core.legacy_mapping$q$, NULL, '削除できません');
SELECT pc_test.throws('LEGACY_MAPPING', 'RETIRED の PP へは対応付けできない',
    $q$INSERT INTO product_core.legacy_mapping (legacy_p, entity_id, entity_type, mapping_role, legacy_snapshot_ref, created_by)
       VALUES ('P000004', 'PP-100004', 'PHYSICAL_PRODUCT', 'SOLE', 'x', 'ai:x')$q$, 'pctest_ingest', 'RETIRED');
SELECT pc_test.lives('LEGACY_MAPPING', '人が却下できる(記録は残る)',
    $q$UPDATE product_core.legacy_mapping SET record_status = 'REJECTED', approved_by = 'human:op_admin', updated_by = 'human:op_admin' WHERE mapping_id = 'LM-10000002'$q$, 'pctest_reviewer');
SELECT pc_test.throws('LEGACY_MAPPING', 'REJECTED の行は変更できない',
    $q$UPDATE product_core.legacy_mapping SET record_status = 'ACTIVE', approved_by = 'human:op_admin', updated_by = 'human:op_admin' WHERE mapping_id = 'LM-10000002'$q$, 'pctest_reviewer');

-- ------------------------------------------------------------
--  LISTING_REF: 出品の参照先(S2・S4・C8・C11)
-- ------------------------------------------------------------
SELECT pc_test.lives('LISTING_REF', 'P000002 の対応付けを PROPOSED のまま用意', ARRAY[
    $q$INSERT INTO product_core.legacy_mapping (mapping_id, legacy_p, entity_id, entity_type, mapping_role, legacy_snapshot_ref, created_by)
       VALUES ('LM-10000003', 'P000002', 'PP-100002', 'PHYSICAL_PRODUCT', 'SOLE', 'sha#row3', 'rule:legacy-map/1.0')$q$], 'pctest_ingest');
SELECT pc_test.throws('LISTING_REF', 'S4: 未承認(PROPOSED)の対応付けからは参照を作れない',
    $q$INSERT INTO product_core.listing_reference_history (listing_id, entity_id, entity_type, valid_from, derived_from_legacy_p, reason, created_by)
       VALUES ('LS-100003', 'PP-100002', 'PHYSICAL_PRODUCT', '2026-01-01', 'P000002', '初期移行', 'ai:x')$q$, 'pctest_ingest', 'legacy_mapping');
SELECT pc_test.lives('LISTING_REF', 'S4: ACTIVE な対応付けからは PROPOSED の参照を作れる',
    $q$INSERT INTO product_core.listing_reference_history (ref_id, listing_id, entity_id, entity_type, valid_from, derived_from_legacy_p, reason, created_by)
       VALUES ('LR-10000001', 'LS-100001', 'PP-100001', 'PHYSICAL_PRODUCT', '2026-01-01', 'P000001', '初期移行', 'rule:ref/1.0')$q$, 'pctest_ingest');
SELECT pc_test.lives('LISTING_REF', '権限者が参照を ACTIVE 化できる',
    $q$UPDATE product_core.listing_reference_history SET record_status = 'ACTIVE', approved_by = 'human:op_admin', updated_by = 'human:op_admin' WHERE ref_id = 'LR-10000001'$q$, 'pctest_reviewer');
SELECT pc_test.throws('LISTING_REF', 'C8: 同じ出品で ACTIVE の期間が重なる参照を拒否(btree_gist の EXCLUDE)',
    $q$INSERT INTO product_core.listing_reference_history (listing_id, entity_id, entity_type, valid_from, reason, created_by, record_status, approved_by)
       VALUES ('LS-100001', 'CP-100001', 'COMPOSITION', '2026-06-01', '誤紐付け修正', 'human:op_admin', 'ACTIVE', 'human:op_admin')$q$, 'pctest_reviewer', '23P01');
SELECT pc_test.lives('LISTING_REF', '期間を閉じてから、重ならない新しい参照を ACTIVE で作れる', ARRAY[
    $q$UPDATE product_core.listing_reference_history SET valid_to = '2026-05-31', updated_by = 'human:op_admin' WHERE ref_id = 'LR-10000001'$q$,
    $q$INSERT INTO product_core.listing_reference_history (listing_id, entity_id, entity_type, valid_from, reason, created_by, record_status, approved_by)
       VALUES ('LS-100001', 'CP-100001', 'COMPOSITION', '2026-06-01', '誤紐付け修正', 'human:op_admin', 'ACTIVE', 'human:op_admin')$q$], 'pctest_reviewer');
SELECT pc_test.throws('LISTING_REF', 'valid_to は一度だけ設定できる',
    $q$UPDATE product_core.listing_reference_history SET valid_to = '2026-06-30', updated_by = 'human:op_admin' WHERE ref_id = 'LR-10000001'$q$, 'pctest_reviewer', '一度だけ');
SELECT pc_test.throws('LISTING_REF', 'S2: RETIRED の PP へ新しい参照(提案)を作れない',
    $q$INSERT INTO product_core.listing_reference_history (listing_id, entity_id, entity_type, valid_from, reason, created_by)
       VALUES ('LS-100002', 'PP-100004', 'PHYSICAL_PRODUCT', '2026-01-01', 'x', 'ai:x')$q$, 'pctest_ingest', 'RETIRED');
SELECT pc_test.throws('LISTING_REF', 'PROVISIONAL の PP へ ACTIVE な参照は作れない',
    $q$INSERT INTO product_core.listing_reference_history (listing_id, entity_id, entity_type, valid_from, reason, created_by, record_status, approved_by)
       VALUES ('LS-100002', 'PP-100003', 'PHYSICAL_PRODUCT', '2026-01-01', 'x', 'human:op_admin', 'ACTIVE', 'human:op_admin')$q$, 'pctest_reviewer', 'PROVISIONAL');
SELECT pc_test.lives('LISTING_REF', '(準備)ACTIVE の PP-100009 への提案を作ってから PP を退役', ARRAY[
    $q$SET LOCAL ROLE pctest_ingest$q$,
    $q$INSERT INTO product_core.listing_reference_history (ref_id, listing_id, entity_id, entity_type, valid_from, reason, created_by)
       VALUES ('LR-10000009', 'LS-100002', 'PP-100009', 'PHYSICAL_PRODUCT', '2026-01-01', 'x', 'ai:x')$q$,
    $q$SET LOCAL ROLE pctest_reviewer$q$,
    $q$UPDATE product_core.physical_product SET status = 'RETIRED', retired_by = 'human:op_prod', retire_reason = 'テスト', updated_by = 'human:op_prod' WHERE pp_id = 'PP-100009'$q$]);
SELECT pc_test.throws('LISTING_REF', 'S2: 提案後に RETIRED になった対象への参照は ACTIVE 化できない',
    $q$UPDATE product_core.listing_reference_history SET record_status = 'ACTIVE', approved_by = 'human:op_admin', updated_by = 'human:op_admin' WHERE ref_id = 'LR-10000009'$q$, 'pctest_reviewer', 'RETIRED');
SELECT pc_test.throws('LISTING_REF', 'C11: 未 review の「整合性高」判定を根拠に参照を ACTIVE で作れない',
    $q$INSERT INTO product_core.listing_reference_history (listing_id, entity_id, entity_type, valid_from, reason, created_by, record_status, approved_by, assessment_id)
       VALUES ('LS-100003', 'PP-100002', 'PHYSICAL_PRODUCT', '2026-01-01', 'x', 'human:op_admin', 'ACTIVE', 'human:op_admin', 'AS-100000003')$q$, 'pctest_reviewer', '承認されていません');
SELECT pc_test.throws('LISTING_REF', 'S5: AI は参照を ACTIVE 化できない(列権限)',
    $q$UPDATE product_core.listing_reference_history SET record_status = 'ACTIVE' WHERE ref_id = 'LR-10000009'$q$, 'pctest_ingest', '42501');
