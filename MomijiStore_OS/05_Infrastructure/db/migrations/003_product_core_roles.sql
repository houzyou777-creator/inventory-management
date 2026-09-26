-- =============================================================
--  003_product_core_roles.sql — Product Core の DB ロールと権限(S5 の土台)
-- =============================================================
--  ロールはログインできないグループロールとして作る。
--  ログインロールとパスワードは人が作成し、NAS の .env にだけ置く(SECURITY.md)。
--  このファイルに認証情報は書かない。
--
--  | ロール       | 使う主体                          | できること                                    |
--  |--------------|-----------------------------------|-----------------------------------------------|
--  | pc_owner     | (予約)将来のスキーマ所有者         | Phase 1 では権限を与えない                     |
--  | pc_ingest    | 取込ジョブ・AI 判定                | 事実の写し・観測値・PROVISIONAL・PROPOSED・AI 判定の追加。**承認列には一切書けない** |
--  | pc_reviewer  | Excel 確認表の取込(人が起動)      | review 列・承認・ACTIVE 化・正式原価。トリガーが operator の権限を確認する |
--  | pc_reader    | 検証レポート・将来の参照           | SELECT のみ                                    |
--
--  ⚠️ AI のプロセスには pc_reviewer もスーパーユーザーも渡さない。
--     スーパーユーザーは列の権限を素通りする(CHECK とトリガーは効く)。
--  ⚠️ どのロールにも正本 Excel への書き込み経路は無い(D1)。
--  ⚠️ pc_ingest は Product Core の ID 列(pp_id・cp_id・listing_id など)を明示指定できない(1.3)。
--     ID は DB のシーケンスが採番し、取込は INSERT ... RETURNING か取込関数(import_*)で受け取る。
-- =============================================================

DO $$
DECLARE r text;
BEGIN
    FOREACH r IN ARRAY ARRAY['pc_owner', 'pc_ingest', 'pc_reviewer', 'pc_reader'] LOOP
        IF NOT EXISTS (SELECT 1 FROM pg_roles WHERE rolname = r) THEN
            EXECUTE format('CREATE ROLE %I NOLOGIN NOSUPERUSER NOCREATEDB NOCREATEROLE NOINHERIT', r);
        END IF;
    END LOOP;
END;
$$;

COMMENT ON ROLE pc_ingest IS 'Product Core 取込・AI 判定。承認列には書けない。';
COMMENT ON ROLE pc_reviewer IS 'Product Core 人の承認(Excel 確認表の取込)。operator の権限はトリガーで確認。';
COMMENT ON ROLE pc_reader IS 'Product Core 参照のみ。';
COMMENT ON ROLE pc_owner IS 'Product Core 予約(Phase 1 では未使用)。';

-- スキーマの利用と参照(トリガーは呼び出し元の権限で他の表を読むため、書き込みロールにも SELECT を与える)
GRANT USAGE ON SCHEMA product_core, legacy_ingest TO pc_ingest, pc_reviewer, pc_reader;
GRANT SELECT ON ALL TABLES IN SCHEMA product_core TO pc_ingest, pc_reviewer, pc_reader;
GRANT SELECT ON ALL TABLES IN SCHEMA legacy_ingest TO pc_ingest, pc_reviewer, pc_reader;
GRANT USAGE ON ALL SEQUENCES IN SCHEMA product_core TO pc_ingest, pc_reviewer;
GRANT USAGE ON ALL SEQUENCES IN SCHEMA legacy_ingest TO pc_ingest;


-- =============================================================
--  pc_ingest — 取込・AI 判定(承認列を持たない列だけに書ける)
-- =============================================================
-- 取込ステージング
GRANT INSERT ON legacy_ingest.ingest_run TO pc_ingest;
GRANT UPDATE (status, finished_at, files, note) ON legacy_ingest.ingest_run TO pc_ingest;
GRANT INSERT ON legacy_ingest.legacy_product_snapshot, legacy_ingest.legacy_listing_snapshot,
                legacy_ingest.external_file_snapshot TO pc_ingest;
-- 取込の冪等性(file_id・record_id は IDENTITY が採番)
GRANT INSERT (ingest_run_id, source_system, file_path, file_sha256, file_mtime, row_count)
    ON legacy_ingest.import_source_file TO pc_ingest;
GRANT INSERT (source_system, source_record_key, source_record_hash, file_id, ingest_run_id, raw)
    ON legacy_ingest.import_source_record TO pc_ingest;
GRANT INSERT ON legacy_ingest.import_record_map TO pc_ingest;

-- 事実の写し・コード・観測値
-- identifier の origin・派生元は人の修正(HUMAN_CORRECTED)専用なので取込には渡さない(常に OBSERVED)
GRANT INSERT (id_type, value, checkdigit_valid, first_seen_source, first_seen_ref, created_by, ingest_run_id)
    ON product_core.identifier TO pc_ingest;
GRANT INSERT (source, source_ref, source_document_id, source_reliability, observed_at, observed_amount, currency,
              tax_inclusion, tax_inclusion_basis, tax_rate, tax_rate_basis, amount_unit, subject_entity_id,
              subject_entity_type, subject_legacy_p, source_unit_basis, source_record_key, source_record_hash,
              import_record_id, evidence, created_by, ingest_run_id)
    ON product_core.cost_observation TO pc_ingest;
GRANT INSERT (channel, group_type, group_key, source_type, source_ref, created_by, ingest_run_id)
    ON product_core.listing_group TO pc_ingest;
GRANT UPDATE (source_type, source_ref, updated_by, ingest_run_id) ON product_core.listing_group TO pc_ingest;
GRANT INSERT (channel, asin, seller_sku, rakuten_item_id, rakuten_sku, fulfillment, listing_group_id,
              channel_status, first_seen_at, last_seen_at, source_type, source_ref, created_by,
              ingest_run_id)
    ON product_core.listing TO pc_ingest;
GRANT UPDATE (fulfillment, listing_group_id, channel_status, last_seen_at, source_type, source_ref, updated_by,
              ingest_run_id)
    ON product_core.listing TO pc_ingest;

-- PROVISIONAL の実体(status・承認列・tax_category には書けない)
GRANT INSERT (name, kind, origin_key, note, created_by, ingest_run_id) ON product_core.physical_product TO pc_ingest;
GRANT UPDATE (name, kind, note, updated_by) ON product_core.physical_product TO pc_ingest;
GRANT INSERT (name, comp_type, origin_key, note, created_by, ingest_run_id) ON product_core.composition TO pc_ingest;
GRANT UPDATE (name, comp_type, note, updated_by) ON product_core.composition TO pc_ingest;
-- 構成品の下書き(親が PROVISIONAL の間だけ。トリガーで確認)
GRANT INSERT (cp_id, pp_id, quantity, role, sort_no, evidence, created_by) ON product_core.composition_component TO pc_ingest;
GRANT DELETE ON product_core.composition_component TO pc_ingest;
GRANT UPDATE (pp_id, quantity, role, sort_no, evidence) ON product_core.composition_component TO pc_ingest;

-- AI 判定(review 欄には書けない)
GRANT INSERT (assessment_type, subject_entity_id, subject_entity_type, subject_key, verdict,
              proposed_entity_id, proposed_entity_type, confidence, reason, evidence, rule_version, assessed_by,
              supersedes_assessment_id)
    ON product_core.assessment TO pc_ingest;

-- 提案(PROPOSED)。record_status・approved_by・approved_at・superseded_by には書けない
GRANT INSERT (identifier_id, entity_id, entity_type, link_basis, scan_policy, valid_from, valid_to, confidence,
              evidence, assessment_id, created_by, ingest_run_id)
    ON product_core.identifier_link TO pc_ingest;
GRANT INSERT (legacy_p, entity_id, entity_type, mapping_role, valid_from, valid_to, legacy_snapshot_ref,
              assessment_id, created_by, ingest_run_id)
    ON product_core.legacy_mapping TO pc_ingest;
GRANT INSERT (rel_type, from_pp_id, to_pp_id, variant_group_key, variant_axis, evidence, assessment_id,
              created_by)
    ON product_core.product_relationship TO pc_ingest;
GRANT INSERT (listing_id, entity_id, entity_type, valid_from, valid_to, derived_from_legacy_p, reason,
              assessment_id, created_by)
    ON product_core.listing_reference_history TO pc_ingest;
GRANT INSERT (legacy_listing_id, listing_id, match_basis, observed_legacy_p, legacy_snapshot_ref, assessment_id,
              created_by, ingest_run_id)
    ON product_core.legacy_listing_mapping TO pc_ingest;
-- review_batch(人が出力・取込する束)には権限を与えない
-- 正式原価(cost_history / composition_cost_history)・operator には一切権限を与えない


-- =============================================================
--  pc_reviewer — 人の承認(トリガーが operator の権限を確認する)
-- =============================================================
GRANT INSERT ON product_core.operator, product_core.operator_permission TO pc_reviewer;
GRANT UPDATE (display_name, status, updated_by) ON product_core.operator TO pc_reviewer;
GRANT UPDATE (revoked_by, revoked_at) ON product_core.operator_permission TO pc_reviewer;
GRANT USAGE ON SEQUENCE product_core.operator_permission_grant_id_seq TO pc_reviewer;

GRANT UPDATE (review_status, reviewed_by, review_note, corrected_entity_id, corrected_entity_type, corrected_value, review_channel,
              review_ref, review_reason_code, review_info_needed, review_batch_id)
    ON product_core.assessment TO pc_reviewer;
GRANT INSERT (review_type, channel, file_name, file_sha256, note, created_by, updated_by)
    ON product_core.review_batch TO pc_reviewer;
GRANT UPDATE (status, reviewed_file_sha256, note, updated_by) ON product_core.review_batch TO pc_reviewer;

-- 12桁 JAN の先頭0補完など、人が review で修正したコード(HUMAN_CORRECTED)を登録する
GRANT INSERT (id_type, value, checkdigit_valid, first_seen_source, first_seen_ref, created_by, origin,
              derived_from_identifier_id, correction_assessment_id)
    ON product_core.identifier TO pc_reviewer;

GRANT UPDATE (status, activated_by, retired_by, retire_reason, tax_category, updated_by)
    ON product_core.physical_product TO pc_reviewer;
GRANT UPDATE (status, activated_by, retired_by, retire_reason, comp_type, updated_by)
    ON product_core.composition TO pc_reviewer;

-- CORRECTED の流れでは、人が修正先を指定した新しい行を ACTIVE で作る
GRANT INSERT ON product_core.identifier_link, product_core.legacy_mapping, product_core.product_relationship,
                product_core.listing_reference_history, product_core.legacy_listing_mapping TO pc_reviewer;
GRANT UPDATE (record_status, approved_by, superseded_by, valid_to, scan_policy, updated_by)
    ON product_core.identifier_link TO pc_reviewer;
GRANT UPDATE (record_status, approved_by, superseded_by, valid_to, updated_by)
    ON product_core.legacy_mapping TO pc_reviewer;
GRANT UPDATE (record_status, approved_by, superseded_by, updated_by)
    ON product_core.product_relationship TO pc_reviewer;
GRANT UPDATE (record_status, approved_by, superseded_by, valid_to, updated_by)
    ON product_core.listing_reference_history TO pc_reviewer;
GRANT UPDATE (record_status, approved_by, superseded_by, updated_by)
    ON product_core.legacy_listing_mapping TO pc_reviewer;

GRANT INSERT ON product_core.cost_history, product_core.composition_cost_history TO pc_reviewer;
GRANT UPDATE (valid_to, record_status, updated_by)
    ON product_core.cost_history, product_core.composition_cost_history TO pc_reviewer;
