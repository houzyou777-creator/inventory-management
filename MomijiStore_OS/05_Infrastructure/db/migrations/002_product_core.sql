-- =============================================================
--  002_product_core.sql — AIKOS Product Core(Shadow Model)Phase 1
-- =============================================================
--  仕様: 04_Manual/DeveloperGuide/ProductCore/Product_Core_Phase1_Specification_v1.0.md
--  設計基準: 04_Manual/DeveloperGuide/ProductCore/Product_Data_Model_v2.2.md
--
--  【この層の位置付け】
--    既存 Excel(商品マスター・出品テーブル)が引き続き正本。
--    ここは Shadow Model であり、Excel への書き戻し経路は持たない。
--    業務データの Phase4 移行先(public)・知識層(intelligence)とはスキーマで分ける。
--
--  【守りを構造で持つ理由】
--    「AI は提案まで、承認は人」を運用の作法に任せると、どこかで崩れる。
--    承認列は3重に守る:
--      ① DB ロールの列単位の権限(003_product_core_roles.sql)
--      ② CHECK 制約: 承認者は human:<operator_id> 形式のみ
--      ③ トリガー: operator が ACTIVE で、その操作の権限を承認時点で持つこと
--    追記専用・状態遷移・期間の重なり・原価の元データ保全もトリガーと制約で強制する。
--
--  ⚠️ このファイルには行を消す SQL を書かない(Migrate.sh の破壊的SQL検査の対象)。
--     削除の拒否はトリガー内の TG_OP 判定で行う。
-- =============================================================

-- 期間の重なり禁止(EXCLUDE)で text の等値比較を GiST に載せるために必要
CREATE EXTENSION IF NOT EXISTS btree_gist;

CREATE SCHEMA IF NOT EXISTS product_core;
CREATE SCHEMA IF NOT EXISTS legacy_ingest;

COMMENT ON SCHEMA product_core IS
    'AIKOS Product Core(Shadow Model)。既存Excelが正本。Excelへの書き戻しはしない。';
COMMENT ON SCHEMA legacy_ingest IS
    '正本Excel・CSVの読み取り専用の写し。Product Core はここからだけ作る。';


-- =============================================================
--  共通の関数
-- =============================================================

-- ID 採番。桁が足りなくなったら桁を増やす(lpad は長い値を切り詰めるので使い分ける)
CREATE OR REPLACE FUNCTION product_core.fmt_id(p_prefix text, p_n bigint, p_width integer)
RETURNS text LANGUAGE sql IMMUTABLE AS $$
    SELECT p_prefix || CASE WHEN length(p_n::text) >= p_width THEN p_n::text
                            ELSE lpad(p_n::text, p_width, '0') END
$$;

-- 承認・review に使える行為者は人だけ。表示名ではなく永続の operator_id を使う
CREATE OR REPLACE FUNCTION product_core.is_human(p_actor text)
RETURNS boolean LANGUAGE sql IMMUTABLE AS $$
    SELECT p_actor ~ '^human:[a-z0-9_]{3,32}$'
$$;

CREATE OR REPLACE FUNCTION product_core.is_actor(p_actor text)
RETURNS boolean LANGUAGE sql IMMUTABLE AS $$
    SELECT p_actor ~ '^(human:[a-z0-9_]{3,32}|(ai|rule|ingest):[^[:space:]]+)$'
$$;

CREATE SEQUENCE IF NOT EXISTS product_core.pp_seq;
CREATE SEQUENCE IF NOT EXISTS product_core.cp_seq;
CREATE SEQUENCE IF NOT EXISTS product_core.ls_seq;
CREATE SEQUENCE IF NOT EXISTS product_core.lg_seq;
CREATE SEQUENCE IF NOT EXISTS product_core.id_seq;
CREATE SEQUENCE IF NOT EXISTS product_core.il_seq;
CREATE SEQUENCE IF NOT EXISTS product_core.lm_seq;
CREATE SEQUENCE IF NOT EXISTS product_core.pr_seq;
CREATE SEQUENCE IF NOT EXISTS product_core.cc_seq;
CREATE SEQUENCE IF NOT EXISTS product_core.lr_seq;
CREATE SEQUENCE IF NOT EXISTS product_core.ch_seq;
CREATE SEQUENCE IF NOT EXISTS product_core.cx_seq;
CREATE SEQUENCE IF NOT EXISTS product_core.co_seq;
CREATE SEQUENCE IF NOT EXISTS product_core.as_seq;

-- 追記専用の強制(intelligence.deny_mutation と同じ考え方)
CREATE OR REPLACE FUNCTION product_core.deny_mutation()
RETURNS trigger LANGUAGE plpgsql AS $$
BEGIN
    RAISE EXCEPTION '% は追記専用です。% は許可されていません。訂正は新しい行で行ってください。',
        TG_TABLE_NAME, TG_OP USING ERRCODE = 'restrict_violation';
END;
$$;


-- =============================================================
--  operator / operator_permission — 人と承認権限(D4・D15)
-- =============================================================
CREATE TABLE IF NOT EXISTS product_core.operator (
    operator_id   text        PRIMARY KEY CHECK (operator_id ~ '^[a-z0-9_]{3,32}$'),
    display_name  text        NOT NULL CHECK (btrim(display_name) <> ''),
    status        text        NOT NULL DEFAULT 'ACTIVE' CHECK (status IN ('ACTIVE', 'INACTIVE')),
    created_at    timestamptz NOT NULL DEFAULT now(),
    created_by    text        NOT NULL CHECK (product_core.is_human(created_by)),
    updated_at    timestamptz NOT NULL DEFAULT now(),
    updated_by    text        NOT NULL CHECK (product_core.is_human(updated_by))
);
COMMENT ON TABLE product_core.operator IS
    '承認者。承認記録には human:<operator_id> を使う。表示名は変えてもよい(主キーではない)。';

CREATE TABLE IF NOT EXISTS product_core.operator_permission (
    grant_id      bigserial   PRIMARY KEY,
    operator_id   text        NOT NULL REFERENCES product_core.operator (operator_id),
    permission    text        NOT NULL CHECK (permission IN (
                      'OPERATOR_ADMIN', 'PRODUCT_APPROVE', 'LEGACY_MAPPING_APPROVE', 'RELATIONSHIP_APPROVE',
                      'IDENTIFIER_APPROVE', 'LISTING_REFERENCE_APPROVE', 'COST_APPROVE', 'REVIEW')),
    granted_by    text        NOT NULL CHECK (product_core.is_human(granted_by)),
    granted_at    timestamptz NOT NULL DEFAULT now(),
    revoked_by    text        CHECK (revoked_by IS NULL OR product_core.is_human(revoked_by)),
    revoked_at    timestamptz,
    CHECK ((revoked_by IS NULL) = (revoked_at IS NULL))
);
CREATE UNIQUE INDEX IF NOT EXISTS operator_permission_active_uq
    ON product_core.operator_permission (operator_id, permission) WHERE revoked_at IS NULL;

-- 行為者がその権限を持つか(承認時点で有効な付与があるか)
CREATE OR REPLACE FUNCTION product_core.has_permission(p_actor text, p_permission text,
                                                       p_at timestamptz DEFAULT now())
RETURNS boolean LANGUAGE sql STABLE AS $$
    SELECT product_core.is_human(p_actor) AND EXISTS (
        SELECT 1
          FROM product_core.operator o
          JOIN product_core.operator_permission g ON g.operator_id = o.operator_id
         WHERE o.operator_id = substr(p_actor, 7)
           AND o.status = 'ACTIVE'
           AND g.permission = p_permission
           AND g.granted_at <= p_at
           AND (g.revoked_at IS NULL OR g.revoked_at > p_at))
$$;

-- 承認の強制。DB ロール(pc_reviewer)と operator の権限の両方を確認する。
-- ⚠️ SECURITY INVOKER のまま置く。current_user で「誰の接続か」を見るため。
CREATE OR REPLACE FUNCTION product_core.assert_approval(p_actor text, p_permission text)
RETURNS void LANGUAGE plpgsql STABLE AS $$
BEGIN
    IF NOT pg_has_role(current_user, 'pc_reviewer', 'MEMBER') THEN
        RAISE EXCEPTION '承認操作は pc_reviewer ロールの接続でのみ行えます(現在: %)', current_user
            USING ERRCODE = 'insufficient_privilege';
    END IF;
    IF p_actor IS NULL OR NOT product_core.is_human(p_actor) THEN
        RAISE EXCEPTION '承認者は human:<operator_id> である必要があります(値: %)', coalesce(p_actor, 'NULL')
            USING ERRCODE = 'insufficient_privilege';
    END IF;
    IF NOT product_core.has_permission(p_actor, p_permission) THEN
        RAISE EXCEPTION '% は有効な % 権限を持っていません', p_actor, p_permission
            USING ERRCODE = 'insufficient_privilege';
    END IF;
END;
$$;

CREATE OR REPLACE FUNCTION product_core.operator_guard()
RETURNS trigger LANGUAGE plpgsql AS $$
BEGIN
    IF TG_OP = 'DELETE' THEN
        RAISE EXCEPTION 'operator は削除できません(INACTIVE にしてください)' USING ERRCODE = 'restrict_violation';
    END IF;
    IF NOT pg_has_role(current_user, 'pc_reviewer', 'MEMBER') THEN
        RAISE EXCEPTION 'operator の登録・変更は pc_reviewer ロールの接続でのみ行えます' USING ERRCODE = 'insufficient_privilege';
    END IF;
    IF TG_OP = 'INSERT' THEN
        NEW.updated_by := NEW.created_by;
        NEW.updated_at := now();
        -- 最初の1名(管理者)だけは自分自身を登録できる(ブートストラップ)
        IF EXISTS (SELECT 1 FROM product_core.operator) THEN
            PERFORM product_core.assert_approval(NEW.created_by, 'OPERATOR_ADMIN');
        ELSIF NEW.created_by <> 'human:' || NEW.operator_id THEN
            RAISE EXCEPTION '最初の operator は本人の名義で登録してください' USING ERRCODE = 'insufficient_privilege';
        END IF;
        RETURN NEW;
    END IF;
    IF NEW.operator_id <> OLD.operator_id OR NEW.created_at <> OLD.created_at OR NEW.created_by <> OLD.created_by THEN
        RAISE EXCEPTION 'operator_id・作成情報は変更できません' USING ERRCODE = 'restrict_violation';
    END IF;
    PERFORM product_core.assert_approval(NEW.updated_by, 'OPERATOR_ADMIN');
    NEW.updated_at := now();
    RETURN NEW;
END;
$$;
DROP TRIGGER IF EXISTS operator_guard ON product_core.operator;
CREATE TRIGGER operator_guard BEFORE INSERT OR UPDATE OR DELETE ON product_core.operator
    FOR EACH ROW EXECUTE FUNCTION product_core.operator_guard();

CREATE OR REPLACE FUNCTION product_core.operator_permission_guard()
RETURNS trigger LANGUAGE plpgsql AS $$
BEGIN
    IF TG_OP = 'DELETE' THEN
        RAISE EXCEPTION '権限の付与記録は削除できません(revoked_at で取り消してください)' USING ERRCODE = 'restrict_violation';
    END IF;
    IF NOT pg_has_role(current_user, 'pc_reviewer', 'MEMBER') THEN
        RAISE EXCEPTION '権限の付与・取消は pc_reviewer ロールの接続でのみ行えます' USING ERRCODE = 'insufficient_privilege';
    END IF;
    IF TG_OP = 'INSERT' THEN
        NEW.granted_at := now();
        IF NEW.revoked_by IS NOT NULL THEN
            RAISE EXCEPTION '取り消し済みの付与は作れません' USING ERRCODE = 'check_violation';
        END IF;
        -- 有効な OPERATOR_ADMIN がまだ誰にも無いときだけ、本人への自己付与を許す(ブートストラップ)
        IF EXISTS (SELECT 1 FROM product_core.operator_permission
                    WHERE permission = 'OPERATOR_ADMIN' AND revoked_at IS NULL) THEN
            PERFORM product_core.assert_approval(NEW.granted_by, 'OPERATOR_ADMIN');
        ELSIF NEW.granted_by <> 'human:' || NEW.operator_id THEN
            RAISE EXCEPTION '最初の管理者権限は本人の名義で付与してください' USING ERRCODE = 'insufficient_privilege';
        END IF;
        RETURN NEW;
    END IF;
    IF (to_jsonb(NEW) - ARRAY['revoked_by', 'revoked_at']) IS DISTINCT FROM (to_jsonb(OLD) - ARRAY['revoked_by', 'revoked_at'])
       OR OLD.revoked_at IS NOT NULL OR NEW.revoked_by IS NULL THEN
        RAISE EXCEPTION '変更できるのは未取消の付与の取り消しだけです' USING ERRCODE = 'restrict_violation';
    END IF;
    PERFORM product_core.assert_approval(NEW.revoked_by, 'OPERATOR_ADMIN');
    NEW.revoked_at := now();
    RETURN NEW;
END;
$$;
DROP TRIGGER IF EXISTS operator_permission_guard ON product_core.operator_permission;
CREATE TRIGGER operator_permission_guard BEFORE INSERT OR UPDATE OR DELETE ON product_core.operator_permission
    FOR EACH ROW EXECUTE FUNCTION product_core.operator_permission_guard();


-- =============================================================
--  取込ステージング(legacy_ingest)
-- =============================================================
CREATE TABLE IF NOT EXISTS legacy_ingest.ingest_run (
    ingest_run_id text        PRIMARY KEY CHECK (ingest_run_id ~ '^IR-[0-9]{8}-[0-9]{6}-[0-9a-f]{4,}$'),
    started_at    timestamptz NOT NULL DEFAULT now(),
    finished_at   timestamptz,
    actor         text        NOT NULL CHECK (product_core.is_human(actor)),  -- 手動実行した人(D7)
    status        text        NOT NULL DEFAULT 'RUNNING' CHECK (status IN ('RUNNING', 'SUCCEEDED', 'FAILED')),
    files         jsonb       NOT NULL DEFAULT '[]'::jsonb,  -- パス・sha256(取込前後)・mtime・行数
    note          text,
    CHECK ((status = 'RUNNING') = (finished_at IS NULL))
);

CREATE OR REPLACE FUNCTION legacy_ingest.ingest_run_guard()
RETURNS trigger LANGUAGE plpgsql AS $$
BEGIN
    IF TG_OP = 'DELETE' THEN
        RAISE EXCEPTION 'ingest_run は削除できません' USING ERRCODE = 'restrict_violation';
    END IF;
    IF OLD.status <> 'RUNNING' THEN
        RAISE EXCEPTION '完了した ingest_run は変更できません' USING ERRCODE = 'restrict_violation';
    END IF;
    IF NEW.ingest_run_id <> OLD.ingest_run_id OR NEW.started_at <> OLD.started_at OR NEW.actor <> OLD.actor THEN
        RAISE EXCEPTION 'ingest_run の識別情報は変更できません' USING ERRCODE = 'restrict_violation';
    END IF;
    RETURN NEW;
END;
$$;
DROP TRIGGER IF EXISTS ingest_run_guard ON legacy_ingest.ingest_run;
CREATE TRIGGER ingest_run_guard BEFORE UPDATE OR DELETE ON legacy_ingest.ingest_run
    FOR EACH ROW EXECUTE FUNCTION legacy_ingest.ingest_run_guard();

CREATE TABLE IF NOT EXISTS legacy_ingest.legacy_product_snapshot (
    ingest_run_id text    NOT NULL REFERENCES legacy_ingest.ingest_run (ingest_run_id),
    row_no        integer NOT NULL,
    legacy_p      text,
    jan_raw       text,
    name          text,
    kind          text,
    standard_cost numeric(18, 6),
    channel       text,
    status        text,
    raw           jsonb   NOT NULL,
    PRIMARY KEY (ingest_run_id, row_no)
);

CREATE TABLE IF NOT EXISTS legacy_ingest.legacy_listing_snapshot (
    ingest_run_id     text    NOT NULL REFERENCES legacy_ingest.ingest_run (ingest_run_id),
    row_no            integer NOT NULL,
    legacy_listing_id text,
    legacy_p          text,
    channel           text,
    rakuten_item_id   text,
    rakuten_sku       text,
    asin              text,
    amazon_sku        text,
    price             numeric(18, 6),
    pack              integer,
    cost_unit         text,
    proof             text,
    raw               jsonb   NOT NULL,
    PRIMARY KEY (ingest_run_id, row_no)
);

CREATE TABLE IF NOT EXISTS legacy_ingest.external_file_snapshot (
    ingest_run_id text    NOT NULL REFERENCES legacy_ingest.ingest_run (ingest_run_id),
    file_kind     text    NOT NULL CHECK (btrim(file_kind) <> ''),
    row_no        integer NOT NULL,
    raw           jsonb   NOT NULL,
    PRIMARY KEY (ingest_run_id, file_kind, row_no)
);

DROP TRIGGER IF EXISTS legacy_product_snapshot_append_only ON legacy_ingest.legacy_product_snapshot;
CREATE TRIGGER legacy_product_snapshot_append_only BEFORE UPDATE OR DELETE ON legacy_ingest.legacy_product_snapshot
    FOR EACH ROW EXECUTE FUNCTION product_core.deny_mutation();
DROP TRIGGER IF EXISTS legacy_listing_snapshot_append_only ON legacy_ingest.legacy_listing_snapshot;
CREATE TRIGGER legacy_listing_snapshot_append_only BEFORE UPDATE OR DELETE ON legacy_ingest.legacy_listing_snapshot
    FOR EACH ROW EXECUTE FUNCTION product_core.deny_mutation();
DROP TRIGGER IF EXISTS external_file_snapshot_append_only ON legacy_ingest.external_file_snapshot;
CREATE TRIGGER external_file_snapshot_append_only BEFORE UPDATE OR DELETE ON legacy_ingest.external_file_snapshot
    FOR EACH ROW EXECUTE FUNCTION product_core.deny_mutation();


-- =============================================================
--  core_entity — 対象の登録簿(Identifier A 案・S1)
-- =============================================================
--  identifier_link などが PP / CP / Listing を実 FK で参照するための登録簿。
--  Phase 2 で STOCK_FORM を加えるときは、この表の CHECK に種類を足すだけで、
--  identifier / identifier_link の定義は変えない。
CREATE TABLE IF NOT EXISTS product_core.core_entity (
    entity_id   text        PRIMARY KEY,
    entity_type text        NOT NULL CHECK (entity_type IN ('PHYSICAL_PRODUCT', 'COMPOSITION', 'LISTING')),
    created_at  timestamptz NOT NULL DEFAULT now(),
    created_by  text        NOT NULL CHECK (product_core.is_actor(created_by)),
    CONSTRAINT core_entity_id_type_uq UNIQUE (entity_id, entity_type),
    -- ID の接頭辞と種類の食い違いを構造で防ぐ
    CONSTRAINT core_entity_prefix_matches_type CHECK (
           (entity_type = 'PHYSICAL_PRODUCT' AND entity_id ~ '^PP-[0-9]{6,}$')
        OR (entity_type = 'COMPOSITION'      AND entity_id ~ '^CP-[0-9]{6,}$')
        OR (entity_type = 'LISTING'          AND entity_id ~ '^LS-[0-9]{6,}$'))
);
COMMENT ON TABLE product_core.core_entity IS
    '対象の登録簿。実体表(physical_product / composition / listing)の挿入時にトリガーで登録される。';

DROP TRIGGER IF EXISTS core_entity_append_only ON product_core.core_entity;
CREATE TRIGGER core_entity_append_only BEFORE UPDATE OR DELETE ON product_core.core_entity
    FOR EACH ROW EXECUTE FUNCTION product_core.deny_mutation();

-- 実体表の挿入と同時に登録簿へ登録する。
-- SECURITY DEFINER: 取込ロールに core_entity への直接の INSERT 権限を与えないため。
CREATE OR REPLACE FUNCTION product_core.register_entity()
RETURNS trigger LANGUAGE plpgsql SECURITY DEFINER
SET search_path = pg_catalog, product_core AS $$
BEGIN
    INSERT INTO product_core.core_entity (entity_id, entity_type, created_by)
    VALUES (to_jsonb(NEW) ->> TG_ARGV[1], TG_ARGV[0], NEW.created_by);
    RETURN NEW;
END;
$$;

-- 登録簿だけあって実体が無い行を COMMIT 時に拒否する(S1)
CREATE OR REPLACE FUNCTION product_core.check_entity_body()
RETURNS trigger LANGUAGE plpgsql AS $$
DECLARE v_found boolean;
BEGIN
    v_found := CASE NEW.entity_type
        WHEN 'PHYSICAL_PRODUCT' THEN EXISTS (SELECT 1 FROM product_core.physical_product WHERE pp_id = NEW.entity_id)
        WHEN 'COMPOSITION'      THEN EXISTS (SELECT 1 FROM product_core.composition WHERE cp_id = NEW.entity_id)
        WHEN 'LISTING'          THEN EXISTS (SELECT 1 FROM product_core.listing WHERE listing_id = NEW.entity_id)
        ELSE false END;
    IF NOT v_found THEN
        RAISE EXCEPTION 'core_entity % (%) に対応する実体がありません', NEW.entity_id, NEW.entity_type
            USING ERRCODE = 'foreign_key_violation';
    END IF;
    RETURN NULL;
END;
$$;


-- =============================================================
--  physical_product — 倉庫で扱う最小の物理商品(P5 実体)
-- =============================================================
CREATE TABLE IF NOT EXISTS product_core.physical_product (
    pp_id          text        PRIMARY KEY DEFAULT product_core.fmt_id('PP-', nextval('product_core.pp_seq'), 6)
                               CHECK (pp_id ~ '^PP-[0-9]{6,}$'),
    entity_type    text        NOT NULL DEFAULT 'PHYSICAL_PRODUCT' CHECK (entity_type = 'PHYSICAL_PRODUCT'),
    name           text        NOT NULL CHECK (btrim(name) <> ''),
    kind           text        NOT NULL DEFAULT 'GOODS' CHECK (kind IN ('GOODS', 'ACCESSORY', 'MATERIAL')),
    status         text        NOT NULL DEFAULT 'PROVISIONAL' CHECK (status IN ('PROVISIONAL', 'ACTIVE', 'RETIRED')),
    -- 税率が書かれていない観測値を税抜に換算するときに使う(10% / 8% 軽減税率)
    tax_category   text        NOT NULL DEFAULT 'UNKNOWN' CHECK (tax_category IN ('STANDARD', 'REDUCED', 'UNKNOWN')),
    origin_key     text        NOT NULL UNIQUE CHECK (btrim(origin_key) <> ''),
    activated_by   text        CHECK (activated_by IS NULL OR product_core.is_human(activated_by)),
    activated_at   timestamptz,
    retired_by     text        CHECK (retired_by IS NULL OR product_core.is_human(retired_by)),
    retired_at     timestamptz,
    retire_reason  text,
    note           text,
    created_at     timestamptz NOT NULL DEFAULT now(),
    created_by     text        NOT NULL CHECK (product_core.is_actor(created_by)),
    updated_at     timestamptz NOT NULL DEFAULT now(),
    updated_by     text        NOT NULL CHECK (product_core.is_actor(updated_by)),
    ingest_run_id  text        REFERENCES legacy_ingest.ingest_run (ingest_run_id),
    CONSTRAINT physical_product_entity_fk FOREIGN KEY (pp_id, entity_type)
        REFERENCES product_core.core_entity (entity_id, entity_type),
    CHECK ((activated_by IS NULL) = (activated_at IS NULL)),
    CHECK ((retired_by IS NULL) = (retired_at IS NULL)),
    CHECK (status <> 'ACTIVE' OR activated_by IS NOT NULL),
    CHECK (status <> 'RETIRED' OR (retired_by IS NOT NULL AND retire_reason IS NOT NULL))
);
COMMENT ON TABLE product_core.physical_product IS
    '物理商品。AI・取込は PROVISIONAL でのみ作成。ACTIVE / RETIRED は人の承認(PRODUCT_APPROVE)。JAN・原価は持たない。';


-- =============================================================
--  composition / composition_component — 販売構成(P5 実体)
-- =============================================================
CREATE TABLE IF NOT EXISTS product_core.composition (
    cp_id          text        PRIMARY KEY DEFAULT product_core.fmt_id('CP-', nextval('product_core.cp_seq'), 6)
                               CHECK (cp_id ~ '^CP-[0-9]{6,}$'),
    entity_type    text        NOT NULL DEFAULT 'COMPOSITION' CHECK (entity_type = 'COMPOSITION'),
    name           text        NOT NULL CHECK (btrim(name) <> ''),
    -- 福袋など構成が未確認のものは UNCLASSIFIED(D10)
    comp_type      text        NOT NULL CHECK (comp_type IN ('FIXED', 'SELECTION', 'UNCLASSIFIED')),
    status         text        NOT NULL DEFAULT 'PROVISIONAL' CHECK (status IN ('PROVISIONAL', 'ACTIVE', 'RETIRED')),
    origin_key     text        NOT NULL UNIQUE CHECK (btrim(origin_key) <> ''),
    activated_by   text        CHECK (activated_by IS NULL OR product_core.is_human(activated_by)),
    activated_at   timestamptz,
    retired_by     text        CHECK (retired_by IS NULL OR product_core.is_human(retired_by)),
    retired_at     timestamptz,
    retire_reason  text,
    note           text,
    created_at     timestamptz NOT NULL DEFAULT now(),
    created_by     text        NOT NULL CHECK (product_core.is_actor(created_by)),
    updated_at     timestamptz NOT NULL DEFAULT now(),
    updated_by     text        NOT NULL CHECK (product_core.is_actor(updated_by)),
    ingest_run_id  text        REFERENCES legacy_ingest.ingest_run (ingest_run_id),
    CONSTRAINT composition_entity_fk FOREIGN KEY (cp_id, entity_type)
        REFERENCES product_core.core_entity (entity_id, entity_type),
    CHECK ((activated_by IS NULL) = (activated_at IS NULL)),
    CHECK ((retired_by IS NULL) = (retired_at IS NULL)),
    CHECK (status <> 'ACTIVE' OR activated_by IS NOT NULL),
    CHECK (status <> 'RETIRED' OR (retired_by IS NOT NULL AND retire_reason IS NOT NULL))
);

CREATE TABLE IF NOT EXISTS product_core.composition_component (
    component_id  text        PRIMARY KEY DEFAULT product_core.fmt_id('CC-', nextval('product_core.cc_seq'), 8)
                              CHECK (component_id ~ '^CC-[0-9]{8,}$'),
    cp_id         text        NOT NULL REFERENCES product_core.composition (cp_id),
    pp_id         text        NOT NULL REFERENCES product_core.physical_product (pp_id),
    quantity      integer     NOT NULL CHECK (quantity > 0),
    role          text        NOT NULL DEFAULT 'GOODS' CHECK (role IN ('GOODS', 'ACCESSORY')),
    sort_no       integer     NOT NULL DEFAULT 1,
    evidence      jsonb       NOT NULL,
    created_at    timestamptz NOT NULL DEFAULT now(),
    created_by    text        NOT NULL CHECK (product_core.is_actor(created_by)),
    CONSTRAINT composition_component_uq UNIQUE (cp_id, pp_id, role)
);

-- 状態の確認に使う(PP / CP / Listing 共通)
CREATE OR REPLACE FUNCTION product_core.entity_status(p_entity_id text, p_entity_type text)
RETURNS text LANGUAGE sql STABLE AS $$
    SELECT CASE p_entity_type
        WHEN 'PHYSICAL_PRODUCT' THEN (SELECT status FROM product_core.physical_product WHERE pp_id = p_entity_id)
        WHEN 'COMPOSITION'      THEN (SELECT status FROM product_core.composition WHERE cp_id = p_entity_id)
        WHEN 'LISTING'          THEN 'ACTIVE'
    END
$$;

-- PP / CP 共通の挿入規則: PROVISIONAL でのみ作る。承認情報は持たせない
CREATE OR REPLACE FUNCTION product_core.entity_before_insert()
RETURNS trigger LANGUAGE plpgsql AS $$
BEGIN
    IF NEW.status <> 'PROVISIONAL' OR NEW.activated_by IS NOT NULL OR NEW.retired_by IS NOT NULL THEN
        RAISE EXCEPTION '% は PROVISIONAL・承認情報なしで作成してください(承認は後から人が行う)', TG_TABLE_NAME
            USING ERRCODE = 'check_violation';
    END IF;
    -- ⚠️ 表に固有の列は、表名の判定を外側の IF に分けてから参照する
    --    (PL/pgSQL は式全体を解釈するため、同じ式に書くと他方の表で「列が無い」エラーになる)
    IF TG_TABLE_NAME = 'physical_product' THEN
        IF NEW.tax_category <> 'UNKNOWN' THEN
            RAISE EXCEPTION 'tax_category は UNKNOWN で作成し、人の承認で設定してください' USING ERRCODE = 'check_violation';
        END IF;
    ELSE
        IF NEW.comp_type = 'SELECTION' THEN
            RAISE EXCEPTION 'SELECTION は人の承認でのみ設定できます' USING ERRCODE = 'check_violation';
        END IF;
    END IF;
    NEW.updated_by := NEW.created_by;
    NEW.updated_at := now();
    RETURN NEW;
END;
$$;

-- PP / CP 共通の更新規則(C6)
CREATE OR REPLACE FUNCTION product_core.entity_before_update()
RETURNS trigger LANGUAGE plpgsql AS $$
DECLARE
    v_new jsonb := to_jsonb(NEW);
    v_old jsonb := to_jsonb(OLD);
    v_id  text  := CASE TG_TABLE_NAME WHEN 'physical_product' THEN 'pp_id' ELSE 'cp_id' END;
BEGIN
    IF TG_OP = 'DELETE' THEN
        RAISE EXCEPTION '% は削除できません(RETIRED にしてください)', TG_TABLE_NAME USING ERRCODE = 'restrict_violation';
    END IF;
    IF v_new -> v_id IS DISTINCT FROM v_old -> v_id OR NEW.entity_type <> OLD.entity_type
       OR NEW.origin_key <> OLD.origin_key OR NEW.created_at <> OLD.created_at OR NEW.created_by <> OLD.created_by THEN
        RAISE EXCEPTION 'ID・origin_key・作成情報は変更できません' USING ERRCODE = 'restrict_violation';
    END IF;

    IF NEW.status IS DISTINCT FROM OLD.status THEN
        IF OLD.status = 'PROVISIONAL' AND NEW.status = 'ACTIVE' THEN
            PERFORM product_core.assert_approval(NEW.activated_by, 'PRODUCT_APPROVE');
            NEW.activated_at := now();
            IF TG_TABLE_NAME = 'composition' THEN
                -- C9: 固定構成で、構成品があり、構成品がすべて ACTIVE のときだけ
                IF NEW.comp_type <> 'FIXED' THEN
                    RAISE EXCEPTION 'ACTIVE にできるのは FIXED の Composition だけです' USING ERRCODE = 'check_violation';
                END IF;
                IF NOT EXISTS (SELECT 1 FROM product_core.composition_component WHERE cp_id = NEW.cp_id) THEN
                    RAISE EXCEPTION '構成品が無い Composition は ACTIVE にできません' USING ERRCODE = 'check_violation';
                END IF;
                IF EXISTS (SELECT 1 FROM product_core.composition_component c
                             JOIN product_core.physical_product p ON p.pp_id = c.pp_id
                            WHERE c.cp_id = NEW.cp_id AND p.status <> 'ACTIVE') THEN
                    RAISE EXCEPTION '構成品に ACTIVE でない物理商品があります' USING ERRCODE = 'check_violation';
                END IF;
            END IF;
        ELSIF OLD.status IN ('PROVISIONAL', 'ACTIVE') AND NEW.status = 'RETIRED' THEN
            PERFORM product_core.assert_approval(NEW.retired_by, 'PRODUCT_APPROVE');
            NEW.retired_at := now();
        ELSE
            RAISE EXCEPTION '状態 % → % には変更できません', OLD.status, NEW.status USING ERRCODE = 'check_violation';
        END IF;
    END IF;
    -- 状態を変えない更新で承認情報を書き換えさせない
    IF NEW.status = OLD.status AND (NEW.activated_by IS DISTINCT FROM OLD.activated_by
                                    OR NEW.retired_by IS DISTINCT FROM OLD.retired_by) THEN
        RAISE EXCEPTION '承認情報だけを変更することはできません' USING ERRCODE = 'restrict_violation';
    END IF;
    IF NEW.status = OLD.status AND NEW.activated_at IS DISTINCT FROM OLD.activated_at THEN
        NEW.activated_at := OLD.activated_at;
    END IF;
    IF NEW.status = OLD.status AND NEW.retired_at IS DISTINCT FROM OLD.retired_at THEN
        NEW.retired_at := OLD.retired_at;
    END IF;

    IF TG_TABLE_NAME = 'physical_product' THEN
        IF NEW.tax_category IS DISTINCT FROM OLD.tax_category THEN
            PERFORM product_core.assert_approval(NEW.updated_by, 'PRODUCT_APPROVE');
        END IF;
    ELSE
        IF NEW.comp_type IS DISTINCT FROM OLD.comp_type THEN
            IF OLD.status <> 'PROVISIONAL' THEN
                RAISE EXCEPTION 'comp_type は PROVISIONAL の間だけ変更できます' USING ERRCODE = 'restrict_violation';
            END IF;
            IF NEW.comp_type = 'SELECTION' THEN
                PERFORM product_core.assert_approval(NEW.updated_by, 'PRODUCT_APPROVE');
            END IF;
            IF NEW.comp_type <> 'FIXED'
               AND EXISTS (SELECT 1 FROM product_core.composition_component WHERE cp_id = NEW.cp_id) THEN
                RAISE EXCEPTION '構成品がある Composition を FIXED 以外にはできません' USING ERRCODE = 'check_violation';
            END IF;
        END IF;
    END IF;
    NEW.updated_at := now();
    RETURN NEW;
END;
$$;

DROP TRIGGER IF EXISTS physical_product_register ON product_core.physical_product;
CREATE TRIGGER physical_product_register BEFORE INSERT ON product_core.physical_product
    FOR EACH ROW EXECUTE FUNCTION product_core.register_entity('PHYSICAL_PRODUCT', 'pp_id');
DROP TRIGGER IF EXISTS physical_product_insert_rules ON product_core.physical_product;
CREATE TRIGGER physical_product_insert_rules BEFORE INSERT ON product_core.physical_product
    FOR EACH ROW EXECUTE FUNCTION product_core.entity_before_insert();
DROP TRIGGER IF EXISTS physical_product_update_rules ON product_core.physical_product;
CREATE TRIGGER physical_product_update_rules BEFORE UPDATE OR DELETE ON product_core.physical_product
    FOR EACH ROW EXECUTE FUNCTION product_core.entity_before_update();

DROP TRIGGER IF EXISTS composition_register ON product_core.composition;
CREATE TRIGGER composition_register BEFORE INSERT ON product_core.composition
    FOR EACH ROW EXECUTE FUNCTION product_core.register_entity('COMPOSITION', 'cp_id');
DROP TRIGGER IF EXISTS composition_insert_rules ON product_core.composition;
CREATE TRIGGER composition_insert_rules BEFORE INSERT ON product_core.composition
    FOR EACH ROW EXECUTE FUNCTION product_core.entity_before_insert();
DROP TRIGGER IF EXISTS composition_update_rules ON product_core.composition;
CREATE TRIGGER composition_update_rules BEFORE UPDATE OR DELETE ON product_core.composition
    FOR EACH ROW EXECUTE FUNCTION product_core.entity_before_update();

-- 構成品は親が PROVISIONAL かつ FIXED の間だけ追加・変更・取り外しできる(構成の確定 = 親の ACTIVE 化)
CREATE OR REPLACE FUNCTION product_core.component_guard()
RETURNS trigger LANGUAGE plpgsql AS $$
DECLARE
    v_cp  text := CASE WHEN TG_OP = 'DELETE' THEN OLD.cp_id ELSE NEW.cp_id END;
    v_parent product_core.composition%ROWTYPE;
BEGIN
    SELECT * INTO v_parent FROM product_core.composition WHERE cp_id = v_cp;
    IF v_parent.status <> 'PROVISIONAL' THEN
        RAISE EXCEPTION 'Composition % は % のため構成品を変更できません', v_cp, v_parent.status
            USING ERRCODE = 'restrict_violation';
    END IF;
    IF TG_OP = 'UPDATE' AND (NEW.cp_id <> OLD.cp_id OR NEW.component_id <> OLD.component_id) THEN
        RAISE EXCEPTION '構成品の所属は変更できません' USING ERRCODE = 'restrict_violation';
    END IF;
    IF TG_OP <> 'DELETE' AND v_parent.comp_type <> 'FIXED' THEN
        RAISE EXCEPTION '構成品を持てるのは FIXED の Composition だけです(現在 %)', v_parent.comp_type
            USING ERRCODE = 'check_violation';
    END IF;
    IF TG_OP = 'DELETE' THEN RETURN OLD; END IF;
    RETURN NEW;
END;
$$;
DROP TRIGGER IF EXISTS composition_component_guard ON product_core.composition_component;
CREATE TRIGGER composition_component_guard BEFORE INSERT OR UPDATE OR DELETE ON product_core.composition_component
    FOR EACH ROW EXECUTE FUNCTION product_core.component_guard();


-- =============================================================
--  listing_group / listing — 販売口(P4 事実の写し)
-- =============================================================
CREATE TABLE IF NOT EXISTS product_core.listing_group (
    listing_group_id text        PRIMARY KEY DEFAULT product_core.fmt_id('LG-', nextval('product_core.lg_seq'), 6)
                                 CHECK (listing_group_id ~ '^LG-[0-9]{6,}$'),
    channel          text        NOT NULL CHECK (channel IN ('AMAZON', 'RAKUTEN', 'OTHER')),
    group_type       text        NOT NULL CHECK (group_type IN ('RAKUTEN_ITEM', 'AMAZON_PARENT_ASIN')),
    group_key        text        NOT NULL CHECK (btrim(group_key) <> ''),
    source_type      text        NOT NULL,
    source_ref       text        NOT NULL,
    created_at       timestamptz NOT NULL DEFAULT now(),
    created_by       text        NOT NULL CHECK (product_core.is_actor(created_by)),
    updated_at       timestamptz NOT NULL DEFAULT now(),
    updated_by       text        NOT NULL CHECK (product_core.is_actor(updated_by)),
    ingest_run_id    text        REFERENCES legacy_ingest.ingest_run (ingest_run_id),
    CONSTRAINT listing_group_natural_uq UNIQUE (channel, group_type, group_key)
);

CREATE TABLE IF NOT EXISTS product_core.listing (
    listing_id        text        PRIMARY KEY DEFAULT product_core.fmt_id('LS-', nextval('product_core.ls_seq'), 6)
                                  CHECK (listing_id ~ '^LS-[0-9]{6,}$'),
    entity_type       text        NOT NULL DEFAULT 'LISTING' CHECK (entity_type = 'LISTING'),
    channel           text        NOT NULL CHECK (channel IN ('AMAZON', 'RAKUTEN', 'OTHER')),
    asin              text        CHECK (asin IS NULL OR asin ~ '^B0[0-9A-Z]{8}$'),  -- 一意にしない
    seller_sku        text,
    rakuten_item_id   text,
    rakuten_sku       text,
    fulfillment       text        NOT NULL DEFAULT 'UNKNOWN' CHECK (fulfillment IN ('FBA', 'SELF', 'UNKNOWN')),
    listing_group_id  text        REFERENCES product_core.listing_group (listing_group_id),
    legacy_listing_id text        CHECK (legacy_listing_id IS NULL OR legacy_listing_id ~ '^C[0-9]{6}$'),
    channel_status    text        NOT NULL DEFAULT 'UNKNOWN' CHECK (channel_status IN ('ACTIVE', 'INACTIVE', 'UNKNOWN')),
    first_seen_at     timestamptz NOT NULL DEFAULT now(),
    last_seen_at      timestamptz NOT NULL DEFAULT now(),
    source_type       text        NOT NULL,
    source_ref        text        NOT NULL,
    created_at        timestamptz NOT NULL DEFAULT now(),
    created_by        text        NOT NULL CHECK (product_core.is_actor(created_by)),
    updated_at        timestamptz NOT NULL DEFAULT now(),
    updated_by        text        NOT NULL CHECK (product_core.is_actor(updated_by)),
    ingest_run_id     text        REFERENCES legacy_ingest.ingest_run (ingest_run_id),
    CONSTRAINT listing_entity_fk FOREIGN KEY (listing_id, entity_type)
        REFERENCES product_core.core_entity (entity_id, entity_type),
    CHECK (channel <> 'RAKUTEN' OR rakuten_item_id IS NOT NULL),
    CHECK (channel <> 'AMAZON' OR (rakuten_item_id IS NULL AND rakuten_sku IS NULL))
);
CREATE UNIQUE INDEX IF NOT EXISTS listing_legacy_uq ON product_core.listing (legacy_listing_id)
    WHERE legacy_listing_id IS NOT NULL;
CREATE UNIQUE INDEX IF NOT EXISTS listing_amazon_sku_uq ON product_core.listing (channel, seller_sku)
    WHERE channel = 'AMAZON' AND seller_sku IS NOT NULL;
CREATE UNIQUE INDEX IF NOT EXISTS listing_rakuten_uq ON product_core.listing (channel, rakuten_item_id, rakuten_sku)
    NULLS NOT DISTINCT WHERE channel = 'RAKUTEN';

-- 事実の写し: 自然キーは変えない(SKU が変わったら新しい listing)。削除しない
CREATE OR REPLACE FUNCTION product_core.fact_copy_guard()
RETURNS trigger LANGUAGE plpgsql AS $$
DECLARE v_keys text[];
BEGIN
    IF TG_OP = 'DELETE' THEN
        RAISE EXCEPTION '% は削除できません', TG_TABLE_NAME USING ERRCODE = 'restrict_violation';
    END IF;
    IF TG_OP = 'INSERT' THEN
        NEW.updated_by := NEW.created_by;
        NEW.updated_at := now();
        RETURN NEW;
    END IF;
    v_keys := CASE TG_TABLE_NAME
        WHEN 'listing' THEN ARRAY['listing_id', 'entity_type', 'channel', 'asin', 'seller_sku', 'rakuten_item_id',
                                  'rakuten_sku', 'legacy_listing_id', 'first_seen_at', 'created_at', 'created_by']
        ELSE ARRAY['listing_group_id', 'channel', 'group_type', 'group_key', 'created_at', 'created_by'] END;
    IF EXISTS (SELECT 1 FROM unnest(v_keys) k WHERE to_jsonb(NEW) -> k IS DISTINCT FROM to_jsonb(OLD) -> k) THEN
        RAISE EXCEPTION '% の自然キー・作成情報は変更できません', TG_TABLE_NAME USING ERRCODE = 'restrict_violation';
    END IF;
    NEW.updated_at := now();
    RETURN NEW;
END;
$$;
DROP TRIGGER IF EXISTS listing_group_guard ON product_core.listing_group;
CREATE TRIGGER listing_group_guard BEFORE INSERT OR UPDATE OR DELETE ON product_core.listing_group
    FOR EACH ROW EXECUTE FUNCTION product_core.fact_copy_guard();
DROP TRIGGER IF EXISTS listing_register ON product_core.listing;
CREATE TRIGGER listing_register BEFORE INSERT ON product_core.listing
    FOR EACH ROW EXECUTE FUNCTION product_core.register_entity('LISTING', 'listing_id');
DROP TRIGGER IF EXISTS listing_guard ON product_core.listing;
CREATE TRIGGER listing_guard BEFORE INSERT OR UPDATE OR DELETE ON product_core.listing
    FOR EACH ROW EXECUTE FUNCTION product_core.fact_copy_guard();

-- 実体表を作った後で登録簿の遅延検査を張る(S1)
DROP TRIGGER IF EXISTS core_entity_has_body ON product_core.core_entity;
CREATE CONSTRAINT TRIGGER core_entity_has_body AFTER INSERT ON product_core.core_entity
    DEFERRABLE INITIALLY DEFERRED
    FOR EACH ROW EXECUTE FUNCTION product_core.check_entity_body();


-- =============================================================
--  identifier — コード登録簿(P3 追記専用・対象を持たない)
-- =============================================================
CREATE TABLE IF NOT EXISTS product_core.identifier (
    identifier_id     text        PRIMARY KEY DEFAULT product_core.fmt_id('ID-', nextval('product_core.id_seq'), 8)
                                  CHECK (identifier_id ~ '^ID-[0-9]{8,}$'),
    id_type           text        NOT NULL CHECK (id_type IN ('JAN', 'GTIN14', 'MAKER_CODE', 'JAN_PARTIAL4',
                                                               'ASIN', 'FNSKU', 'KIT_LABEL', 'TEMP_ID')),
    value             text        NOT NULL,
    checkdigit_valid  boolean,
    first_seen_source text        NOT NULL CHECK (btrim(first_seen_source) <> ''),
    first_seen_ref    text        NOT NULL CHECK (btrim(first_seen_ref) <> ''),
    created_at        timestamptz NOT NULL DEFAULT now(),
    created_by        text        NOT NULL CHECK (product_core.is_actor(created_by)),
    ingest_run_id     text        REFERENCES legacy_ingest.ingest_run (ingest_run_id),
    CONSTRAINT identifier_code_uq UNIQUE (id_type, value),
    CONSTRAINT identifier_value_format CHECK (CASE id_type
        WHEN 'JAN'          THEN value ~ '^([0-9]{8}|[0-9]{12,13})$'
        WHEN 'GTIN14'       THEN value ~ '^[0-9]{14}$'
        WHEN 'JAN_PARTIAL4' THEN value ~ '^[0-9]{4}$'
        WHEN 'ASIN'         THEN value ~ '^B0[0-9A-Z]{8}$'
        WHEN 'TEMP_ID'      THEN value ~ '^T-[0-9]{6}$'
        ELSE btrim(value) <> '' END)
);
DROP TRIGGER IF EXISTS identifier_append_only ON product_core.identifier;
CREATE TRIGGER identifier_append_only BEFORE UPDATE OR DELETE ON product_core.identifier
    FOR EACH ROW EXECUTE FUNCTION product_core.deny_mutation();


-- =============================================================
--  assessment — AI 判定(追記専用)+ 人の review 欄(D11)
-- =============================================================
CREATE TABLE IF NOT EXISTS product_core.assessment (
    assessment_id            text        PRIMARY KEY DEFAULT product_core.fmt_id('AS-', nextval('product_core.as_seq'), 9)
                                         CHECK (assessment_id ~ '^AS-[0-9]{9,}$'),
    assessment_type          text        NOT NULL CHECK (assessment_type IN ('LEGACY_MAPPING', 'LISTING_REFERENCE',
                                             'COST_VARIANCE', 'TAX_BASIS', 'DUPLICATE_PP', 'IDENTIFIER_LINK',
                                             'COMPOSITION_CLASS')),
    subject_entity_id        text,
    subject_entity_type      text,
    subject_key              text,
    verdict                  text        NOT NULL,
    proposed_entity_id       text,
    proposed_entity_type     text,
    confidence               text        NOT NULL CHECK (confidence IN ('HIGH', 'MEDIUM', 'LOW')),
    reason                   text        NOT NULL CHECK (btrim(reason) <> ''),
    evidence                 jsonb       NOT NULL,
    rule_version             text        NOT NULL CHECK (btrim(rule_version) <> ''),
    -- 人は assessment を作らない。人の判断は review 欄に書く
    assessed_by              text        NOT NULL CHECK (assessed_by ~ '^(ai|rule):[^[:space:]]+$'),
    assessed_at              timestamptz NOT NULL DEFAULT now(),
    supersedes_assessment_id text        REFERENCES product_core.assessment (assessment_id),
    content_hash             text        NOT NULL,  -- AI 部分の sha256(Excel 確認表の改ざん検知)。トリガーで計算
    -- ---- 人の review 欄 ----
    review_status            text        NOT NULL DEFAULT 'UNREVIEWED'
                                         CHECK (review_status IN ('UNREVIEWED', 'ON_HOLD', 'APPROVED', 'REJECTED', 'CORRECTED')),
    reviewed_by              text        CHECK (reviewed_by IS NULL OR product_core.is_human(reviewed_by)),
    reviewed_at              timestamptz,
    review_note              text,
    corrected_entity_id      text,
    corrected_entity_type    text,
    corrected_value          text,        -- 対象ではなく値を人が修正したとき(例: TAX_BASIS の INCLUDED / EXCLUDED)
    review_channel           text        CHECK (review_channel IS NULL OR review_channel IN ('EXCEL', 'UI')),
    review_ref               text,
    FOREIGN KEY (subject_entity_id, subject_entity_type) REFERENCES product_core.core_entity (entity_id, entity_type),
    FOREIGN KEY (proposed_entity_id, proposed_entity_type) REFERENCES product_core.core_entity (entity_id, entity_type),
    FOREIGN KEY (corrected_entity_id, corrected_entity_type) REFERENCES product_core.core_entity (entity_id, entity_type),
    CHECK (subject_entity_id IS NOT NULL OR subject_key IS NOT NULL),
    CHECK ((subject_entity_id IS NULL) = (subject_entity_type IS NULL)),
    CHECK ((proposed_entity_id IS NULL) = (proposed_entity_type IS NULL)),
    CHECK ((corrected_entity_id IS NULL) = (corrected_entity_type IS NULL)),
    CHECK ((review_status = 'UNREVIEWED') = (reviewed_by IS NULL)),
    CHECK ((reviewed_by IS NULL) = (reviewed_at IS NULL)),
    CHECK (review_status NOT IN ('ON_HOLD', 'REJECTED', 'CORRECTED') OR btrim(coalesce(review_note, '')) <> ''),
    CHECK (review_status <> 'CORRECTED' OR corrected_entity_id IS NOT NULL OR corrected_value IS NOT NULL),
    CHECK (assessment_type <> 'TAX_BASIS' OR corrected_value IS NULL OR corrected_value IN ('INCLUDED', 'EXCLUDED', 'UNKNOWN')),
    CONSTRAINT assessment_verdict_by_type CHECK (
           (assessment_type = 'LEGACY_MAPPING'    AND verdict IN ('PHYSICAL_PRODUCT_CANDIDATE', 'COMPOSITION_CANDIDATE',
                                                                  'ALIAS_CANDIDATE', 'NEEDS_REVIEW'))
        OR (assessment_type = 'LISTING_REFERENCE' AND verdict IN ('CONSISTENT_HIGH', 'SUSPECT_MISLINK', 'AWAITING_DATA',
                                                                  'HUMAN_REVIEW_CANDIDATE'))
        OR (assessment_type = 'COST_VARIANCE'     AND verdict IN ('OK', 'CONFIRMED_WARN', 'REFERENCE_WARN',
                                                                  'TAX_BASIS_UNKNOWN', 'TAX_RATE_UNKNOWN',
                                                                  'UNIT_UNKNOWN', 'NO_BASELINE'))
        OR (assessment_type = 'TAX_BASIS'         AND verdict IN ('INCLUDED', 'EXCLUDED', 'UNKNOWN'))
        OR (assessment_type = 'DUPLICATE_PP'      AND verdict IN ('DUPLICATE_CANDIDATE', 'VARIANT_CANDIDATE', 'NOT_DUPLICATE'))
        OR (assessment_type = 'IDENTIFIER_LINK'   AND verdict IN ('VALID', 'CHECKDIGIT_ERROR', 'MISFILED',
                                                                  'SHARED_ACROSS_ENTITIES'))
        OR (assessment_type = 'COMPOSITION_CLASS' AND verdict IN ('FIXED_CANDIDATE', 'SELECTION_CANDIDATE',
                                                                  'HUMAN_REVIEW_REQUIRED')))
);
COMMENT ON TABLE product_core.assessment IS
    'AI 判定(追記専用)と人の review 欄。verdict が CONSISTENT_HIGH でも review_status は UNREVIEWED のまま。';

CREATE OR REPLACE FUNCTION product_core.assessment_review_columns()
RETURNS text[] LANGUAGE sql IMMUTABLE AS $$
    SELECT ARRAY['review_status', 'reviewed_by', 'reviewed_at', 'review_note', 'corrected_entity_id',
                 'corrected_entity_type', 'corrected_value', 'review_channel', 'review_ref']
$$;

CREATE OR REPLACE FUNCTION product_core.assessment_guard()
RETURNS trigger LANGUAGE plpgsql AS $$
DECLARE v_ai jsonb;
BEGIN
    IF TG_OP = 'DELETE' THEN
        RAISE EXCEPTION 'assessment は削除できません' USING ERRCODE = 'restrict_violation';
    END IF;
    IF TG_OP = 'INSERT' THEN
        IF NEW.review_status <> 'UNREVIEWED' OR NEW.reviewed_by IS NOT NULL OR NEW.corrected_entity_id IS NOT NULL
           OR NEW.corrected_value IS NOT NULL
           OR NEW.review_channel IS NOT NULL OR NEW.review_ref IS NOT NULL OR NEW.review_note IS NOT NULL THEN
            RAISE EXCEPTION 'assessment は UNREVIEWED・review 欄空欄で作成してください' USING ERRCODE = 'check_violation';
        END IF;
        NEW.assessed_at := now();
        v_ai := to_jsonb(NEW) - product_core.assessment_review_columns() - 'content_hash';
        NEW.content_hash := encode(sha256(convert_to(v_ai::text, 'UTF8')), 'hex');
        RETURN NEW;
    END IF;
    -- AI 部分は変更不可
    IF (to_jsonb(NEW) - product_core.assessment_review_columns())
       IS DISTINCT FROM (to_jsonb(OLD) - product_core.assessment_review_columns()) THEN
        RAISE EXCEPTION 'assessment の AI 部分は変更できません(再判定は新しい assessment で)' USING ERRCODE = 'restrict_violation';
    END IF;
    IF OLD.review_status IN ('APPROVED', 'REJECTED', 'CORRECTED') THEN
        RAISE EXCEPTION 'review は確定済みです(%)。やり直しは新しい assessment で行ってください', OLD.review_status
            USING ERRCODE = 'restrict_violation';
    END IF;
    IF NEW.review_status = 'UNREVIEWED' THEN
        RAISE EXCEPTION 'UNREVIEWED には戻せません' USING ERRCODE = 'check_violation';
    END IF;
    PERFORM product_core.assert_approval(NEW.reviewed_by, 'REVIEW');
    NEW.reviewed_at := now();
    RETURN NEW;
END;
$$;
DROP TRIGGER IF EXISTS assessment_guard ON product_core.assessment;
CREATE TRIGGER assessment_guard BEFORE INSERT OR UPDATE OR DELETE ON product_core.assessment
    FOR EACH ROW EXECUTE FUNCTION product_core.assessment_guard();

-- C11: ACTIVE 化の根拠になる assessment は、人の review が APPROVED / CORRECTED であること
CREATE OR REPLACE FUNCTION product_core.assert_reviewed_assessment(p_assessment_id text)
RETURNS void LANGUAGE plpgsql STABLE AS $$
DECLARE v_status text;
BEGIN
    IF p_assessment_id IS NULL THEN RETURN; END IF;
    SELECT review_status INTO v_status FROM product_core.assessment WHERE assessment_id = p_assessment_id;
    IF v_status IS NULL OR v_status NOT IN ('APPROVED', 'CORRECTED') THEN
        RAISE EXCEPTION 'assessment % は人の review で承認されていません(%)', p_assessment_id, coalesce(v_status, '存在しない')
            USING ERRCODE = 'check_violation';
    END IF;
END;
$$;


-- =============================================================
--  提案 → 承認パターン(P1)の共通ガード
-- =============================================================
--  TG_ARGV[0] = ACTIVE 化に必要な権限 / TG_ARGV[1] = 承認時だけ変えてよい追加列(カンマ区切り)
CREATE OR REPLACE FUNCTION product_core.p1_guard()
RETURNS trigger LANGUAGE plpgsql AS $$
DECLARE
    v_perm  text   := TG_ARGV[0];
    v_extra text[] := CASE WHEN TG_NARGS > 1 AND TG_ARGV[1] <> '' THEN string_to_array(TG_ARGV[1], ',') ELSE ARRAY[]::text[] END;
    v_mut   text[];
    v_extra_changed boolean;
BEGIN
    IF TG_OP = 'DELETE' THEN
        RAISE EXCEPTION '% は削除できません(REJECTED / SUPERSEDED で表します)', TG_TABLE_NAME USING ERRCODE = 'restrict_violation';
    END IF;
    IF TG_OP = 'INSERT' THEN
        NEW.updated_by := NEW.created_by;
        NEW.updated_at := now();
        IF NEW.superseded_by IS NOT NULL THEN
            RAISE EXCEPTION 'superseded_by を持つ行は作れません' USING ERRCODE = 'check_violation';
        END IF;
        IF NEW.record_status = 'PROPOSED' THEN
            IF NEW.approved_by IS NOT NULL OR NEW.approved_at IS NOT NULL THEN
                RAISE EXCEPTION 'PROPOSED の行に承認情報は書けません' USING ERRCODE = 'check_violation';
            END IF;
        ELSIF NEW.record_status = 'ACTIVE' THEN
            -- 人が修正先を指定して直接 ACTIVE で作る場合(CORRECTED の流れ)
            PERFORM product_core.assert_approval(NEW.approved_by, v_perm);
            PERFORM product_core.assert_reviewed_assessment(NEW.assessment_id);
            NEW.approved_at := now();
        ELSE
            RAISE EXCEPTION '% は PROPOSED か ACTIVE で作成してください', TG_TABLE_NAME USING ERRCODE = 'check_violation';
        END IF;
        RETURN NEW;
    END IF;

    -- C3: 内容列は変更不可
    v_mut := ARRAY['record_status', 'approved_by', 'approved_at', 'superseded_by', 'valid_to', 'updated_at', 'updated_by']
             || v_extra;
    IF (to_jsonb(NEW) - v_mut) IS DISTINCT FROM (to_jsonb(OLD) - v_mut) THEN
        RAISE EXCEPTION '% の内容列は変更できません(新しい行を作ってください)', TG_TABLE_NAME USING ERRCODE = 'restrict_violation';
    END IF;
    v_extra_changed := EXISTS (SELECT 1 FROM unnest(v_extra) k WHERE to_jsonb(NEW) -> k IS DISTINCT FROM to_jsonb(OLD) -> k);

    IF OLD.record_status = 'PROPOSED' AND NEW.record_status IN ('ACTIVE', 'REJECTED') THEN
        PERFORM product_core.assert_approval(NEW.approved_by, v_perm);
        IF NEW.record_status = 'ACTIVE' THEN
            PERFORM product_core.assert_reviewed_assessment(NEW.assessment_id);
        END IF;
        NEW.approved_at := now();
    ELSIF OLD.record_status = 'ACTIVE' AND NEW.record_status = 'SUPERSEDED' THEN
        IF NEW.superseded_by IS NULL THEN
            RAISE EXCEPTION 'SUPERSEDED には superseded_by が必要です' USING ERRCODE = 'check_violation';
        END IF;
        PERFORM product_core.assert_approval(NEW.updated_by, v_perm);
    ELSIF OLD.record_status = NEW.record_status AND OLD.record_status = 'ACTIVE' THEN
        -- 期間終了(valid_to の一度だけの設定)と、承認時に確定する追加列だけ
        IF (to_jsonb(OLD) ->> 'valid_to') IS NOT NULL
           AND (to_jsonb(NEW) ->> 'valid_to') IS DISTINCT FROM (to_jsonb(OLD) ->> 'valid_to') THEN
            RAISE EXCEPTION 'valid_to は一度だけ設定できます' USING ERRCODE = 'restrict_violation';
        END IF;
        IF (to_jsonb(NEW) ->> 'valid_to') IS DISTINCT FROM (to_jsonb(OLD) ->> 'valid_to') OR v_extra_changed THEN
            PERFORM product_core.assert_approval(NEW.updated_by, v_perm);
        END IF;
    ELSIF OLD.record_status = NEW.record_status AND OLD.record_status = 'PROPOSED' THEN
        IF v_extra_changed OR (to_jsonb(NEW) ->> 'valid_to') IS DISTINCT FROM (to_jsonb(OLD) ->> 'valid_to') THEN
            RAISE EXCEPTION 'PROPOSED の行は変更できません(新しい提案を作ってください)' USING ERRCODE = 'restrict_violation';
        END IF;
    ELSIF OLD.record_status = NEW.record_status THEN
        RAISE EXCEPTION '% の行は変更できません', OLD.record_status USING ERRCODE = 'restrict_violation';
    ELSE
        RAISE EXCEPTION '状態 % → % には変更できません', OLD.record_status, NEW.record_status USING ERRCODE = 'check_violation';
    END IF;
    -- 承認情報を変えてよいのは PROPOSED → ACTIVE / REJECTED の遷移だけ
    IF NOT (OLD.record_status = 'PROPOSED' AND NEW.record_status IN ('ACTIVE', 'REJECTED'))
       AND (NEW.approved_by IS DISTINCT FROM OLD.approved_by OR NEW.approved_at IS DISTINCT FROM OLD.approved_at) THEN
        RAISE EXCEPTION '承認情報だけを変更することはできません' USING ERRCODE = 'restrict_violation';
    END IF;
    NEW.updated_at := now();
    RETURN NEW;
END;
$$;


-- =============================================================
--  identifier_link — コード → 対象の候補(P1・多対多)
-- =============================================================
CREATE TABLE IF NOT EXISTS product_core.identifier_link (
    link_id        text        PRIMARY KEY DEFAULT product_core.fmt_id('IL-', nextval('product_core.il_seq'), 8)
                               CHECK (link_id ~ '^IL-[0-9]{8,}$'),
    identifier_id  text        NOT NULL REFERENCES product_core.identifier (identifier_id),
    entity_id      text        NOT NULL,
    entity_type    text        NOT NULL,
    link_basis     text        NOT NULL CHECK (link_basis IN ('PRINTED_ON_ITEM', 'PRINTED_ON_PACKAGE', 'CASE_OUTER',
                                                              'SELLER_REGISTERED', 'DERIVED_FROM_SKU', 'LEGACY_MASTER',
                                                              'MISFILED')),
    scan_policy    text        NOT NULL DEFAULT 'CONFIRM_ALWAYS'
                               CHECK (scan_policy IN ('AUTO_IF_UNIQUE', 'CONFIRM_ALWAYS', 'NOT_FOR_SCAN')),
    valid_from     date,
    valid_to       date,
    confidence     text        NOT NULL CHECK (confidence IN ('HIGH', 'MEDIUM', 'LOW')),
    evidence       jsonb       NOT NULL,
    record_status  text        NOT NULL DEFAULT 'PROPOSED'
                               CHECK (record_status IN ('PROPOSED', 'ACTIVE', 'REJECTED', 'SUPERSEDED')),
    approved_by    text        CHECK (approved_by IS NULL OR product_core.is_human(approved_by)),
    approved_at    timestamptz,
    assessment_id  text        REFERENCES product_core.assessment (assessment_id),
    superseded_by  text        REFERENCES product_core.identifier_link (link_id),
    created_at     timestamptz NOT NULL DEFAULT now(),
    created_by     text        NOT NULL CHECK (product_core.is_actor(created_by)),
    updated_at     timestamptz NOT NULL DEFAULT now(),
    updated_by     text        NOT NULL CHECK (product_core.is_actor(updated_by)),
    ingest_run_id  text        REFERENCES legacy_ingest.ingest_run (ingest_run_id),
    -- 種類を制限しない(Phase 2 で STOCK_FORM が加わっても定義は不変)
    FOREIGN KEY (entity_id, entity_type) REFERENCES product_core.core_entity (entity_id, entity_type),
    CHECK ((approved_by IS NULL) = (approved_at IS NULL)),
    CHECK (record_status NOT IN ('ACTIVE', 'REJECTED', 'SUPERSEDED') OR approved_by IS NOT NULL),
    CHECK (record_status <> 'SUPERSEDED' OR superseded_by IS NOT NULL),
    CHECK (valid_to IS NULL OR valid_from IS NULL OR valid_from <= valid_to),
    -- S6: 自動確定の許可は人の承認後だけ。誤記由来の紐付けはスキャンに使わない
    CHECK (scan_policy <> 'AUTO_IF_UNIQUE' OR record_status IN ('ACTIVE', 'SUPERSEDED')),
    CHECK (link_basis <> 'MISFILED' OR scan_policy = 'NOT_FOR_SCAN')
);
CREATE UNIQUE INDEX IF NOT EXISTS identifier_link_open_uq ON product_core.identifier_link (identifier_id, entity_id)
    WHERE record_status IN ('PROPOSED', 'ACTIVE');

CREATE OR REPLACE FUNCTION product_core.identifier_link_rules()
RETURNS trigger LANGUAGE plpgsql AS $$
DECLARE v_type text;
BEGIN
    SELECT id_type INTO v_type FROM product_core.identifier WHERE identifier_id = NEW.identifier_id;
    -- 部分識別子・ASIN は物を特定できないのでスキャンに使わない
    IF v_type IN ('JAN_PARTIAL4', 'ASIN') AND NEW.scan_policy <> 'NOT_FOR_SCAN' THEN
        RAISE EXCEPTION '% の紐付けは NOT_FOR_SCAN にしてください', v_type USING ERRCODE = 'check_violation';
    END IF;
    RETURN NEW;
END;
$$;
DROP TRIGGER IF EXISTS a_identifier_link_p1 ON product_core.identifier_link;
CREATE TRIGGER a_identifier_link_p1 BEFORE INSERT OR UPDATE OR DELETE ON product_core.identifier_link
    FOR EACH ROW EXECUTE FUNCTION product_core.p1_guard('IDENTIFIER_APPROVE', 'scan_policy');
DROP TRIGGER IF EXISTS b_identifier_link_rules ON product_core.identifier_link;
CREATE TRIGGER b_identifier_link_rules BEFORE INSERT OR UPDATE ON product_core.identifier_link
    FOR EACH ROW EXECUTE FUNCTION product_core.identifier_link_rules();


-- =============================================================
--  legacy_mapping — 既存 P が何を表しているか(P1)
-- =============================================================
CREATE TABLE IF NOT EXISTS product_core.legacy_mapping (
    mapping_id          text        PRIMARY KEY DEFAULT product_core.fmt_id('LM-', nextval('product_core.lm_seq'), 8)
                                    CHECK (mapping_id ~ '^LM-[0-9]{8,}$'),
    legacy_p            text        NOT NULL CHECK (legacy_p ~ '^P[0-9]{6}$'),  -- 正本は Excel(FK なし)
    entity_id           text        NOT NULL,
    entity_type         text        NOT NULL CHECK (entity_type IN ('PHYSICAL_PRODUCT', 'COMPOSITION')),
    mapping_role        text        NOT NULL CHECK (mapping_role IN ('SOLE', 'MERGED')),  -- 代表 P は決めない
    valid_from          date,
    valid_to            date,
    legacy_snapshot_ref text        NOT NULL CHECK (btrim(legacy_snapshot_ref) <> ''),
    record_status       text        NOT NULL DEFAULT 'PROPOSED'
                                    CHECK (record_status IN ('PROPOSED', 'ACTIVE', 'REJECTED', 'SUPERSEDED')),
    approved_by         text        CHECK (approved_by IS NULL OR product_core.is_human(approved_by)),
    approved_at         timestamptz,
    assessment_id       text        REFERENCES product_core.assessment (assessment_id),
    superseded_by       text        REFERENCES product_core.legacy_mapping (mapping_id),
    created_at          timestamptz NOT NULL DEFAULT now(),
    created_by          text        NOT NULL CHECK (product_core.is_actor(created_by)),
    updated_at          timestamptz NOT NULL DEFAULT now(),
    updated_by          text        NOT NULL CHECK (product_core.is_actor(updated_by)),
    ingest_run_id       text        REFERENCES legacy_ingest.ingest_run (ingest_run_id),
    FOREIGN KEY (entity_id, entity_type) REFERENCES product_core.core_entity (entity_id, entity_type),
    CHECK ((approved_by IS NULL) = (approved_at IS NULL)),
    CHECK (record_status NOT IN ('ACTIVE', 'REJECTED', 'SUPERSEDED') OR approved_by IS NOT NULL),
    CHECK (record_status <> 'SUPERSEDED' OR superseded_by IS NOT NULL),
    CHECK (valid_to IS NULL OR valid_from IS NULL OR valid_from <= valid_to)
);
-- legacy P ごとに ACTIVE は1行だけ
CREATE UNIQUE INDEX IF NOT EXISTS legacy_mapping_active_uq ON product_core.legacy_mapping (legacy_p)
    WHERE record_status = 'ACTIVE';

CREATE OR REPLACE FUNCTION product_core.legacy_mapping_rules()
RETURNS trigger LANGUAGE plpgsql AS $$
BEGIN
    IF product_core.entity_status(NEW.entity_id, NEW.entity_type) = 'RETIRED'
       AND (TG_OP = 'INSERT' OR (NEW.record_status = 'ACTIVE' AND OLD.record_status <> 'ACTIVE')) THEN
        RAISE EXCEPTION 'RETIRED の % へは対応付けできません', NEW.entity_id USING ERRCODE = 'check_violation';
    END IF;
    RETURN NEW;
END;
$$;
DROP TRIGGER IF EXISTS a_legacy_mapping_p1 ON product_core.legacy_mapping;
CREATE TRIGGER a_legacy_mapping_p1 BEFORE INSERT OR UPDATE OR DELETE ON product_core.legacy_mapping
    FOR EACH ROW EXECUTE FUNCTION product_core.p1_guard('LEGACY_MAPPING_APPROVE');
DROP TRIGGER IF EXISTS b_legacy_mapping_rules ON product_core.legacy_mapping;
CREATE TRIGGER b_legacy_mapping_rules BEFORE INSERT OR UPDATE ON product_core.legacy_mapping
    FOR EACH ROW EXECUTE FUNCTION product_core.legacy_mapping_rules();


-- =============================================================
--  product_relationship — PP 同士の関係(P1)
-- =============================================================
CREATE TABLE IF NOT EXISTS product_core.product_relationship (
    relationship_id   text        PRIMARY KEY DEFAULT product_core.fmt_id('PR-', nextval('product_core.pr_seq'), 8)
                                  CHECK (relationship_id ~ '^PR-[0-9]{8,}$'),
    rel_type          text        NOT NULL CHECK (rel_type IN ('DUPLICATE_OF', 'SUCCESSOR_OF', 'VARIANT_MEMBER')),
    from_pp_id        text        NOT NULL REFERENCES product_core.physical_product (pp_id),
    to_pp_id          text        REFERENCES product_core.physical_product (pp_id),
    variant_group_key text,
    variant_axis      text        CHECK (variant_axis IS NULL OR variant_axis IN ('COLOR', 'SIZE', 'SCENT', 'CAPACITY',
                                                                                  'DESIGN', 'OTHER')),
    evidence          jsonb       NOT NULL,
    record_status     text        NOT NULL DEFAULT 'PROPOSED'
                                  CHECK (record_status IN ('PROPOSED', 'ACTIVE', 'REJECTED', 'SUPERSEDED')),
    approved_by       text        CHECK (approved_by IS NULL OR product_core.is_human(approved_by)),
    approved_at       timestamptz,
    assessment_id     text        REFERENCES product_core.assessment (assessment_id),
    superseded_by     text        REFERENCES product_core.product_relationship (relationship_id),
    created_at        timestamptz NOT NULL DEFAULT now(),
    created_by        text        NOT NULL CHECK (product_core.is_actor(created_by)),
    updated_at        timestamptz NOT NULL DEFAULT now(),
    updated_by        text        NOT NULL CHECK (product_core.is_actor(updated_by)),
    CHECK ((approved_by IS NULL) = (approved_at IS NULL)),
    CHECK (record_status NOT IN ('ACTIVE', 'REJECTED', 'SUPERSEDED') OR approved_by IS NOT NULL),
    CHECK (record_status <> 'SUPERSEDED' OR superseded_by IS NOT NULL),
    CHECK (from_pp_id <> to_pp_id),
    CHECK ((rel_type = 'VARIANT_MEMBER') = (to_pp_id IS NULL)),
    CHECK (rel_type <> 'VARIANT_MEMBER' OR variant_group_key IS NOT NULL)
);
CREATE UNIQUE INDEX IF NOT EXISTS product_relationship_active_uq
    ON product_core.product_relationship (rel_type, from_pp_id, (coalesce(to_pp_id, variant_group_key)))
    WHERE record_status = 'ACTIVE';
DROP TRIGGER IF EXISTS a_product_relationship_p1 ON product_core.product_relationship;
CREATE TRIGGER a_product_relationship_p1 BEFORE INSERT OR UPDATE OR DELETE ON product_core.product_relationship
    FOR EACH ROW EXECUTE FUNCTION product_core.p1_guard('RELATIONSHIP_APPROVE');


-- =============================================================
--  listing_reference_history — 出品の参照先(P1・期間付き)
-- =============================================================
CREATE TABLE IF NOT EXISTS product_core.listing_reference_history (
    ref_id                text        PRIMARY KEY DEFAULT product_core.fmt_id('LR-', nextval('product_core.lr_seq'), 8)
                                      CHECK (ref_id ~ '^LR-[0-9]{8,}$'),
    listing_id            text        NOT NULL REFERENCES product_core.listing (listing_id),
    entity_id             text        NOT NULL,
    entity_type           text        NOT NULL CHECK (entity_type IN ('PHYSICAL_PRODUCT', 'COMPOSITION')),
    valid_from            date        NOT NULL,
    valid_to              date,
    derived_from_legacy_p text        CHECK (derived_from_legacy_p IS NULL OR derived_from_legacy_p ~ '^P[0-9]{6}$'),
    reason                text        NOT NULL CHECK (btrim(reason) <> ''),
    record_status         text        NOT NULL DEFAULT 'PROPOSED'
                                      CHECK (record_status IN ('PROPOSED', 'ACTIVE', 'REJECTED', 'SUPERSEDED')),
    approved_by           text        CHECK (approved_by IS NULL OR product_core.is_human(approved_by)),
    approved_at           timestamptz,
    assessment_id         text        REFERENCES product_core.assessment (assessment_id),
    superseded_by         text        REFERENCES product_core.listing_reference_history (ref_id),
    created_at            timestamptz NOT NULL DEFAULT now(),
    created_by            text        NOT NULL CHECK (product_core.is_actor(created_by)),
    updated_at            timestamptz NOT NULL DEFAULT now(),
    updated_by            text        NOT NULL CHECK (product_core.is_actor(updated_by)),
    FOREIGN KEY (entity_id, entity_type) REFERENCES product_core.core_entity (entity_id, entity_type),
    CHECK ((approved_by IS NULL) = (approved_at IS NULL)),
    CHECK (record_status NOT IN ('ACTIVE', 'REJECTED', 'SUPERSEDED') OR approved_by IS NOT NULL),
    CHECK (record_status <> 'SUPERSEDED' OR superseded_by IS NOT NULL),
    CHECK (valid_to IS NULL OR valid_from <= valid_to),
    -- C8: 同じ出品の ACTIVE な参照先は、期間が重ならない(注文日で一意に解決できる)
    CONSTRAINT listing_reference_no_overlap EXCLUDE USING gist (
        listing_id WITH =, daterange(valid_from, valid_to, '[]') WITH &&) WHERE (record_status = 'ACTIVE')
);

CREATE OR REPLACE FUNCTION product_core.listing_reference_rules()
RETURNS trigger LANGUAGE plpgsql AS $$
DECLARE
    v_status   text := product_core.entity_status(NEW.entity_id, NEW.entity_type);
    v_new_ref  boolean := TG_OP = 'INSERT';
    v_activate boolean := NEW.record_status = 'ACTIVE' AND (TG_OP = 'INSERT' OR OLD.record_status <> 'ACTIVE');
BEGIN
    -- S2: RETIRED への新しい参照は作れない(提案も不可)。ACTIVE 化の時点でも再確認する
    IF (v_new_ref OR v_activate) AND v_status = 'RETIRED' THEN
        RAISE EXCEPTION 'RETIRED の % へ出品参照は作れません', NEW.entity_id USING ERRCODE = 'check_violation';
    END IF;
    -- 正式な参照(ACTIVE)は ACTIVE の対象にだけ
    IF v_activate AND v_status <> 'ACTIVE' THEN
        RAISE EXCEPTION '% は % のため ACTIVE な出品参照を作れません', NEW.entity_id, v_status USING ERRCODE = 'check_violation';
    END IF;
    -- S4: 旧紐付け由来の参照は、同じ legacy P → 同じ対象の ACTIVE な legacy_mapping がある場合だけ
    IF (v_new_ref OR v_activate) AND NEW.derived_from_legacy_p IS NOT NULL
       AND NOT EXISTS (SELECT 1 FROM product_core.legacy_mapping m
                        WHERE m.legacy_p = NEW.derived_from_legacy_p AND m.entity_id = NEW.entity_id
                          AND m.record_status = 'ACTIVE') THEN
        RAISE EXCEPTION '% の ACTIVE な legacy_mapping が無いため、旧紐付けから参照を作れません', NEW.derived_from_legacy_p
            USING ERRCODE = 'check_violation';
    END IF;
    RETURN NEW;
END;
$$;
DROP TRIGGER IF EXISTS a_listing_reference_p1 ON product_core.listing_reference_history;
CREATE TRIGGER a_listing_reference_p1 BEFORE INSERT OR UPDATE OR DELETE ON product_core.listing_reference_history
    FOR EACH ROW EXECUTE FUNCTION product_core.p1_guard('LISTING_REFERENCE_APPROVE');
DROP TRIGGER IF EXISTS b_listing_reference_rules ON product_core.listing_reference_history;
CREATE TRIGGER b_listing_reference_rules BEFORE INSERT OR UPDATE ON product_core.listing_reference_history
    FOR EACH ROW EXECUTE FUNCTION product_core.listing_reference_rules();


-- =============================================================
--  cost_observation — 原価の観測値(P3 追記専用・元資料の値のまま)
-- =============================================================
--  税込・税抜が混在する前提(D12)。元資料に書かれた金額は変換しない。
CREATE TABLE IF NOT EXISTS product_core.cost_observation (
    observation_id      text           PRIMARY KEY DEFAULT product_core.fmt_id('CO-', nextval('product_core.co_seq'), 9)
                                       CHECK (observation_id ~ '^CO-[0-9]{9,}$'),
    source              text           NOT NULL CHECK (source IN (
                            'legacy_product_master', 'pricetar_inventory_csv', 'sku_string', 'rakuten_item_number',
                            'amazon_kpi', 'rakuten_kpi', 'cost_confirmation_record', 'purchase_record',
                            'delivery_note', 'invoice', 'maker_document', 'wholesaler_document', 'price_list',
                            'supplier_csv', 'other_document')),
    source_ref          text           NOT NULL CHECK (btrim(source_ref) <> ''),
    source_document_id  text           CHECK (source_document_id IS NULL OR source_document_id ~ '^DOC-[0-9]{8}-[0-9]{6}$'),
    source_reliability  text           NOT NULL CHECK (source_reliability IN ('HIGH', 'MEDIUM', 'LOW')),
    observed_at         timestamptz    NOT NULL,
    observed_amount     numeric(18, 6) NOT NULL CHECK (observed_amount >= 0),
    currency            text           NOT NULL DEFAULT 'JPY' CHECK (currency = 'JPY'),
    tax_inclusion       text           NOT NULL CHECK (tax_inclusion IN ('INCLUDED', 'EXCLUDED', 'UNKNOWN')),
    tax_inclusion_basis text           NOT NULL CHECK (tax_inclusion_basis IN ('DOCUMENT_STATED', 'SOURCE_SPEC',
                                           'OPERATIONAL_CONVENTION', 'HUMAN_CONFIRMED', 'NONE')),
    tax_rate            numeric(7, 6)  CHECK (tax_rate IS NULL OR (tax_rate >= 0 AND tax_rate < 1)),
    tax_rate_basis      text           NOT NULL CHECK (tax_rate_basis IN ('DOCUMENT_STATED', 'SOURCE_SPEC', 'NONE')),
    amount_unit         text           NOT NULL CHECK (amount_unit IN ('PER_LISTING_UNIT', 'PER_PP', 'PER_COMPOSITION',
                                                                      'UNKNOWN')),
    subject_entity_id   text,
    subject_entity_type text           CHECK (subject_entity_type IS NULL
                                              OR subject_entity_type IN ('PHYSICAL_PRODUCT', 'COMPOSITION', 'LISTING')),
    subject_legacy_p    text           CHECK (subject_legacy_p IS NULL OR subject_legacy_p ~ '^P[0-9]{6}$'),
    evidence            jsonb          NOT NULL,
    created_at          timestamptz    NOT NULL DEFAULT now(),
    created_by          text           NOT NULL CHECK (product_core.is_actor(created_by)),
    ingest_run_id       text           NOT NULL REFERENCES legacy_ingest.ingest_run (ingest_run_id),
    FOREIGN KEY (subject_entity_id, subject_entity_type) REFERENCES product_core.core_entity (entity_id, entity_type),
    CHECK ((subject_entity_id IS NULL) = (subject_entity_type IS NULL)),
    CHECK ((subject_entity_id IS NULL) <> (subject_legacy_p IS NULL)),  -- どちらか一方だけ
    CHECK ((tax_inclusion = 'UNKNOWN') = (tax_inclusion_basis = 'NONE')),
    CHECK ((tax_rate IS NULL) = (tax_rate_basis = 'NONE')),
    -- 書類から読んだ値は AIKOS の DocID(原本)に必ず結び付ける
    CHECK (source NOT IN ('delivery_note', 'invoice', 'maker_document', 'wholesaler_document', 'price_list',
                          'other_document') OR source_document_id IS NOT NULL)
);
CREATE UNIQUE INDEX IF NOT EXISTS cost_observation_source_uq ON product_core.cost_observation
    (source, source_ref, (coalesce(subject_entity_id, subject_legacy_p)));
DROP TRIGGER IF EXISTS cost_observation_append_only ON product_core.cost_observation;
CREATE TRIGGER cost_observation_append_only BEFORE UPDATE OR DELETE ON product_core.cost_observation
    FOR EACH ROW EXECUTE FUNCTION product_core.deny_mutation();


-- =============================================================
--  税抜への標準化と原価差異の判定(D8・D12)
-- =============================================================
--  観測値の行は書き換えず、計算で標準化する。途中で丸めない(D13)。
--  certainty:
--    CONFIRMED  … 税区分・税率とも十分な証拠(書類の明記・仕様・人の確認)
--    REFERENCE  … 運用慣行・PP の税区分など推定を含む
--    TAX_BASIS_UNKNOWN / TAX_RATE_UNKNOWN … 標準化できない
CREATE OR REPLACE FUNCTION product_core.normalize_observation(p_observation_id text, p_pp_id text DEFAULT NULL)
RETURNS TABLE (excl_tax_amount numeric, certainty text, rate_used numeric, amount_unit text, detail text)
LANGUAGE plpgsql STABLE AS $$
DECLARE
    o          product_core.cost_observation%ROWTYPE;
    v_cat      text;
    v_rate     numeric;
    v_strong_i boolean;
    v_strong_r boolean;
BEGIN
    SELECT * INTO o FROM product_core.cost_observation WHERE observation_id = p_observation_id;
    IF NOT FOUND THEN
        RAISE EXCEPTION '観測値 % がありません', p_observation_id;
    END IF;
    amount_unit := o.amount_unit;
    IF o.tax_inclusion = 'UNKNOWN' THEN
        excl_tax_amount := NULL; certainty := 'TAX_BASIS_UNKNOWN'; rate_used := NULL;
        detail := '税区分が不明'; RETURN NEXT; RETURN;
    END IF;
    v_strong_i := o.tax_inclusion_basis IN ('DOCUMENT_STATED', 'SOURCE_SPEC', 'HUMAN_CONFIRMED');
    IF o.tax_inclusion = 'EXCLUDED' THEN
        excl_tax_amount := o.observed_amount;
        certainty := CASE WHEN v_strong_i THEN 'CONFIRMED' ELSE 'REFERENCE' END;
        rate_used := o.tax_rate; detail := '税抜の値をそのまま使用';
        RETURN NEXT; RETURN;
    END IF;
    -- INCLUDED: 税率は観測値の記載を優先。無ければ PP の税区分(推定扱い)
    IF o.tax_rate IS NOT NULL THEN
        v_rate := o.tax_rate; v_strong_r := true;
    ELSIF p_pp_id IS NOT NULL THEN
        SELECT tax_category INTO v_cat FROM product_core.physical_product WHERE pp_id = p_pp_id;
        v_rate := CASE v_cat WHEN 'STANDARD' THEN 0.10 WHEN 'REDUCED' THEN 0.08 END;
        v_strong_r := false;
    END IF;
    IF v_rate IS NULL THEN
        excl_tax_amount := NULL; certainty := 'TAX_RATE_UNKNOWN'; rate_used := NULL;
        detail := '税込だが税率が決まらない'; RETURN NEXT; RETURN;
    END IF;
    excl_tax_amount := o.observed_amount / (1 + v_rate);
    certainty := CASE WHEN v_strong_i AND v_strong_r THEN 'CONFIRMED' ELSE 'REFERENCE' END;
    rate_used := v_rate;
    detail := CASE WHEN v_strong_r THEN '記載税率で税抜換算' ELSE 'PP の税区分の税率で税抜換算(推定)' END;
    RETURN NEXT;
END;
$$;

-- 差額 ≥ 50円 かつ 差率 ≥ 5% を WARN とする。双方が CONFIRMED のときだけ CONFIRMED_WARN。
-- 標準化できない観測値を含む比較では正式な警告を出さない。
CREATE OR REPLACE FUNCTION product_core.cost_variance(p_observation_id text, p_baseline_observation_id text,
                                                      p_pp_id text DEFAULT NULL)
RETURNS TABLE (verdict text, diff numeric, diff_ratio numeric)
LANGUAGE plpgsql STABLE AS $$
DECLARE
    a record;
    b record;
BEGIN
    SELECT * INTO a FROM product_core.normalize_observation(p_observation_id, p_pp_id);
    SELECT * INTO b FROM product_core.normalize_observation(p_baseline_observation_id, p_pp_id);
    diff := NULL; diff_ratio := NULL;
    IF a.certainty IN ('TAX_BASIS_UNKNOWN', 'TAX_RATE_UNKNOWN') OR b.certainty IN ('TAX_BASIS_UNKNOWN', 'TAX_RATE_UNKNOWN') THEN
        verdict := CASE WHEN 'TAX_BASIS_UNKNOWN' IN (a.certainty, b.certainty) THEN 'TAX_BASIS_UNKNOWN'
                        ELSE 'TAX_RATE_UNKNOWN' END;
        RETURN NEXT; RETURN;
    END IF;
    IF a.amount_unit = 'UNKNOWN' OR b.amount_unit = 'UNKNOWN' OR a.amount_unit <> b.amount_unit THEN
        verdict := 'UNIT_UNKNOWN'; RETURN NEXT; RETURN;
    END IF;
    IF b.excl_tax_amount = 0 THEN
        verdict := 'NO_BASELINE'; RETURN NEXT; RETURN;
    END IF;
    diff := a.excl_tax_amount - b.excl_tax_amount;
    diff_ratio := abs(diff) / b.excl_tax_amount;
    IF abs(diff) >= 50 AND diff_ratio >= 0.05 THEN
        verdict := CASE WHEN a.certainty = 'CONFIRMED' AND b.certainty = 'CONFIRMED' THEN 'CONFIRMED_WARN'
                        ELSE 'REFERENCE_WARN' END;
    ELSE
        verdict := 'OK';
    END IF;
    RETURN NEXT;
END;
$$;


-- =============================================================
--  cost_history / composition_cost_history — 正式原価(P2・税抜)
-- =============================================================
--  normalization の必須キー:
--    observed_amount / tax_inclusion / tax_inclusion_basis / tax_rate … 観測値の写し(S8: 一致を検証)
--    tax_inclusion_applied(INCLUDED/EXCLUDED)/ tax_rate_applied /
--    tax_rate_basis_applied(OBSERVED / PP_TAX_CATEGORY / HUMAN_CONFIRMED / NOT_APPLICABLE)/
--    unit_divisor(単位換算。1 = 換算なし)/ formula(人が読む式)
CREATE OR REPLACE FUNCTION product_core.normalization_keys()
RETURNS text[] LANGUAGE sql IMMUTABLE AS $$
    SELECT ARRAY['observed_amount', 'tax_inclusion', 'tax_inclusion_basis', 'tax_rate', 'tax_inclusion_applied',
                 'tax_rate_applied', 'tax_rate_basis_applied', 'unit_divisor', 'formula']
$$;

CREATE TABLE IF NOT EXISTS product_core.cost_history (
    cost_id                 text           PRIMARY KEY DEFAULT product_core.fmt_id('CH-', nextval('product_core.ch_seq'), 8)
                                           CHECK (cost_id ~ '^CH-[0-9]{8,}$'),
    pp_id                   text           NOT NULL REFERENCES product_core.physical_product (pp_id),
    unit_cost_excl_tax      numeric(18, 6) NOT NULL CHECK (unit_cost_excl_tax >= 0),
    currency                text           NOT NULL DEFAULT 'JPY' CHECK (currency = 'JPY'),
    valid_from              date           NOT NULL,
    valid_to                date,
    source_observation_id   text           NOT NULL REFERENCES product_core.cost_observation (observation_id),
    normalization           jsonb          NOT NULL CHECK (normalization ?& product_core.normalization_keys()),
    basis                   text           NOT NULL CHECK (btrim(basis) <> ''),
    assessment_id           text           NOT NULL REFERENCES product_core.assessment (assessment_id),
    tax_basis_assessment_id text           REFERENCES product_core.assessment (assessment_id),
    approved_by             text           NOT NULL CHECK (product_core.is_human(approved_by)),
    approved_at             timestamptz    NOT NULL DEFAULT now(),
    record_status           text           NOT NULL DEFAULT 'ACTIVE' CHECK (record_status IN ('ACTIVE', 'SUPERSEDED', 'VOID')),
    supersedes              text           REFERENCES product_core.cost_history (cost_id),
    created_at              timestamptz    NOT NULL DEFAULT now(),
    created_by              text           NOT NULL CHECK (product_core.is_human(created_by)),
    updated_at              timestamptz    NOT NULL DEFAULT now(),
    updated_by              text           NOT NULL CHECK (product_core.is_human(updated_by)),
    CHECK (valid_to IS NULL OR valid_from <= valid_to),
    CONSTRAINT cost_history_no_overlap EXCLUDE USING gist (
        pp_id WITH =, daterange(valid_from, valid_to, '[]') WITH &&) WHERE (record_status = 'ACTIVE')
);
COMMENT ON TABLE product_core.cost_history IS
    'PP の正式原価(税抜・1個あたり)。人の承認でのみ作成。元の観測値は source_observation_id と normalization に残る。';

CREATE TABLE IF NOT EXISTS product_core.composition_cost_history (
    cost_id                 text           PRIMARY KEY DEFAULT product_core.fmt_id('CX-', nextval('product_core.cx_seq'), 8)
                                           CHECK (cost_id ~ '^CX-[0-9]{8,}$'),
    cp_id                   text           NOT NULL REFERENCES product_core.composition (cp_id),
    set_cost_excl_tax       numeric(18, 6) NOT NULL CHECK (set_cost_excl_tax >= 0),
    currency                text           NOT NULL DEFAULT 'JPY' CHECK (currency = 'JPY'),
    cost_origin             text           NOT NULL DEFAULT 'DIRECT_PURCHASE' CHECK (cost_origin = 'DIRECT_PURCHASE'),
    purchase_evidence       jsonb          NOT NULL,
    valid_from              date           NOT NULL,
    valid_to                date,
    source_observation_id   text           NOT NULL REFERENCES product_core.cost_observation (observation_id),
    normalization           jsonb          NOT NULL CHECK (normalization ?& product_core.normalization_keys()),
    basis                   text           NOT NULL CHECK (btrim(basis) <> ''),
    assessment_id           text           NOT NULL REFERENCES product_core.assessment (assessment_id),
    tax_basis_assessment_id text           REFERENCES product_core.assessment (assessment_id),
    approved_by             text           NOT NULL CHECK (product_core.is_human(approved_by)),
    approved_at             timestamptz    NOT NULL DEFAULT now(),
    record_status           text           NOT NULL DEFAULT 'ACTIVE' CHECK (record_status IN ('ACTIVE', 'SUPERSEDED', 'VOID')),
    supersedes              text           REFERENCES product_core.composition_cost_history (cost_id),
    created_at              timestamptz    NOT NULL DEFAULT now(),
    created_by              text           NOT NULL CHECK (product_core.is_human(created_by)),
    updated_at              timestamptz    NOT NULL DEFAULT now(),
    updated_by              text           NOT NULL CHECK (product_core.is_human(updated_by)),
    CHECK (valid_to IS NULL OR valid_from <= valid_to),
    CONSTRAINT composition_cost_history_no_overlap EXCLUDE USING gist (
        cp_id WITH =, daterange(valid_from, valid_to, '[]') WITH &&) WHERE (record_status = 'ACTIVE')
);

-- 正式原価の作成規則(C10・C11・S3・S7・S8)
CREATE OR REPLACE FUNCTION product_core.cost_history_guard()
RETURNS trigger LANGUAGE plpgsql AS $$
DECLARE
    v_is_cp    boolean := TG_TABLE_NAME = 'composition_cost_history';
    v_entity   text;
    v_status   text;
    v_cat      text;
    v_amount   numeric;
    o          product_core.cost_observation%ROWTYPE;
    n          jsonb;
    v_applied  text;
    v_rate     numeric;
    v_rbasis   text;
    v_divisor  numeric;
    v_expected numeric;
    v_weak     boolean;
    v_confirmed text;
    t          product_core.assessment%ROWTYPE;
BEGIN
    IF TG_OP = 'DELETE' THEN
        RAISE EXCEPTION '% は削除できません(VOID で取り消してください)', TG_TABLE_NAME USING ERRCODE = 'restrict_violation';
    END IF;
    IF TG_OP = 'UPDATE' THEN
        -- 許すのは valid_to の一度だけの設定と、ACTIVE → SUPERSEDED / VOID だけ
        IF (to_jsonb(NEW) - ARRAY['valid_to', 'record_status', 'updated_at', 'updated_by'])
           IS DISTINCT FROM (to_jsonb(OLD) - ARRAY['valid_to', 'record_status', 'updated_at', 'updated_by']) THEN
            RAISE EXCEPTION '% の内容は変更できません(新しい行で訂正してください)', TG_TABLE_NAME USING ERRCODE = 'restrict_violation';
        END IF;
        IF OLD.valid_to IS NOT NULL AND NEW.valid_to IS DISTINCT FROM OLD.valid_to THEN
            RAISE EXCEPTION 'valid_to は一度だけ設定できます' USING ERRCODE = 'restrict_violation';
        END IF;
        IF NEW.record_status IS DISTINCT FROM OLD.record_status
           AND NOT (OLD.record_status = 'ACTIVE' AND NEW.record_status IN ('SUPERSEDED', 'VOID')) THEN
            RAISE EXCEPTION '状態 % → % には変更できません', OLD.record_status, NEW.record_status USING ERRCODE = 'check_violation';
        END IF;
        PERFORM product_core.assert_approval(NEW.updated_by, 'COST_APPROVE');
        NEW.updated_at := now();
        RETURN NEW;
    END IF;

    -- ---- INSERT ----
    PERFORM product_core.assert_approval(NEW.approved_by, 'COST_APPROVE');
    IF NEW.record_status <> 'ACTIVE' THEN
        RAISE EXCEPTION '正式原価は ACTIVE で作成してください' USING ERRCODE = 'check_violation';
    END IF;
    NEW.approved_at := now();
    NEW.created_by := NEW.approved_by;
    NEW.updated_by := NEW.approved_by;
    NEW.updated_at := now();

    IF v_is_cp THEN
        v_entity := NEW.cp_id;
        SELECT status INTO v_status FROM product_core.composition WHERE cp_id = NEW.cp_id;
        v_amount := NEW.set_cost_excl_tax;
    ELSE
        v_entity := NEW.pp_id;
        SELECT status, tax_category INTO v_status, v_cat FROM product_core.physical_product WHERE pp_id = NEW.pp_id;
        v_amount := NEW.unit_cost_excl_tax;
    END IF;
    -- S3: 未確定(PROVISIONAL)・退役(RETIRED)の対象には正式原価を作らない
    IF v_status IS DISTINCT FROM 'ACTIVE' THEN
        RAISE EXCEPTION '% は % のため正式原価を登録できません', v_entity, coalesce(v_status, '存在しない')
            USING ERRCODE = 'check_violation';
    END IF;
    -- C11: 経由した AI 判定に人の承認があること
    PERFORM product_core.assert_reviewed_assessment(NEW.assessment_id);

    SELECT * INTO o FROM product_core.cost_observation WHERE observation_id = NEW.source_observation_id;
    -- 観測値がこの対象のものであること(直接 / 承認済みの legacy_mapping 経由 / 承認済みの出品参照経由)
    IF NOT (o.subject_entity_id = v_entity
            OR (o.subject_legacy_p IS NOT NULL AND EXISTS (
                    SELECT 1 FROM product_core.legacy_mapping m
                     WHERE m.legacy_p = o.subject_legacy_p AND m.entity_id = v_entity AND m.record_status = 'ACTIVE'))
            OR (o.subject_entity_type = 'LISTING' AND EXISTS (
                    SELECT 1 FROM product_core.listing_reference_history r
                     WHERE r.listing_id = o.subject_entity_id AND r.entity_id = v_entity AND r.record_status = 'ACTIVE'))) THEN
        RAISE EXCEPTION '観測値 % は % のものではありません(承認済みの対応付けもありません)', o.observation_id, v_entity
            USING ERRCODE = 'check_violation';
    END IF;
    n := NEW.normalization;
    -- S8: 換算記録が元の観測値と一致すること(元の金額・税区分・税率を失わない)
    IF (n ->> 'observed_amount')::numeric IS DISTINCT FROM o.observed_amount
       OR (n ->> 'tax_inclusion') IS DISTINCT FROM o.tax_inclusion
       OR (n ->> 'tax_inclusion_basis') IS DISTINCT FROM o.tax_inclusion_basis
       OR (n ->> 'tax_rate')::numeric IS DISTINCT FROM o.tax_rate THEN
        RAISE EXCEPTION 'normalization の観測値の写しが元の観測値 % と一致しません', o.observation_id
            USING ERRCODE = 'check_violation';
    END IF;

    v_applied := n ->> 'tax_inclusion_applied';
    v_rate    := (n ->> 'tax_rate_applied')::numeric;
    v_rbasis  := n ->> 'tax_rate_basis_applied';
    v_divisor := (n ->> 'unit_divisor')::numeric;
    IF v_applied NOT IN ('INCLUDED', 'EXCLUDED') OR v_applied IS NULL THEN
        RAISE EXCEPTION '適用する税区分は INCLUDED / EXCLUDED のどちらかです' USING ERRCODE = 'check_violation';
    END IF;
    IF v_divisor IS NULL OR v_divisor <= 0 THEN
        RAISE EXCEPTION 'unit_divisor は正の数です' USING ERRCODE = 'check_violation';
    END IF;

    -- S7: 税区分が不明、根拠が運用慣行だけ、観測と違う税区分を適用、税率を人が決めた
    --     のいずれかなら、人が review で確定した TAX_BASIS 判定を必須にする
    v_weak := o.tax_inclusion = 'UNKNOWN' OR o.tax_inclusion_basis = 'OPERATIONAL_CONVENTION'
              OR v_applied <> o.tax_inclusion OR v_rbasis = 'HUMAN_CONFIRMED';
    IF v_weak THEN
        IF NEW.tax_basis_assessment_id IS NULL THEN
            RAISE EXCEPTION '観測値 % の税区分は確定していません(人の TAX_BASIS 確定が必要)', o.observation_id
                USING ERRCODE = 'check_violation';
        END IF;
        SELECT * INTO t FROM product_core.assessment WHERE assessment_id = NEW.tax_basis_assessment_id;
        IF t.assessment_type <> 'TAX_BASIS' OR t.subject_key IS DISTINCT FROM o.observation_id
           OR t.review_status NOT IN ('APPROVED', 'CORRECTED') THEN
            RAISE EXCEPTION 'TAX_BASIS 判定 % は、この観測値について人が確定したものではありません', NEW.tax_basis_assessment_id
                USING ERRCODE = 'check_violation';
        END IF;
        -- 人が確定した税区分(承認なら AI の判定値、修正なら修正値)と、適用した税区分が一致すること
        v_confirmed := CASE t.review_status WHEN 'APPROVED' THEN t.verdict ELSE t.corrected_value END;
        IF v_confirmed IS DISTINCT FROM v_applied THEN
            RAISE EXCEPTION '人が確定した税区分(%)と適用した税区分(%)が一致しません', coalesce(v_confirmed, 'NULL'), v_applied
                USING ERRCODE = 'check_violation';
        END IF;
    END IF;

    -- 適用税率の根拠
    IF v_applied = 'INCLUDED' THEN
        IF v_rbasis = 'OBSERVED' THEN
            IF o.tax_rate IS NULL OR v_rate IS DISTINCT FROM o.tax_rate THEN
                RAISE EXCEPTION '観測値に記載された税率と適用税率が一致しません' USING ERRCODE = 'check_violation';
            END IF;
        ELSIF v_rbasis = 'PP_TAX_CATEGORY' THEN
            IF v_is_cp OR v_cat NOT IN ('STANDARD', 'REDUCED')
               OR v_rate IS DISTINCT FROM (CASE v_cat WHEN 'STANDARD' THEN 0.10 ELSE 0.08 END) THEN
                RAISE EXCEPTION 'PP の税区分から税率を決められません' USING ERRCODE = 'check_violation';
            END IF;
        ELSIF v_rbasis = 'HUMAN_CONFIRMED' THEN
            IF v_rate IS NULL THEN
                RAISE EXCEPTION '税率がありません' USING ERRCODE = 'check_violation';
            END IF;
        ELSE
            RAISE EXCEPTION '税込の換算には税率の根拠が必要です(%)', coalesce(v_rbasis, 'NULL') USING ERRCODE = 'check_violation';
        END IF;
        IF v_rate < 0 OR v_rate >= 1 THEN
            RAISE EXCEPTION '税率が範囲外です' USING ERRCODE = 'check_violation';
        END IF;
        v_expected := o.observed_amount / (1 + v_rate) / v_divisor;
    ELSE
        v_expected := o.observed_amount / v_divisor;
    END IF;

    -- 保存精度(小数6桁)で換算結果と一致すること。途中では丸めない
    IF round(v_expected, 6) <> v_amount THEN
        RAISE EXCEPTION '正式原価 % が換算結果 % と一致しません', v_amount, round(v_expected, 6) USING ERRCODE = 'check_violation';
    END IF;
    RETURN NEW;
END;
$$;
DROP TRIGGER IF EXISTS cost_history_guard ON product_core.cost_history;
CREATE TRIGGER cost_history_guard BEFORE INSERT OR UPDATE OR DELETE ON product_core.cost_history
    FOR EACH ROW EXECUTE FUNCTION product_core.cost_history_guard();
DROP TRIGGER IF EXISTS composition_cost_history_guard ON product_core.composition_cost_history;
CREATE TRIGGER composition_cost_history_guard BEFORE INSERT OR UPDATE OR DELETE ON product_core.composition_cost_history
    FOR EACH ROW EXECUTE FUNCTION product_core.cost_history_guard();


-- =============================================================
--  スキャン解決(S6)— 曖昧なコードで自動確定しない
-- =============================================================
--  Phase 1 は「コード → 対象」までを解決する。Phase 2 で在庫形態への展開を足す。
--  resolution:
--    UNIQUE_AUTO      … 候補がちょうど1つ・AUTO_IF_UNIQUE・対象が ACTIVE → その1件を確定値として返す
--    NEEDS_SELECTION  … それ以外の候補あり。候補は返すが confirmed_entity_id は NULL(人が選ぶ)
--    NO_ACTIVE_LINK   … コードはあるが使える紐付けが無い
--    NOT_FOUND        … 未登録のコード
CREATE OR REPLACE FUNCTION product_core.resolve_identifier(p_id_type text, p_value text, p_on date DEFAULT current_date)
RETURNS TABLE (resolution text, candidate_entity_id text, candidate_entity_type text,
               confirmed_entity_id text, candidate_count integer)
LANGUAGE plpgsql STABLE AS $$
DECLARE
    v_identifier text;
    v_count      integer;
    r            record;
BEGIN
    SELECT identifier_id INTO v_identifier FROM product_core.identifier WHERE id_type = p_id_type AND value = p_value;
    IF v_identifier IS NULL THEN
        resolution := 'NOT_FOUND'; candidate_count := 0; RETURN NEXT; RETURN;
    END IF;
    SELECT count(*) INTO v_count FROM product_core.identifier_link l
     WHERE l.identifier_id = v_identifier AND l.record_status = 'ACTIVE' AND l.scan_policy <> 'NOT_FOR_SCAN'
       AND (l.valid_from IS NULL OR l.valid_from <= p_on) AND (l.valid_to IS NULL OR p_on <= l.valid_to);
    IF v_count = 0 THEN
        resolution := 'NO_ACTIVE_LINK'; candidate_count := 0; RETURN NEXT; RETURN;
    END IF;
    FOR r IN SELECT l.entity_id, l.entity_type, l.scan_policy,
                    product_core.entity_status(l.entity_id, l.entity_type) AS status
               FROM product_core.identifier_link l
              WHERE l.identifier_id = v_identifier AND l.record_status = 'ACTIVE' AND l.scan_policy <> 'NOT_FOR_SCAN'
                AND (l.valid_from IS NULL OR l.valid_from <= p_on) AND (l.valid_to IS NULL OR p_on <= l.valid_to)
              ORDER BY l.entity_id LOOP
        candidate_entity_id := r.entity_id;
        candidate_entity_type := r.entity_type;
        candidate_count := v_count;
        IF v_count = 1 AND r.scan_policy = 'AUTO_IF_UNIQUE' AND r.status = 'ACTIVE' THEN
            resolution := 'UNIQUE_AUTO'; confirmed_entity_id := r.entity_id;
        ELSE
            resolution := 'NEEDS_SELECTION'; confirmed_entity_id := NULL;
        END IF;
        RETURN NEXT;
    END LOOP;
END;
$$;


-- =============================================================
--  Composition の適用原価(A → B)
-- =============================================================
--  A: 対象日に有効な ACTIVE の composition_cost_history(直接仕入の承認済み原価)
--  B: Σ(構成品 PP の対象日に有効な ACTIVE の cost_history × 数量)。1つでも欠ければ算出不可(0 で埋めない)
--  A を使うときも B を計算し、差額 ≥ 50円 かつ 差率 ≥ 5% なら WARN を返す(警告候補。自動更新はしない)
CREATE OR REPLACE FUNCTION product_core.effective_composition_cost(p_cp_id text, p_on date)
RETURNS TABLE (method text, cost numeric, component_cost numeric, variance_flag text)
LANGUAGE plpgsql STABLE AS $$
DECLARE
    v_a       numeric;
    v_b       numeric;
    v_missing integer;
BEGIN
    SELECT h.set_cost_excl_tax INTO v_a FROM product_core.composition_cost_history h
     WHERE h.cp_id = p_cp_id AND h.record_status = 'ACTIVE'
       AND h.valid_from <= p_on AND (h.valid_to IS NULL OR p_on <= h.valid_to);
    SELECT sum(c.quantity * h.unit_cost_excl_tax), count(*) FILTER (WHERE h.cost_id IS NULL)
      INTO v_b, v_missing
      FROM product_core.composition_component c
      LEFT JOIN product_core.cost_history h
        ON h.pp_id = c.pp_id AND h.record_status = 'ACTIVE'
       AND h.valid_from <= p_on AND (h.valid_to IS NULL OR p_on <= h.valid_to)
     WHERE c.cp_id = p_cp_id;
    IF v_missing > 0 OR v_b IS NULL THEN
        v_b := NULL;
    END IF;
    component_cost := v_b;
    IF v_a IS NOT NULL THEN
        method := 'A'; cost := v_a;
        variance_flag := CASE WHEN v_b IS NULL OR v_b = 0 THEN 'NO_BASELINE'
                              WHEN abs(v_a - v_b) >= 50 AND abs(v_a - v_b) / v_b >= 0.05 THEN 'WARN'
                              ELSE 'OK' END;
    ELSIF v_b IS NOT NULL THEN
        method := 'B'; cost := v_b; variance_flag := NULL;
    ELSE
        method := 'UNAVAILABLE'; cost := NULL; variance_flag := NULL;
    END IF;
    RETURN NEXT;
END;
$$;


-- =============================================================
--  正式処理用のビュー(S4)— 承認済みの行だけを公開する
-- =============================================================
CREATE OR REPLACE VIEW product_core.v_active_legacy_mapping AS
    SELECT legacy_p, entity_id, entity_type, mapping_role, valid_from, valid_to, approved_by, approved_at, mapping_id
      FROM product_core.legacy_mapping
     WHERE record_status = 'ACTIVE';

CREATE OR REPLACE VIEW product_core.v_active_listing_reference AS
    SELECT listing_id, entity_id, entity_type, valid_from, valid_to, approved_by, approved_at, ref_id
      FROM product_core.listing_reference_history
     WHERE record_status = 'ACTIVE';

CREATE OR REPLACE VIEW product_core.v_active_identifier_link AS
    SELECT l.link_id, i.id_type, i.value, l.entity_id, l.entity_type, l.scan_policy, l.valid_from, l.valid_to
      FROM product_core.identifier_link l
      JOIN product_core.identifier i ON i.identifier_id = l.identifier_id
     WHERE l.record_status = 'ACTIVE';


-- =============================================================
--  索引 — 想定する読み方に合わせる
-- =============================================================
CREATE INDEX IF NOT EXISTS legacy_mapping_legacy_p_idx ON product_core.legacy_mapping (legacy_p);
CREATE INDEX IF NOT EXISTS listing_reference_listing_idx ON product_core.listing_reference_history (listing_id);
CREATE INDEX IF NOT EXISTS identifier_link_identifier_idx ON product_core.identifier_link (identifier_id);
CREATE INDEX IF NOT EXISTS identifier_link_entity_idx ON product_core.identifier_link (entity_id);
CREATE INDEX IF NOT EXISTS cost_observation_subject_idx ON product_core.cost_observation (subject_entity_id);
CREATE INDEX IF NOT EXISTS cost_observation_legacy_idx ON product_core.cost_observation (subject_legacy_p);
CREATE INDEX IF NOT EXISTS assessment_subject_key_idx ON product_core.assessment (assessment_type, subject_key);
CREATE INDEX IF NOT EXISTS assessment_review_idx ON product_core.assessment (review_status);
CREATE INDEX IF NOT EXISTS listing_asin_idx ON product_core.listing (asin) WHERE asin IS NOT NULL;
