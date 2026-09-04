-- =============================================================
--  001_intelligence_layer.sql — Intelligence Layer(知識層)
-- =============================================================
--  Architecture.md「6. Intelligence Layer」の実装。
--
--  【この層が答える問い】
--    何を判断したか / なぜそう判断したか / 結果どうなったか
--
--  分析だけではAIは成長できない。判断と結果が対で残って初めて、
--  AIは「その判断が正しかったか」を学習できる。
--
--  【なぜ Excel ではなく PostgreSQL か】
--  この2表は Excel に対応物を持たない。移行データではなく
--  「DBで生まれるデータ」であり、最初からここが正本である。
--  そのため momiji-stack.yml の「Phase4完了までDBは正本ではない」
--  という注意書きの対象外となる(対象は移行途中の業務データ)。
--  BL-5「正本は増やしてはならない」に抵触しない。
--
--  【スキーマを分ける理由】
--  Phase4 で移行してくる業務データ(public)と、この層で生まれる
--  知識(intelligence)を構造として分離する。混ざると
--  「どちらが正本か」が再び曖昧になる。
-- =============================================================

CREATE SCHEMA IF NOT EXISTS intelligence;

COMMENT ON SCHEMA intelligence IS
    '知識層。判断・施策・結果の記録。Excelに対応物を持たず、ここが正本。';


-- -------------------------------------------------------------
--  decisions — 判断と、その結果として実行した施策
-- -------------------------------------------------------------
--  1行 = 1回の判断。Decision_Catalog.md が「判断の型」の一覧で、
--  この表はその型に沿って実際に下した個別の判断を貯める。
--
--  ⚠️ 追記専用。UPDATE / DELETE はトリガーで拒否する(後述)。
--     訂正は supersedes に旧IDを入れた新しい行で行う。
--     「記録なき変更を認めない」(BL-7)を構造で担保する。
-- -------------------------------------------------------------
CREATE TABLE IF NOT EXISTS intelligence.decisions (
    decision_id     bigserial   PRIMARY KEY,

    -- いつ判断したか。記録した日時ではなく、判断した日時を入れる。
    decided_at      timestamptz NOT NULL DEFAULT now(),

    -- 判断の「型」。Decision_Catalog.md の DEC-<領域>-<連番2桁>。
    -- カタログに無い型を記録したくなったら、先にカタログへ追加する。
    decision_type   text        NOT NULL,

    -- 判断の対象。商品・広告キャンペーン・仕入先・システムなど。
    subject_type    text,
    subject_id      text,       -- 内部管理ID / ASIN / 楽天SKU / キャンペーンID 等
    subject_label   text,       -- 人が読むための名称

    -- 実際に取った施策。「変更しない」も必ずここへ書く(BL-11)。
    action          text        NOT NULL,

    -- 施策の種類。
    --   changed   … 変更した(値下げした・発注した・入札を上げた)
    --   unchanged … 検討したうえで変更しなかった(BL-11)
    --   rejected  … 候補を却下した(仕入見送り等)
    -- unchanged / rejected を残さないと、AIは同じ提案を繰り返す。
    action_kind     text        NOT NULL,

    -- なぜその判断をしたか。この列が空の記録には学習価値がない。
    reason          text        NOT NULL,

    -- 検討したが採らなかった案。後から「他に何を考えたか」を辿るため。
    alternatives    text,

    -- 期待した結果。定性(expected)と定量(metric/value)の両方を持つ。
    -- 定量が無い判断もあるため metric / value は NULL を許す。
    expected        text,
    expected_metric text,       -- 利益額 / TACOS / ROAS / 回転率 …
    expected_value  numeric,

    -- いつ結果を確認するか。この日を過ぎて outcomes が無い判断は
    -- 「やりっぱなし」として pending_reviews に現れる。
    review_due      date,

    -- 判断者。AI Constitution 第1条により必ず人。
    -- CHECK制約で 'ai:' 始まりを拒否し、憲法を構造で守る。
    decided_by      text        NOT NULL,

    -- 提案者。AIが提案し人が決めた場合は 'ai:claude' 等を入れる。
    -- 「AI提案の採用率」(Success Metrics)はこの列から測る。
    proposed_by     text,

    -- 根拠にした Business Logic 番号。例: {BL-3,BL-4}
    business_logic  text[],

    -- 記録経路。mcp / manual / batch など。
    source          text        NOT NULL DEFAULT 'mcp',

    -- 訂正時に旧レコードを指す。追記専用を保ったまま訂正するための唯一の手段。
    supersedes      bigint      REFERENCES intelligence.decisions(decision_id),

    note            text,

    -- 行が物理的に作られた時刻(decided_at は人が指定しうるため別に持つ)
    recorded_at     timestamptz NOT NULL DEFAULT now(),

    CONSTRAINT decisions_action_kind_valid
        CHECK (action_kind IN ('changed', 'unchanged', 'rejected')),

    -- AI Constitution 第1条「決めるのは人」。AIを判断者にできなくする。
    CONSTRAINT decisions_decided_by_is_human
        CHECK (decided_by !~* '^ai[:_-]'),

    -- 空文字での骨抜きを防ぐ。理由の無い判断は記録として成立しない。
    CONSTRAINT decisions_required_text_not_blank
        CHECK (
            btrim(decision_type) <> ''
            AND btrim(action)    <> ''
            AND btrim(reason)    <> ''
            AND btrim(decided_by) <> ''
        ),

    -- 自分自身を訂正対象にはできない
    CONSTRAINT decisions_supersedes_not_self
        CHECK (supersedes IS DISTINCT FROM decision_id)
);

COMMENT ON TABLE  intelligence.decisions IS '判断と施策の記録(追記専用)。訂正は supersedes で行う。';
COMMENT ON COLUMN intelligence.decisions.action_kind IS 'changed / unchanged / rejected。「変更しない」も経営判断(BL-11)。';
COMMENT ON COLUMN intelligence.decisions.decided_by IS '判断者。AI Constitution 第1条により必ず人。';


-- -------------------------------------------------------------
--  outcomes — 結果の記録
-- -------------------------------------------------------------
--  判断 1 : 結果 0..n。1週間後・1ヶ月後と複数回測ることがあるため
--  1対1にしない。結果が0件のうちは「まだ学習素材になっていない」。
--
--  ⚠️ こちらも追記専用。測り直しは新しい行を足す。
-- -------------------------------------------------------------
CREATE TABLE IF NOT EXISTS intelligence.outcomes (
    outcome_id      bigserial   PRIMARY KEY,

    decision_id     bigint      NOT NULL
                                REFERENCES intelligence.decisions(decision_id),

    measured_at     timestamptz NOT NULL DEFAULT now(),

    -- 何を対象期間として測ったか
    period_start    date,
    period_end      date,

    -- 実績。判断時の expected_metric / expected_value と対で読む。
    metric          text,
    actual_value    numeric,

    -- 判断の評価。
    --   success  … 期待どおり
    --   partial  … 一部達成
    --   failure  … 期待に届かなかった
    --   unclear  … 他要因が大きく切り分けられない
    -- unclear を用意するのは、無理に成否を付けると学習を誤らせるため。
    assessment      text        NOT NULL,

    -- 何が起きたかの説明
    summary         text        NOT NULL,

    -- 次に活かす学び。AIが最も参照する列。
    -- 「なぜそうなったか」まで書く(結果の数字だけでは学習できない)。
    learning        text,

    measured_by     text        NOT NULL,
    source          text        NOT NULL DEFAULT 'mcp',
    note            text,
    recorded_at     timestamptz NOT NULL DEFAULT now(),

    CONSTRAINT outcomes_assessment_valid
        CHECK (assessment IN ('success', 'partial', 'failure', 'unclear')),

    CONSTRAINT outcomes_required_text_not_blank
        CHECK (btrim(summary) <> '' AND btrim(measured_by) <> ''),

    CONSTRAINT outcomes_period_order
        CHECK (period_start IS NULL OR period_end IS NULL OR period_start <= period_end)
);

COMMENT ON TABLE  intelligence.outcomes IS '判断の結果(追記専用)。1判断に複数回の測定を許す。';
COMMENT ON COLUMN intelligence.outcomes.learning IS '次に活かす学び。AI学習の主素材。';


-- -------------------------------------------------------------
--  追記専用の強制
-- -------------------------------------------------------------
--  アプリ側の作法ではなくDBの構造として守る。
--  DBユーザーは所有者のため GRANT では止められない。トリガーで拒否する。
--
--  ※ 正当な理由で解除する場合は、解除自体を Decision Log へ記録すること。
-- -------------------------------------------------------------
CREATE OR REPLACE FUNCTION intelligence.deny_mutation()
RETURNS trigger
LANGUAGE plpgsql
AS $$
BEGIN
    RAISE EXCEPTION
        '% は追記専用です。%は許可されていません。訂正は supersedes を指定した新しい行で行ってください。',
        TG_TABLE_NAME, TG_OP
        USING ERRCODE = 'restrict_violation';
END;
$$;

DROP TRIGGER IF EXISTS decisions_append_only ON intelligence.decisions;
CREATE TRIGGER decisions_append_only
    BEFORE UPDATE OR DELETE ON intelligence.decisions
    FOR EACH ROW EXECUTE FUNCTION intelligence.deny_mutation();

DROP TRIGGER IF EXISTS outcomes_append_only ON intelligence.outcomes;
CREATE TRIGGER outcomes_append_only
    BEFORE UPDATE OR DELETE ON intelligence.outcomes
    FOR EACH ROW EXECUTE FUNCTION intelligence.deny_mutation();


-- -------------------------------------------------------------
--  索引 — 想定する読み方に合わせる
-- -------------------------------------------------------------
--  ①「この商品について過去どう判断したか」 → subject_id
--  ②「価格改定の判断を時系列で見たい」     → decision_type + decided_at
--  ③「結果待ちのものは何か」               → review_due
CREATE INDEX IF NOT EXISTS decisions_subject_idx
    ON intelligence.decisions (subject_id) WHERE subject_id IS NOT NULL;
CREATE INDEX IF NOT EXISTS decisions_type_decided_at_idx
    ON intelligence.decisions (decision_type, decided_at DESC);
CREATE INDEX IF NOT EXISTS decisions_decided_at_idx
    ON intelligence.decisions (decided_at DESC);
CREATE INDEX IF NOT EXISTS decisions_review_due_idx
    ON intelligence.decisions (review_due) WHERE review_due IS NOT NULL;
CREATE INDEX IF NOT EXISTS outcomes_decision_idx
    ON intelligence.outcomes (decision_id);


-- -------------------------------------------------------------
--  pending_reviews — 結果が未記録の判断
-- -------------------------------------------------------------
--  「やりっぱなし禁止」(Architecture.md §6 設計ルール)の実装。
--  この一覧が常に空に近い状態を保てているかが、学習ループが
--  回っているかどうかの唯一の客観指標になる。
--
--  訂正された(supersedesで上書きされた)判断は結果を待つ必要がないため除く。
-- -------------------------------------------------------------
CREATE OR REPLACE VIEW intelligence.pending_reviews AS
SELECT
    d.decision_id,
    d.decided_at,
    d.decision_type,
    d.subject_id,
    d.subject_label,
    d.action,
    d.action_kind,
    d.expected,
    d.expected_metric,
    d.expected_value,
    d.review_due,
    d.decided_by,
    CASE WHEN d.review_due IS NULL THEN NULL
         ELSE (CURRENT_DATE - d.review_due)
    END AS overdue_days
FROM intelligence.decisions d
WHERE NOT EXISTS (
        SELECT 1 FROM intelligence.outcomes o
         WHERE o.decision_id = d.decision_id
      )
  AND NOT EXISTS (
        SELECT 1 FROM intelligence.decisions s
         WHERE s.supersedes = d.decision_id
      );

COMMENT ON VIEW intelligence.pending_reviews IS
    '結果が未記録の判断。空に近いほど学習ループが回っている。';
