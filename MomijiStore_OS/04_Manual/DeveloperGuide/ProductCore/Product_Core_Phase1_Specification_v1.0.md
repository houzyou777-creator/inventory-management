# AIKOS Product Core — Phase 1 Specification v1.0

| 項目 | 内容 |
|---|---|
| 版 | **v1.0**(2026-09-26 承認・正式配置)。U1〜U4 の決定、追加安全制約 S1〜S8、原価警告の2段階、NUMERIC 精度を反映。DDL: `05_Infrastructure/db/migrations/002_product_core.sql` / `003_product_core_roles.sql` |
| 設計基準 | 商品データモデル v2.2(2026-09-25 承認・凍結) |
| 状態 | DDL・DB ロール・制約テストを作成済み(使い捨て PostgreSQL での検証用)。**本番 DB への適用と実データ投入は未承認・未実施** |
| 正本 | 既存 Excel(商品マスター・出品テーブル等)が引き続き正本。Product Core は **Shadow Model** |

---

## 0. 決定事項(2026-09-26)

| # | 事項 | 決定 |
|---|---|---|
| D1 | Shadow Model | **NAS 上の PostgreSQL(`momiji-stack`)**。スキーマ `product_core`。既存 Excel が正本。**Product Core → Excel の自動書き戻しは禁止**。Google Sheets は Product Core 本体に使わない(将来の人間確認 UI としての利用はありうる) |
| D2 | 原価の標準 | AIKOS 内部の正式原価は**税抜**。**観測値は税込・税抜が混在することを正式な前提とする**(U1・§3.13)。Cost Observation は元資料の金額を変換せずに保存し、Cost History は標準化した税抜原価と元の観測値への参照を持つ |
| D3 | Human Review | Phase 1 は **AIKOS → Excel 確認表 → 人が 承認 / 却下 / 保留(/ 修正)→ AIKOS へ取込**。将来 AIKOS Warehouse / Control Center の承認 UI へ移行できる構造にする(review 列は入力経路に依存させない) |
| D4 | 承認者 | `human:<operator_id>`。表示名は主キーにしない(operator 表で管理) |
| D5 | ID | 既存 P 番号とは別体系。Physical Product **`PP-xxxxxx`**、Composition **`CP-xxxxxx`**。既存 P は legacy_mapping 側でのみ保持し、Product Core の主キーに再利用しない |
| D6 | 現行標準原価 | 商品マスターの標準原価は **Cost History へ一括承認しない**。`source = legacy_product_master` の Cost Observation として取り込み、他の根拠と照合 → AI Assessment → Human Approval を経たものだけ Cost History へ |
| D7 | 取込頻度 | Phase 1 は**手動実行**。安定後に日次同期、将来 API / イベント連携 |
| D8 | 原価差異警告 | **差額 ≥ 50円 かつ 差率 ≥ 5%** を基本条件とし、**2段階**に分ける: 双方の税区分・税率が十分な証拠で確認できる → `CONFIRMED_WARN`、運用慣行など推定を含む → `REFERENCE_WARN`(同じ確度として扱わない)。税抜に標準化できない観測値には正式な警告を出さない。どちらの警告でも Cost History は自動更新しない |
| D9 | 在庫引当(Phase 2) | Kit Stock → バラ在庫の順。単品注文のための Kit 自動解体は禁止(提案のみ・解体は承認後の Stock Event)。FBA 在庫は自社倉庫注文の引当対象外(別 Location) |
| D10 | 福袋 | 既存9件は `HUMAN_REVIEW_REQUIRED / 構成未確認` で保持。必要時に固定 → Composition、可変 → Selection Rule |
| D11 | AI と人 | AI 判定(assessment_type / confidence / evidence / reason)と Human Review(UNREVIEWED / APPROVED / REJECTED / CORRECTED + 保留)を完全分離。「整合性高」でも Review 済みとは扱わない |
| D12(U1) | 原価の税区分 | 既存の商品マスター・Pricetar の手入力仕入単価は**運用上は原則税込**。今後の納品書・請求書・メーカー/問屋資料・価格表・CSV 等は**税抜の場合もある**。観測値には税区分とその**根拠の確からしさ**を持たせ、根拠が弱いデータを無条件に確定値へ変換しない。UNKNOWN は推測で正式原価に採用しない。原価差異 WARN は税抜に正しく標準化できた観測値同士でだけ判定する |
| D13(U2) | 丸め・精度 | 内部計算では途中で丸めない。**保存精度・表示精度・会計/KPI 確定時の丸め**を分離する。金額は **NUMERIC(18,6)**、税率は **NUMERIC(7,6)**(FLOAT は使わない)。境界値テストで確認済み(§2.2)。会計 / KPI 確定時の最終丸め規則は確定処理の設計時に決める |
| D14(U3) | 確認表の保存先 | **NAS 上の AIKOS 管理領域**に置き、Pending / Reviewed / Archive 相当の状態で管理する。**実際のパスは既存の AIKOS 文書管理設計を確認してから確定**し、推測で新しいパスを作らない(§7.1) |
| D15(U4) | 承認者と権限 | 最初の承認者を管理者として登録できる構造。operator の永続 ID と表示名を分離し、承認記録には `human:<operator_id>`。権限(商品マスター承認・Legacy Mapping 承認・Listing 参照先変更・正式原価承認 等)を operator ごとに分離できる(§3.0b / §3.0c) |
| D16 | Identifier | A 案(identifier / identifier_link / core_entity)を**正式採用**。曖昧な Identifier から Stock Form を自動確定しない |

---

## 1. Identifier 設計の最終決定

### 1.1 比較に使った実データ(商品マスター 1,211件・JAN 763種類)

| 事実 | 件数 | 意味 |
|---|---|---|
| 単品Pに1対1の JAN | 621 | 読めば物理商品は決まる(ただし、その商品を含む梱包済みキットが別にあれば**在庫形態は決まらない**) |
| 単品同士で共有される JAN(容量違い等 例: 50本/100本/200本入り) | 10 | 読んでも**物理商品すら決まらない** |
| 単品とセットで共有される JAN | 21 | 梱包済みセットの外から単品の JAN が見える。**バラか梱包済みかはスキャンでは分からない** |
| セットだけに付く JAN | 111(うち複数のセットPに付く 26) | 読めばセット(Composition)は決まることが多いが、26 は候補が複数 |
| ケースコード(GTIN-14 / 外箱) | 登録0(棚卸しは「ケースコード」として判別済み) | 外箱1つ = 単品 n 個という**在庫形態そのもの**を表すコード |
| JAN 欄の問題 | 383(ASIN 入り14・空欄369) | コードの種類の誤記・未取得 |

### 1.2 比較

| 観点 | A: 汎用化(identifier → entity) | B: 分離(product_identifier / stock_form_identifier) |
|---|---|---|
| 同じ JAN が単品とセットで共有(21) | コード1行に紐付けが複数並び、**曖昧さが1か所で見える** | JAN が product 側・stock_form 側の2表に分かれて入り、曖昧さは両表を合わせないと見えない。**表をまたぐ一意性・整合性は制約で守れない** |
| スキャンで在庫形態が決まらない | 「コード → 物(entity)」と「物 → 在庫形態」を**別の段階**として扱える | stock_form_identifier に JAN を入れると「この JAN = この在庫形態」と**読めてしまい、自動確定の誤りを誘う** |
| ケースコード / 外箱 | 対象の種類 STOCK_FORM へ紐付けるだけ | stock_form_identifier に入る(ここは B でも自然) |
| 未登録コード(ケースコード初見・仮ID) | 紐付け0件のコードとして**記録できる** | どちらの表にも入れる場所が無い(別の表が要る) |
| 将来の対象追加(FNSKU ラベル・自社キットラベル・出品) | 対象の種類を登録簿に足すだけ。**identifier 系の表は不変** | 対象ごとに `*_identifier` 表が増え、スキャン時の照合先も増える |
| 外部キーの整合性 | 単純な entity_type + entity_id では守れない → **対象の登録簿(core_entity)への実 FK で守る** | 表ごとに実 FK で守れる(B の長所) |
| 検索のしやすさ | 1回の検索で全候補 | 2表以上の UNION |
| 責務の安定性 | 「コードが指しうる対象の候補を持つ」で一貫 | product 側は安定。ただし stock_form 側に入れる JAN の判断が運用に委ねられる |

### 1.3 決定: **A(汎用化)を採用し、次の3点で実装する**

1. **identifier = コード登録簿。** バーコード・コードそのもの(種類 + 正規化値)を一意に登録する。対象を持たない。
2. **identifier_link = コード → 対象の候補(多対多)。** 対象は **core_entity(対象の登録簿)への実 FK** で持つ。
   Phase 2 で在庫形態を足すときは core_entity の種類に STOCK_FORM を加えるだけで、**identifier / identifier_link の定義は変わらない**。
3. **紐付けの原則: そのコード単体で証明できる最も具体的な対象に紐付ける。**

| コード | 紐付け先 | 理由 |
|---|---|---|
| 商品本体に印刷された JAN | PHYSICAL_PRODUCT | 何の物かは分かるが、バラか梱包済みかは分からない |
| セット外装にだけ付く JAN | COMPOSITION | そのセットであることは分かる |
| ケースコード(GTIN-14)・外箱コード | STOCK_FORM(CASE)※Phase 2 | 在庫形態そのものを証明する |
| 自社キットラベル・FNSKU | STOCK_FORM / LISTING ※Phase 2 | 同上 |
| JAN 欄に入っていた ASIN | LISTING 候補(`link_basis=MISFILED`、スキャン不可) | コードの種類の誤記 |
| 仮ID(`T-000123`) | PROVISIONAL の PHYSICAL_PRODUCT | 棚卸しの未登録品 |

**単品の JAN をキットの在庫形態へ直接紐付けない。** 梱包済みキットの候補は、Phase 2 の解決手順(§1.4 ③)で「その物を含むキット」として導く。これにより、バーコードだけでキットを自動確定する経路が構造的に存在しなくなる。

### 1.4 スキャン解決の手順(Phase 2 で実装。Phase 1 は前提として固定)

```
① scan 生値 → 正規化(棚卸しの jan_lookup.normalize_code と同じ規則)→ (id_type, value)
② identifier を検索。無ければ「未登録コード」として記録(仮ID の流れへ)
③ ACTIVE・有効期間内・scan_policy≠NOT_FOR_SCAN の identifier_link → 対象の候補(entity)
④ 対象 → 在庫形態の候補に展開
     PHYSICAL_PRODUCT → その LOOSE 形態 + 「外から単品のコードが見える」梱包済みキット形態
     COMPOSITION      → その PREPACKED_KIT 形態
     STOCK_FORM       → そのもの
⑤ 在庫形態の候補がちょうど1つ、かつ紐付けが AUTO_IF_UNIQUE、かつ対象が ACTIVE → 確定
   それ以外(複数・PROVISIONAL・CONFIRM_ALWAYS)→ スタッフが選ぶ / 後で確認(自動確定しない)
⑥ 確定した在庫形態で stock_event(選択した場合は選択者・候補一覧を evidence に残す)
```

| 実データ | ③の対象候補 | ④の在庫形態候補 | 結果 |
|---|---|---|---|
| 単品1対1の JAN(621)でキットなし | 1 | 1(LOOSE) | 自動確定 |
| 単品セット共通 JAN(21) | 1〜2 | 2以上(LOOSE + キット) | **スタッフ選択**(棚卸しの紫表示と同じ) |
| 単品同士の共有 JAN(10) | 2以上 | 2以上 | **スタッフ選択**(青表示) |
| セットのみ JAN(85) | 1 | 1(キット) | 自動確定 |
| セットのみ JAN、候補複数(26) | 2以上 | 2以上 | **スタッフ選択** |
| ケースコード | 1(CASE) | 1 | 自動確定(Phase 2 で CASE 登録後) |

---

## 2. 共通規約

### 2.1 スキーマ

| スキーマ | 内容 |
|---|---|
| `product_core` | 本書のテーブル(Shadow Model) |
| `legacy_ingest` | 正本 Excel・CSV の読み取り専用の写し(§3.17) |
| 既存 `public` / `intelligence` | 変更しない |

### 2.2 型

| 論理型 | PostgreSQL |
|---|---|
| ID | `text` + 形式 CHECK |
| 金額(元資料の値・標準化後の原価) | **`numeric(18,6)`**(整数12桁・小数6桁。上限 999,999,999,999.999999) |
| 金額の計算途中 | 桁指定なしの `numeric`(**途中で丸めない**。保存時に小数6桁) |
| 税率 | **`numeric(7,6)`**(0.100000 = 10%、0.080000 = 8%)。**FLOAT は使わない** |

**NUMERIC(18,6) の境界値テストの結果(2026-09-26)**

| 観点 | 結果 |
|---|---|
| 通常原価(税込 1,100・10% → 税抜 1,000、税込 1,080・8% → 1,000) | ちょうど一致 |
| 非常に小さい単価(0.12円 ÷ 100枚 = 0.0012、0.000001) | 正確に保存 |
| 小数7桁目以下(0.0000004 → 0、0.0000005 → 0.000001) | **保存時に四捨五入される**。資料の単価に小数7桁以下は現れない前提。現れる場合は保存前に人へ回す |
| 大きな原価(99,999,999.99、12,000,000 × 999 = 11,988,000,000) | 正確に保存 |
| 上限超え(1兆円) | **黙って丸めず拒否**(SQLSTATE 22003) |
| 税込→税抜の割り切れない換算(1,000 ÷ 1.1) | 保存値 909.090909。税込に戻して円表示すると 1,000 |
| 100個分の合計 | 誤差は保存精度の範囲内(0.00005 以下)、円表示は一致 |
| Composition 原価合計(0.0012 + 99,999,999.99) | 途中で丸めず 99,999,999.9912 |

**丸めの分離(D13):** ① 保存 = 上記の精度で丸めずに保持 ② 表示 = 画面・Excel 確認表で円単位に丸めて見せるだけ(保存値は変えない)
③ 会計 / KPI 確定 = 確定処理の時点で1回だけ、定めた規則(例: 税抜換算は円未満〇〇)で丸め、丸めた値と規則を確定側の記録に残す。
**具体的な精度・最終丸め規則は DDL 設計時に提示し、承認を得る。**
| 数量 | `integer` |
| 適用期間 | `date`(`valid_from` を含み `valid_to` を含む。`valid_to` NULL = 無期限) |
| 日時 | `timestamptz` |
| 列挙 | `text` + CHECK |
| 根拠 | `jsonb` |

### 2.3 ID 形式(D5)

| 対象 | 形式 |
|---|---|
| physical_product | **`PP-` + 6桁**(例 `PP-000001`) |
| composition | **`CP-` + 6桁** |
| listing | `LS-` + 6桁(既存の出品ID `C000001` は `legacy_listing_id`) |
| listing_group | `LG-` + 6桁 |
| identifier / identifier_link | `ID-` + 8桁 / `IL-` + 8桁 |
| legacy_mapping | `LM-` + 8桁 |
| product_relationship | `PR-` + 8桁 |
| composition_component | `CC-` + 8桁 |
| listing_reference_history | `LR-` + 8桁 |
| cost_history / composition_cost_history | `CH-` + 8桁 / `CX-` + 8桁 |
| cost_observation | `CO-` + 9桁 |
| assessment | `AS-` + 9桁 |

- 既存の `P000001` は**文字列の legacy P** としてだけ扱う(legacy_mapping・snapshot・evidence)。Product Core の主キー・FK には使わない。
- 桁が不足したら、同じ接頭辞のまま桁を増やす(既存 ID は変えない)。

### 2.4 行為者(actor)と承認者

| 形式 | 用途 |
|---|---|
| `human:<operator_id>` | 人(例 `human:op001`)。**承認・review 列はこれのみ** |
| `ai:<model>/<rule_version>` | AI 判定 |
| `rule:<rule_name>/<version>` | 決定的な規則による判定 |
| `ingest:<job>/<run_id>` | 取込ジョブ |

- 承認・review 列は CHECK で `^human:` のみ、さらにトリガーで `operator` 表の **ACTIVE** かつ `operator_permission` で**その操作の権限を持つ** operator_id であることを確認する。
- 表示名の変更は operator 表だけで行い、記録済みの行は変わらない。

### 2.5 共通列

| 列 | 型 | NULL | 説明 |
|---|---|---|---|
| `created_at` | timestamptz | NOT NULL | 既定 `now()` |
| `created_by` | text | NOT NULL | actor |
| `updated_at` / `updated_by` | timestamptz / text | 更新可の表のみ NOT NULL | 更新トリガーで自動設定 |
| `ingest_run_id` | text | NULL可 | 取込で作った行は NOT NULL(FK `legacy_ingest.ingest_run`) |

### 2.6 状態パターン

| パターン | 対象 | 規則 |
|---|---|---|
| **P1 提案→承認** | identifier_link / legacy_mapping / product_relationship / listing_reference_history | `record_status`: PROPOSED → ACTIVE / REJECTED、ACTIVE → SUPERSEDED。**内容列は作成後に変更不可**。ACTIVE は `approved_by`(human)必須。下流は ACTIVE のみ参照 |
| **P2 承認済み履歴** | cost_history / composition_cost_history | 人の承認時にだけ作る。`record_status`: ACTIVE / SUPERSEDED / VOID(取消。削除しない)。`valid_to` は NULL → 日付の一度だけ設定可 |
| **P3 追記専用** | cost_observation / identifier / assessment(AI 部分) | UPDATE / DELETE をトリガーで拒否(assessment の review 列のみ例外) |
| **P4 事実の写し** | listing / listing_group | チャネルの事実。自然キーは変更不可。状態・最終確認日時は更新可。判断を含まないので承認不要 |
| **P5 実体** | physical_product / composition | AI・取込は PROVISIONAL でのみ作成。ACTIVE・RETIRED は人の承認。削除不可 |

### 2.7 AI が書ける範囲 / 人間承認が必要な範囲

| 操作 | AI・取込 | 人(`human:`) |
|---|---|---|
| identifier(コード)の登録 | ✅ | — |
| listing / listing_group の写し | ✅ | — |
| cost_observation の追記 | ✅ | — |
| assessment(AI 判定)の追記 | ✅ | review 列のみ |
| PROVISIONAL の PP / CP 作成、構成品の下書き | ✅ | — |
| identifier_link / legacy_mapping / relationship / listing 参照の PROPOSED | ✅ | — |
| 上記の ACTIVE 化(**P 統合・参照先変更を含む**) | ❌ | ✅ 必須 |
| PP / CP の ACTIVE・RETIRED 化 | ❌ | ✅ 必須 |
| cost_history / composition_cost_history の作成(**正式原価**) | ❌ | ✅ 必須 |
| 正本 Excel の変更 | ❌ | Phase 1 では Product Core から行わない |

---

## 3. テーブル仕様

**Phase 1 の論理 13 テーブル**(identifier は「コード登録簿 + 紐付け」の2物理表で実現)**+ 構造用 3 表**(core_entity・operator・operator_permission)**+ 取込ステージング**。

凡例: **PK** 主キー / **UQ** 一意 / **FK** 外部キー / **AI** AI・取込が書ける / **人** 人間承認が必要

### 3.0 `core_entity` — 対象の登録簿(構造用)

- **目的:** identifier_link・legacy_mapping・listing_reference_history・cost_observation・assessment が「PP / CP / Listing(Phase 2 では Stock Form)」を**実 FK で**参照するための登録簿。
- **パターン:** 追記のみ(削除不可)。

| 列 | 型 | NULL | 制約・値 | 説明 |
|---|---|---|---|---|
| `entity_id` | text | NOT NULL | **PK** | `PP-…` / `CP-…` / `LS-…`(/ Phase 2 `SF-…`) |
| `entity_type` | text | NOT NULL | `PHYSICAL_PRODUCT` / `COMPOSITION` / `LISTING`(Phase 2 で `STOCK_FORM` を追加) | |
| | | | **UQ** `(entity_id, entity_type)` | 複合 FK の参照先 |
| `created_at` / `created_by` | | NOT NULL | | |

- 各実体表は `(pp_id, 'PHYSICAL_PRODUCT')` 等の**複合 FK** で core_entity を参照する(同一トランザクションで登録。制約は DEFERRABLE)。
- 参照する側は `(entity_id, entity_type)` の複合 FK + `entity_type IN (...)` の CHECK で、**参照できる種類を表ごとに制限**する。

### 3.0b `operator` — 人の登録簿(構造用)

| 列 | 型 | NULL | 制約・値 | 説明 |
|---|---|---|---|---|
| `operator_id` | text | NOT NULL | **PK** `^[a-z0-9_]{3,32}$` | actor は `human:<operator_id>` |
| `display_name` | text | NOT NULL | | 変更可(主キーではない) |
| `status` | text | NOT NULL | `ACTIVE` / `INACTIVE` | 削除しない |
| 共通列 | | | | 更新可 |

- operator の追加・権限変更は人が行う(AI は書けない)。
- **最初の承認者(管理者)の登録:** operator 表が空のときに限り、人が手動で実行するブートストラップ手順で1名を登録し、§3.0c の `OPERATOR_ADMIN` を含む全権限を付与する。2人目以降の登録・権限付与は `OPERATOR_ADMIN` を持つ operator だけが行える(トリガー)。

### 3.0c `operator_permission` — 承認権限(構造用)

| 列 | 型 | NULL | 制約・値 | 説明 |
|---|---|---|---|---|
| `operator_id` | text | NOT NULL | FK operator | |
| `permission` | text | NOT NULL | 下表 | |
| `granted_by` / `granted_at` | text / timestamptz | NOT NULL | `human:` かつ `OPERATOR_ADMIN` 保持者(ブートストラップ時を除く) | |
| `revoked_by` / `revoked_at` | | NULL可 | 取り消し(行は削除しない) | |
| | | | **UQ** `(operator_id, permission)` の有効行(revoked_at IS NULL)で一意 | |

| permission | 許可する操作 |
|---|---|
| `OPERATOR_ADMIN` | operator の追加・権限の付与と取り消し |
| `PRODUCT_APPROVE` | physical_product / composition の ACTIVE・RETIRED 化(商品マスター承認) |
| `LEGACY_MAPPING_APPROVE` | legacy_mapping の ACTIVE 化(P 統合を含む) |
| `RELATIONSHIP_APPROVE` | product_relationship の ACTIVE 化 |
| `IDENTIFIER_APPROVE` | identifier_link の ACTIVE 化・scan_policy の確定 |
| `LISTING_REFERENCE_APPROVE` | listing_reference_history の ACTIVE 化・期間終了 |
| `COST_APPROVE` | cost_history / composition_cost_history の作成(正式原価)・税区分の人による確定 |
| `REVIEW` | assessment の review 列の記入(承認・却下・保留・修正) |

- 承認列を持つ各表のトリガーは「その操作に対応する permission を、承認時点で有効に持つ operator」であることを確認する。

### 3.1 `physical_product`

- **目的:** 倉庫で扱う最小の物理商品。在庫・入荷・棚卸・検品・原価の基準。**P5 実体**。

| 列 | 型 | NULL | 制約・値 | AI | 人 | 説明 |
|---|---|---|---|---|---|---|
| `pp_id` | text | NOT NULL | **PK** `^PP-\d{6,}$`、複合 FK → core_entity | ✅ | | |
| `name` | text | NOT NULL | 空文字不可 | ✅ | | 表示名 |
| `kind` | text | NOT NULL | `GOODS` / `ACCESSORY` / `MATERIAL` | ✅ | | ACCESSORY = 店舗付属品。MATERIAL は Phase 1 では作らない |
| `status` | text | NOT NULL | `PROVISIONAL`(既定) / `ACTIVE` / `RETIRED` | PROVISIONAL | ACTIVE・RETIRED | |
| `origin_key` | text | NOT NULL | **UQ** | ✅ | | 冪等キー(例 `legacy:P000123`、`dup-group:G016`、`jan:4550391811816`) |
| `tax_category` | text | NOT NULL | `STANDARD`(10%)/ `REDUCED`(8% 軽減税率)/ `UNKNOWN`(既定) | 提案のみ(UNKNOWN 以外は人) | ✅(`PRODUCT_APPROVE`) | 税率が書かれていない観測値を税抜に換算するときに使う。食品・飲料と日用品・ペットフードで税率が違うため |
| `activated_by` / `activated_at` | text / timestamptz | NULL可 | human のみ。status≠PROVISIONAL なら NOT NULL | | ✅ | |
| `retired_by` / `retired_at` / `retire_reason` | | NULL可 | status=RETIRED なら NOT NULL | | ✅ | |
| `note` | text | NULL可 | | ✅ | | |
| 共通列 | | | 更新可 | | | |

- 遷移: PROVISIONAL → ACTIVE → RETIRED、PROVISIONAL → RETIRED。逆方向は不可。
- **JAN・原価はこの表に持たない。** PROVISIONAL の PP は正式原価計算・自動処理に使わない。

### 3.2 `identifier` — コード登録簿

- **目的:** スキャン・取込で出会ったコードそのものを登録する。**対象を持たない。** **P3 追記専用**。

| 列 | 型 | NULL | 制約・値 | AI | 説明 |
|---|---|---|---|---|---|
| `identifier_id` | text | NOT NULL | **PK** `^ID-\d{8,}$` | ✅ | |
| `id_type` | text | NOT NULL | `JAN`(GTIN-13/8)/ `GTIN14` / `MAKER_CODE` / `JAN_PARTIAL4` / `ASIN` / `FNSKU` / `KIT_LABEL` / `TEMP_ID` | ✅ | FNSKU・KIT_LABEL は Phase 2 用に値だけ予約 |
| `value` | text | NOT NULL | 正規化済み(数字は先頭0を保持) | ✅ | |
| | | | **UQ** `(id_type, value)` | | 同じコードは1行 |
| `checkdigit_valid` | boolean | NULL可 | JAN / GTIN14 のみ | ✅ | 不一致でも登録する(誤記の証拠) |
| `first_seen_source` | text | NOT NULL | 例 `legacy_product_master` / `pricetar_sku` / `set_sku` / `stocktake_scan` | ✅ | |
| `first_seen_ref` | text | NOT NULL | ファイル + sha256 + 行 | ✅ | |
| 共通列 | | | created_* / ingest_run_id | | |

### 3.3 `identifier_link` — コード → 対象の候補

- **目的:** 1コード → 複数対象(共有 JAN)、1対象 → 複数コード(旧 JAN)を持つ。**P1 提案→承認**。

| 列 | 型 | NULL | 制約・値 | AI | 人 | 説明 |
|---|---|---|---|---|---|---|
| `link_id` | text | NOT NULL | **PK** `^IL-\d{8,}$` | ✅ | | |
| `identifier_id` | text | NOT NULL | **FK** identifier | ✅ | | |
| `entity_id` / `entity_type` | text | NOT NULL | 複合 **FK** core_entity(種類の制限なし) | ✅ | | **Phase 2 で STOCK_FORM が加わっても定義は不変** |
| `link_basis` | text | NOT NULL | `PRINTED_ON_ITEM` / `PRINTED_ON_PACKAGE` / `CASE_OUTER` / `SELLER_REGISTERED` / `DERIVED_FROM_SKU` / `LEGACY_MASTER` / `MISFILED` | ✅ | | 何を根拠に紐付けたか |
| `scan_policy` | text | NOT NULL | `AUTO_IF_UNIQUE` / `CONFIRM_ALWAYS` / `NOT_FOR_SCAN` | ✅(提案) | 確定 | 既定 `CONFIRM_ALWAYS`。MISFILED・JAN_PARTIAL4 は `NOT_FOR_SCAN` 固定 |
| `valid_from` / `valid_to` | date | NULL可 | | ✅ | | JAN の切替(リニューアル・旧 JAN)を表す |
| `confidence` | text | NOT NULL | `HIGH` / `MEDIUM` / `LOW` | ✅ | | |
| `evidence` | jsonb | NOT NULL | | ✅ | | |
| `record_status` | text | NOT NULL | P1 | PROPOSED | ACTIVE・REJECTED | |
| `approved_by` / `approved_at` | | NULL可 | human。ACTIVE なら NOT NULL | | ✅ | |
| `assessment_id` | text | NULL可 | FK assessment | ✅ | | |
| `superseded_by` | text | NULL可 | FK identifier_link | ✅ | | |

- **UQ:** `(identifier_id, entity_id)` を `record_status IN ('PROPOSED','ACTIVE')` の範囲で一意。
- **1コード → 複数の ACTIVE 紐付けを許す**(共有 JAN の事実)。一意性はスキャン解決(§1.4)で判定し、表の制約では強制しない。
- `scan_policy='AUTO_IF_UNIQUE'` にできるのは人の承認時のみ。

### 3.4 `legacy_mapping` — 既存 P が何を表しているか

- **目的:** 既存 P 1,211件を削除・変更せずに PP / CP へ対応付ける。**P1 提案→承認**。
- AI の候補は **assessment(LEGACY_MAPPING)** に置く。この表の行は、人の review(APPROVED / CORRECTED)を受けて作るのが原則。一覧性のために AI が PROPOSED を作ることも許す(下流は ACTIVE のみ参照)。

| 列 | 型 | NULL | 制約・値 | AI | 人 | 説明 |
|---|---|---|---|---|---|---|
| `mapping_id` | text | NOT NULL | **PK** | ✅ | | |
| `legacy_p` | text | NOT NULL | `^P\d{6}$`(FK なし。正本は Excel) | ✅ | | |
| `entity_id` / `entity_type` | text | NOT NULL | 複合 FK core_entity、`entity_type IN ('PHYSICAL_PRODUCT','COMPOSITION')` | ✅ | | |
| `mapping_role` | text | NOT NULL | `SOLE` / `MERGED` | ✅ | | MERGED = 複数の legacy P が同じ対象の重複登録。**代表 P は決めない** |
| `valid_from` / `valid_to` | date | NULL可 | | ✅ | | |
| `legacy_snapshot_ref` | text | NOT NULL | 取込時のマスター sha256 + 行 | ✅ | | |
| `record_status` / `approved_by` / `approved_at` / `assessment_id` / `superseded_by` | | | P1 | PROPOSED | ACTIVE・REJECTED | |

- **UQ:** `legacy_p` ごとに ACTIVE は最大1行。
- MERGED の ACTIVE 化 = **P 統合と同じ扱い(人間承認必須)**。

### 3.5 `product_relationship` — PP 同士の関係

- **P1 提案→承認**。

| 列 | 型 | NULL | 制約・値 | AI | 人 | 説明 |
|---|---|---|---|---|---|---|
| `relationship_id` | text | NOT NULL | **PK** | ✅ | | |
| `rel_type` | text | NOT NULL | `DUPLICATE_OF` / `SUCCESSOR_OF` / `VARIANT_MEMBER` | ✅ | | |
| `from_pp_id` | text | NOT NULL | FK physical_product | ✅ | | |
| `to_pp_id` | text | NULL可 | FK。VARIANT_MEMBER 以外は NOT NULL | ✅ | | |
| `variant_group_key` | text | NULL可 | VARIANT_MEMBER なら NOT NULL | ✅ | | |
| `variant_axis` | text | NULL可 | `COLOR` / `SIZE` / `SCENT` / `CAPACITY` / `DESIGN` / `OTHER` | ✅ | | |
| `evidence` | jsonb | NOT NULL | | ✅ | | |
| `record_status` / `approved_by` / `approved_at` / `assessment_id` / `superseded_by` | | | P1 | | ✅(ACTIVE 化) | |

- **CHECK:** `from_pp_id <> to_pp_id`。**UQ:** `(rel_type, from_pp_id, coalesce(to_pp_id, variant_group_key))` の ACTIVE 一意。
- 色・サイズ・種類違いは別 PP のまま VARIANT_MEMBER で束ねる(統合しない)。

### 3.6 `composition` — 販売構成

- **P5 実体**。ACTIVE 後は構成を変更しない(変えるときは新しい CP を作り、listing 参照を付け替える)。

| 列 | 型 | NULL | 制約・値 | AI | 人 | 説明 |
|---|---|---|---|---|---|---|
| `cp_id` | text | NOT NULL | **PK** `^CP-\d{6,}$`、複合 FK core_entity | ✅ | | |
| `name` | text | NOT NULL | | ✅ | | |
| `comp_type` | text | NOT NULL | `FIXED` / `SELECTION` / `UNCLASSIFIED` | FIXED・UNCLASSIFIED | SELECTION 化 | 福袋9件は UNCLASSIFIED(D10) |
| `status` | text | NOT NULL | `PROVISIONAL` / `ACTIVE` / `RETIRED` | PROVISIONAL | ACTIVE・RETIRED | |
| `origin_key` | text | NOT NULL | **UQ** | ✅ | | |
| `activated_*` / `retired_*` / `note` / 共通列 | | | physical_product と同じ | | ✅ | |

- **ACTIVE 化の条件(トリガー):** `comp_type='FIXED'`、構成品1行以上、構成品の PP がすべて ACTIVE。
- UNCLASSIFIED は構成品を持てない。SELECTION は Phase 3(selection_rule)まで ACTIVE にしない。

### 3.7 `composition_component` — 構成品

| 列 | 型 | NULL | 制約・値 | AI | 説明 |
|---|---|---|---|---|---|
| `component_id` | text | NOT NULL | **PK** | ✅ | |
| `cp_id` | text | NOT NULL | FK composition | ✅ | |
| `pp_id` | text | NOT NULL | FK physical_product | ✅ | |
| `quantity` | integer | NOT NULL | > 0 | ✅ | |
| `role` | text | NOT NULL | `GOODS` / `ACCESSORY` | ✅ | ACCESSORY = おしぼり・スパウトパウチ・スプーン等の店舗付属品 |
| `sort_no` | integer | NOT NULL | | ✅ | |
| `evidence` | jsonb | NOT NULL | 商品名の入数表記・SKU の `-3` 等 | ✅ | |

- **UQ:** `(cp_id, pp_id, role)`。親の CP が PROVISIONAL の間だけ追加・変更可(トリガー)。**構成の確定 = 親の ACTIVE 化(人間承認)**。
- メーカー同梱品(ケース付き・替刃付き)は PP の一部として扱い、構成品にしない。梱包資材は固定費として扱い、構成品にしない。

### 3.8 `listing` — 販売口

- **P4 事実の写し**。複合 FK → core_entity(`LISTING`)。

| 列 | 型 | NULL | 制約・値 | AI | 説明 |
|---|---|---|---|---|---|
| `listing_id` | text | NOT NULL | **PK** `^LS-\d{6,}$` | ✅ | |
| `channel` | text | NOT NULL | `AMAZON` / `RAKUTEN` / `OTHER` | ✅ | |
| `asin` | text | NULL可 | `^B0[0-9A-Z]{8}$` | ✅ | **一意にしない** |
| `seller_sku` | text | NULL可 | | ✅ | |
| `rakuten_item_id` / `rakuten_sku` | text | NULL可 | | ✅ | |
| `fulfillment` | text | NOT NULL | `FBA` / `SELF` / `UNKNOWN` | ✅ | |
| `listing_group_id` | text | NULL可 | FK listing_group | ✅ | |
| `legacy_listing_id` | text | NULL可 | **UQ**(非NULL時)`^C\d{6}$` | ✅ | |
| `channel_status` | text | NOT NULL | `ACTIVE` / `INACTIVE` / `UNKNOWN` | ✅ | |
| `first_seen_at` / `last_seen_at` | timestamptz | NOT NULL | | ✅ | |
| `source_type` / `source_ref` | text | NOT NULL | | ✅ | |

- **UQ(部分):** Amazon は `(channel, seller_sku)`、楽天は `(channel, rakuten_item_id, rakuten_sku)`。自然キーは変更不可(SKU 変更は新しい listing)。
- **listing は参照先を持たない**(listing_reference_history のみ)。

### 3.9 `listing_group`

| 列 | 型 | NULL | 制約・値 |
|---|---|---|---|
| `listing_group_id` | text | NOT NULL | **PK** |
| `channel` | text | NOT NULL | |
| `group_type` | text | NOT NULL | `RAKUTEN_ITEM` / `AMAZON_PARENT_ASIN` |
| `group_key` | text | NOT NULL | |
| `source_type` / `source_ref` / 共通列 | | | |

- **UQ:** `(channel, group_type, group_key)`。P4。

### 3.10 `listing_reference_history` — 出品の参照先(期間付き)

- **P1 提案→承認。参照先の変更は人間承認必須。**

| 列 | 型 | NULL | 制約・値 | AI | 人 | 説明 |
|---|---|---|---|---|---|---|
| `ref_id` | text | NOT NULL | **PK** | ✅ | | |
| `listing_id` | text | NOT NULL | FK listing | ✅ | | |
| `entity_id` / `entity_type` | text | NOT NULL | 複合 FK、`IN ('PHYSICAL_PRODUCT','COMPOSITION')` | ✅ | | |
| `valid_from` | date | NOT NULL | | ✅ | | |
| `valid_to` | date | NULL可 | 一度だけ設定可 | | ✅ | |
| `derived_from_legacy_p` | text | NULL可 | | ✅ | | 旧紐付け(監査用) |
| `reason` | text | NOT NULL | `初期移行` / `誤紐付け修正` / `SKU差し替え` 等 | ✅ | | |
| `record_status` / `approved_by` / `approved_at` / `assessment_id` / `superseded_by` | | | P1 | PROPOSED | ACTIVE・REJECTED | |

- **排他制約:** 同じ listing の ACTIVE 行で期間が重ならない(`EXCLUDE USING gist (listing_id WITH =, daterange(valid_from, valid_to, '[]') WITH &&) WHERE (record_status='ACTIVE')`)。
- ASIN 誤紐付け疑い 41件は、人の確認前に PROPOSED 行を作らない(assessment のみ)。

### 3.11 `cost_history` — PP の正式原価(税抜)

- **P2 承認済み履歴。AI は書けない。**

| 列 | 型 | NULL | 制約・値 | 説明 |
|---|---|---|---|---|
| `cost_id` | text | NOT NULL | **PK** `^CH-\d{8,}$` | |
| `pp_id` | text | NOT NULL | FK physical_product(ACTIVE のみ・トリガー) | |
| `unit_cost_excl_tax` | numeric(§2.2) | NOT NULL | ≥ 0 | **税抜・PP 1個あたり**(D2) |
| `currency` | text | NOT NULL | `JPY` | |
| `valid_from` / `valid_to` | date | NOT NULL / NULL可 | | |
| `source_observation_id` | text | NOT NULL | FK cost_observation | 採用した観測値(元の金額・税区分はそちらに残る) |
| `normalization` | jsonb | NOT NULL | 必須キー: `observed_amount`・`tax_inclusion`・`tax_inclusion_basis`・`tax_rate`・`tax_rate_basis`・`formula`・`unit_conversion`・`excl_tax_amount`。例 `{"observed_amount":1100,"tax_inclusion":"INCLUDED","tax_inclusion_basis":"OPERATIONAL_CONVENTION","tax_rate":0.10,"tax_rate_basis":"PP_TAX_CATEGORY","formula":"1100/1.10","unit_conversion":"なし","excl_tax_amount":1000}` | 税抜・単位への換算の記録 |
| `tax_basis_assessment_id` | text | NULL可 | FK assessment(`TAX_BASIS`) | 観測値の税区分が UNKNOWN、または根拠が運用慣行だけのときに、人が税区分を確定した記録 |
| `basis` | text | NOT NULL | 人が読む根拠 | |
| `assessment_id` | text | NOT NULL | FK assessment | どの AI 判定を経たか(D6) |
| `approved_by` / `approved_at` | | NOT NULL | human | |
| `record_status` | text | NOT NULL | `ACTIVE` / `SUPERSEDED` / `VOID` | |
| `supersedes` | text | NULL可 | FK cost_history | 訂正元 |

- **排他制約:** 同じ PP の ACTIVE 行で期間が重ならない。
- **商品マスターの標準原価を一括でこの表に入れない**(D6)。
- **元の観測値を失わない(S8):** `source_observation_id` は NOT NULL で、観測値は削除・変更できない(P3 + RESTRICT)。
  さらにトリガーで、`normalization` の `observed_amount`・`tax_inclusion`・`tax_rate` が参照先の観測値と一致することを確認する(食い違う換算記録は作れない)。
- **UNKNOWN を推測で採用しない(S7):** 参照先の観測値の `tax_inclusion='UNKNOWN'` なら、`tax_basis_assessment_id`(人が review で APPROVED / CORRECTED した `TAX_BASIS` 判定)が無い限り作成を拒否する。
  `tax_inclusion_basis='OPERATIONAL_CONVENTION'`(運用上は税込、というだけの根拠)の観測値も同じ扱いとする。
- **PROVISIONAL の PP には作れない(S3):** 対象 PP が ACTIVE でなければ拒否する。
- **人が確定した税区分と適用する税区分の一致:** `TAX_BASIS` 判定の review が APPROVED なら AI の判定値、CORRECTED なら `corrected_value` を「人が確定した税区分」とし、`normalization.tax_inclusion_applied` と一致しなければ拒否する。
- **観測値がその商品のものであること:** 観測値の対象が、同じ PP / CP、または承認済み(ACTIVE)の legacy_mapping・出品参照でその PP / CP に結び付く legacy P・出品でなければ拒否する(他の商品の観測値を流用させない)。

**税抜への標準化の規則(D12)**

| 観測値の状態 | 税抜額 | 標準化できたか |
|---|---|---|
| `EXCLUDED` | observed_amount | ✅ |
| `INCLUDED` かつ tax_rate あり | observed_amount ÷ (1 + tax_rate) | ✅ |
| `INCLUDED` で tax_rate なし、PP の tax_category が STANDARD / REDUCED | observed_amount ÷ (1 + 区分の税率)。`tax_rate_basis='PP_TAX_CATEGORY'` | ✅(根拠を記録) |
| `INCLUDED` で税率が決まらない | — | ❌ `TAX_RATE_UNKNOWN` |
| `UNKNOWN` | — | ❌ `TAX_BASIS_UNKNOWN` |

- 例: 商品マスター 1,100円(INCLUDED・10%)と納品書 1,000円(EXCLUDED・10%)は、どちらも税抜 1,000円として比較できる。
- 標準化は計算で行い、**観測値の行そのものは書き換えない。** 標準化した値は、assessment の evidence と cost_history の `normalization` に記録する。
- 途中で丸めない(D13)。

### 3.12 `composition_cost_history` — CP の正式原価(直接仕入・税抜)

- **P2**。列構成は cost_history と同じで、`pp_id` を `cp_id` に、`unit_cost_excl_tax` を `set_cost_excl_tax`(1組あたり)に置き換える。追加列:

| 列 | 型 | NULL | 制約・値 | 説明 |
|---|---|---|---|---|
| `cost_id` | text | NOT NULL | **PK** `^CX-\d{8,}$` | |
| `cost_origin` | text | NOT NULL | `DIRECT_PURCHASE` | Phase 1 は直接仕入のみ |
| `purchase_evidence` | jsonb | NOT NULL | 仕入先・見積 / 伝票番号 | |

**原価の適用規則**
1. 対象日に有効な ACTIVE の composition_cost_history があれば **A** を使う。
2. 無ければ **B** = Σ(構成品 PP の対象日に有効な cost_history × quantity)。1つでも欠ければ「算出不可」(0 や推定で埋めない)。
3. A を使うときも B を計算し、**差額 ≥ 50円 かつ 差率 ≥ 5%** なら警告候補とする(関数 `effective_composition_cost()` が `variance_flag='WARN'` を返す。記録は assessment `COST_VARIANCE`)。
4. cost_observation は 1・2 の計算に使わない。

### 3.13 `cost_observation` — 原価の観測値(元データのまま)

- **P3 追記専用**(D2・D6)。

| 列 | 型 | NULL | 制約・値 | 説明 |
|---|---|---|---|---|
| `observation_id` | text | NOT NULL | **PK** `^CO-\d{9,}$` | |
| `source` | text | NOT NULL | `legacy_product_master` / `pricetar_inventory_csv` / `sku_string` / `rakuten_item_number` / `amazon_kpi` / `rakuten_kpi` / `cost_confirmation_record` / `purchase_record` / `delivery_note` / `invoice` / `maker_document` / `wholesaler_document` / `price_list` / `supplier_csv` / `other_document` | |
| `source_ref` | text | NOT NULL | ファイル + sha256 + 行 | |
| `source_document_id` | text | NULL可 | `^DOC-\d{8}-\d{6}$` | AIKOS Document Intelligence の DocID(原本は `35_Documents/10_Archive`)。書類から読んだ観測値で必須 |
| `source_reliability` | text | NOT NULL | `HIGH` / `MEDIUM` / `LOW` | 入力経路の確からしさ(書類の明記 > 仕入先 CSV > 手入力の運用値 > 文字列埋め込み) |
| `observed_at` | timestamptz | NOT NULL | そのデータの基準日時(書類は発行日) | |
| `observed_amount` | numeric(§2.2) | NOT NULL | **元資料に書かれた金額をそのまま**(換算しない) | |
| `currency` | text | NOT NULL | `JPY` | |
| `tax_inclusion` | text | NOT NULL | `INCLUDED` / `EXCLUDED` / `UNKNOWN` | 分からなければ UNKNOWN |
| `tax_inclusion_basis` | text | NOT NULL | `DOCUMENT_STATED`(資料に明記)/ `SOURCE_SPEC`(CSV 等の仕様で明示)/ `OPERATIONAL_CONVENTION`(運用上の慣行のみ)/ `HUMAN_CONFIRMED`(取込時に人が確認)/ `NONE` | 税区分の根拠。UNKNOWN のときは NONE |
| `tax_rate` | numeric(§2.2) | NULL可 | 資料に書かれていれば記録 | |
| `tax_rate_basis` | text | NOT NULL | `DOCUMENT_STATED` / `SOURCE_SPEC` / `NONE` | 税率の根拠(PP の税区分からの推定は観測値には書かない。標準化の記録にだけ書く) |
| `amount_unit` | text | NOT NULL | `PER_LISTING_UNIT` / `PER_PP` / `PER_COMPOSITION` / `UNKNOWN` | 単品かセットかを決めつけない |
| `subject_entity_id` / `subject_entity_type` | text | NULL可 | 複合 FK core_entity(PP / CP / LISTING) | |
| `subject_legacy_p` | text | NULL可 | `^P\d{6}$` | legacy P 単位の観測(標準原価等) |
| | | | CHECK: subject_entity か subject_legacy_p の**ちょうど一方** | |
| `evidence` | jsonb | NOT NULL | | |
| `ingest_run_id` | text | NOT NULL | | |

- **UQ:** `(source, source_ref, 対象)`(再取込で重複しない)。
- 原価確認記録(人の確認)も、観測値(`cost_confirmation_record`)として取り込む。cost_history への登録は、AI Assessment → Human Approval を経て行う。

**既存データの初期値(D12)**

| source | tax_inclusion | tax_inclusion_basis | source_reliability | 理由 |
|---|---|---|---|---|
| `legacy_product_master`(標準原価) | `INCLUDED` | `OPERATIONAL_CONVENTION` | MEDIUM | 運用上は原則税込だが、行ごとの証拠は無い |
| `pricetar_inventory_csv`(cost) | `INCLUDED` | `OPERATIONAL_CONVENTION` | MEDIUM | 同上(手入力の運用値) |
| `rakuten_item_number`(「/仕入値」) | `INCLUDED` | `OPERATIONAL_CONVENTION` | LOW | 文字列への埋め込み |
| `sku_string`(`@値`) | `UNKNOWN` | `NONE` | LOW | 入力規則が確認できない |
| `amazon_kpi` / `rakuten_kpi`(L列) | `UNKNOWN` | `NONE` | LOW | 他の値の転記で、出所が混在する |
| `cost_confirmation_record` | `UNKNOWN` | `NONE` | MEDIUM | 人の確認だが税区分の記載が無い |
| 書類(`delivery_note` / `invoice` 等) | 書類の表記どおり。読めなければ `UNKNOWN` | `DOCUMENT_STATED` / `NONE` | 表記があれば HIGH | |

- `OPERATIONAL_CONVENTION` は**比較と警告には使える**が、それだけを根拠に正式原価を作ることはできない(§3.11 S7: 人の `TAX_BASIS` 確定が必要)。

**書類から取り込む場合の evidence(将来の Document Intelligence 連携)**

| キー | 内容 |
|---|---|
| `line_unit_price` | 明細単価(書かれたまま) |
| `line_quantity` / `line_amount` | 数量・明細金額 |
| `tax_label` | 税込 / 税抜の表記(原文) |
| `tax_rate` | 税率(明細または書類単位、8% / 10% 混在可) |
| `tax_amount` | 消費税額 |
| `subtotal` / `total` | 小計・合計 |
| `issuer` | 発行元(仕入先・メーカー・問屋) |
| `document_date` | 資料の日付 |
| `page` / `line_no` / `bbox` | 原本上の位置(再確認用) |
| `extraction_method` / `extraction_confidence` | 読み取り方法と確からしさ |

- **税区分を判定できない書類は `UNKNOWN`** とし、assessment(`TAX_BASIS` / `UNKNOWN`)で人の確認または追加証拠待ちにする。
- 明細の合計と書類の小計・税額・合計が整合するかも、assessment の evidence に記録する(税区分の判定材料)。

### 3.14 `assessment` — AI 判定 + 人の review

- AI 部分: **P3 追記専用**。review 部分: 人だけが更新。

**AI 部分(作成後は変更不可)**

| 列 | 型 | NULL | 制約・値 | 説明 |
|---|---|---|---|---|
| `assessment_id` | text | NOT NULL | **PK** `^AS-\d{9,}$` | |
| `assessment_type` | text | NOT NULL | `LEGACY_MAPPING` / `LISTING_REFERENCE` / `COST_VARIANCE` / `TAX_BASIS` / `DUPLICATE_PP` / `IDENTIFIER_LINK` / `COMPOSITION_CLASS` | |
| `subject_entity_id` / `subject_entity_type` | text | NULL可 | 複合 FK core_entity | |
| `subject_key` | text | NULL可 | 例 legacy P・ASIN・observation_id | |
| | | | CHECK: 上の2つの**少なくとも一方** | |
| `verdict` | text | NOT NULL | 種類ごとの許可値(下表) | |
| `proposed_entity_id` / `proposed_entity_type` | text | NULL可 | 複合 FK core_entity | 候補の参照先(PROVISIONAL の PP / CP 等) |
| `confidence` | text | NOT NULL | `HIGH` / `MEDIUM` / `LOW` | |
| `reason` | text | NOT NULL | 空文字不可 | 判定理由 |
| `evidence` | jsonb | NOT NULL | | |
| `rule_version` | text | NOT NULL | 例 `legacy-map/1.0` | |
| `assessed_by` | text | NOT NULL | `^(ai\|rule):` のみ | 人は assessment を作らない |
| `assessed_at` | timestamptz | NOT NULL | | |
| `supersedes_assessment_id` | text | NULL可 | FK assessment | 再判定 |
| `content_hash` | text | NOT NULL | AI 部分の sha256 | Excel 確認表の改ざん検知 |

**Human Review 部分(人だけが更新・入力経路に依存しない)**

| 列 | 型 | NULL | 制約・値 |
|---|---|---|---|
| `review_status` | text | NOT NULL | `UNREVIEWED`(既定) / `ON_HOLD`(保留) / `APPROVED` / `REJECTED` / `CORRECTED` |
| `reviewed_by` / `reviewed_at` | | NULL可 | human。UNREVIEWED 以外は NOT NULL |
| `review_note` | text | NULL可 | ON_HOLD・REJECTED・CORRECTED のとき NOT NULL |
| `corrected_entity_id` / `corrected_entity_type` | text | NULL可 | CORRECTED のとき NOT NULL(複合 FK) |
| `corrected_value` | text | NULL可 | 対象ではなく**値**を人が修正したとき(例: TAX_BASIS の INCLUDED / EXCLUDED)。CORRECTED では corrected_entity か corrected_value のどちらかが必須 |
| `review_channel` | text | NULL可 | `EXCEL` / `UI` |
| `review_ref` | text | NULL可 | 確認表のファイル sha256 + 行、または UI のセッション ID |

- **遷移(トリガー):** UNREVIEWED → ON_HOLD / APPROVED / REJECTED / CORRECTED、ON_HOLD → APPROVED / REJECTED / CORRECTED。**APPROVED / REJECTED / CORRECTED は最終状態**(やり直しは新しい assessment で行う)。
- **`verdict='CONSISTENT_HIGH'` でも review_status は UNREVIEWED のまま。** 下流の ACTIVE 化は review が APPROVED / CORRECTED の assessment にだけ許す。

| assessment_type | verdict |
|---|---|
| LEGACY_MAPPING | `PHYSICAL_PRODUCT_CANDIDATE` / `COMPOSITION_CANDIDATE` / `ALIAS_CANDIDATE` / `NEEDS_REVIEW` |
| LISTING_REFERENCE | `CONSISTENT_HIGH` / `SUSPECT_MISLINK` / `AWAITING_DATA` / `HUMAN_REVIEW_CANDIDATE` |
| COST_VARIANCE | `OK` / `CONFIRMED_WARN` / `REFERENCE_WARN` / `TAX_BASIS_UNKNOWN` / `TAX_RATE_UNKNOWN` / `UNIT_UNKNOWN` / `NO_BASELINE` |
| TAX_BASIS | `INCLUDED` / `EXCLUDED` / `UNKNOWN`(人が review で確定・修正する。CORRECTED の修正値は `review_note` と evidence に記録) |
| DUPLICATE_PP | `DUPLICATE_CANDIDATE` / `VARIANT_CANDIDATE` / `NOT_DUPLICATE` |
| IDENTIFIER_LINK | `VALID` / `CHECKDIGIT_ERROR` / `MISFILED` / `SHARED_ACROSS_ENTITIES` |
| COMPOSITION_CLASS | `FIXED_CANDIDATE` / `SELECTION_CANDIDATE` / `HUMAN_REVIEW_REQUIRED` |

### 3.15 取込ステージング(`legacy_ingest`)

| テーブル | 主な列 | 説明 |
|---|---|---|
| `ingest_run` | `ingest_run_id`(PK)・`started_at`・`finished_at`・`actor`(`human:` が起動)・`status`(RUNNING / SUCCEEDED / FAILED)・`files` jsonb(パス・sha256・mtime・行数・取込前後の sha256) | 1回の手動取込(D7) |
| `legacy_product_snapshot` | `ingest_run_id`・`row_no`・`legacy_p`・`jan_raw`・`name`・`kind`・`standard_cost`・`channel`・`status`・`raw` jsonb | 商品マスター |
| `legacy_listing_snapshot` | `ingest_run_id`・`row_no`・`legacy_listing_id`・`legacy_p`・`channel`・楽天/Amazon の各キー・`price`・`pack`・`cost_unit`・`proof`・`raw` | 出品テーブル |
| `external_file_snapshot` | `ingest_run_id`・`file_kind`・`row_no`・`raw` jsonb | Pricetar CSV・RMS CSV・KPI・原価確認記録 |

- 追記のみ。古い run の削除は別承認。

---

## 4. 制約一覧(DDL + 制約テストの対象)

| # | 制約 | 対象 |
|---|---|---|
| C1 | 承認・review 列は `^human:` かつ ACTIVE の operator で、その操作の permission(§3.0c)を承認時点で有効に持つ | 全表の approved_by / activated_by / retired_by / reviewed_by |
| C2 | 承認者列と承認日時列は同時に NULL / 非 NULL | 同上 |
| C3 | P1: 内容列は変更不可、状態遷移は定義のとおり、ACTIVE は承認必須 | identifier_link / legacy_mapping / product_relationship / listing_reference_history |
| C4 | P2: 人の承認でのみ作成、`valid_to` は一度だけ、削除不可 | cost_history / composition_cost_history |
| C5 | P3: UPDATE / DELETE 拒否(assessment は review 列のみ・定義の遷移のみ) | identifier / cost_observation / assessment |
| C6 | P5: PROVISIONAL → ACTIVE / RETIRED は人、逆方向不可、削除不可 | physical_product / composition |
| C7 | 実体と core_entity の複合 FK・参照側の entity_type 制限 | 全参照 |
| C8 | 期間の重なり禁止(ACTIVE) | listing_reference_history(listing 単位)、cost_history(PP 単位)、composition_cost_history(CP 単位)、legacy_mapping(legacy_p 単位の ACTIVE 1行) |
| C9 | CP の ACTIVE 化条件(FIXED・構成品あり・構成品 PP が ACTIVE)、ACTIVE 後の構成変更禁止 | composition / composition_component |
| C10 | cost_history の PP は ACTIVE、`assessment_id` と `source_observation_id` 必須 | cost_history |
| C11 | ACTIVE 化の元になる assessment は APPROVED / CORRECTED | P1 表・P2 表 |
| C12 | 自然キー・ID 形式・列挙値 | 全表 |
| C13 | 既存 P 番号は主キー・FK に使わない(`^P\d{6}$` は legacy 列にのみ出現) | 全表 |
| **S1** | core_entity の entity_type と実体表が一致する(実体表側の固定値列 + 複合 FK。core_entity 側に対応する実体行が無い登録は COMMIT 時に拒否する遅延制約トリガー) | core_entity / physical_product / composition / listing |
| **S2** | RETIRED の PP / CP へ、新しい listing 参照(PROPOSED・ACTIVE とも)を作れない。ACTIVE 化の時点でも再確認する | listing_reference_history |
| **S3** | PROVISIONAL の PP を正式原価に使えない(cost_history の作成拒否。CP の B 計算も構成品が ACTIVE のときだけ) | cost_history / composition |
| **S4** | 未承認の legacy_mapping を正式処理に使えない(`derived_from_legacy_p` を持つ listing 参照は、同じ legacy P → 同じ対象の ACTIVE な legacy_mapping が無ければ作れない。正式処理用のビューは ACTIVE のみを公開) | listing_reference_history / ビュー |
| **S5** | AI プロセスは Human Approval を書けない(下の DB ロールの権限 + `^human:` CHECK + operator の permission 確認トリガーの三重) | 承認列・review 列・cost_history / composition_cost_history |
| **S6** | 曖昧な Identifier で自動確定できない(解決関数 `resolve_identifier()` は、候補が複数・PROVISIONAL・`CONFIRM_ALWAYS` のとき `NEEDS_SELECTION` を返し、確定 ID を返さない。`AUTO_IF_UNIQUE` は承認時だけ設定できる) | identifier_link / 解決関数 |
| **S7** | 税区分 UNKNOWN(および運用慣行のみ)の観測値を、人の `TAX_BASIS` 確定なしに正式原価へ採用できない | cost_history / composition_cost_history |
| **S8** | 正式原価へ変換しても、観測値の元の金額・税区分・evidence が失われない(観測値は追記専用・RESTRICT、`normalization` と観測値の一致をトリガーで検証) | cost_observation / cost_history |

**DB ロールの分離(S5 の土台)**

| ロール(NOLOGIN のグループロール) | 使う主体 | 権限 |
|---|---|---|
| `pc_owner` | マイグレーション(Migrate.sh) | スキーマの所有者。アプリからは使わない |
| `pc_ingest` | 取込ジョブ・AI 判定 | P3・P4 表への INSERT、PROVISIONAL / PROPOSED の INSERT、assessment の AI 部分の INSERT。**承認列・review 列・cost_history / composition_cost_history には権限なし**(列単位の権限) |
| `pc_reviewer` | Excel 確認表の取込(人が起動) | review 列の UPDATE、ACTIVE 化の UPDATE、cost_history / composition_cost_history の INSERT。**ただしトリガーで operator の permission を確認** |
| `pc_reader` | 検証レポート・将来の MCP 参照 | SELECT のみ |

- ログインロールとパスワードは人が作成し、NAS の `.env` だけに置く(リポジトリには置かない。SECURITY.md の方針どおり)。
- どのロールにも、正本 Excel への書き込み経路は無い(D1)。

---

## 5. 関係一覧(ER 相当)

```
operator ─(human:<operator_id>)─ 全表の承認・review 列

core_entity ◄─1:1─ physical_product(PP-) / composition(CP-) / listing(LS-) / [Phase 2] stock_form(SF-)
core_entity ─1:N─ identifier_link / legacy_mapping / listing_reference_history / cost_observation / assessment

identifier ─1:N─ identifier_link ─N:1─ core_entity(PP / CP / LS / [Phase 2] SF)

legacy P(Excel・FK なし)─1:N(ACTIVE は1)─ legacy_mapping ─N:1─ core_entity(PP / CP)
physical_product ─1:N─ product_relationship(from / to)
composition ─1:N─ composition_component ─N:1─ physical_product
listing_group ─1:N─ listing ─1:N─ listing_reference_history ─N:1─ core_entity(PP / CP)
physical_product ─1:N─ cost_history ─N:1─ cost_observation(source_observation_id)
composition ─1:N─ composition_cost_history ─N:1─ cost_observation
assessment ─0..1:N─ identifier_link / legacy_mapping / product_relationship / listing_reference_history / cost_history / composition_cost_history
legacy_ingest.ingest_run ─1:N─ snapshots / 取込で作った各行
```

- 親の削除は**すべて禁止(RESTRICT)**。使わなくなったものは RETIRED / SUPERSEDED / VOID で表す。

---

## 6. 移行フロー(既存 Excel → Shadow Model)

```
[正本 Excel / CSV](読み取りのみ。取込の前後で sha256 を照合。開いている可能性があるファイルはコピーを読む)
   │  手動実行(D7。起動者 human:<operator_id> を ingest_run に記録)
   ▼
legacy_ingest(ingest_run + snapshots)
   ├─① 事実の写し     listing_group / listing / identifier(コード登録)
   ├─② 観測値         cost_observation(legacy_product_master・pricetar_inventory_csv・sku_string・
   │                   rakuten_item_number・amazon_kpi・rakuten_kpi・cost_confirmation_record)
   ├─③ AI 判定         PROVISIONAL の PP / CP(origin_key で冪等)
   │                   assessment: LEGACY_MAPPING ×1,211 / IDENTIFIER_LINK / LISTING_REFERENCE(ASIN 41)/
   │                               COMPOSITION_CLASS(福袋9 = HUMAN_REVIEW_REQUIRED)/ COST_VARIANCE / DUPLICATE_PP
   │                   identifier_link / legacy_mapping の PROPOSED(任意)
   ├─④ Human Review    Excel 確認表(§7)→ review 列
   │                   → 承認分だけ ACTIVE 化・PP / CP の ACTIVE 化・cost_history / composition_cost_history の作成
   └─⑤ 検証レポート     件数照合・制約違反0件・未確認一覧・WARN 一覧
        ✕ Product Core → Excel の書き戻し(D1)
```

- **冪等性:** 同じ入力で再実行しても行が増えない(identifier の `(id_type, value)`・origin_key・listing の自然キー・observation の UQ)。
- 取込で作るのは P3 / P4 の行・PROVISIONAL・PROPOSED・assessment だけ。ACTIVE と正式原価は作らない。
- 出品テーブルの旧紐付け(出品 → legacy P)は、legacy_mapping が ACTIVE になった後に、listing_reference_history の PROPOSED(`reason='初期移行'`)として提案し、人の承認を経て ACTIVE にする。

---

## 7. AI Assessment / Human Approval フロー(D3・D11)

```
AI 判定(assessment, review_status=UNREVIEWED)
   │  「整合性高」でもここで止まる
   ▼
Excel 確認表の出力(AIKOS → Output/review_<type>_<日時>.xlsx)
   各行: assessment_id・content_hash・AI 判定(読み取り専用の列)・判断欄・修正先欄・メモ欄・operator_id 欄
   │
   ▼  人が記入: 承認 / 却下 / 保留 / 修正(修正先を指定)
Excel 確認表の取込(手動実行)
   検証: ファイル sha256 を記録 / assessment_id が存在 / content_hash が一致(AI 部分の改ざんなし)/
         operator が ACTIVE で `REVIEW`(と対象操作の permission)を持つ / 判断値が定義どおり / 遷移が許される
   → review 列を更新(review_channel=EXCEL, review_ref=sha256+行)
   │
   ├─ APPROVED  → 提案の ACTIVE 化 / 実体の ACTIVE 化 / 正式原価の作成(approved_by = reviewed_by)
   ├─ CORRECTED → 人が指定した先で新しい行を作って ACTIVE 化(AI の提案は REJECTED)
   ├─ REJECTED  → 提案を REJECTED(記録は残る)
   └─ ON_HOLD   → 何もしない(次回の確認表に再掲)
```

- 将来の承認 UI は、同じ review 列を `review_channel=UI` で更新するだけにする(表の変更は不要)。

### 7.1 確認表の保存先(D14)

- 保存先は **NAS 上の AIKOS 管理領域**とし、Pending(記入待ち)→ Reviewed(記入済み・取込待ち / 取込済み)→ Archive(取込完了・保管)の状態で管理する。
- **既存の AIKOS 文書管理設計**(`06_DocumentIntelligence/docs/Phase1_Intake_Spec.md`)を確認した結果:
  - doc root は `35_Documents/`。定義済みの領域は `00_Inbox`・`10_Archive`・`80_Ledger`・`85_Staging`・`90_Logs`・`99_Sandbox`。
  - `20_Catalogs` / `30_Extracted` / `40_Exports` は「Phase 2 以降で使用」と名前だけがあり、**用途は未定義**。確認表の置き場所として定義された領域は無い。
  - CLAUDE.md の例外承認(2026-09-24)は、`35_Documents` への書き込みを `06_DocumentIntelligence` のコード経由に限っている。
- したがって **確認表の NAS パスは本書では確定しない。** 次のどちらかを、AIKOS 文書管理設計の改訂として別途決める:
  - 案1: `40_Exports/` の用途を「AIKOS からの出力(確認表を含む)」と定義し、その下に Pending / Reviewed / Archive を置く
  - 案2: 確認表専用の新しい番号領域を文書管理設計に追加する
  いずれも CLAUDE.md の NAS 例外(書き込み主体・範囲)の改訂承認が必要。
- 保存先が決まるまでは、確認表の往復(実装順序の段階8)に着手しない。段階1〜7は保存先に依存しない。
- **一括承認:** 同じ `rule_version`・同じ `verdict`・`confidence=HIGH` の範囲に限る。各行に reviewed_by / reviewed_at を記録する。
- **人間承認が必要な変更:** P 統合(MERGED / DUPLICATE_OF)、正式原価(cost_history / composition_cost_history)、Listing 参照先、PP / CP の ACTIVE・RETIRED、scan_policy の AUTO_IF_UNIQUE 化。

**原価差異の判定(D8・D12)**
1. 観測値を税抜に揃える(§3.11 の標準化規則。関数 `normalize_observation()`)。標準化の確度:
   - **CONFIRMED** … 税区分の根拠が 書類の明記 / 仕様 / 人の確認、かつ税率が観測値に記載されている
   - **REFERENCE** … 運用慣行(原則税込)や PP の税区分からの税率など、推定を含む
   - 標準化できない → `TAX_BASIS_UNKNOWN` / `TAX_RATE_UNKNOWN`(**正式な差異警告を出さない**)
2. 単位を揃える: amount_unit が UNKNOWN、または換算に使う入数が未確定 → `UNIT_UNKNOWN`。
3. 対象日の適用原価(§3.12 の A → B)が無い → `NO_BASELINE`。
4. |差額| ≥ 50円 **かつ** |差額| ÷ 基準値 ≥ 5%(境界を含む)のとき、双方が CONFIRMED なら **`CONFIRMED_WARN`**、どちらかが REFERENCE なら **`REFERENCE_WARN`**。それ以外は `OK`(関数 `cost_variance()`)。REFERENCE_WARN は CONFIRMED_WARN と同じ確度として扱わない。
5. WARN でも Cost History は変更しない。承認を経た場合だけ、新しい cost_history を作る。

---

## 8. Legacy Mapping 1,211件 生成仕様

- **入力:** legacy_product_snapshot(1,211)・legacy_listing_snapshot・Pricetar CSV・DUPLICATE_PP assessment・棚卸しと同じ JAN 正規化規則。
- **出力:**
  - 既存 P ごとに `assessment(LEGACY_MAPPING)` を**最新1件**(`review_status=UNREVIEWED`、`rule_version=legacy-map/1.0`、`assessed_by=rule:legacy-map/1.0`)。
  - 候補先として PROVISIONAL の PP(`PP-`)/ CP(`CP-`)(origin_key で冪等)。
  - **既存 P 番号は新しい ID に使わない。legacy_mapping の ACTIVE 行は作らない。**

**判定順(先に当たったものを採用)**

| # | 条件 | verdict | confidence | 候補先 |
|---|---|---|---|---|
| 1 | 商品名に 福袋 / ランダム / おまかせ | NEEDS_REVIEW | — | CP(`UNCLASSIFIED`)+ COMPOSITION_CLASS=`HUMAN_REVIEW_REQUIRED`(構成未確認) |
| 2 | 商品名が空 | NEEDS_REVIEW | — | なし |
| 3 | 重複候補グループで、原価の食い違い / 名称に違いの語 / 種別の混在 | NEEDS_REVIEW | — | なし(evidence にグループと食い違いを記録) |
| 4 | 重複候補グループ(全員セット) | COMPOSITION_CANDIDATE | MEDIUM | グループ共有の CP(`dup-group:<G>`)・MERGED |
| 5 | 重複候補グループ(単品同士) | ALIAS_CANDIDATE | MEDIUM | グループ共有の PP・MERGED(**代表 P は決めない**) |
| 6 | 種別=セット | COMPOSITION_CANDIDATE | 名称からセット数が読めれば MEDIUM、読めなければ LOW | CP(`legacy:<P>`) |
| 7 | 店舗付属品(おしぼり / スパウトパウチ / スプーン / おまけ) | COMPOSITION_CANDIDATE | MEDIUM | CP = 本体 PP + ACCESSORY |
| 8 | 種別=単品だが名称がまとめ売り | COMPOSITION_CANDIDATE | LOW | CP(`legacy:<P>`) |
| 9 | JAN 欄に ASIN | PHYSICAL_PRODUCT_CANDIDATE | LOW | PP + IDENTIFIER_LINK=`MISFILED` |
| 10 | 同じ JAN の単品が他にもある | PHYSICAL_PRODUCT_CANDIDATE | LOW | PP + `SHARED_ACROSS_ENTITIES` |
| 11 | JAN があり単品で一意 | PHYSICAL_PRODUCT_CANDIDATE | HIGH | PP(`legacy:<P>`) |
| 12 | JAN なしの単品 | PHYSICAL_PRODUCT_CANDIDATE | MEDIUM | PP(`legacy:<P>`) |

- **JAN 一致だけで同一 PP と判定しない**(#3〜#5 は JAN・ASIN・名称・規格・種別の複合条件)。
- 単品セット共通 JAN(21)は、単品側が #10 / #11、セット側が #6 に入り、evidence に `shared_jan` を記録する。
- **evidence の必須キー:** `legacy_snapshot_ref`・`name`・`kind`・`jan_raw`・`standard_cost`(観測値 ID)・`matched_rule`・`listings`・`duplicate_group`・`set_qty_from_name`・`flags`・`pricetar_cost`。
- **試算(v2.1 データ・参考値):**

| verdict | 件数 | 内訳 |
|---|---|---|
| PHYSICAL_PRODUCT_CANDIDATE | 671 | HIGH 566 / MEDIUM 95 / LOW 10 |
| COMPOSITION_CANDIDATE | 452 | MEDIUM 305 / LOW 147 |
| ALIAS_CANDIDATE | 33 | |
| NEEDS_REVIEW | 55 | 重複の食い違い44・福袋9・商品名なし2 |

---

## 9. Phase 1 実装順序(実装可否は ChatGPT 最終レビュー後に判断)

| 段階 | 内容 | 完了の目安 |
|---|---|---|
| 1 | **空の Shadow DB 上で** DDL(`product_core` + `legacy_ingest`)→ 制約 → 自動テスト(C1〜C13・S1〜S8)→ 安全性確認。**既存 P 1,211件などの本投入は行わない** | 全制約テストが通る(違反が拒否される)。この段階の完了を確認してから次へ進む |
| 2 | operator の初期登録(人が実施) | |
| 3 | 取込: ingest_run + snapshots(sha256 照合・冪等) | 同じ入力の再実行で差分0 |
| 4 | ① listing_group / listing / identifier | 出品テーブル1,430行が listing と1対1 |
| 5 | ② cost_observation(全ソース。標準原価は legacy_product_master) | ソースごとの行数と一致 |
| 6 | ③ LEGACY_MAPPING 1,211件 + PROVISIONAL の PP / CP + IDENTIFIER_LINK | 全既存 P に最新 assessment が1件ずつ |
| 7 | ③ LISTING_REFERENCE(41)・COMPOSITION_CLASS(福袋9)・DUPLICATE_PP・COST_VARIANCE | 件数照合 |
| 8 | ④ Excel 確認表の出力 / 取込と ACTIVE 化 | 改ざん検知・承認者検証・遷移制約が機能する |
| 9 | ⑤ 検証レポート・運用手順書 | §10 を満たす |

---

## 10. Phase 1 完了条件

1. §3 のテーブルと §4 の制約が作られ、制約テストがすべて通る。
2. 取込の前後で正本 Excel・CSV の sha256 が一致している(書き込みゼロ)。
3. 同じ入力で再取込しても行が増えない。
4. 既存 P 1,211件すべてに LEGACY_MAPPING assessment が1件ずつあり、`UNREVIEWED` から開始している。
5. 出品テーブル 1,430件が listing に `legacy_listing_id` で1対1に対応している。
6. ACTIVE・正式原価の行はすべて承認者が `human:<operator_id>`(ACTIVE・対象操作の permission あり)である。
7. `verdict='CONSISTENT_HIGH'` の assessment が、人の review 無しで下流を変えていない(検証クエリで0件)。
8. PROVISIONAL の PP / CP に cost_history / composition_cost_history が無い。商品マスターの標準原価が一括で cost_history に入っていない。
9. Excel 確認表の往復(出力 → 記入 → 取込)で、改ざん行が拒否され、正しい行だけが反映される。
10. Product Core → Excel の書き戻し処理が存在しない(コード上・DB 権限上)。

---

## 11. Phase 2(AIKOS Warehouse)への接続

| 接続点 | Phase 2 で追加するもの | Phase 1 側の変更 |
|---|---|---|
| 在庫形態 | `stock_form`(`SF-`: LOOSE → PP / PREPACKED_KIT → CP / CASE → PP × 入数)。core_entity の種類に `STOCK_FORM` を追加 | **identifier / identifier_link は変更なし** |
| スキャン | §1.4 の解決手順。ケースコード・キットラベル・FNSKU は STOCK_FORM へ紐付け | なし |
| 在庫 | `location`(WAREHOUSE / FBA / IN_TRANSIT / OTHER)・`location_event_rule`・`stock_balance`(lot_id は NULL 可で最初から用意)・`stock_event`(追記のみ) | なし |
| 引当 | `order_line` → listing → **注文日の ACTIVE な listing_reference_history** → PP / CP → stock_form → `allocation`(D9: Kit → バラ、自動解体なし、FBA は対象外) | なし(参照先が期間付きで一意) |
| 出荷原価 | 出荷日の A(composition_cost_history)→ B(構成品 cost_history × 数量) | なし |
| 可変福袋・ロット | Phase 3: `selection_rule`・`shipment_component_actual`・`lot` | なし |

**既存の JAN スキャン棚卸し記録の接続**
- `events/*.csv` の `正規化コード` → identifier。`内部管理ID` + `マスター版`(sha)→ **その時点の ACTIVE な legacy_mapping** で PP / CP に変換する。
- セット P で数えた行(「セット換算待ち」)は、legacy_mapping 先の CP の PREPACKED_KIT 在庫形態として扱う(入数を推測して単品へ換算しない)。
- 仮ID(`T-` + 6桁)→ identifier(`TEMP_ID`)+ PROVISIONAL の PP。紐付けの確認後に legacy_mapping / product_relationship で正式な PP へ接続する。

---

## 12. 未決事項(DDL 設計時・実装判断時に確認)

| # | 事項 | 状態 |
|---|---|---|
| U1 | 原価の税区分 | **決定済み**(D12・§3.11・§3.13) |
| U2 | NUMERIC の精度 | **決定済み**(NUMERIC(18,6) / NUMERIC(7,6)、境界値テスト済み)。会計 / KPI 確定時の最終丸め規則は確定処理の設計時に決める |
| U3 | 確認表の NAS パス | **AIKOS 文書管理設計の改訂で決める**(§7.1)。段階8の前提 |
| U4 | 承認者と権限 | **決定済み**(D15・§3.0b・§3.0c)。最初の管理者の実際の operator_id は人が登録時に決める |
| U5 | PP の tax_category(10% / 8%)の初期値 | 既定は UNKNOWN。AI が商品名から提案し、人が承認する(`PRODUCT_APPROVE`) |

## 13. DDL 実装で仕様から追加・変更した点(2026-09-26)

| 点 | 内容 | 理由 |
|---|---|---|
| `assessment.corrected_value` 列を追加 | 対象(entity)ではなく値(税区分など)を人が修正した記録 | TAX_BASIS の CORRECTED を表せなかった |
| 正式原価の作成時に、観測値の対象の一致を検査 | 直接 / ACTIVE な legacy_mapping / ACTIVE な出品参照 経由のみ | 他の商品の観測値を流用できてしまった |
| 人が確定した税区分と適用税区分の一致を検査 | TAX_BASIS の review 結果と `tax_inclusion_applied` を照合 | 確定済みの判定があれば、どの税区分でも通ってしまった |
| `normalization` の必須キーを確定 | 観測値の写し(observed_amount・tax_inclusion・tax_inclusion_basis・tax_rate)+ 適用値(tax_inclusion_applied・tax_rate_applied・tax_rate_basis_applied)+ unit_divisor + formula | 元データの保全(S8)と換算の再計算を DB で検証するため |
| 関数を追加 | `normalize_observation()`・`cost_variance()`・`effective_composition_cost()`・`resolve_identifier()` | 警告の2段階・A→B の優先順位・スキャンの自動確定禁止を DB で一貫して判定するため |
| 構成品の取り外し | 親の Composition が PROVISIONAL の間だけ可(取込ロールに構成品の DELETE を許可) | 下書きの修正に必要。ACTIVE 後はトリガーで拒否 |
| 楽天出品の一意制約 | `NULLS NOT DISTINCT`(PostgreSQL 15 以降) | SKU 空の楽天出品の重複を防ぐ |
| `pc_owner` | ロールだけ作成し、Phase 1 では権限・所有権を与えない | 所有権の移し替えは本番適用の手順で別途決める |
| 拡張 `btree_gist` | データベース単位で有効化(スキーマは既定) | 期間の重なり禁止(EXCLUDE)に必要 |

## 付記
- 本書は v2.2(凍結)と 2026-09-26 の決定事項 D1〜D11 に基づく。草案(`product_core_phase1_spec_20260925.md`)からの主な変更点:
  - Identifier を「コード登録簿 + 紐付け + 対象の登録簿」に変更
  - ID 形式を `PP-` / `CP-` に変更
  - 税抜原価と観測値の税区分を追加
  - 商品マスター標準原価の一括承認を取りやめ
  - review に保留・入力経路・改ざん検知を追加
  - 承認者を operator 表で管理
- 確定版での追加: 税込・税抜混在の前提と観測値の税区分の根拠(D12)、丸めの分離(D13)、確認表の保存先の扱い(D14)、operator の権限分離(D15)、Identifier A 案の正式採用(D16)、追加安全制約 S1〜S8、DB ロールの分離。
- 根拠となる検証成果物(`Output/`、途中版を含めすべて保持): `pricetar_cost_diff_*`・`product_model_survey_*`・`product_model_v2_validation_*`・`product_model_v21_validation_*`・`product_model_v22_validation_*`・`product_core_phase1_spec_20260925.md`(草案)。
