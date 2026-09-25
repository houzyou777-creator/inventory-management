# MomijiStore OS / AIKOS Product Core — 商品データモデル v2.2(設計基準)

| 項目 | 内容 |
|---|---|
| 状態 | **設計基準として凍結(2026-09-25 承認)** |
| 実装仕様 | [Product_Core_Phase1_Specification_v1.0.md](Product_Core_Phase1_Specification_v1.0.md) |
| 根拠 | 読み取り専用検証 v2 → v2.1 → v2.2(2026-09-25)。既存データ(商品マスター 1,211件・出品テーブル 1,430件・Pricetar CSV・棚卸しレーンの共通 JAN 判定)を仮想テーブルへ当てはめて確認した |
| 原則 | **既存 P 番号は削除・統合・変更しない。** 新しい関係構造を追加して移行する(Shadow Model) |

---

## 1. なぜ新しいモデルが要るのか(実データで確認した事実)

| 事実 | 件数(2026-09-25) | 旧構造で起きること |
|---|---|---|
| 同じ物理商品に複数の P 番号(重複候補) | 59 グループ(うち原価の食い違い 16) | 原価・在庫が分散する |
| 商品マスターの「セット」に構成情報が無い | セット種別 372件・構成定義 0件 | セット原価を単品原価から計算・検証できない |
| 出品の販売入数が未登録 | 1,430件中 12件のみ登録 | まとめ売りの原価計算が止まる |
| JAN が物理商品を一意に表さない | 単品同士で共有 10・単品とセットで共有 21・セットのみで複数 P 26 | JAN を読んだだけでは物が決まらない |
| 同じ ASIN に複数の P | 122 ASIN(色違い等のバリエーション 81・紐付け誤りの疑い 41) | 出品 → 商品の対応が一意でない |
| 原価が文字列・慣行に埋め込まれている | Pricetar SKU の `@原価` 53件・楽天商品番号「/仕入値」 | 原価の正本が分散し、税区分も分からない |
| 棚卸しで梱包済みセットがセットのまま数えられる | 共通 JAN 21件 | 「販売構成」と「倉庫での存在形態」を分けないと数量が壊れる |

---

## 2. 基本構造

```
                 ┌──────────────── Identifier(コード登録簿)─ identifier_link ─┐
                 │                                                          ▼
Legacy P ─ legacy_mapping ─► Physical Product ◄── product_relationship ──► Physical Product
(既存 P 番号)        │          ▲   │                (alias / successor / variant)
                    │          │   └─ cost_history(税抜・期間付き)
                    └────► Composition ─ composition_component(PP × 数量)
                              │   └─ composition_cost_history(直接仕入のみ)
                              │   └─ selection_rule(可変構成のみ・Phase 3)
Listing ─ listing_reference_history(期間付き)─► Physical Product | Composition
   └─ listing_group(楽天管理番号・親 ASIN)

[Phase 2] Stock Form(LOOSE / PREPACKED_KIT / CASE)─ stock_balance ─ stock_event ─ allocation
[横断] cost_observation(観測値・元の値のまま)/ assessment(AI 判定 + 人の review)
```

| 層 | 責務 | 持たないもの |
|---|---|---|
| **Physical Product** | 倉庫で扱う最小の物。棚卸・入荷・検品・原価の基準 | JAN(→ Identifier)、原価(→ Cost History) |
| **Identifier** | バーコード・コードそのもの(多対多・有効期間・確からしさ) | 対象そのもの(→ identifier_link) |
| **Product Relationship** | 重複(alias)・後継・バリエーション | 既存 P の重複(→ Legacy Mapping) |
| **Composition** | 販売構成。固定セット・まとめ売り・アソート = 構成品 PP × 数量 | 倉庫での形態(→ Stock Form) |
| **Listing** | 販売口(Amazon SKU / ASIN・楽天管理番号 / SKU) | 原価・在庫 |
| **Stock Form**(Phase 2) | 今、倉庫にどの形で存在するか(バラ・事前梱包キット・ケース) | 販売構成 |
| **Legacy Mapping** | 既存 P 番号が何を表しているか(PP / Composition / 重複の統合先) | 既存 P の変更 |
| **Cost History / Composition Cost History** | 承認済みの正式原価(税抜・有効期間) | 候補値(→ Cost Observation) |
| **Cost Observation** | 観測した原価候補(Pricetar・書類・商品マスター等)を元の値のまま | 正式原価 |
| **Assessment** | AI 判定(追記専用)と人の review 欄 | 人の判断を AI 判定として記録すること |

---

## 3. 設計規則

1. **JAN 一致だけで同一の物理商品と確定しない。** 同一性は JAN・ASIN・名称・規格・種別の複合判断 + 人の承認。
2. **代表 P(canonical P)を選ばない。** 重複した既存 P は、新しい物理商品へ対応付ける(MERGED)。既存 P は残る。
3. **色・サイズ・種類違いは別の物理商品**として保持し、必要なら variant group で束ねる(統合しない)。
4. **販売構成(Composition)と倉庫での存在形態(Stock Form)を混同しない。** 梱包済みセットは「キット在庫」という Stock Form であり、販売構成とは別に数える。
5. **曖昧なコードで在庫形態を自動確定しない。** コード → 対象候補 → 在庫形態候補 → 一意なら確定、複数ならスタッフが選ぶ。
6. **原価は税抜が標準。観測値は元の値のまま**(税込・税抜が混在する前提)。税区分が不明なものを推測で正式原価にしない。
7. **セット原価の優先順位:** A = Composition として直接仕入れた承認済み原価 → 無ければ B = 構成品 PP の承認済み原価 × 数量の合計。A を使うときも B と比べ、差額 ≥ 50円 かつ 差率 ≥ 5% を警告候補にする。
8. **可変アソート(福袋等)** は固定構成と区別し、Selection Rule(候補集合から n 個・条件)で表す。原価は出荷時に記録した実構成から計算する。固定内容なら通常の Composition。
9. **AI 判定と人の確認を混同しない。** AI の「整合性高」は人の確認済みではない。正本変更・P 統合・正式原価・出品参照先の変更は人の承認が必要。
10. **付属品の境界:** 店舗が付けるもの(おしぼり・スパウトパウチ等)= Composition の構成品(役割 ACCESSORY)/ メーカー同梱品 = 物理商品の一部 / 梱包資材 = 固定費。

---

## 4. 在庫と引当(Phase 2 以降の設計基準)

- **On Hand − Allocated = Available**(在庫形態 × 場所ごと)。出品ごとの「販売可能数」は単独で売った場合の最大値であり、出品間で足し合わせない。
- 注文: Listing → その日の参照先(PP / Composition)→ 必要な Stock Form へ引当。キャンセルで引当解除、出荷は引当から Stock Event で出庫。
- **引当の優先順位:** 事前梱包済みキット → バラ在庫。単品注文のためにキットを自動で解体しない(解体は承認後の Stock Event)。
- **場所の種類:** WAREHOUSE / FBA / IN_TRANSIT / OTHER。場所ごとに許すイベントを制御する(例: FBA では梱包・解体しない)。FBA 在庫は自社倉庫注文の引当対象外。
- **ロット・賞味期限:** Phase 3 以降。在庫系の表にロット列(空欄可)を最初から持たせ、既存の行を壊さずに追加できるようにする。

---

## 5. 検証の結果(読み取り専用・2026-09-25)

| 版 | 内容 | 結果 |
|---|---|---|
| v2 | 5層で 374 ケースを当てはめ | 表現不可(追加概念が必要)160 → 追加概念8つを特定 |
| v2.1 | Identifier・Stock Form・Legacy Mapping・Cost Observation・状態 等を追加して再当てはめ | 前回の「表現不可」160 → すべて「表現可 51 / データ不足 105 / 規則未定義 4」 |
| v2.2 | 仮想シミュレーターでシナリオ A〜E(同時販売・キット優先・本体+詰替・可変福袋・直接仕入原価)を実行 | 全ステップで Available ≥ 0・梱包/解体で単品換算総数が不変。表現不能 0 |

**規則未定義として残し、人が決めた事項(2026-09-25〜26)**
- 原価警告: 差額 50円以上 かつ 差率 5%以上。証拠で確認できる比較は CONFIRMED_WARN、推定を含む比較は REFERENCE_WARN
- 引当: キット → バラ。自動解体は禁止(提案のみ)。FBA は別の場所
- 福袋 9件: 構成未確認(HUMAN_REVIEW_REQUIRED)のまま保持し、必要時に固定 / 可変へ分類
- 税区分: 既存の手入力原価は原則税込だが、根拠の弱いデータを無条件に確定しない

---

## 6. Phase 分割

| Phase | 対象 |
|---|---|
| **1** | physical_product・identifier(+ identifier_link・core_entity)・legacy_mapping・product_relationship・composition・composition_component・listing・listing_group・listing_reference_history・cost_history・composition_cost_history・cost_observation・assessment(+ operator・operator_permission・取込ステージング) |
| **2**(AIKOS Warehouse) | location・location_event_rule・stock_form・stock_balance・stock_event・allocation・order_line |
| **3** | selection_rule・shipment_component_actual(可変福袋)・lot(ロット / 賞味期限) |

---

## 付記
- 検証の成果物(Git 管理外 `01_InventoryManagement/SourceData/Output/`): `product_model_survey_*`・`product_model_v2_validation_*`・`product_model_v21_validation_*`・`product_model_v22_validation_*`・`pricetar_cost_diff_*`。
- 本書は設計基準であり、列定義・制約は実装仕様(Phase 1 Specification v1.0)を正とする。
