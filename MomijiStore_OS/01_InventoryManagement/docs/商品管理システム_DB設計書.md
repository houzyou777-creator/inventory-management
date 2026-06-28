# 商品管理システム DB設計書

**作成日:** 2026-06-28  
**対象チャネル:** 楽天・Amazon FBA・Amazon自己発送  
**設計方針:** 内部管理IDをシステム全体の主キーとして全チャネルを統一管理

---

## 前提・設計方針

| 項目 | 内容 |
|---|---|
| 楽天商品管理番号（新形式） | B始まり10桁英数字（ASIN形式）。AmazonのASINと同一値のため `asin` カラムに統一 |
| 楽天商品管理番号（旧形式） | JAN・EAN・独自番号。Amazonに存在しない。`rakuten_item_id` として保持 |
| 内部管理ID | 全チャネルを横断する主キー。形式: `P000001`（連番） |
| Amazon FBA SKU | 形式: `pr_1296040_F_YYYYMMDD_注文番号_連番` |
| Amazon自己発送 SKU | 形式: ランダム（`06-BEWY-859M`）・旧形式（`2022_0816_...`）・JAN番号など混在 |

---

## テーブル一覧

| テーブル名 | 概要 |
|---|---|
| `product_master` | 商品マスター（主キー: `internal_id`） |
| `purchase_order` | 仕入れ管理表（FK: `internal_id`） |
| `inventory` | 在庫管理表（FK: `internal_id`） |
| `inventory_log` | 在庫異動ログ（推奨オプション） |

---

## 1. 商品マスター（`product_master`）

### テーブル定義

| # | カラム名 | 型 | 制約 | 説明 |
|---|---|---|---|---|
| 1 | `internal_id` | VARCHAR(20) | **PK** | 内部管理ID（例: `P000001`） |
| 2 | `product_name` | VARCHAR(500) | NOT NULL | 商品名 |
| 3 | `jan_code` | VARCHAR(20) | | JANコード / EANコード |
| 4 | `asin` | VARCHAR(12) | UNIQUE | Amazon ASIN ＝ **楽天新形式と共通** |
| 5 | `rakuten_item_id` | VARCHAR(100) | | 楽天商品管理番号（**旧形式のみ**） |
| 6 | `rakuten_item_id_type` | CHAR(3) | | `new`=ASIN形式 / `old`=旧形式 |
| 7 | `rakuten_sku` | VARCHAR(100) | | 楽天SKU管理番号（システム連携用SKU番号） |
| 8 | `amazon_sku_fba` | VARCHAR(200) | | Amazon FBA用SKU |
| 9 | `amazon_sku_self` | VARCHAR(200) | | Amazon自己発送用SKU |
| 10 | `brand` | VARCHAR(200) | | ブランド名 |
| 11 | `category` | VARCHAR(100) | | カテゴリ |
| 12 | `supplier_name` | VARCHAR(200) | | 主仕入先名 |
| 13 | `standard_cost` | INTEGER | | 標準仕入原価（基準値） |
| 14 | `condition` | VARCHAR(10) | | `new` / `used` |
| 15 | `status` | VARCHAR(20) | DEFAULT `active` | `active` / `inactive` / `discontinued` |
| 16 | `notes` | TEXT | | 備考 |
| 17 | `created_at` | DATETIME | NOT NULL | 登録日時 |
| 18 | `updated_at` | DATETIME | NOT NULL | 更新日時 |

### インデックス

| インデックス名 | 対象カラム | 用途 |
|---|---|---|
| `idx_pm_jan` | `jan_code` | JANコード検索 |
| `idx_pm_asin` | `asin` | ASIN / 楽天新形式検索 |
| `idx_pm_rakuten_item` | `rakuten_item_id` | 楽天旧形式検索 |
| `idx_pm_sku_fba` | `amazon_sku_fba` | FBA SKU検索 |
| `idx_pm_sku_self` | `amazon_sku_self` | 自己発送SKU検索 |

### DDL（MySQL想定）

```sql
CREATE TABLE product_master (
    internal_id             VARCHAR(20)  NOT NULL,
    product_name            VARCHAR(500) NOT NULL,
    jan_code                VARCHAR(20),
    asin                    VARCHAR(12),
    rakuten_item_id         VARCHAR(100),
    rakuten_item_id_type    CHAR(3)      COMMENT 'new=ASIN形式, old=旧形式',
    rakuten_sku             VARCHAR(100),
    amazon_sku_fba          VARCHAR(200),
    amazon_sku_self         VARCHAR(200),
    brand                   VARCHAR(200),
    category                VARCHAR(100),
    supplier_name           VARCHAR(200),
    standard_cost           INTEGER,
    condition               VARCHAR(10)  COMMENT 'new / used',
    status                  VARCHAR(20)  NOT NULL DEFAULT 'active'
                            COMMENT 'active / inactive / discontinued',
    notes                   TEXT,
    created_at              DATETIME     NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at              DATETIME     NOT NULL DEFAULT CURRENT_TIMESTAMP
                            ON UPDATE CURRENT_TIMESTAMP,
    PRIMARY KEY (internal_id),
    UNIQUE KEY uq_pm_asin (asin),
    INDEX idx_pm_jan        (jan_code),
    INDEX idx_pm_rakuten    (rakuten_item_id),
    INDEX idx_pm_sku_fba    (amazon_sku_fba),
    INDEX idx_pm_sku_self   (amazon_sku_self)
);
```

### 設計ポイント

- **楽天新形式（B始まり）はASINと同一値**のため `asin` カラムで一元管理し、`rakuten_item_id` への二重持ちを排除
- **旧形式**（JAN・`14214143s1` のような独自番号）は `rakuten_item_id` に格納し `rakuten_item_id_type = 'old'` で識別
- `standard_cost` は基準となる仕入原価。実際の仕入価格は `purchase_order.unit_cost` で管理し原価変動を追跡可能にする

---

## 2. 仕入れ管理表（`purchase_order`）

### テーブル定義

| # | カラム名 | 型 | 制約 | 説明 |
|---|---|---|---|---|
| 1 | `purchase_id` | VARCHAR(20) | **PK** | 仕入れID（例: `PO-20260628-001`） |
| 2 | `internal_id` | VARCHAR(20) | FK → product_master | 内部管理ID |
| 3 | `order_date` | DATE | NOT NULL | 発注日 |
| 4 | `arrival_date` | DATE | | 入荷予定日 / 入荷実績日 |
| 5 | `supplier_name` | VARCHAR(200) | | 仕入先名 |
| 6 | `supplier_order_no` | VARCHAR(100) | | 仕入先注文番号 / 伝票番号 |
| 7 | `quantity_ordered` | INTEGER | NOT NULL | 発注数 |
| 8 | `quantity_received` | INTEGER | | 入荷数（検品後確定値） |
| 9 | `unit_cost` | INTEGER | NOT NULL | 仕入単価（税抜） |
| 10 | `tax_rate` | DECIMAL(4,2) | | 消費税率（例: `0.10`） |
| 11 | `total_cost` | INTEGER | | 仕入総額（税抜）= unit_cost × qty |
| 12 | `shipping_cost` | INTEGER | DEFAULT 0 | 送料 |
| 13 | `destination` | VARCHAR(20) | | 入庫先: `self` / `fba` / `both` |
| 14 | `status` | VARCHAR(20) | | `ordered` / `received` / `cancelled` |
| 15 | `invoice_no` | VARCHAR(100) | | 請求書番号 |
| 16 | `payment_date` | DATE | | 支払日 |
| 17 | `notes` | TEXT | | 備考 |
| 18 | `created_at` | DATETIME | NOT NULL | 登録日時 |
| 19 | `updated_at` | DATETIME | NOT NULL | 更新日時 |

### インデックス

| インデックス名 | 対象カラム | 用途 |
|---|---|---|
| `idx_po_internal` | `internal_id` | 商品単位の仕入れ履歴検索 |
| `idx_po_date` | `order_date` | 期間絞り込み |
| `idx_po_status` | `status` | 未入荷・支払い待ちの抽出 |

### DDL

```sql
CREATE TABLE purchase_order (
    purchase_id         VARCHAR(20)     NOT NULL,
    internal_id         VARCHAR(20)     NOT NULL,
    order_date          DATE            NOT NULL,
    arrival_date        DATE,
    supplier_name       VARCHAR(200),
    supplier_order_no   VARCHAR(100),
    quantity_ordered    INTEGER         NOT NULL,
    quantity_received   INTEGER,
    unit_cost           INTEGER         NOT NULL,
    tax_rate            DECIMAL(4,2),
    total_cost          INTEGER,
    shipping_cost       INTEGER         NOT NULL DEFAULT 0,
    destination         VARCHAR(20)     COMMENT 'self / fba / both',
    status              VARCHAR(20)     COMMENT 'ordered / received / cancelled',
    invoice_no          VARCHAR(100),
    payment_date        DATE,
    notes               TEXT,
    created_at          DATETIME        NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at          DATETIME        NOT NULL DEFAULT CURRENT_TIMESTAMP
                        ON UPDATE CURRENT_TIMESTAMP,
    PRIMARY KEY (purchase_id),
    CONSTRAINT fk_po_product FOREIGN KEY (internal_id)
        REFERENCES product_master (internal_id),
    INDEX idx_po_internal   (internal_id),
    INDEX idx_po_date       (order_date),
    INDEX idx_po_status     (status)
);
```

### 設計ポイント

- `destination` で **FBA納品か自社倉庫か**を区別し、後続の在庫更新先チャネルを明確にする
- `unit_cost`（実仕入原価）を毎回記録することで、商品マスターの `standard_cost` との乖離を追跡し原価管理に活用できる
- `quantity_ordered` と `quantity_received` を分離し、**検品ロス・数量差異**を記録可能にする

---

## 3. 在庫管理表（`inventory`）

### テーブル定義

| # | カラム名 | 型 | 制約 | 説明 |
|---|---|---|---|---|
| 1 | `inventory_id` | VARCHAR(30) | **PK** | 在庫ID（例: `P000001-fba`） |
| 2 | `internal_id` | VARCHAR(20) | FK → product_master | 内部管理ID |
| 3 | `channel` | VARCHAR(20) | NOT NULL | `self`（自社）/ `fba` / `rakuten` |
| 4 | `stock_quantity` | INTEGER | DEFAULT 0 | 現在庫数 |
| 5 | `reserved_quantity` | INTEGER | DEFAULT 0 | 受注済・未出荷数 |
| 6 | `available_quantity` | INTEGER | 計算値 | 販売可能数 = stock − reserved |
| 7 | `reorder_point` | INTEGER | | 発注点（これを下回ったら発注） |
| 8 | `reorder_quantity` | INTEGER | | 発注点到達時の標準発注数 |
| 9 | `location` | VARCHAR(100) | | 保管場所・棚番（自社倉庫用） |
| 10 | `last_stocked_at` | DATETIME | | 最終入庫日時 |
| 11 | `last_sold_at` | DATETIME | | 最終出荷日時 |
| 12 | `updated_at` | DATETIME | NOT NULL | 更新日時 |

### インデックス・制約

| 種別 | 対象カラム | 目的 |
|---|---|---|
| UNIQUE | (`internal_id`, `channel`) | 同一商品×チャネルの重複防止 |
| INDEX | `internal_id` | 商品単位の在庫集計 |
| INDEX | `channel` | チャネル単位の在庫一覧 |

### DDL

```sql
CREATE TABLE inventory (
    inventory_id        VARCHAR(30)     NOT NULL,
    internal_id         VARCHAR(20)     NOT NULL,
    channel             VARCHAR(20)     NOT NULL COMMENT 'self / fba / rakuten',
    stock_quantity      INTEGER         NOT NULL DEFAULT 0,
    reserved_quantity   INTEGER         NOT NULL DEFAULT 0,
    available_quantity  INTEGER         GENERATED ALWAYS AS
                        (stock_quantity - reserved_quantity) STORED,
    reorder_point       INTEGER,
    reorder_quantity    INTEGER,
    location            VARCHAR(100),
    last_stocked_at     DATETIME,
    last_sold_at        DATETIME,
    updated_at          DATETIME        NOT NULL DEFAULT CURRENT_TIMESTAMP
                        ON UPDATE CURRENT_TIMESTAMP,
    PRIMARY KEY (inventory_id),
    UNIQUE KEY uq_inv_product_channel (internal_id, channel),
    CONSTRAINT fk_inv_product FOREIGN KEY (internal_id)
        REFERENCES product_master (internal_id),
    INDEX idx_inv_internal  (internal_id),
    INDEX idx_inv_channel   (channel)
);
```

### 設計ポイント

- **チャネルごとに1行**で管理（自社倉庫・FBA・楽天を分離）。1商品が全チャネルに出品されている場合は最大3行
- `available_quantity` は生成カラム（GENERATED）として自動計算し、アプリ側での計算ミスを防止
- `reorder_point` / `reorder_quantity` を持つことで**発注アラート機能**を後から実装しやすくする

---

## 4. 在庫異動ログ（`inventory_log`）— 推奨オプション

在庫の増減履歴をすべて記録し、棚卸しや問い合わせ対応に使用する。

### テーブル定義

| # | カラム名 | 型 | 制約 | 説明 |
|---|---|---|---|---|
| 1 | `log_id` | BIGINT | PK AUTO_INCREMENT | ログID |
| 2 | `internal_id` | VARCHAR(20) | FK | 内部管理ID |
| 3 | `channel` | VARCHAR(20) | NOT NULL | チャネル |
| 4 | `transaction_type` | VARCHAR(30) | NOT NULL | 種別（下記参照） |
| 5 | `quantity_change` | INTEGER | NOT NULL | 変動数（＋入庫 / −出庫） |
| 6 | `quantity_after` | INTEGER | NOT NULL | 変動後在庫数 |
| 7 | `reference_id` | VARCHAR(50) | | 参照ID（purchase_id・注文番号） |
| 8 | `notes` | TEXT | | 備考 |
| 9 | `created_at` | DATETIME | NOT NULL | 記録日時 |

### `transaction_type` の値

| 値 | 説明 |
|---|---|
| `purchase_in` | 仕入れ入庫 |
| `sale_out` | 販売出庫 |
| `return_in` | 返品入庫 |
| `adjustment` | 棚卸し調整 |
| `fba_transfer` | 自社倉庫 → FBA転送 |
| `disposal` | 廃棄・処分 |

---

## チャネル横断の検索クエリ例

### ASINで全チャネルの在庫を確認

```sql
SELECT
    pm.internal_id,
    pm.product_name,
    i.channel,
    i.stock_quantity,
    i.reserved_quantity,
    i.available_quantity
FROM product_master pm
JOIN inventory i ON pm.internal_id = i.internal_id
WHERE pm.asin = 'B0BV9FG7LH';
```

### 楽天旧形式IDで仕入れ履歴を確認

```sql
SELECT po.*
FROM product_master pm
JOIN purchase_order po ON pm.internal_id = po.internal_id
WHERE pm.rakuten_item_id = '4550391011919'
ORDER BY po.order_date DESC;
```

### 発注点を下回った商品を抽出

```sql
SELECT
    pm.internal_id,
    pm.product_name,
    i.channel,
    i.available_quantity,
    i.reorder_point,
    i.reorder_quantity
FROM inventory i
JOIN product_master pm ON i.internal_id = pm.internal_id
WHERE i.available_quantity <= i.reorder_point
  AND pm.status = 'active';
```

### 原価と実仕入価格の乖離チェック

```sql
SELECT
    pm.internal_id,
    pm.product_name,
    pm.standard_cost,
    po.unit_cost          AS actual_cost,
    po.unit_cost - pm.standard_cost AS cost_diff,
    po.order_date
FROM purchase_order po
JOIN product_master pm ON po.internal_id = pm.internal_id
WHERE po.status = 'received'
  AND po.order_date >= DATE_SUB(CURDATE(), INTERVAL 3 MONTH)
ORDER BY cost_diff DESC;
```

---

## 楽天商品管理番号の新旧判定・データ移行ロジック

```sql
-- ① 新形式（B始まり ASIN形式）→ asinカラムに統一、rakuten_item_idは空に
UPDATE product_master
SET
    asin                 = rakuten_item_id,
    rakuten_item_id      = NULL,
    rakuten_item_id_type = 'new'
WHERE rakuten_item_id REGEXP '^B[0-9A-Z]{9}$';

-- ② 旧形式（JAN・独自番号）→ rakuten_item_idをそのまま保持
UPDATE product_master
SET rakuten_item_id_type = 'old'
WHERE rakuten_item_id_type IS NULL
  AND rakuten_item_id IS NOT NULL;
```

---

## テーブル間のリレーション

```
product_master (internal_id) ──< purchase_order (internal_id)
                              ──< inventory      (internal_id)
                              ──< inventory_log  (internal_id)
```

- `product_master` が中心。他3テーブルすべてが `internal_id` で紐付く
- `inventory` は `(internal_id, channel)` のユニーク制約により1商品×1チャネル = 1行を保証
- `inventory_log` はアペンドオンリーで在庫の全変動履歴を保持

---

*この設計書は参考CSVファイル（楽天商品CSV・Amazon FBA CSV・Amazon自己発送CSV）を元に 2026-06-28 に作成*
