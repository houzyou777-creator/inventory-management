# 月次KPI更新手順

毎月の楽天売上分析(KPI管理シート)の更新フロー。原価は商品マスターに一元管理する。

## 前提

- 対象ファイル: `02_Analytics/SourceData/楽天運営 KPI管理シート.xlsx`
- 原価の正: `01_InventoryManagement/SourceData/商品マスター_単品_v1.0.xlsx`
- スクリプト: `02_Analytics/Python/`(macOS + Microsoft Excel が必要)

## 毎月の手順

```
cd MomijiStore_OS/02_Analytics/Python
```

1. **RMSからCSVをダウンロード**
   データ分析 → SKU別売上 → 対象月で絞ってCSVダウンロード
   (ファイル名例: `202608_SKU_SalesList.csv`)。`02_Analytics/SourceData/` に置く。

2. **月次シートを生成**
   ```
   python3 build_kpi_month.py ../SourceData/202608_SKU_SalesList.csv
   ```
   - シート名はファイル名から自動判定(`202608` → `8月`)
   - バックアップ取得 → シート生成 → Excel再計算 → CSVとの合計突合まで自動
   - 仕入値は 商品マスター → RMS埋め込み → 過去月 の順で解決。未解決は黄色塗り

3. **黄色セルを入力**
   - 仕入値が未解決の行(新商品)
   - 月次経費: 広告費・ポイント費用・クーポン利用額・送料・梱包資材
   - 入力後にExcelで保存すれば限界利益・限界利益率が確定

4. **新商品をマスターへ還流**
   ```
   python3 sync_cost_master.py register 8月
   ```
   - 未登録の商品を商品マスター+出品テーブルへ追加登録
   - 原価不一致が表示されたら内容を確認し、KPI側に合わせてよければ
     `python3 sync_cost_master.py adopt 8月`

5. **在庫集計ツールへ配信**
   ```
   python3 sync_cost_master.py push
   ```
   - 商品マスターの原価を在庫金額集計ツールの「原価マスター」シートへ反映
   - 以後、在庫集計の原価未登録(要確認一覧)が減る

## Amazon(ビジネスレポート)

1. セラーセントラル → ビジネスレポート → 詳細ページ売上・トラフィック(子商品別)
   → 対象月で期間指定してCSVダウンロード
2. ```
   python3 build_amazon_kpi_month.py ../SourceData/BusinessReport-XX-XX-XX.csv 8月
   ```
   - 出力先: `Amazon運営 KPI管理シート.xlsx`(月別シート)
   - 仕入値の照合順: 商品マスター(ASIN→AmazonSKU) → Amazon在庫リスト
3. 販売手数料は暫定一律15%。実額はトランザクションレポート
   (ペイメント → レポートリポジトリ)を入手して置き換える
4. 黄色セル(仕入値・経費)を入力して確定

## トラブル時

- 各ステップは実行前にバックアップを `SourceData/Backup/` に作る
- シートが既に存在するエラー → Excelで該当シートを削除して再実行
- Excelが応答しない → マクロ有効化などのダイアログが出ていないか画面を確認
