# Human Review 保存先(45_HumanReview)— CLAUDE.md 変更案

| 項目 | 内容 |
|---|---|
| 状態 | **変更案(レビュー待ち)**。CLAUDE.md 本体・NAS はまだ変更しない |
| 決定 | D18(2026-09-26): B 案 `45_HumanReview/` を独立した共有として採用。35_Documents は証憑・Document Intelligence の領域として維持 |
| 前提 | AI Assessment → Human Review → Human Approval → Product Core 反映 の**意思決定・承認ワークフローの資産**を置く場所。将来 AIKOS Control Center の承認 UI に置き換える |
| 関連 | 仕様 v1.0 §3.14b(review_batch)・§7.1、Human Review Pilot 設計 |

---

## 1. CLAUDE.md への追記案(「作業範囲」の例外に追加)

現在の「例外(AIKOS Document Intelligence・2026-09-24 承認)」の後ろに、次の節を追加する案。

```markdown
**例外(AIKOS Human Review・<承認日> 承認):** NAS `/Volumes/<共有名>/`(45_HumanReview)配下に限り、
`MomijiStore_OS/05_Infrastructure/`(Product Core)の Human Review 出力・取込ツール経由での読み書きを許可する。
- 書き込み前に SMB マウント(mount 表)と `.aikos_humanreview_marker` を確認し、失敗時は停止する(Mac ローカルへの代替保存禁止)
- ファイルは削除・上書きしない。状態の移動(10_Pending → 20_Reviewed → 30_Imported)はツールが「コピー → sha256 照合 → 台帳記録」で行い、元の位置のファイルは残す(整理は別承認)
- 人が記入するのは 20_Reviewed に置いたファイルだけ。ツールは 10_Pending のファイルを書き換えない
- 実 NAS でのテストは `45_HumanReview/99_Sandbox/` のみで行い、本番の台帳・バッチをテストデータで汚さない
- 45_HumanReview は証憑ではない。DocID を振らず、35_Documents の台帳にも記録しない
- 業務データを含む確認表・台帳は Git に入れない
```

- 共有の作成・アクセス権の設定は UGOS 上で**人が実施**する(Claude は行わない)。
- `<共有名>` は UGOS 上の共有名(NFC 正規化に注意。既存共有と同じ規則)で、作成後に確定する。

---

## 2. フォルダ構造

```
45_HumanReview/
├── .aikos_humanreview_marker     … マウント確認用の目印(内容: 共有名・作成日。人が作成)
├── 10_Pending/<review_batch_id>/  … AIKOS が出力した確認表(記入前)。ツールだけが書く
├── 20_Reviewed/<review_batch_id>/ … 人が記入したファイル(記入済み・取込待ち)
├── 30_Imported/<review_batch_id>/ … 取込が完了したファイルと取込結果(取り込めなかった行の一覧を含む)
├── 80_Ledger/                     … 台帳(バッチ・ファイル sha256・状態遷移・取込結果)。追記のみ
├── 90_Logs/                       … 出力・取込ツールの実行ログ
└── 99_Sandbox/                    … 実 NAS でのテスト専用
```

- 番号は既存の命名規約(2桁 + PascalCase・10 刻み)に合わせる。
- `<review_batch_id>` は DB の `review_batch.review_batch_id`(`RB-` + 8桁)。フォルダ名と DB の束を1対1にする。

## 3. 状態管理(Pending / Reviewed / Imported)

| フォルダ | review_batch.status | 入る条件 | 誰が |
|---|---|---|---|
| 10_Pending | `PENDING` | 確認表を出力し、`file_sha256` を DB と台帳に記録した | 出力ツール(人が起動) |
| 20_Reviewed | `REVIEWED` | 人が記入を終え、ファイルを置いた。ツールが `reviewed_file_sha256` を記録 | 人(配置)+ ツール(記録) |
| 30_Imported | `IMPORTED` | 取込が完了した(取り込めなかった行は一覧として同じフォルダに保存) | 取込ツール(人が起動) |

- 状態は **DB の review_batch が正**。フォルダは写し。食い違いは取込ツールが検出して停止する(Fail Closed)。
- 逆行しない(DB のトリガーで禁止済み)。やり直しは新しい review_batch で行う。
- 取込は `content_hash`(AI 部分の改ざん検知)・reviewer の権限・判断値・遷移を行ごとに検証し、不正な行は取り込まない。

## 4. ファイル命名規則

| 種類 | 形式 | 例 |
|---|---|---|
| 確認表(出力) | `review_<review_type>_<review_batch_id>_<YYYYMMDD-HHMMSS>.xlsx` | `review_LEGACY_MAPPING_RB-00000001_20261001-101500.xlsx` |
| 記入済み | 出力と同じ名前 + `_reviewed_<operator_id>` | `..._reviewed_op_admin.xlsx` |
| 取込結果 | `import_result_<review_batch_id>_<YYYYMMDD-HHMMSS>.xlsx` | |
| 台帳 | `review_ledger.csv`(追記のみ)| |
| ログ | `<tool>_<YYYYMMDD>.log` | |

- ファイル名は英数字・`_`・`-` のみ(CLAUDE.md「ファイル名は英語で統一」)。日本語はシート内に書く。
- ファイル名に商品名・金額などの業務データを入れない。

## 5. アクセス権(UGOS で人が設定する案)

| 主体 | 10_Pending | 20_Reviewed | 30_Imported | 80_Ledger / 90_Logs |
|---|---|---|---|---|
| 出力・取込ツール(Mac から SMB) | 読み書き(作成のみ) | 読み取り | 読み書き(作成のみ) | 追記 |
| レビュー担当者(人) | 読み取り | 読み書き | 読み取り | 読み取り |
| その他(MCP・コンテナ等) | なし | なし | なし | なし |

- UGOS の共有権限は「削除不可」を細かく分けられない可能性がある。その場合は **スナップショット(下記)で削除・上書きからの復旧を担保**し、ツール側の規則(削除・上書きしない)と併用する。
- Docker コンテナ(momiji-postgres・momiji-mcp)にはマウントしない。

## 6. バックアップ対象

| 対象 | 方法 | 頻度 |
|---|---|---|
| 45_HumanReview 全体 | UGOS のスナップショット(共有単位) | 日次・保持 30日(案) |
| 45_HumanReview 全体 | `20_Backup/` への定期コピー(既存のバックアップ手順に追加) | 週次(案) |
| review_batch・assessment(DB) | 既存の PostgreSQL バックアップに含まれる(product_core スキーマ適用後) | 既存どおり |

- 確認表は「人の判断の原本」なので、DB への取込後も削除しない(整理は別承認)。
- 新しい共有を作ったら、**バックアップ・スナップショット対象へ追加したことを確認してから**運用を始める。

## 7. 導入手順(レビュー後)

1. 本変更案のレビュー・修正
2. CLAUDE.md へ §1 の例外を追記(承認後)
3. UGOS で共有 45_HumanReview を作成・権限設定・スナップショット設定(人が実施)
4. 目印ファイル `.aikos_humanreview_marker` を作成(人が実施)
5. 99_Sandbox で出力・取込ツールを検証
6. Human Review Pilot(別文書)を開始

## 8. 決めてほしいこと

| # | 事項 | 案 |
|---|---|---|
| Q1 | スナップショットの頻度・保持期間 | 日次・30日 |
| Q2 | 20_Reviewed への配置を人が手で行うか、ツールが Pending からコピーして人が記入するか | 人が手で配置(ツールは人の記入ファイルを作らない) |
| Q3 | 30_Imported への移動後、20_Reviewed のファイルを残すか | 残す(削除は別承認) |
| Q4 | レビュー担当者のアカウント(UGOS)と operator_id の対応表の置き場所 | 80_Ledger に置かず、DB の operator 表を正とする |
