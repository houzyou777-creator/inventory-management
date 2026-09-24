# AIKOS Document Intelligence Ver.1 — Phase 1 仕様(原本の取り込みと保全)

## 目的
`35_Documents/00_Inbox/` に入れた資料を、原本のまま `10_Archive/` へ保全し、DocID・ハッシュ・台帳・ログで追跡できるようにする。
OCR・分類・Money Forward 連携・JAN 抽出は後続 Phase(本 Phase では実装しない)。

## 配置
| 対象 | 場所 |
|---|---|
| 資料の正本 | NAS `/Volumes/MomijiStore/35_Documents/` |
| コード・設定・仕様・テスト | `MomijiStore_OS/06_DocumentIntelligence/`(Git) |
| マウント異常の記録のみ | `06_DocumentIntelligence/logs/`(Git 管理外。資料は保存しない) |

```
35_Documents/
├── .aikos_mount_marker        目印(中身は config の marker_id)。init だけが作る
├── 00_Inbox/                  投入口(直下のファイルのみ対象)
│   ├── _imported/YYYY-MM/     取り込み済みの投入ファイル(DocID__元名)。削除しない
│   └── _duplicate_candidates/ 重複候補(既存DocID__IntakeEventID__元名)。人が判断する
├── 10_Archive/YYYY/YYYY-MM/DOC-YYYYMMDD-NNNNNN/元ファイル名   原本(読み取り専用)
├── 20_Catalogs/  30_Extracted/  40_Exports/    Phase 2 以降で使用
├── 80_Ledger/document_ledger.csv               台帳
├── 85_Staging/INT-…__元名.aikos_partial        複製途中・失敗した一時ファイル(正式原本ではない)
├── 90_Logs/aikos_intake_YYYYMM.jsonl           処理ログ
│   └── .aikos_ingest.lock                      常設ロックファイル(削除しない)
└── 99_Sandbox/                                 検証用(同じ構成・別の目印・別台帳)
```

## コマンド
```
cd MomijiStore_OS/06_DocumentIntelligence/Python
python3 aikos_intake.py check  --target production            # マウント・目印の確認のみ
python3 aikos_intake.py init   --target production --confirm 本番   # 初回のみ
python3 aikos_intake.py ingest --target production            # 取り込み
python3 aikos_intake.py verify --target production            # 原本の再ハッシュ検証(読み取りのみ)
```
終了コード: 0 正常 / 1 失敗・問題あり / 2 停止(ロック等) / 3 Fail Closed(マウント異常) / 4 verify で未完了・要確認のみあり

## 安全仕様
- **Fail Closed**: 開始時と 1 件ごとの書き込み直前に、①マウントポイントが実在 ②`ismount` ③mount 表で `smbfs` かつ許可された NAS 共有 ④目印ファイルの中身一致 を確認。1 つでも失敗したら何も書かずに停止し、ローカルにエラーだけ記録する。Mac ローカルへの代替保存はしない。
- 書き込み先は doc root 配下に限定(`_assert_inside`)。既存の NAS 領域(data/products・mcp-server・Docker・.env・20_Backup 等)には触れない。
- 原本の保護(Ver.1): ①アプリは上書きしない ②既存 DocID フォルダへは書かない(新規作成が既存なら失敗) ③SHA-256 ④処理ログ ⑤`verify`。SMB 上の読み取り専用属性は付与するが**保全手段とはみなさない**(SMB では無視されうる)。NAS スナップショットは Phase 2 で追加する。
- Inbox からの移動は **Archive 保存 → 保存後ハッシュ一致 → 台帳登録 → ログ記録** がすべて成功した後だけ行う。途中で異常終了したら投入ファイルは Inbox に残る。
- **NAS 上では一切削除しない。** UGOS のごみ箱は SMB 経由の削除を拾って 35_Documents 外の `#recycle` に書き込むため(2026-09-25 実 NAS で確認)、削除を伴う後片付けを持たない。Inbox のファイルも `_imported/` か `_duplicate_candidates/` へ rename で移すだけ(rename はごみ箱に入らない)。
- 同時実行は常設の `90_Logs/.aikos_ingest.lock` への排他ロック(`flock`)で 1 本に限定。ロックファイルは作成後削除しない。ロックはプロセス終了(SIGKILL 等の異常終了を含む)で OS/SMB が解放するため、人手での解除は不要。
- 複製は `85_Staging/` で行い、ハッシュ一致を確認してから DocID を確定し、`10_Archive/` へ rename で移す。失敗した一時ファイルは `85_Staging/` に残し(自動削除しない)、正式原本・DocID・台帳件数には含めない。`verify` が「未完了/要確認」として報告し、ログ `INGEST_FAIL` に `staging_path`・`inbox_path`・`intake_event_id` を残す。元ファイルは Inbox に残るので再取り込みで復旧できる。

## 保全状態と Inbox 整理状態の分離
- 保全(Archive 保存 → ハッシュ一致 → 台帳 → ログ)が完了した後の Inbox 移動だけが失敗しても、正式登録は取り消さない。
- 移動は SMB 由来の一時的エラー(EDEVERR 83・EBUSY・EIO・EAGAIN・ETIMEDOUT)に限り、`intake.move_retry_delays_seconds`(既定 1・2・4 秒)の間隔で有限回だけ再試行する(最大 4 回)。上限到達時は元ファイルを Inbox に残し、`INBOX_MOVE_PENDING`(`move_pending`・`move_error`・`attempts`)を記録する。上書き・削除はしない。
- 次回実行で、同一ハッシュ・同一元ファイル名・`_imported/` 未移動のファイルは「登録済み文書の残存」と判定し、重複候補ではなく `INBOX_CLEANUP` として `_imported/` へ移す(新 DocID なし)。`_imported/` へ移動済みなら再投入とみなし重複候補にする。
- 重複候補の移動が上限に達した場合も Inbox に残し(`DUPLICATE_CANDIDATE` の `move_pending: true`)、次回再試行する。

### SMB の rename エラー(Errno 83 Device error)について(2026-09-25 調査)
- 実運用と同じ流れ(NFD 名で取り込み・移動 → 同名を NFC で再投入 → 重複候補として移動)で、**3 回中 3 回、最初の移動だけが Errno 83 で失敗**した。1 秒後の再試行で成功。
- 切り分け(Sandbox `_smb_diag/`・各 3 回): NFD→NFC 同名 / NFC→NFD 同名 / 書き込み直後(ASCII・NFC・NFD・更新日時設定あり) / 安定後 — **36 回すべて成功し再現せず**。
- 結論: Unicode 正規化単独でも書き込み直後単独でもない。「同名 NFD ファイルを取り込み処理で移動した後の、同名 NFC ファイルの移動」という組み合わせで起きる一時的な SMB クライアント側の状態と推定(断定はしない)。本番コードは再試行と move_pending で安全側に処理する。
- 書き込み中の可能性があるファイル(更新から `min_age_seconds` 未満)は見送る。0 バイトは失敗扱いで Inbox に残す。

## DocID
`DOC-YYYYMMDD-NNNNNN`(取り込み日・6 桁連番)。**AIKOS へ正式登録された文書(Archive に原本がある文書)にだけ発行する。** 同日の最大連番(台帳と 10_Archive の両方から算出)+1。ロックで直列化し、さらに DocID フォルダの排他作成で衝突を防ぐ。

## 重複(同一 SHA-256 の再投入)
新しい Archive 原本・新しい DocID・台帳行は作らない。既存 DocID を参照し、ログに `DUPLICATE_CANDIDATE`(`duplicate_of` = 既存 DocID、`intake_event_id`、`moved_to`)を記録する。投入ファイルは削除せず `_duplicate_candidates/` へ退避する。

## ID の種類
| ID | 形式 | 対象 |
|---|---|---|
| DocID | `DOC-YYYYMMDD-NNNNNN` | 正式登録文書(台帳の主キー) |
| IntakeEventID | `INT-YYYYMMDD-HHMMSS-xxxxxxxx` | Inbox の 1 ファイルの取り込み試行。成功・重複・失敗の関連ログを束ねる(将来の intake_events テーブルの主キー) |
| EventID | `LOG-YYYYMMDD-HHMMSS-xxxxxxxx` | ログ 1 行 |

## 台帳 `document_ledger.csv`(UTF-8 BOM・追記のみ)
正式登録文書だけを 1 DocID = 1 行で記録する。列は将来 DB の `documents` テーブルにそのまま移す想定。会計年度・事業年度は物理フォルダではなく、将来の台帳/DB 側で管理する。**列の追加は末尾のみ・既存列の変更禁止**。

| 列 | 内容 |
|---|---|
| schema_version | 台帳の版(1) |
| doc_id | DocID |
| status | `ARCHIVED`(将来 VOID 等を追加しうる) |
| ingested_at | 取込日時(ISO 8601・+09:00) |
| original_filename | 元ファイル名(NFC 正規化) |
| extension / file_size / sha256 | 拡張子・バイト数・SHA-256 |
| archive_path | doc root からの相対パス |
| intake_event_id | 登録時の取り込み試行 ID(ログと突き合わせる) |
| run_id | 実行 ID(ログと突き合わせる) |
| doc_type / note | Phase 2 以降で使用(空) |

## ログ `90_Logs/aikos_intake_YYYYMM.jsonl`
イベント: `RUN_START` `INGEST_OK` `INBOX_MOVED` `INBOX_MOVE_PENDING` `INBOX_CLEANUP` `MOVE_RETRY` `INGEST_FAIL` `DUPLICATE_CANDIDATE` `SKIPPED` `RUN_LOCKED` `RUN_END`。全行に `event_id`、ファイル単位の行に `intake_event_id` を持つ。
`MOUNT_ERROR` は NAS に書けないため `06_DocumentIntelligence/logs/aikos_local_errors.jsonl` に記録する。

## 日本語ファイル名
NAS(SMB)は NFC、Mac(APFS)は NFD で返すことがあるため、台帳・保存名は NFC に統一する。重複判定は名前ではなく SHA-256 で行う。

## テスト
- 単体: `python3 -m unittest discover -s MomijiStore_OS/06_DocumentIntelligence/tests -v`(一時フォルダを NAS に見立てる。実 NAS には触れない)
- 実 NAS: `python3 tests/nas_sandbox_e2e.py`(**99_Sandbox 専用**。本番 doc root には書かない)。更新日時を人工的に変更せず、SMB 上で安定したことを確認してから取り込む。L(前回残存)・A〜F・G 同時実行・H SIGKILL 後のロック解放・I partial・R 移動リトライ/上限・S SMB 切り分け
