# MomijiStore NAS Infrastructure Phase1 — 調査レポート

作成日: 2026-08-05(v1.1 — レビュー反映版)
ステータス: **調査のみ実施(実装・変更・復元は一切なし)**

---

## 1. 調査サマリー(結論)

| 項目 | 結果 |
|------|------|
| ① NAS接続確認 | **Claude Code実行環境(Mac側)からNASへ接続できていない。** NAS本体は構築済み(NAS設定・固定IP・ファイアウォール・UGREEN Link・iPhone 4G接続まで完了) |
| ② 共有フォルダ調査 | 接続未確立のため**保留**(NASの固定IPが分かり次第、Phase1.5で実施) |
| ③ MomijiStore_OS構成 | ローカル(`Desktop/Claude Code`)に存在。合計約40MB。構造は §4 に記載 |
| ④ Git | 正本はGitHub。**Mac(作業)→ GitHub(正本)→ NAS Mirror(災害対策)** の構成とする |
| ⑤ Claude Code | Macのデスクトップアプリとして稼働。現段階は「実行はMac、資産はNAS」、**最終ゴールはNAS中心のDocker統合基盤**(§6 長期構想) |

**Phase1.5(次ステップ):** NAS接続確立 + 開発環境調査(§7)。実装はPhase2以降・承認後。

---

## 2. ① NAS接続確認 — 詳細

**前提(確定事項):** UGREEN NASはセットアップ完了済み。NAS設定・固定IP・ファイアウォール・UGREEN Link・iPhone 4G接続まで動作確認済み。
**問題は「Claude Code実行環境(このMac)からNASへ到達できていない」ことのみ。**

読み取り専用の調査結果(接続試行・認証・書き込みは一切なし):

- **マウント状況:** SMB/AFP/NFSのネットワークボリュームは未マウント(`/Volumes` にはローカルディスクとDMGのみ)
- **LANスキャン(ARPテーブル記載の8台にポート応答確認):** SMB(445)応答は `192.168.0.1`(NEC Atermルーター)のみ。UGOS管理ポート(9999)応答なし
- **Bonjour(mDNS):** NAS系サービスの広告なし
- **ホスト名解決:** `UGREEN.local` 等は解決不可

**接続できない原因の候補(Phase1.5で切り分け):**

| 候補 | 切り分け方法 |
|------|--------------|
| NASのファイアウォールがこのMacからのSMB/探索をブロック | 固定IPへ直接 `ping` / ポート確認 |
| NASが別セグメント(このMacは192.168.0.x) | 固定IPのセグメント確認 |
| mDNS(Bonjour広告)が無効設定 | UGOS設定確認(IP直指定なら問題なし) |
| アクセス経路がUGREEN Link(リレー)前提になっている | LAN内直結(SMB)経路を確立するか方針決定 |

**✅ ユーザーから提供された接続情報(2026-08-05):**

| 項目 | 値 |
|------|-----|
| 固定IP | `192.168.0.8` |
| SMB | 有効 |
| UGREEN Link | 設定済み |
| リモート接続 | iPhoneから4G接続確認済み |
| 管理画面 | `http://192.168.0.8:9999` |

※ 補足: Phase1調査時のARPテーブルに `192.168.0.8`(MAC `6c:1f:f7:ac:c7:32`)が記録されており、**NASはこのMacと同一LAN上に存在していた**。初回スキャンでポート応答が取れなかったのはファイアウォールまたはタイムアウト(1秒)の可能性が高い。Phase1.5で再確認する。

## 3. ② MomijiStore共有フォルダ調査

**保留(接続未確立)。** NASへ到達でき次第、以下を読み取り専用で調査する:
共有フォルダ一覧 / ディレクトリ構造 / 権限(ユーザー・グループ) / 容量・空き / 既存データの有無 / MomijiStore_OSとの差分。

---

## 4. ③ MomijiStore_OS プロジェクト構成(ローカル)

場所: `/Users/hide0726/Desktop/Claude Code`(合計 約40MB、うち `.git` 14MB)

```
Claude Code/
├── CLAUDE.md                     … 開発標準ルール
├── PROJECT_CHARTER.md            … プロジェクト憲章(v0.2ドラフト)
├── .claude/                      … Claude Code設定
├── .git/                         … Gitリポジトリ(14MB)
├── MomijiStore_OS/
│   ├── 01_InventoryManagement/   … 18MB(Excel / VBA / Python / SourceData / import / tests / docs)
│   ├── 02_Analytics/             … 4.1MB(Excel / Python / SourceData / docs)
│   ├── 03_BrandReuse/            … 12KB(ほぼ未着手)
│   ├── 04_Manual/                … 3.4MB(DeveloperGuide / OperationManual / ReleaseNotes)
│   ├── リリースチェックリスト_v1.1.xlsx
│   └── 改善管理表_v1.1.xlsx
└── _backup/                      … daily / weekly / monthly / release(releaseのみ実ファイルあり)
```

**NASとの差分比較:** Phase1.5(NAS接続後)で実施。

### ⚠️ 本番用xlsm 所在調査の結果(重要)

**運用ルール: 復元・移動・削除は一切実施しない。すべてユーザー承認後のみ。今回は調査のみ実施した。**

対象: `MomijiStore_OS/01_InventoryManagement/Excel/在庫管理システム_v1.0_本番用.xlsm`

| 調査項目 | 結果 |
|----------|------|
| ワークスペース内検索 | **なし** |
| Mac全体のSpotlight検索(ホーム配下) | **なし** |
| ゴミ箱(`~/.Trash`) | **空(なし)** |
| gitリポジトリ内 | **あり。** HEADに blob として存在(156KB)。完全に復元可能 |

**git履歴から分かった事実:**
- 初回コミット `c44aa1a`(2026-06-28 10:57)で追加、`c981b07`(同日13:24「プロジェクト構成を整理」)以降**一度も更新されていない**(中身は6/28時点のまま)
- 現存する `在庫管理システム_v1.0.xlsm`(7/8更新)とは**中身が異なる別ファイル**(blobハッシュ不一致)
- 働き木(working tree)上で未コミットの削除状態 = **6/28以降のどこかの時点でディスクから消えた**(git操作ではなくFinder等での削除/移動の可能性。ゴミ箱にもないため、ゴミ箱を空にしたか、ワークスペース外への移動→その後の消失か、外部ドライブへの移動が考えられる)

**現状の安全性:** 中身はgitに完全に残っているため、**データは失われていない**。

**✅ ユーザー決定(2026-08-05):**
- **現時点では復元しない。** Git上に履歴が存在することを確認できたため、バックアップは確保されている
- **Phase1.5終了まで保留とする**
- **復元・削除・移動・リネームは禁止**(理由: 今復元すると、どちらが最新なのか混乱する可能性があるため)

## 5. ④ Git — 現状と方針

**現状:** ローカル `Desktop/Claude Code/.git` + GitHubリモート(`houzyou777-creator/inventory-management`、mainのみ、c7bcc94まで同期済み)。

**採用方針(決定):**

```
Mac(作業場所)
  ↓ push
GitHub(正本・単一の真実)
  ↓ mirror
NAS(災害対策・会社資産としての完全複製)
```

- **GitHubが正本。** 開発・履歴・リリース管理はすべてGitHub基準
- **NASは災害対策ミラー。** GitHub障害・アカウント喪失・サービス終了に備えた完全複製(Phase3で `git clone --mirror` + 定期fetchを構築予定)
- 作業ツリーをNAS上に置くSMB越しgit操作は**禁止**(破損リスク)

## 6. ⑤ Claude Code — 実行場所の評価

**現状:** Macのデスクトップアプリとして稼働(CLIはPATH未登録)。作業対象はローカルの `Desktop/Claude Code`。

### 現状方針(Phase3まで)

**「実行はMac、資産はNAS」。** 開発・Excel運用はMacローカルで行い、NASへはGitミラー・定期バックアップ・データ集約の形で移行する。SMBマウント越しのxlsm編集・git操作はしない。

### 長期構想(最終ゴール)

**NASを「会社を動かすOS」の実行基盤にする。** Macは開発端末に徹し、運営システム本体はNAS上のDockerに集約する:

```
NAS
 └─ Docker
     ├─ Claude Code(自動化ジョブ実行)
     ├─ Git(ミラー/サーバー)
     ├─ AI(MCP Server・Ollama・AI Agent)
     ├─ 在庫管理(DB化: PostgreSQL / Redis)
     └─ バックアップ(Snapshot・世代管理)
```

- 在庫管理・商品マスターは将来的にExcel→DB(PostgreSQL)へ移行し、AIエージェントが直接読み書きできる形にする
- Claude Codeの自動化ジョブ(夜間バッチ・定期集計)をNAS上のDockerで常駐実行
- 実現可否・構築方法は**Phase1.5の開発環境調査**(§7)で評価する

## 7. フェーズ計画(改訂)

| フェーズ | 内容 | 前提 |
|----------|------|------|
| Phase1(今回) | 現状調査・レポート ✅ | — |
| **Phase1.5** | **NAS接続確認 + 開発環境調査(調査のみ・実装禁止)** | NASの固定IP提供 |
| Phase2 | 共有フォルダ設計の承認・作成・権限設定 ※現在の正式名称は **MomijiStore OS Core**(最新のフェーズ定義は `Architecture.md` の Roadmap を正とする) | Phase1.5完了・承認 |
| Phase3 | Gitミラー(NAS)構築・バックアップ自動化・データ集約 | Phase2完了 |
| Phase4 | Docker統合基盤(長期構想の実装)・監視・リストア訓練 | Phase3完了 |

### Phase1.5 調査項目(実装禁止・調査のみ)

NAS接続確立後、UGOS上で以下の運用可否・構築方法・制約を調査する:

1. Docker運用可否(UGOSのDocker対応状況・リソース制約)
2. Container Manager(UGOS標準のコンテナ管理機能)
3. Git Server構築方法(bareミラー vs Gitea等のセルフホスト)
4. Claude Code実行方法(コンテナ内でのCLI実行・認証・ジョブ化)
5. MCP Server構築方法
6. Playwright(ブラウザ自動化コンテナの動作可否)
7. Python(実行環境・バージョン管理)
8. PostgreSQL(在庫管理DB化の受け皿)
9. Redis(キャッシュ/キュー)
10. Ollama(ローカルLLMの動作可否 — NASのCPU/RAM制約評価)
11. バックアップ方式(Mac→NAS、NAS内世代管理、対外バックアップ)
12. Snapshot(UGOSのスナップショット機能・世代設定)
13. AI Agent構成(上記を組み合わせた運営自動化の全体設計)

**追加: Storage Architecture(2026-08-05 ユーザー指示)**

14. RAID構成
15. Volume設計
16. Snapshot設計
17. SSD Cache
18. Docker保存場所
19. Git保存場所
20. PostgreSQL保存場所
21. AIモデル保存場所
22. バックアップ保存場所
23. 将来8TB追加時の拡張方法

**追加: AI Infrastructure(2026-08-05 ユーザー指示・最重要)**

24. Claude Code / MCP / Playwright / Keepa / Python / Docker / Ollama / Redis / PostgreSQL を **Docker Composeで統合管理できるか** — メリット・デメリット・構成図まで作成する

### 危険箇所(今後の注意)

1. **xlsmをNAS上で直接開いて編集しない**(ロック・破損リスク)
2. **SMBマウント越しのgit操作をしない**(インデックス破損リスク)
3. 本番用xlsmの扱い(§4)— ユーザー判断待ち。**復元・削除・移動は承認後のみ**
4. 未コミット変更が複数あり(`Module_Inventory.bas` 等)— NAS作業開始前にコミットで確定推奨
5. バックアップの「取りっぱなし」防止 — Phase4でリストア訓練を実施

### 削除候補(※提示のみ・削除は実施しない)

- `~$在庫管理システム_v1.0.xlsm`(Excelの一時ロックファイル、7/1のもの)
- 各所の `.DS_Store`(macOSメタデータ。`.gitignore` 追加も検討)

---

## 8. ユーザーへの確認事項

1. **NASの固定IPアドレス**を教えてください(UGOS管理画面に表示)→ Phase1.5の接続確認・切り分けを開始します
2. **本番用xlsmの扱い**(§4の選択肢1〜3)をご判断ください — それまで何も操作しません
3. 本レポート(v1.1)と `PROJECT_CHARTER.md`(v0.2)の承認・コミット可否をお願いします
