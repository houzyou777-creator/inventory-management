# PROJECT_CHARTER — MomijiStore OS

> **This document follows FOUNDATION.md.**
> **If any conflict exists, FOUNDATION.md takes precedence.**
>
> **This project follows FOUNDATION.md.**

ステータス: **v0.8 — 承認済み(2026-08-05)**
作成日: 2026-08-05(v0.8 — FOUNDATION.md優先宣言を冒頭へ追加)

> 本憲章は `FOUNDATION.md`(会社の憲法)に従う。
> 文書階層は FOUNDATION → PROJECT CHARTER → ARCHITECTURE → OPERATIONS → IMPLEMENTATION。
> 本憲章とFOUNDATIONが矛盾した場合はFOUNDATIONが優先する。

---

## 0. MomijiStore Philosophy

**目的**

会社ではなく、**会社を動かすOSを作る。**
**AIが運営できる会社を作る。**
人ではなく、**仕組みで利益が出る会社を目指す。**

このプロジェクトのすべての設計判断は、この思想に従う。
「人がいないと回らない」ものを作りそうになったら、立ち止まって設計をやり直す。

## 1. プロジェクトの目的

Momiji Store(EC事業)の運営システム「MomijiStore OS」を、一つのプロダクトとして育てる。

- 在庫管理・商品マスター・分析・自動化を統合した運営基盤を構築する
- 会社の資産(データ・コード・履歴)をUGREEN NASへ集約し、Mac1台への依存を解消する
- 最終的にNAS上のDocker基盤で運営システム本体が自律稼働する状態を目指す
- 事業売却時にも「再現性のある運営システム」として提示できる状態を維持する(5年構想)

## 2. 設計原則

1. **シンプルさを優先する。** 高度な構成より、**長期間運用できる構成**を優先する。売却できる会社とは「誰でも引き継げる構成」の会社である。構成を複雑にする提案は、シンプルな代替案と比較してから採用する
2. **複雑さは必要になった時だけ追加する。最初から追加しない。Always Simple.**
3. **フォルダ構成は設計の結果であり、目的ではない。** 目的は会社全体をシステム化すること。フォルダは`Architecture.md`(会社全体の設計図)が確定してから、その結果として最後に作る
4. **段階移行** — 現段階(Phase3のInfrastructure完成まで)は「実行はMac、資産はNAS」。最終ゴールはNAS上のDocker統合基盤(Claude Code / Git / AI / 在庫管理DB / バックアップ)で、Macは開発端末に徹する
5. **安全性 > 保守性 > 拡張性 > 速度** の優先順位
6. **AIが読み書きできる形式を優先する** — 属人的なファイルより構造化データ(将来的にExcel→DB移行)
7. **Googleスプレッドシート移行を見据える** — VBA依存最小化・数式優先・シート構造/列名/内部ID不変(詳細はCLAUDE.md)
8. **ハードコーディング禁止**、コメントはWHY中心
9. 単一障害点を作らない — 重要データは「Macローカル + GitHub + NAS」の3か所に存在させる

## 3. 命名規則

- ファイル名・フォルダ名は**英語**で統一(既存の日本語ファイル名は互換のため維持)
- バックアップは `名前_backup_YYYYMMDD_内容` 形式、`Backup/` フォルダへ集約
- NAS共有フォルダは番号プレフィックス(`10_Git/` `20_Backup/` …)で用途を明示
- import対象のファイル名(モール由来CSV等)は変更しない

## 4. セキュリティ方針

- NASの認証情報(パスワード・APIキー)をリポジトリ・ドキュメントに**平文で書かない**
- NASへの外部公開(ポート開放)は原則行わない。リモートアクセスはUGREEN Link等の正規経路のみ。変更は都度承認制
- 権限は最小限 — 共有フォルダごとに必要なユーザーにのみ付与
- 認証情報の入力はユーザー自身が行う(Claudeは代行しない)

## 5. バックアップ方針

- 対象: Gitリポジトリ / Excelファイル / SourceData / マニュアル類
- 世代: daily(7世代) / weekly(4世代) / monthly(12世代) / release(無期限)※現行 `_backup/` 構成を踏襲
- 保存先: NAS `20_Backup/` を正、Macローカル `_backup/` を作業用とする
- NASのSnapshot機能を併用する(方式はPhase1で調査済み — `NAS_PHASE15_SURVEY_REPORT.md`)
- **古いバックアップの自動削除は禁止** — 削除候補の提示のみ
- 年1回以上リストア訓練を実施し、復旧手順を `90_System/` に文書化する

## 6. Git運用

**データフロー(確定方針):**

```
Mac(作業場所)→ GitHub(正本)→ NAS Mirror(災害対策)
```

- **GitHubが正本(単一の真実)。** 開発・履歴・PR・リリース管理はすべてGitHub基準
- **NASは災害対策ミラー。** GitHub障害・アカウント喪失に備えた完全複製(Phase3で構築)
- 変更前にコミット(バックアップ兼用)・機能単位でコミット・mainへの直接大規模変更禁止
- SMBマウント越しのgit操作は禁止(破損リスク)
- リリース時は `CHANGELOG.md` 更新とGitHub Release作成を検討

## 7. フェーズ管理

**Phase番号は `Architecture.md` の Roadmap を正とする。**

| フェーズ | テーマ | 内容 | 状態 |
|----------|--------|------|------|
| Phase1 | **NAS** | 現状調査(`NAS_PHASE1_SURVEY_REPORT.md`)+接続確認・開発環境調査(`NAS_PHASE15_SURVEY_REPORT.md`) | 完了 |
| Phase2 | **MomijiStore OS Core** | 会社OSの中核を**論理として**確立する。成果物: `MomijiStore_OS Logical Design v1.0` | ✅ **Logical Design Complete**(2026-08-05) |
| Phase2.5 | **Business Catalog** | 会社の全業務を一覧化(`Business_Catalog.md`)。Infrastructureは Business を実現するために存在するため先に行う | 承認待ち |
| Phase3 | **MomijiStore OS Infrastructure** | 共有フォルダ作成・権限設定・Docker実行基盤・GitミラーNAS構築・バックアップ自動化 | 開始条件待ち |
| Phase4 | **Database** | Excel→PostgreSQL移行(商品マスター・在庫・広告・会計) | 未着手 |
| Phase5 | **AI Agent** | MCP経由でAIが業務を実行(Automation Rule準拠) | 未着手 |
| Phase6 | **Autonomous Company** | 人は判断のみ。仕組みで利益が出る会社 | 未着手 |

各フェーズは「調査・提案 → ユーザー承認 → バックアップ → 実装 → テスト・報告」の順で進める。承認前に次フェーズの実装へ進まない。

**Phase2 開始条件** — 論理設計とInfrastructure実装で必要な条件は異なるため、2つに分離する。

**A. 論理設計開始条件(`MomijiStore_OS Logical Design v1.0`)— ✅ すべて充足(2026-08-05)**

- [x] SMB確認(実機で接続確認済み — SMB 3.1.1・共有`MomijiStore`マウント成功)
- [x] **NAS永続マウント方式の決定** → Finderの「ログイン項目」を正式採用(Always Simple優先。launchd / autofs は将来必要になった時だけ検討)
- [x] Phase番号統一(Architecture.md Roadmapを正として統一済み)
- [x] NAS機種・ストレージ方式の把握(DXP4800 GT / M.2 SSD RAID1 約1.85TB)
- [x] GitHub運用確認済

**B. Phase3(MomijiStore OS Infrastructure)開始条件 — ⬜ 未充足**
※ 共有フォルダ作成・権限設定・Docker・バックアップ自動化はすべてPhase3で実施する

- [ ] RAM容量の確定(現在8GB・将来増設予定 → 増設後の容量と時期)
- [ ] HDD構成の確定(本数・容量・RAIDレベル・空きベイ)
- [ ] Volume設計の確定(SSD/HDDの役割分担・ファイルシステム)

**論理設計はBの確定を待たずに進められる。** RAM・HDD構成はInfrastructure Layerの実装方式に影響するが、データ構造・業務プロセス・正本の定義といった論理設計には影響しないため。

**Phase2 完成条件(5つすべてを満たした時に完成)** — 定義は `MomijiStore_OS_Logical_Design_v1.0.md` §10 を正とする。

1. 会社の動きを誰でも説明できる
2. 実装方法を変更しても設計は変わらない
3. AIが読める
4. 人が読める
5. 5年後でも理解できる

**完成判定はフォルダを作ったかどうかとは無関係である。**

## 8. 承認制ルール(最優先)

以下はすべて**ユーザー承認前の実施を禁止**する:

- 削除(ファイル・フォルダ・設定・バックアップ)
- 移動・リネーム
- **復元**(git履歴・バックアップからの復元を含む — 現在の場所の調査までは可、実施は承認後のみ)
- 上書き(既存ファイル・既存設定・既存Git)
- 新規作成(NAS上のフォルダ・Docker・Git・サービス)
- 権限変更・設定変更
- `_本番用` ファイルの編集

違反しそうな操作が必要になった場合は、内容・影響範囲・戻し方を提示して承認を待つ。
