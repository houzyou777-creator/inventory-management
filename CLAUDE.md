# CLAUDE.md

## Momiji Store AI Development Rules

# プロジェクト概要

このプロジェクトは「もみじストアOS」の開発環境です。

目的

* EC運営の自動化
* 在庫管理システムの構築
* 楽天・Amazon・メルカリとの連携
* AIエージェントによる業務効率化

---

# 作業範囲

Claude Code が作業してよい場所

* Workspace/MomijiOS
* Workspace/AI
* Workspace/Automation
* Workspace/InventorySystem
* Workspace/Rakuten
* Workspace/Amazon
* Workspace/Prompt
* Workspace/Docs

作業禁止

* ~/Desktop
* ~/Documents
* ~/Downloads
* ~/Pictures
* ~/Movies
* ~/Music
* 外付けSSD
* NAS
* iCloud Drive

これらへアクセスする場合は必ず確認すること。

---

# 安全ルール

削除は禁止。

既存ファイルを書き換える前には

・バックアップ
または
・Gitコミット

を行うこと。

大規模変更を行う場合は必ず事前に内容を説明すること。

---

# 実装ルール

コードは可読性を重視する。

コメントを適切に記述する。

ファイル名は英語で統一する。

関数は責務を小さくする。

ハードコーディングを避ける。

---

# 作業手順

新しい機能を作る時は

1. 要件整理
2. 設計
3. 実装
4. テスト
5. ドキュメント更新

の順番で進める。

---

# Git運用

変更前にコミット。

大きな機能ごとにコミット。

mainへ直接大規模変更しない。

---

# 不明点

推測で実装しない。

分からないことは質問する。

---

# AIの役割

Claudeは共同開発者である。

勝手な仕様変更は禁止。

保守性を最優先する。

常に品質、安全性、再利用性を意識して実装する。
