# SECURITY — MomijiStore OS セキュリティ運用

> **This document follows FOUNDATION.md.**
> **If any conflict exists, FOUNDATION.md takes precedence.**

ステータス: v1.0
作成日: 2026-08-23
階層: OPERATIONS層

---

## 1. 秘密情報の管理

**原則: 秘密情報をリポジトリに置かない。**(`PROJECT_CHARTER.md` §4)

| 秘密情報 | 保管場所 | Git管理 |
|----------|----------|---------|
| MCP APIトークン | NAS `/volume1/MomijiStore/.env` の `MOMIJI_API_KEYS` | ❌ 対象外 |
| PostgreSQLパスワード | 同上 `MOMIJI_DB_PASSWORD` | ❌ 対象外 |
| NAS認証情報 | macOS Keychain / 人の記憶 | ❌ 対象外 |
| SSH秘密鍵 | `~/.ssh/id_ed25519`(権限600) | ❌ リポジトリ外 |

**リポジトリに置いてよいのはテンプレートのみ** — `docker/.env.example` には**キー名だけ**を書き、値は空にする。

### 守るべきこと

- `.env` は `.gitignore` 済み。**`git add -f` で強制追加しない**
- 秘密情報を**チャット・Issue・コミットメッセージに貼らない**
- **AIに秘密情報を入力させない。** 認証情報の入力は人が行う(FOUNDATION AI Constitution 第9条)
- ログに秘密情報を出さない(実装で担保 — §5)

### 混入していないかの確認

```bash
git log -p --all | grep -iE 'MOMIJI_API_KEY|POSTGRES_PASSWORD' | grep -vE '\$\{|=$|例:'
```

出力があれば履歴に秘密が入っている。その場合は**鍵を即座にローテーション**する(§4)。

---

## 2. Bearer認証の仕様

MCPサーバーは **OAuth 2.1 と同じ Bearer トークン方式**で保護されている(MCP仕様が定める形式)。

```
Authorization: Bearer <TOKEN>
```

| 対象 | 認証 |
|------|------|
| `/mcp` 配下すべて(全MCPツール) | **必須** |
| `GET /health`(監視用) | 不要 |

**MCPのツール呼び出しはすべて `/mcp` を通るため、今後ツールを追加しても自動的に保護される。** 個別の対応は不要。

### 応答仕様

| 状況 | ステータス | ヘッダー / 本文 |
|------|-----------|----------------|
| トークン不正・未提示 | **401 Unauthorized** | `WWW-Authenticate: Bearer error="invalid_token"` / `{"error":"invalid_token",…}` |
| レート制限超過 | **429 Too Many Requests** | `Retry-After: 60` / `{"error":"rate_limit_exceeded",…}` |
| 正常 | 200 | — |

- 比較は `secrets.compare_digest()` による**定数時間比較**(タイミング攻撃対策)
- scheme判定は大文字小文字を区別しない(RFC 6750準拠)
- **トークン未設定ならコンテナは起動しない**(認証なしで公開される事故を防ぐ)

### レート制限

**IP単位で 1分あたり100回**(既定値。`MOMIJI_RATE_LIMIT_PER_MINUTE` で変更可)。認証の**前段**に置いており、総当たり攻撃も抑止する。

※ 現在は直接接続前提のため接続元IPをそのまま使う。**`X-Forwarded-For` は信用しない**(偽装可能なため)。将来リバースプロキシを挟む場合は、信頼できるプロキシからのみ当該ヘッダーを採用する設計に変更すること。

---

## 2.5 データへの書き込み範囲(2026-09-05 追加)

Intelligence Layer の実装により、**MCPサーバーに初めて書き込みが発生した。** 範囲を明確に定める。

| 対象 | 権限 | 根拠 |
|------|------|------|
| NAS上のExcel(`data/products/`) | **読み取りのみ**(`:ro` マウント) | 正本はMac側。ここは検索用スナップショット |
| `intelligence.decisions` / `intelligence.outcomes` | **追記のみ** | Excelに対応物を持たない「DBで生まれるデータ」。ここが正本 |
| `public` スキーマ(Phase4の移行先) | **触らない** | 移行は別の手順で行う |
| その他すべて | なし | — |

**追記のみという制約は、アプリの作法ではなくDBの構造で守られている。**

- `intelligence` の2表には `BEFORE UPDATE OR DELETE` トリガーがあり、UPDATE / DELETE は**必ず例外になる**
- 訂正は `supersedes` に旧IDを入れた**新しい行**で行う(BL-7「記録なき変更を認めない」)
- したがって、**トークンが漏れても過去の判断記録を書き換えたり消したりはできない**(追記はできる)

### スキーマ変更の扱い

`deploy/Migrate.sh` は **`DROP TABLE` / `DROP SCHEMA` / `DROP COLUMN` / `TRUNCATE` / `DELETE FROM` を含むSQLの適用を拒否する**(`lib.sh` の `self_check` と同じ考え方)。
データを削除する変更は、影響範囲と復元方法を提示し、承認を得たうえで個別に実施する。

適用履歴は `public.schema_migrations` に**チェックサム付きで**残る。適用済みファイルが後から変更されていれば、Migrate.sh は適用も無視もせず**停止する**。

---

## 3. APIキーの更新方法

**キーはNASの `.env` にのみ置く。** カンマ区切りで複数指定できる。

```bash
# 1. 新しいキーを生成
openssl rand -hex 32
```

```bash
# 2. .env を編集(既存キーの後ろにカンマ区切りで追加、または置換)
sudo vi /volume1/MomijiStore/.env
```

```bash
# 3. 反映(コンテナ再作成のみ。ビルド不要)
cd /volume1/MomijiStore && sudo docker compose -f momiji-stack.yml up -d momiji-mcp
```

**`down -v` は使わない。** PostgreSQLのボリュームを消してしまう。

---

## 4. ローテーション方法

**無停止で切り替えられる。** 複数キーを同時に有効にできるため。

```
① 新キーを生成し、.env に「新,旧」の順で並べる
       MOMIJI_API_KEYS=<新キー>,<旧キー>
   ↓ コンテナ再作成 → この時点で新旧どちらでも認証が通る
② 各クライアント(Claude Code / ChatGPT / アプリ)の設定を新キーへ更新
   ↓
③ ログで key_id を確認し、旧キーが使われていないことを見届ける
       sudo docker logs momiji-mcp | grep key_id
   ↓
④ .env から旧キーを削除 → コンテナ再作成
       MOMIJI_API_KEYS=<新キー>
```

**③を飛ばさない。** どのクライアントがまだ旧キーを使っているか分からないまま削除すると接続が切れる。ログには**キーの先頭8文字(`key_id`)だけ**が出るので、これで判別する。

**推奨頻度:** 年1回、および担当者の交代時・漏洩の疑いが生じた時。

---

## 5. ログに出さないもの

実装で担保している。

| 出す | 出さない |
|------|---------|
| HTTPメソッド | **トークン全文** |
| パス | パスワード・APIキー |
| 認証成功/失敗 | 環境変数の値 |
| `key_id`(先頭8文字のみ) | 個人情報 |
| レート制限の発生とIP | — |

**認証失敗時はトークンに関する情報を一切記録しない**(攻撃者が送った値も残さない)。

---

## 6. 漏洩時の対応

**発見したら、調査より先に無効化する。**

```
① 直ちに新キーを生成し、.env を新キーのみに置換
       MOMIJI_API_KEYS=<新キー>
   ↓ コンテナ再作成 → 漏洩キーは即座に無効になる
② 影響範囲を確認
       sudo docker logs momiji-mcp | grep -E 'key_id|認証失敗'
       → 漏洩キーの key_id による利用と、不審なIPを確認する
   ↓
③ 各クライアントを新キーへ更新
   ↓
④ Decision Log へ記録(いつ・何が漏れた可能性があるか・対応内容)
   ↓
⑤ 漏洩経路を塞ぐ(Gitに入っていたなら .gitignore 見直し等)
```

### Gitの履歴に入ってしまった場合

**キーの無効化が最優先。** 履歴の書き換え(`git filter-repo` 等)はリモートを巻き込むため、**鍵を無効化したうえで**、必要性を判断してから実施する。**無効化済みのキーは履歴に残っていても実害がない。**

### NASが侵害された疑いがある場合

MCPは `192.168.0.8:8000` にバインドされており**LAN内からのみ到達可能**。外部公開はしていない。LAN内に侵入された疑いがある場合は、キーのローテーションに加えてNASの管理画面パスワードとSMB認証情報も変更する。

---

## 7. OAuth2 / JWT への移行方針

**現在は静的トークン方式。** 単一組織・少数クライアントには十分で、運用が単純(Always Simple)。

移行が必要になる条件は次のいずれか。

- クライアントが増え、**利用者ごとに権限を分けたい**とき
- **外部組織へ公開**するとき
- トークンに**有効期限**を持たせたいとき

### 移行の実装点

`TokenVerifier` インターフェースを定義済みで、**検証ロジックだけを差し替えられる**。

```python
class TokenVerifier(Protocol):
    def verify(self, token: str) -> str | None: ...   # 有効ならkey_id、無効ならNone

class StaticTokenVerifier: ...   # 現在
class JWTTokenVerifier: ...      # 署名検証へ置換
class OAuthTokenVerifier: ...    # イントロスペクションへ置換
```

**ミドルウェア本体・401応答形式・クライアント側の設定はいずれも変更不要。** ヘッダー形式が `Authorization: Bearer` のまま変わらないため、クライアントから見れば透過的に移行できる。

さらに本格的に移行する場合は、SDK標準の `TokenVerifier` + `AuthSettings` を `MCPServer()` へ渡す方式に移せる。現在それを使っていないのは、OAuth用の issuer URL 等の設定が必要で、静的トークン1本には過剰なため。

**移行時は必ずDecision Logへ記録する。**
