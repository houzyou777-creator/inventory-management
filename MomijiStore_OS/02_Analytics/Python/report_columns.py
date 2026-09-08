# -*- coding: utf-8 -*-
"""report_columns.py — レポートの列を「名前」で解決する共通処理

2026-09-08、Amazonビジネスレポートから SKU 列が消え、以降の列が1つずつ
左へずれた。固定の列番号で読んでいたため、売上・個数・注文数がすべて
隣の「- B2B」列を指し、8月売上が ¥5,610,675 → ¥227,580 と報告された。
売上が落ちたのではなく、読み取りが壊れていた。

以来「レポートは見出し名で読む」が標準仕様である
(決定事項1 / 02_Analytics/docs/月次KPI更新手順.md)。

設計方針:
- **必須列が1つでも欠けたら異常終了する。** 黙って0や空で埋めない。
  間違った数字が出るより、動かない方がよい
- 欠けたときは**実際の見出し一覧を表示する。** 人がレポート側を見に行けるように
- 任意列は欠けてもよい(例: AmazonのSKU列は2026年8月から消えたが、
  ASINで代替できるため処理は続行できる)
"""
import sys


def resolve(header, required, optional=None, source='レポート'):
    """見出し行から {キー: 列位置} を作る。

    required / optional は {キー: [期待する見出し名, ...]}。
    見出し名が複数あるのは、レポート側の表記ゆれ(全角スペース等)を
    吸収するため。**先に書いたものが優先される。**

    🚫 部分一致で引かないこと。
       「セッション数 - 合計」は「セッション数 - 合計 - B2B」にも含まれる。
       部分一致にすると B2B 列を掴み、上記の事故がそのまま再現する。
    """
    pos = {}
    for i, h in enumerate(header):
        # 同じ見出しが2度出たら先勝ち。後ろは集計列であることが多い
        pos.setdefault(str(h).strip(), i)

    idx, missing = {}, []
    for key, names in required.items():
        hit = next((pos[n] for n in names if n in pos), None)
        if hit is None:
            missing.append(f'  {key} … 期待した見出し: ' + ' / '.join(names))
        else:
            idx[key] = hit

    if missing:
        sys.exit(
            f'❌ {source} に必須列が見つかりません。\n'
            + '\n'.join(missing)
            + '\n\n実際の見出し(' + str(len(header)) + '列):\n'
            + '\n'.join(f'  [{i:2}] {h}' for i, h in enumerate(header))
            + '\n\nレポートの仕様が変わった可能性があります。'
              '見出し名を確認し、スクリプトの期待値を更新してください。')

    for key, names in (optional or {}).items():
        idx[key] = next((pos[n] for n in names if n in pos), None)
    return idx


def get(row, idx, key, default=None):
    """解決済みの位置から値を取り出す。任意列が無い場合は default を返す。"""
    i = idx.get(key)
    if i is None or i >= len(row):
        return default
    v = row[i]
    return default if v is None else v


def report(idx, header, source):
    """解決結果を実行ログへ残す。

    「今回どの列を読んだか」を毎回記録しておく。
    次に事故が起きたとき、いつから列が変わったのかを遡れる。
    """
    print(f'  {source} 列構成: {len(header)}列')
    for key, i in idx.items():
        where = f'[{i:2}] {header[i]}' if i is not None else '(なし — 任意列)'
        print(f'    {key:12} = {where}')
