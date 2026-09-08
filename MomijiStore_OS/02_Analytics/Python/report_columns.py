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
import json
import os
import sys
from datetime import date

# 前月までに見た列構成の記録。**Git で追えるようにコードの隣へ置く。**
# レポートの仕様がいつ変わったのかを後から遡れることが目的(BL-7 記録なき変更を認めない)
LAYOUT_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                           'column_layouts.json')

# 列構成が変わったとき、内容を確認したうえで先へ進めるための逃げ道。
#   ACCEPT_LAYOUT=1 python3 build_xxx.py ...
# これが無いと、レポート側が正当に変わった月に一切処理できなくなる
ACCEPT_ENV = 'ACCEPT_LAYOUT'


def _load_layouts():
    if not os.path.exists(LAYOUT_FILE):
        return {}
    with open(LAYOUT_FILE, encoding='utf-8') as f:
        return json.load(f)


def _save_layouts(data):
    with open(LAYOUT_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
        f.write('\n')


def check_layout(source, header, month=''):
    """前回見た列構成と突き合わせ、違っていたら差分を出して**処理を止める**。

    2026-09-08 の事故(Amazonビジネスレポートから SKU 列が消え、8月売上が
    ¥5,610,675 → ¥227,580 と誤集計された)を二度と起こさないための関門。

    resolve() は「必須列があるか」しか見ない。列が増えただけ・並び替わっただけの
    ときは resolve() を素通りしてしまうが、**それはレポート仕様が変わった合図**で
    あり、人が一度見るべきである。ここで止める。

    止まったあと、内容を確認して問題なければ次のように再実行して記録を更新する:
        ACCEPT_LAYOUT=1 python3 build_xxx.py ...
    """
    header = [str(h).strip() for h in header]
    layouts = _load_layouts()
    rec = layouts.get(source)
    accept = os.environ.get(ACCEPT_ENV, '') not in ('', '0')

    def _record(note):
        layouts[source] = {
            'header': header,
            'columns': len(header),
            'recorded': f'{date.today():%Y-%m-%d}',
            'month': month,
            'history': (rec or {}).get('history', []) + [{
                'date': f'{date.today():%Y-%m-%d}', 'month': month,
                'columns': len(header), 'note': note}],
        }
        _save_layouts(layouts)

    if rec is None:
        print(f'  {source} 列構成: 初回登録 {len(header)}列 — 次回からこれと比較する')
        _record('初回登録')
        return

    old = rec['header']
    if old == header:
        print(f'  {source} 列構成: 前回({rec.get("month") or rec["recorded"]})と同じ '
              f'{len(header)}列 ✓')
        # 履歴は汚さないが「いつ確認したか」は残す。
        # 記録が「7月」のまま止まっていると、確認していないのか変わっていないのか
        # 区別できなくなるため
        rec['last_seen'] = {'date': f'{date.today():%Y-%m-%d}', 'month': month}
        _save_layouts(layouts)
        return

    added = [h for h in header if h not in old]
    removed = [h for h in old if h not in header]
    reordered = not added and not removed

    lines = [f'{"⚠️" if accept else "❌"} {source} の列構成が変わりました。',
             f'  前回: {len(old)}列 ({rec.get("month") or rec["recorded"]})'
             f' → 今回: {len(header)}列']
    if added:
        lines.append(f'  追加された列({len(added)}件):')
        lines += [f'    + [{header.index(h):2}] {h}' for h in added]
    if removed:
        lines.append(f'  削除された列({len(removed)}件):')
        lines += [f'    - [{old.index(h):2}] {h}' for h in removed]
    if reordered:
        lines.append('  列の増減はなく、**並び順だけ**変わりました')
        lines += [f'    [{i:2}] {o}  →  {n}'
                  for i, (o, n) in enumerate(zip(old, header)) if o != n]
    lines.append(f'  現在の見出し({len(header)}列):')
    lines += [f'    [{i:2}] {h}' for i, h in enumerate(header)]

    if accept:
        note = (f'追加{len(added)} / 削除{len(removed)}' if not reordered else '並び順変更')
        lines.append(f'  → {ACCEPT_ENV} が指定されているため、この構成を新しい基準として記録します')
        print('\n'.join(lines))
        _record(note)
        return

    lines += [
        '',
        '  処理を停止しました。数字が静かに壊れるより、動かない方がよいためです。',
        '  1. 上の差分を確認する',
        '  2. 読んでいる列(下に表示される「列構成」)が意図どおりか確かめる',
        f'  3. 問題なければ {ACCEPT_ENV}=1 を付けて再実行する',
        f'     例: {ACCEPT_ENV}=1 python3 <このスクリプト> <引数>',
    ]
    sys.exit('\n'.join(lines))


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
