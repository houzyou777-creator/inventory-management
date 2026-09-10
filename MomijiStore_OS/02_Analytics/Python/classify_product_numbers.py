# -*- coding: utf-8 -*-
"""classify_product_numbers.py — RMS商品番号の意味を分類する(読み取り専用)

⚠️ **このスクリプトはパイプラインへ接続していない。** 分類結果を見るためだけのもの。
   build_kpi_month の原価解決へ組み込むかは、分類の妥当性を確認してから決める。

使い方: python3 classify_product_numbers.py

英樹が確認した意味だけをルールにする。**それ以外へ一般化しない。**
原価の取得可否と販売入数の確認状態は**別々に**持つ
(1販売分の原価が確認できている番号を、入数不明だけを理由に原価まで不明にしない)。

確認済みの意味(2026-09-10 英樹確認):
  4987176309099-4-3516-a  JAN / 販売入数4 / セット全体原価3,516円(掛けない)
  4987176286284-2-814     JAN / 販売入数2 / セット全体原価814円(掛けない)
  SET-8668-4304-a         JAN下4桁8668と4304の商品を各1個。**数字は原価ではない**
  SET-12-242              販売入数12 / 各商品の単価242円 → 1販売あたり2,904円
  7378/1642-a             7378=JAN下4桁 / 1642=**1販売分の原価**(掛けない)。入数は不明
  4901301451217-6-a       6=販売入数。**番号内に原価は無い**
  h = 廃盤 / 末尾a = Amazon共有在庫(無ければ楽天のみ)

⚠️ SET形式は「構成JAN下4桁」と「入数-単価」が同じ形になりうる。
   **桁数や大小で自動確定しない。** 確認済みの対応表にあるものだけ読む。
⚠️ 廃盤(h)があっても 在庫ゼロ・原価ゼロ・即販売停止とは解釈しない。
⚠️ 末尾aが無いことは「楽天のみ販売」の意味であり、入数1や在庫単位までは推定しない。
"""
import re, sys, csv, warnings
warnings.filterwarnings('ignore')
sys.path.insert(0, '/Users/hide0726/Desktop/Claude Code/MomijiStore_OS/02_Analytics/Python')
from openpyxl import load_workbook
import sync_cost_master as S, build_todo_lists as T

# 確認済みのSET形式の対応表。**桁数や大小で自動判定しない**(承認 判定方法2)
SET_KNOWN = {
    'SET-8668-4304': ('構成JAN下4桁', 'JAN下4桁8668と4304の商品を各1個。数字は原価ではない'),
    'SET-12-242':    ('入数-単価',    '販売入数12・各商品の単価242円 → 1販売あたり2,904円'),
}

# 2026-09-11 英樹回答(A-1〜A-9)を反映。**回答された形だけ**をルールにする
#   A-1 N,N/N       カンマ区切り=構成JAN下4桁 / 「/」右=1販売分の原価
#   A-2 N,N/N,N-a   「/」右もカンマ区切り=構成ごとの原価 → 合計が1販売分
#   A-3 N,N,N,N-a   「/」無し=全部JAN下4桁。原価なし
#   A-4 AST-N       アソート商品。原価・入数は番号から読めない
#   A-5 BR-N-a      バリエーション商品。同上
#   A-6 13桁JAN-a   入数1。原価なし
#   A-7 N 単独      基本はJAN下4桁。**ただし原価のこともある** → 自動確定しない
#   A-8 SET-0014/6668.6637  「/」以降が構成JAN下4桁(修正候補。変更はしない)
#   A-9 SET-N-N     意味が未解決のまま

def parse(pn, jan4=None):
    """(区分, 原価, 原価の根拠, 入数, 入数の根拠, 廃盤, 共有在庫, 構成JAN下4桁, 備考)"""
    jan4 = jan4 or {}
    raw = pn.strip()
    s = raw
    shared = s.endswith('-a')                     # 末尾a = Amazon共有在庫
    if shared: s = s[:-2]
    disc = bool(re.search(r'h$', s))              # h = 廃盤(在庫/原価の解釈はしない)
    if disc: s = re.sub(r'h$', '', s)
    s = s.strip('-')

    def R(kind, cost, csrc, n, nsrc, comps, note):
        return (kind, cost, csrc, n, nsrc, disc, shared, comps, note)

    # A-4 / A-5: 接頭辞で意味は確定するが、原価も入数も番号からは読めない
    if s.upper().startswith('AST-'):
        return R('確認済み(原価なし)', None, '番号内に原価が無い', None, '', [],
                 'アソート商品(AST)。原価・入数は別途確認')
    if s.upper().startswith('BR-'):
        return R('確認済み(原価なし)', None, '番号内に原価が無い', None, '', [],
                 'バリエーション商品(BR)。原価・入数は別途確認')

    # SET形式
    if s.upper().startswith('SET-'):
        k = s.upper()
        if k in SET_KNOWN:
            kind, note = SET_KNOWN[k]
            if kind == '入数-単価':
                n = int(k.split('-')[1]); u = int(k.split('-')[2])
                return R('確認済み', n*u, 'SET(入数×単価)', n, 'SET形式', [], note)
            return R('確認済み(原価なし)', None, '数字は原価ではない', None, '',
                     k.split('-')[1:], note)
        # A-8: SET-0014/6668.6637 は「/」以降が構成JAN下4桁
        m = re.fullmatch(r'SET-(\d+)/([\d.]+)', k)
        if m:
            return R('確認済み(原価なし)', None, '番号内に原価が無い', None, '',
                     m.group(2).split('.'), 'A-8形式。「/」以降が構成JAN下4桁(表記の修正候補)')
        return R('意味が複数ある', None, '', None, '', [],
                 'SET形式(A-9)は意味が未解決。対応表に無い')

    # 13桁JAN-入数-原価
    m = re.fullmatch(r'(\d{13})-(\d+)-(\d+)', s)
    if m:
        return R('確認済み', int(m.group(3)), 'セット全体原価(掛けない)',
                 int(m.group(2)), '商品番号', [m.group(1)], '')
    # 13桁JAN-入数
    m = re.fullmatch(r'(\d{13})-(\d+)', s)
    if m:
        return R('確認済み(原価なし)', None, '番号内に原価が無い', int(m.group(2)),
                 '商品番号', [m.group(1)], '単品原価が別途確認できれば入数分を計算できる')
    # A-6: 13桁JANのみ → 入数1(英樹確認)
    if re.fullmatch(r'\d{13}', s):
        return R('確認済み(原価なし)', None, '番号内に原価が無い', 1, '商品番号(A-6)',
                 [s], '13桁JAN単独=入数1(2026-09-11 確認)')

    # A-1 / A-2: 構成JAN下4桁(カンマ区切り) / 原価(カンマ区切りなら合計)
    m = re.fullmatch(r'([\d,]+)/([\d,]+)', s)
    if m:
        comps = [x for x in m.group(1).split(',') if x]
        costs = [int(x) for x in m.group(2).split(',') if x]
        cost = sum(costs)
        why = ('構成ごとの原価を合計(A-2)' if len(costs) > 1 else '1販売分の原価(掛けない)')
        return R('確認済み', cost, why, None, '', comps,
                 f'構成{len(comps)}品。入数はこの番号から確定できない')

    # A-3: カンマ区切りのみ = 全部JAN下4桁。原価なし
    m = re.fullmatch(r'\d+(,\d+)+', s)
    if m:
        return R('確認済み(原価なし)', None, '番号内に原価が無い', None, '',
                 s.split(','), 'A-3形式。全部JAN下4桁。入数は構成数と推定しない')

    # A-7: 数字単独 — 基本はJAN下4桁だが原価のこともある → **自動確定しない**
    if re.fullmatch(r'\d{1,6}', s):
        c = jan4.get(s.zfill(4), []) if len(s) <= 4 else []
        if len(c) == 1:
            return R('要確認(A-7)', None, '', None, '', [s],
                     f'JAN下4桁として商品マスターに一意に該当({c[0]})。ただし原価の可能性も残る')
        return R('要確認(A-7)', None, '', None, '', [],
                 f'JAN下4桁の候補{len(c)}件。原価の可能性もあり自動確定しない')

    return R('未登録形式', None, '', None, '', [], f'確認済みルールに無い形: {s}')


# --- データ読み込み ---
wb = load_workbook(T.RAKUTEN_STOCK, read_only=True, data_only=True)
ws = wb['楽天CSV取込']
hi = next(i for i,r in enumerate(ws.iter_rows(min_row=1,max_row=6,values_only=True),1) if r and r[0]=='JAN')
rows = [(str(r[2]).strip(), r[0], str(r[4] or '')[:44])
        for r in ws.iter_rows(min_row=hi+1, values_only=True) if r and r[2]]
wb.close()

# JAN下4桁 → フルJAN の照合表(商品マスター)
wbm = load_workbook(S.MASTER_FILE, read_only=True, data_only=True)
jan4 = {}
for r in wbm[S.SH_MASTER].iter_rows(min_row=2, values_only=True):
    if r[1]:
        j = str(r[1]).strip()
        jan4.setdefault(j[-4:], []).append(j)
wbm.close()

out = []
for pn, jan, name in rows:
    kind, cost, csrc, n, nsrc, disc, shared, comps, note = parse(pn, jan4)
    # 構成JAN下4桁を商品マスターのフルJANへ照合(承認 判定方法3)。
    # 4桁のものだけ照合し、候補が0件または複数なら保留にする
    unresolved = []
    for c4 in comps:
        if len(c4) == 4:
            cands = jan4.get(c4, [])
            if len(cands) != 1:
                unresolved.append(f'{c4}({len(cands)}件)')
    if unresolved and kind.startswith('確認済み'):
        kind = '保留(JAN照合)'
        note += ' / 照合できない構成JAN下4桁: ' + ' '.join(unresolved)
    out.append([pn, jan, kind, cost if cost is not None else '',
                csrc, n if n is not None else '', nsrc,
                '廃盤' if disc else '', 'Amazon共有' if shared else '楽天のみ',
                ','.join(comps), note, name])

from collections import Counter
c = Counter(x[2] for x in out)
print("■ RMS商品番号の分類(楽天在庫CSV 858件)\n")
for k in ['確認済み','確認済み(原価なし)','保留(JAN照合)','要確認(A-7)','意味が複数ある','未登録形式']:
    print(f"   {k:20} {c.get(k,0):>4}件")
print(f"\n   原価が読める      : {sum(1 for x in out if x[3] != ''):>4}件")
print(f"   販売入数が読める  : {sum(1 for x in out if x[5] != ''):>4}件")
print(f"   両方読める        : {sum(1 for x in out if x[3] != '' and x[5] != ''):>4}件")
print(f"   廃盤(h)           : {sum(1 for x in out if x[7]):>4}件")
print(f"   Amazon共有(-a)    : {sum(1 for x in out if x[8]=='Amazon共有'):>4}件")

OUT='/Users/hide0726/Desktop/Claude Code/MomijiStore_OS/01_InventoryManagement/SourceData/Output'
with open(f'{OUT}/RMS商品番号_分類_20260911.csv','w',encoding='utf-8-sig',newline='') as f:
    w = csv.writer(f)
    w.writerow(['商品番号','JAN','分類','原価','原価の根拠','販売入数','入数の根拠',
                '廃盤','在庫区分','構成JAN下4桁','備考','商品名'])
    w.writerows(out)
print(f"\n   → {OUT}/RMS商品番号_分類_20260911.csv")
