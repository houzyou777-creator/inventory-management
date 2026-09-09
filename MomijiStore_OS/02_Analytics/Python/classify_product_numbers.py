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

def parse(pn):
    """(区分, 原価, 原価の根拠, 入数, 入数の根拠, 廃盤, 共有在庫, 備考)"""
    raw = pn.strip()
    s = raw
    shared = s.endswith('-a')                     # 末尾a = Amazon共有在庫
    if shared: s = s[:-2]
    disc = bool(re.search(r'h$', s))              # h = 廃盤
    if disc: s = re.sub(r'h$', '', s)
    s = s.strip('-')

    # ① SET形式 — 確認済みの対応表にあるものだけ読む
    if s.upper().startswith('SET-'):
        k = s.upper()
        if k in SET_KNOWN:
            kind, note = SET_KNOWN[k]
            if kind == '入数-単価':
                n = int(k.split('-')[1]); u = int(k.split('-')[2])
                return ('確認済み', n*u, 'SET(入数×単価)', n, 'SET形式', disc, shared, note)
            return ('確認済み', None, '数字は原価ではない', None, '', disc, shared, note)
        return ('意味が複数ある', None, '', None, '',  disc, shared,
                'SET形式は「構成JAN下4桁」と「入数-単価」の両方がありうる。対応表に無い')

    # ② 13桁JAN-入数-原価  (4987176309099-4-3516)
    m = re.fullmatch(r'(\d{13})-(\d+)-(\d+)', s)
    if m:
        return ('確認済み', int(m.group(3)), 'セット全体原価(掛けない)',
                int(m.group(2)), '商品番号', disc, shared, '')

    # ③ 13桁JAN-入数  (4901301451217-6)
    m = re.fullmatch(r'(\d{13})-(\d+)', s)
    if m:
        return ('確認済み', None, '番号内に原価が無い', int(m.group(2)), '商品番号',
                disc, shared, '単品原価が別途確認できれば入数分を計算できる')

    # ④ JAN下4桁/原価  (7378/1642 ・ 8507/2283)
    m = re.fullmatch(r'(\d{4})/(\d+)', s)
    if m:
        return ('確認済み', int(m.group(2)), '1販売分の原価(掛けない)',
                None, '', disc, shared, 'JAN下4桁で商品照合が要る。入数はこの番号から確定できない')

    # ⑤ 13桁JANのみ
    if re.fullmatch(r'\d{13}', s):
        return ('未登録形式', None, '', None, '', disc, shared, '13桁JANのみ。原価も入数も無い')

    return ('未登録形式', None, '', None, '', disc, shared, f'確認済みルールに無い形: {s}')

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
    kind, cost, csrc, n, nsrc, disc, shared, note = parse(pn)
    # JAN下4桁の照合(承認 判定方法3)
    m = re.fullmatch(r'(\d{4})/(\d+)', pn.replace('-a','').rstrip('h'))
    if m:
        c = jan4.get(m.group(1), [])
        if len(c) != 1:
            kind = '保留(JAN照合)'
            note += f' / JAN下4桁 {m.group(1)} の候補が {len(c)}件'
    out.append([pn, jan, kind, cost if cost is not None else '',
                csrc, n if n is not None else '', nsrc,
                '廃盤' if disc else '', 'Amazon共有' if shared else '楽天のみ', note, name])

from collections import Counter
c = Counter(x[2] for x in out)
print("■ RMS商品番号の分類(楽天在庫CSV 858件)\n")
for k in ['確認済み','保留(JAN照合)','意味が複数ある','未登録形式']:
    print(f"   {k:16} {c.get(k,0):>4}件")
print(f"\n   原価が読める      : {sum(1 for x in out if x[3] != ''):>4}件")
print(f"   販売入数が読める  : {sum(1 for x in out if x[5] != ''):>4}件")
print(f"   両方読める        : {sum(1 for x in out if x[3] != '' and x[5] != ''):>4}件")
print(f"   廃盤(h)           : {sum(1 for x in out if x[7]):>4}件")
print(f"   Amazon共有(-a)    : {sum(1 for x in out if x[8]=='Amazon共有'):>4}件")

OUT='/Users/hide0726/Desktop/Claude Code/MomijiStore_OS/01_InventoryManagement/SourceData/Output'
with open(f'{OUT}/RMS商品番号_分類_20260910.csv','w',encoding='utf-8-sig',newline='') as f:
    w = csv.writer(f)
    w.writerow(['商品番号','JAN','分類','原価','原価の根拠','販売入数','入数の根拠',
                '廃盤','在庫区分','備考','商品名'])
    w.writerows(out)
print(f"\n   → {OUT}/RMS商品番号_分類_20260910.csv")
