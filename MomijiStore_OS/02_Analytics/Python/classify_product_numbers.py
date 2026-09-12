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

確認済みの意味(2026-09-12 英樹確認):
  SET-4916-5050-a 等7件    SET=セット品 / 4916・5050=構成JAN下4桁 / a=Amazon共有。
                          **7件を個別に対応表へ登録**(「各区切りが4桁ならJAN下4桁」という
                          一般化は不採用・ChatGPT 2026-09-12)。各商品の数量は別確認
  4901133863318-2-w       末尾 -w = **訳あり品**。原価・価格の扱いは別途確認
  14993499s1-3            14993499s1 = この商品の**型番** / -3 = 数量(販売入数3)
  M-8991/730              カタログが既に無く**解明不能**。自動処理対象外。新資料があるときだけ再確認
  M-0089/1599-a           M は意味なし → 0089/1599-a(JAN下4桁/1販売分の原価/Amazon共有)。この番号だけ

分類名は「形式確認済み」。**番号の形が読めたという意味であり、原価を採用したという意味ではない。**

⚠️ SET形式は「構成JAN下4桁」と「入数-単価」が同じ形になりうる。
   対応表にあるものだけ読む。4桁区切りは**候補の抽出にだけ**使い、未登録の番号を自動確定しない。
⚠️ 構成JANが判明しても各商品の数量は別確認。「各1個」は SET-8668-4304-a だけの確認事項。
⚠️ 廃盤(h)があっても 在庫ゼロ・原価ゼロ・即販売停止とは解釈しない。
⚠️ 訳あり品(w)があっても 原価が違う・同じ とは解釈しない。区別して持つだけ。
⚠️ 末尾aが無いことは「楽天のみ販売」の意味であり、入数1や在庫単位までは推定しない。
"""
import re, sys, csv, warnings
warnings.filterwarnings('ignore')
sys.path.insert(0, '/Users/hide0726/Desktop/Claude Code/MomijiStore_OS/02_Analytics/Python')
from openpyxl import load_workbook
import sync_cost_master as S, build_todo_lists as T

# 確認済みのSET形式の対応表。**桁数や大小で自動判定しない**(承認 判定方法2)
SET_KNOWN = {
    'SET-8668-4304': ('構成JAN下4桁', 'JAN下4桁8668と4304の商品を**各1個**(数量まで確認済み)。数字は原価ではない'),
    'SET-12-242':    ('入数-単価',    '販売入数12・各商品の単価242円 → 1販売あたり2,904円'),
    # 2026-09-12 英樹確認(A-9)。構成JAN下4桁。**各商品の数量は別確認**(各1個と一般化しない)
    'SET-4916-5050': ('構成JAN下4桁', '構成JAN下4桁 4916・5050(2026-09-12 確認)。数量は別確認'),
    'SET-2509-2516': ('構成JAN下4桁', '構成JAN下4桁 2509・2516(2026-09-12 確認)。数量は別確認'),
    'SET-0002':      ('構成JAN下4桁', '構成JAN下4桁 0002(2026-09-12 確認)。数量は別確認'),
    'SET-3414-6317': ('構成JAN下4桁', '構成JAN下4桁 3414・6317(2026-09-12 確認)。数量は別確認'),
    'SET-3018-3001': ('構成JAN下4桁', '構成JAN下4桁 3018・3001(2026-09-12 確認)。数量は別確認'),
    'SET-5931-5924': ('構成JAN下4桁', '構成JAN下4桁 5931・5924(2026-09-12 確認)。数量は別確認'),
    'SET-4879-4893': ('構成JAN下4桁', '構成JAN下4桁 4879・4893(2026-09-12 確認)。数量は別確認'),
}

# 型番を含む商品番号。**型番は英樹が確認したものだけ**(2026-09-12)。英数字混在を型番と一般化しない
MODEL_KNOWN = {
    '14993499s1': 'スクイーズ もっちりバター の型番',
}

# 接頭辞が「意味なし」と英樹が確認した番号。**その番号だけ**接頭辞を外して読む(2026-09-12)。
#   M-0089/1599-a → 0089/1599-a(JAN下4桁 0089 / 1販売分の原価 1,599 / Amazon共有)
#   M接頭辞を一般化しない。M-8991/730 はカタログが無く確認できないので下の UNRESOLVED_KNOWN のまま
PREFIX_IGNORED_KNOWN = {
    'M-0089/1599-a': ('0089/1599-a', 'M接頭辞は意味なし(2026-09-12 英樹確認・この番号のみ)'),
}

# 英樹に確認したが意味が確定できなかった番号。**再確認しない**ために記録する(2026-09-12)
UNRESOLVED_KNOWN = {
    'M-8991/730': 'カタログが既に無く解明不能(2026-09-12 英樹確認)。新しい資料が得られた場合のみ再確認。'
                  'M接頭辞の意味は未確定で、M-0089/1599-a へは一般化しない',
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
#   A-9 SET-N-N     7件を個別登録(2026-09-12 英樹確認)。一般化しない
# 2026-09-12 追加回答:
#   -w              訳あり品(接尾辞)
#   14993499s1-3    型番-数量
#   M-8991/730      確認済・解明不能(自動処理対象外)

def parse(pn, jan4=None):
    """(区分, 原価, 原価の根拠, 入数, 入数の根拠, 廃盤, 共有在庫, 構成JAN下4桁, 備考)

    訳あり品(-w)は 備考 の先頭に「訳あり品(w)。」を付けて返す。
    呼び出し側(build_kpi_month.rms_cost)が9要素で受けているため、戻り値の形は変えない。
    """
    jan4 = jan4 or {}
    raw = pn.strip()
    s = raw
    shared = s.endswith('-a')                     # 末尾a = Amazon共有在庫
    if shared: s = s[:-2]
    wake = bool(re.search(r'-w$', s, re.I))       # -w = 訳あり品(原価の解釈はしない)
    if wake: s = re.sub(r'-w$', '', s, flags=re.I)
    disc = bool(re.search(r'h$', s))              # h = 廃盤(在庫/原価の解釈はしない)
    if disc: s = re.sub(r'h$', '', s)
    s = s.strip('-')

    def R(kind, cost, csrc, n, nsrc, comps, note):
        if wake:
            note = '訳あり品(w)。' + note
        return (kind, cost, csrc, n, nsrc, disc, shared, comps, note)

    # 英樹に確認したが意味が確定できなかった番号 → 再確認の対象から外す
    if raw in UNRESOLVED_KNOWN:
        return R('確認済・解明不能(自動処理対象外)', None, '', None, '', [], UNRESOLVED_KNOWN[raw])
    # 接頭辞が意味なしと確認された番号 → 接頭辞を外した形で読み、備考にその旨を残す
    if raw in PREFIX_IGNORED_KNOWN:
        base, why = PREFIX_IGNORED_KNOWN[raw]
        k = parse(base, jan4)
        return k[:8] + ((why + ('。' + k[8] if k[8] else '')),)

    # A-4 / A-5: 接頭辞で意味は確定するが、原価も入数も番号からは読めない
    if s.upper().startswith('AST-'):
        return R('形式確認済み(原価なし)', None, '番号内に原価が無い', None, '', [],
                 'アソート商品(AST)。原価・入数は別途確認')
    if s.upper().startswith('BR-'):
        return R('形式確認済み(原価なし)', None, '番号内に原価が無い', None, '', [],
                 'バリエーション商品(BR)。原価・入数は別途確認')

    # SET形式
    if s.upper().startswith('SET-'):
        k = s.upper()
        if k in SET_KNOWN:
            kind, note = SET_KNOWN[k]
            if kind == '入数-単価':
                n = int(k.split('-')[1]); u = int(k.split('-')[2])
                return R('形式確認済み', n*u, 'SET(入数×単価)', n, 'SET形式', [], note)
            return R('形式確認済み(原価なし)', None, '数字は原価ではない', None, '',
                     k.split('-')[1:], note)
        # A-8: SET-0014/6668.6637 は「/」以降が構成JAN下4桁
        m = re.fullmatch(r'SET-(\d+)/([\d.]+)', k)
        if m:
            return R('形式確認済み(原価なし)', None, '番号内に原価が無い', None, '',
                     m.group(2).split('.'), 'A-8形式。「/」以降が構成JAN下4桁(表記の修正候補)')
        # 対応表に無いSET形式は**自動確定しない**(ChatGPT 2026-09-12: 4桁区切りの一般化は不採用)。
        #   各区切りが4桁なら「構成JAN下4桁の候補」として備考に残すだけ。分類は未解決のまま
        if re.fullmatch(r'SET(-\d{4})+', k):
            cand = k.split('-')[1:]
            return R('意味が複数ある', None, '', None, '', [],
                     f'SET形式・対応表に無い。候補: 構成JAN下4桁 {cand}(未確定。英樹確認後に個別登録)')
        return R('意味が複数ある', None, '', None, '', [],
                 'SET形式だが対応表に無い')

    # 型番-数量(英樹確認の型番だけ。英数字混在を型番と一般化しない)
    m = re.fullmatch(r'(.+)-(\d+)', s)
    if m and m.group(1) in MODEL_KNOWN:
        return R('形式確認済み(原価なし)', None, '番号内に原価が無い', int(m.group(2)),
                 '商品番号(型番-数量)', [], f'{MODEL_KNOWN[m.group(1)]}(2026-09-12 確認)')

    # 13桁JAN-入数-原価
    m = re.fullmatch(r'(\d{13})-(\d+)-(\d+)', s)
    if m:
        return R('形式確認済み', int(m.group(3)), 'セット全体原価(掛けない)',
                 int(m.group(2)), '商品番号', [m.group(1)], '')
    # 13桁JAN-入数
    m = re.fullmatch(r'(\d{13})-(\d+)', s)
    if m:
        return R('形式確認済み(原価なし)', None, '番号内に原価が無い', int(m.group(2)),
                 '商品番号', [m.group(1)], '単品原価が別途確認できれば入数分を計算できる')
    # A-6: 13桁JANのみ → 入数1(英樹確認)
    if re.fullmatch(r'\d{13}', s):
        return R('形式確認済み(原価なし)', None, '番号内に原価が無い', 1, '商品番号(A-6)',
                 [s], '13桁JAN単独=入数1(2026-09-11 確認)')

    # A-1: 構成JAN下4桁 / **1販売分の原価が明記**(明記型)
    # A-2: 構成JAN下4桁 / 構成ごとの原価(**積み上げ型**)
    #      積み上げ型は 構成商品・数量・各原価 が未確認なら算定しない(2026-09-11 Z-5)。
    #      番号には数量が無いため、現時点では**算定保留**
    m = re.fullmatch(r'([\d,]+)/([\d,]+)', s)
    if m:
        comps = [x for x in m.group(1).split(',') if x]
        costs = [int(x) for x in m.group(2).split(',') if x]
        if len(costs) == 1:
            return R('形式確認済み', costs[0], '1販売分の原価(明記型・掛けない)', None, '', comps,
                     f'構成{len(comps)}品。入数はこの番号から確定できない')
        return R('形式確認済み(積み上げ型・算定保留)', None,
                 f'構成ごとの原価{costs}(積み上げ型)。数量が未確認のため算定しない',
                 None, '', comps,
                 f'構成{len(comps)}品×原価{len(costs)}件。各構成の数量を確認してから合計する')

    # A-3: カンマ区切りのみ = 全部JAN下4桁。原価なし
    m = re.fullmatch(r'\d+(,\d+)+', s)
    if m:
        return R('形式確認済み(原価なし)', None, '番号内に原価が無い', None, '',
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


def main():
    # --- データ読み込み ---
    #   1行 = 1 SKU管理番号。行ごとに 楽天商品管理番号・SKU管理番号・元データ行 を持つ
    #   (同じ商品番号を複数の出品が持つため、商品番号→管理番号 の辞書で代表させない)
    wb = load_workbook(T.RAKUTEN_STOCK, read_only=True, data_only=True)
    ws = wb['楽天CSV取込']
    hi = next(i for i,r in enumerate(ws.iter_rows(min_row=1,max_row=6,values_only=True),1) if r and r[0]=='JAN')
    rows = []
    for rno, r in enumerate(ws.iter_rows(min_row=hi+1, values_only=True), hi+1):
        if r and r[2] not in (None, ''):
            rows.append((rno, S.norm(r[1]), S.norm(r[3]), str(r[2]).strip(), r[0], str(r[4] or '')[:44]))
    wb.close()

    # JAN下4桁 → フルJAN の照合表(商品マスター)
    wbm = load_workbook(S.MASTER_FILE, read_only=True, data_only=True)
    jan4 = {}
    for r in wbm[S.SH_MASTER].iter_rows(min_row=2, values_only=True):
        if r[1]:
            j = str(r[1]).strip()
            jan4.setdefault(j[-4:], []).append(j)
    wbm.close()

    # 出品特定: 管理番号×SKU → 出品テーブル → 内部管理ID(既存の仕組み)。ペアで引けなければ管理番号単独
    _, pair2pid, ctrl2pid = S.load_master()

    out = []
    for rno, ctrl, sku, pn, jan, name in rows:
        kind, cost, csrc, n, nsrc, disc, shared, comps, note = parse(pn, jan4)

        # ── 判定①: 出品を一意に特定できるか(管理番号×SKU→出品テーブル) ──
        pid = (pair2pid.get((ctrl, sku)) or ctrl2pid.get(ctrl)) if ctrl else None
        listing_ok = '可' if pid else '不可'

        # ── 判定②: 構成商品を特定できるか(構成JAN下4桁→商品マスター) ──
        unresolved = []
        for c4 in comps:
            if len(c4) == 4:
                cands = jan4.get(c4, [])
                if len(cands) != 1:
                    unresolved.append(f'{c4}({len(cands)}件)')
        if not comps:
            comp_ok = '—'
        elif not unresolved:
            comp_ok = '可'
        elif len(unresolved) < len([c for c in comps if len(c) == 4]):
            comp_ok = '一部'
        else:
            comp_ok = '不可'

        # ── 判定③: 原価として採用できるか ──
        #   明記型で出品が特定できれば、構成単品がマスターに無いだけでは除外しない(Z-5)。
        #   ただし 対象時点・含有範囲 は別途確認が要る(採用「候補」であって確定ではない)
        if cost is not None and listing_ok == '可':
            cost_use = '候補(対象時点・含有範囲は要確認)'
        elif cost is not None:
            cost_use = '不可(出品を特定できない)'
        elif kind.startswith('形式確認済み(積み上げ型'):
            cost_use = '算定保留(数量未確認)'
        else:
            cost_use = '—(番号に原価が無い)'

        # ── 判定④: 在庫換算に使えるか(構成と入数が確認できて初めて) ──
        stock_use = ('可' if (comp_ok in ('可', '—') and n is not None) else
                     '不可(構成または入数が未確認)')

        out.append([rno, ctrl, sku, pn, jan, kind, listing_ok, pid or '', comp_ok,
                    cost if cost is not None else '', csrc, cost_use,
                    n if n is not None else '', nsrc, stock_use,
                    '廃盤' if disc else '', '訳あり' if note.startswith('訳あり品(w)') else '',
                    'Amazon共有' if shared else '楽天のみ',
                    ','.join(comps), (' / '.join(unresolved) if unresolved else ''), note, name])

    HEAD = ['元データ行','楽天商品管理番号','SKU管理番号','商品番号','JAN','分類','①出品特定','内部管理ID','②構成特定',
            '原価(番号から)','原価の根拠','③原価採用','販売入数','入数の根拠','④在庫換算',
            '廃盤','訳あり','在庫区分','構成JAN下4桁','照合できない構成','備考','商品名']
    I = {h: i for i, h in enumerate(HEAD)}            # 列は名前で引く(位置の決め打ちで数え間違えた)

    from collections import Counter
    c = Counter(x[I['分類']] for x in out)
    print(f"■ RMS商品番号の分類(楽天在庫CSV {len(out)}件)\n")
    kinds = ['形式確認済み','形式確認済み(積み上げ型・算定保留)','形式確認済み(原価なし)','要確認(A-7)',
             '意味が複数ある','確認済・解明不能(自動処理対象外)','未登録形式']
    for k in kinds:
        print(f"   {k:28} {c.get(k,0):>4}件")
    other = {k: v for k, v in c.items() if k not in kinds}
    if other:
        print(f"   ⚠️ 集計表に無い分類: {other}")
    print(f"\n■ 判定を分けた集計")
    for label in ('①出品特定', '②構成特定', '③原価採用', '④在庫換算'):
        cc = Counter(x[I[label]] for x in out)
        print(f"   {label}: " + " / ".join(f"{k}={v}" for k, v in cc.most_common()))
    print(f"\n   原価が番号から読める: {sum(1 for x in out if x[I['原価(番号から)']] != ''):>4}件")
    print(f"   販売入数が読める    : {sum(1 for x in out if x[I['販売入数']] != ''):>4}件")
    print(f"   廃盤(h)             : {sum(1 for x in out if x[I['廃盤']]):>4}件")
    print(f"   訳あり(w)           : {sum(1 for x in out if x[I['訳あり']]):>4}件")
    print(f"   Amazon共有(-a)      : {sum(1 for x in out if x[I['在庫区分']]=='Amazon共有'):>4}件")

    from datetime import date
    OUT='/Users/hide0726/Desktop/Claude Code/MomijiStore_OS/01_InventoryManagement/SourceData/Output'
    path = f'{OUT}/RMS商品番号_分類_{date.today():%Y%m%d}.csv'
    with open(path,'w',encoding='utf-8-sig',newline='') as f:
        w = csv.writer(f)
        w.writerow(HEAD)
        w.writerows(out)
    print(f"\n   → {path}")


if __name__ == '__main__':
    main()
