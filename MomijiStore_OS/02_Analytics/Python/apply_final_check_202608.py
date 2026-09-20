# -*- coding: utf-8 -*-
"""apply_final_check_202608.py — 8月最終確認(英樹 F-01〜F-16・ChatGPT承認 2026-09-21)を8月KPIへ反映する

順序: バックアップ → 反映 → Excelで再計算 → 再読込 → 数値照合 → 結果報告。想定外はそこで停止(終了コード1)。

反映する内容:
    F-01〜03  8月KPIの L列(仕入値)を変更(楽天 IROKA 7,164→7,536 / Amazon B0HD7N75QV 2,642→2,862 / B0HB36RZPY 1,152→1,232)
    F-04〜12  数値は変えない。「英樹確認済み(2026-09-21)・現在値・一次根拠未照合」と P列に記録し、塗りを黄色(未確認)から水色(確認済み・一次根拠未照合)へ
    F-13/14   経費 0 を「確認済みの0円」に(黄色を外し P列に理由)。Amazon 梱包資材は配賦注記(主として楽天側へ計上)
    Amazon    チャネル全体KPI(実額手数料 ¥759,756)を **Q:R 列**に別管理で置く(A〜O列は触らない=ルール56。行の挿入もしない)
    共有在庫  例外一覧CSV の F-15/F-16 に承認者・承認日を入れる(元データの商品番号は変えない)
    在庫      build_stock_summary.py を本番実行(別スクリプト。本スクリプトの後に実行)

使い方:
    python3 apply_final_check_202608.py --target copy          作業用フォルダの複製で全手順を通す(本番は読むだけ)
    python3 apply_final_check_202608.py --target production --confirm 本番
"""
import argparse
import csv
import os
import shutil
import sys
import warnings
from datetime import date, datetime

warnings.filterwarnings('ignore')
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import excel_bridge as XB
import sync_cost_master as S

BASE = '/Users/hide0726/Desktop/Claude Code/MomijiStore_OS'
SD = BASE + '/02_Analytics/SourceData'
PROD = {'楽天': SD + '/楽天運営 KPI管理シート.xlsx', 'Amazon': SD + '/Amazon運営 KPI管理シート.xlsx',
        'exc': BASE + '/01_InventoryManagement/SourceData/共有在庫_例外一覧.csv'}
COPY_DIR = '/private/tmp/claude-502/-Users-hide0726-Desktop-Claude-Code/c6dbcd1e-f559-44f8-8c0a-cae8afcb1973/scratchpad/final_check_copy'
COPY = {'楽天': COPY_DIR + '/楽天運営 KPI管理シート.xlsx', 'Amazon': COPY_DIR + '/Amazon運営 KPI管理シート.xlsx', 'exc': COPY_DIR + '/共有在庫_例外一覧.csv'}
OUT = BASE + '/01_InventoryManagement/SourceData/Output'
MONTH = '8月'
DAY = '2026-09-21'
BLUE = PatternFill('solid', fgColor='DDEBF7')     # 英樹確認済み・一次根拠未照合
NO_FILL = PatternFill(fill_type=None)
AMZ_FEE_ACTUAL = 759756                            # 2026AugMonthlyTransaction.csv 販売手数料+FBA手数料+その他(引継ぎ150)

# F-01〜F-03: L列を変更。(チャネル, 識別子, SKU, 現在値, 新値, 記録)
CHANGES = [
    ('楽天', 'b0915ngd2j-6', 'b0915ngd2j-6', 7164, 7536, 'F-01 英樹確認 2026-09-21: 1個 1,256円 × 6個セット = 7,536円(8月実績KPI原価の確認。商品マスターは変更しない。一次根拠未照合)'),
    ('Amazon', 'B0HD7N75QV', '', 2642, 2862, 'F-02 英樹確認 2026-09-21: 本体722円＋詰め替え2,140円 = 2,862円(仕入先により変動あり。8月実績KPI原価の確認。一次根拠未照合)'),
    ('Amazon', 'B0HB36RZPY', '', 1152, 1232, 'F-03 英樹確認 2026-09-21: 1,232円(仕入先により変動あり。8月実績KPI原価の確認。一次根拠未照合)'),
]
# F-04〜F-12: 値は変えない。記録だけ
CONFIRMS = [
    ('楽天', 'b0hd77wvwr', 'b0hd77wvwr', 2395, 'F-04 英樹確認 2026-09-21: 現在値で正しい(補足「セット販売価格」= 用語混在に注意: 原価の単位はセット1販売分の意味と解釈、販売価格ではない。一次根拠未照合)'),
    ('楽天', 'b0gvxq7xpv', 'b0gvxq7xpv', 768, 'F-05 英樹確認 2026-09-21: 現在値で正しい(単品価格。一次根拠未照合)'),
    ('楽天', 'b0g4mfq7cw-2', 'b0g4mfq7cw-2', 3376, 'F-06 英樹確認 2026-09-21: 現在値で正しい(単価1,688円×2。一次根拠未照合)'),
    ('楽天', 'b0d6g9rw9y-2', 'b0d6g9rw9y-2', 3376, 'F-07 英樹確認 2026-09-21: 現在値で正しい(単価1,688円×2。一次根拠未照合)'),
    ('楽天', 'b0f9dzl2mg', 'b0f9dzl2mg', 1316, 'F-08 英樹確認 2026-09-21: 現在値で正しい(単価658円×2。一次根拠未照合)'),
    ('Amazon', 'B0HGY7NBDJ', '', 1406, 'F-09 英樹確認 2026-09-21: 現在値で正しい(Amazon在庫リストの単価1,178との差は残す。在庫リストは変更しない。一次根拠未照合)'),
    ('楽天', 'b0dfyjnwjt', 'b0dfyjnwjt', 1856, 'F-10 英樹確認 2026-09-21: 現在値で正しい(単価928円×2。一次根拠未照合)'),
    ('楽天', 'b0ckj4fcmh-2', 'b0ckj4fcmh-2', 3264, 'F-11 英樹確認 2026-09-21: 現在値で正しい(単価1,632円×2。楽天在庫リストの単価3,770との差は残す。在庫リストは変更しない。一次根拠未照合)'),
    ('楽天', '4550391012541', '4550391012541', 1530, 'F-12 英樹確認 2026-09-21: 現在値で正しい(単価255円×6。一次根拠未照合)'),
]
EXPENSES = [  # (チャネル, ラベル, 現在値, 記録)
    ('Amazon', '梱包資材', 0, 'F-13 英樹確認 2026-09-21: 確認済みの0円。理由=梱包資材費は主として楽天側へまとめて計上している(Amazonで未使用の意味ではない)。チャネル比較時は配賦注記として扱う'),
    ('楽天', 'クーポン利用手数料', 0, 'F-14 英樹確認 2026-09-21: 確認済みの0円。理由=手数料発生条件(利用数)に達していない'),
]
CHANNEL_KPI = [  # Amazon Q:R 列(P列以降=自動記録の領域)。行は経費ブロックに揃える
    ('■ チャネル全体 経営KPI(手数料はトランザクション実額総額)', None),
    ('手数料 実額総額(トランザクション)', AMZ_FEE_ACTUAL),
    ('粗利(チャネル・実額手数料)', '=I{t}-R{r1}-M{t}'),
    ('限界利益(チャネル・実額手数料)', '=R{r2}-N{exp}-N{ship}'),
    ('限界利益率(チャネル・実額手数料)', '=IF(I{t}=0,"",R{r3}/I{t})'),
    ('※ 商品別のK列合計(実額配賦+15%推定)は商品分析用として N列に残す。未配賦53SKU ¥90,017 は暫定', None),
]


def backup(path, label):
    d = os.path.join(os.path.dirname(path), 'Backup')
    os.makedirs(d, exist_ok=True)
    base, ext = os.path.splitext(os.path.basename(path))
    dst = os.path.join(d, f'{base}_backup_{datetime.now():%Y%m%d_%H%M%S}_{label}{ext}')
    if os.path.exists(dst):
        sys.exit(f'❌ バックアップ名が既にあります: {dst}')
    shutil.copy2(path, dst)
    return dst


def find_row(ws, ch, ident, sku):
    """8月シートで出品の行を1つに特定する(楽天=管理番号×SKU、Amazon=ASIN)。複数なら停止。"""
    hits = []
    for r in range(8, ws.max_row + 1):
        if ws.cell(r, 3).value is None:
            break
        if ch == '楽天':
            if S.norm(ws.cell(r, 3).value) == S.norm(ident) and S.norm(ws.cell(r, 5).value) == S.norm(sku):
                hits.append(r)
        else:
            if S.norm(ws.cell(r, 3).value) == S.norm(ident):
                hits.append(r)
    if len(hits) != 1:
        sys.exit(f'❌ {ch} {ident}/{sku} の行が一意でない: {hits}')
    return hits[0]


def block_rows(ws):
    """合計行と M列ラベル→行"""
    r = 8
    while ws.cell(r, 3).value is not None:
        r += 1
    labels = {}
    for rr in range(r + 1, ws.max_row + 1):
        lab = ws.cell(rr, 13).value
        if lab and str(lab).strip() not in labels:
            labels[str(lab).strip()] = rr
    return r, labels


def formulas_of(path):
    wb = load_workbook(path)
    ws = wb[MONTH]
    return {c.coordinate: c.value for row in ws.iter_rows() for c in row if isinstance(c.value, str) and c.value.startswith('=')}


def apply_kpi(P, log):
    fx_before = {ch: formulas_of(P[ch]) for ch in ('楽天', 'Amazon')}
    wbs = {ch: load_workbook(P[ch]) for ch in ('楽天', 'Amazon')}
    # F-01〜03
    for ch, ident, sku, cur, new, note in CHANGES:
        ws = wbs[ch][MONTH]
        r = find_row(ws, ch, ident, sku)
        c = ws.cell(r, 12)
        if c.value != cur:
            sys.exit(f'❌ {ch} {ident} 行{r} の現在値が想定と違う: {c.value} ≠ {cur}(停止)')
        c.value = new
        c.fill = BLUE
        ws.cell(r, 16).value = note
        log.append(('原価変更', ch, ident, f'L{r}', cur, new, note))
        print(f'   {ch} 行{r} L: {cur:,} → {new:,}')
    # F-04〜12
    for ch, ident, sku, cur, note in CONFIRMS:
        ws = wbs[ch][MONTH]
        r = find_row(ws, ch, ident, sku)
        c = ws.cell(r, 12)
        if c.value != cur:
            sys.exit(f'❌ {ch} {ident} 行{r} の現在値が想定と違う: {c.value} ≠ {cur}(停止)')
        c.fill = BLUE
        ws.cell(r, 16).value = note
        log.append(('原価確認(値不変)', ch, ident, f'L{r}', cur, cur, note))
    # 経費
    for ch, lab, cur, note in EXPENSES:
        ws = wbs[ch][MONTH]
        t, labels = block_rows(ws)
        r = labels[lab]
        c = ws.cell(r, 14)
        if c.value != cur:
            sys.exit(f'❌ {ch} {lab} N{r} の現在値が想定と違う: {c.value} ≠ {cur}(停止)')
        c.fill = NO_FILL
        ws.cell(r, 16).value = note
        log.append(('経費確認(0円確定)', ch, lab, f'N{r}', cur, cur, note))
        print(f'   {ch} {lab} N{r}: 0 = 確認済み0円')
    # Amazon チャネル全体KPI(Q:R)
    ws = wbs['Amazon'][MONTH]
    t, labels = block_rows(ws)
    r0 = labels['広告費']                              # 経費ブロックの先頭行に揃える
    exp_total, ship_total = labels['合計'], None
    # 「合計」は2つ(販売費・送料)。送料ブロックの合計は 送料 の後ろ
    for rr in range(labels['送料'], ws.max_row + 1):
        if ws.cell(rr, 13).value == '合計':
            ship_total = rr
            break
    for k in range(6):
        for col in (17, 18):
            if ws.cell(r0 + k, col).value not in (None, ''):
                sys.exit(f'❌ Amazon Q/R 列 行{r0 + k} に既に値がある(停止): {ws.cell(r0 + k, col).value!r}')
    rows_q = {}
    for k, (lab, val) in enumerate(CHANNEL_KPI):
        rr = r0 + k
        ws.cell(rr, 17).value = lab
        rows_q[k] = rr
    ws.cell(rows_q[1], 18).value = CHANNEL_KPI[1][1]
    ws.cell(rows_q[2], 18).value = CHANNEL_KPI[2][1].format(t=t, r1=rows_q[1])
    ws.cell(rows_q[3], 18).value = CHANNEL_KPI[3][1].format(r2=rows_q[2], exp=exp_total, ship=ship_total)
    ws.cell(rows_q[4], 18).value = CHANNEL_KPI[4][1].format(t=t, r3=rows_q[3])
    for k in (1, 2, 3):
        ws.cell(rows_q[k], 18).number_format = '#,##0'
    ws.cell(rows_q[4], 18).number_format = '0.0%'
    ws.column_dimensions['Q'].width = max(ws.column_dimensions['Q'].width or 0, 44)
    ws.column_dimensions['R'].width = max(ws.column_dimensions['R'].width or 0, 14)
    log.append(('チャネルKPI追加', 'Amazon', 'Q:R', f'Q{rows_q[0]}:R{rows_q[5]}', '', AMZ_FEE_ACTUAL, '手数料実額総額・粗利・限界利益・率(A〜O列不変・行挿入なし)'))
    print(f'   Amazon チャネル全体KPI: Q{rows_q[0]}:R{rows_q[5]}(合計行 {t}・販売費合計 N{exp_total}・送料合計 N{ship_total})')
    for ch in ('楽天', 'Amazon'):
        wbs[ch].save(P[ch])
    # 数式が意図した追加以外で変わっていないこと
    for ch in ('楽天', 'Amazon'):
        after = formulas_of(P[ch])
        added = {k: v for k, v in after.items() if k not in fx_before[ch]}
        changed = {k: (fx_before[ch][k], after[k]) for k in fx_before[ch] if after.get(k) != fx_before[ch][k]}
        if changed or (ch == '楽天' and added) or (ch == 'Amazon' and any(not k.startswith('R') for k in added)):
            sys.exit(f'❌ {ch}: 数式が意図せず変わった/増えた: 変更 {changed} 追加 {list(added)[:5]}(停止)')
        print(f'   {ch}: 既存数式 {len(fx_before[ch])} 個は不変' + (f'・追加 {sorted(added)}' if added else ''))
    return t, rows_q, exp_total, ship_total


def verify(P, t_amz, rows_q):
    res = {}
    for ch in ('楽天', 'Amazon'):
        ws = load_workbook(P[ch], data_only=True)[MONTH]
        t, labels = block_rows(ws)
        m = labels['限界利益']
        res[ch] = dict(sales=ws.cell(t, 9).value, gp=ws.cell(t, 14).value, gp_rate=ws.cell(t, 15).value, fee=ws.cell(t, 11).value, cost=ws.cell(t, 13).value,
                       margin=ws.cell(m, 14).value, rate=ws.cell(m + 1, 14).value, total_row=t)
        if ch == 'Amazon':
            res[ch].update(ch_fee=ws.cell(rows_q[1], 18).value, ch_gp=ws.cell(rows_q[2], 18).value, ch_margin=ws.cell(rows_q[3], 18).value, ch_rate=ws.cell(rows_q[4], 18).value)
        for chg in CHANGES:
            if chg[0] == ch:
                r = find_row(ws, ch, chg[1], chg[2])
                res.setdefault('L', {})[chg[1]] = ws.cell(r, 12).value
    return res


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--target', choices=['copy', 'production'], default='copy')
    ap.add_argument('--confirm', default='')
    a = ap.parse_args()
    if a.target == 'production' and a.confirm != '本番':
        sys.exit('❌ 本番は --confirm 本番 が必要')
    opened = XB.open_workbooks()
    if opened:
        sys.exit(f'⛔ Excel にブックが開いています(人が作業中の可能性)。閉じてから: {opened}')
    if a.target == 'copy':
        os.makedirs(COPY_DIR, exist_ok=True)
        for k in ('楽天', 'Amazon', 'exc'):
            shutil.copy2(PROD[k], COPY[k])
        P = COPY
    else:
        P = PROD
    print(f'■ 対象: {a.target} / {P["楽天"]} / {P["Amazon"]}')
    log = []
    if a.target == 'production':
        for k in ('楽天', 'Amazon'):
            print('   バックアップ:', backup(P[k], '8月最終確認反映前'))
    t_amz, rows_q, exp_total, ship_total = apply_kpi(P, log)
    for ch in ('楽天', 'Amazon'):
        print(f'   Excel再計算: {ch} …', XB.recalc(P[ch]))
    res = verify(P, t_amz, rows_q)
    exp = {'L': {'b0915ngd2j-6': 7536, 'B0HD7N75QV': 2862, 'B0HB36RZPY': 1232},
           '楽天': dict(sales=2579795, gp=706377, margin=353895),
           'Amazon': dict(sales=5610675, gp=1649976, margin=884213, ch_fee=759756, ch_gp=1669256, ch_margin=903493)}
    ok = True

    def chk(cond, msg):
        nonlocal ok
        print(('   ✅ ' if cond else '   ❌ ') + msg)
        ok = ok and bool(cond)
    for k, v in exp['L'].items():
        chk(res['L'].get(k) == v, f'L列 {k} = {res["L"].get(k)}(想定 {v})')
    rk, am = res['楽天'], res['Amazon']
    chk(round(rk['sales']) == exp['楽天']['sales'], f'楽天 売上 {rk["sales"]:,}')
    chk(round(rk['gp']) == exp['楽天']['gp'], f'楽天 粗利 {rk["gp"]:,.0f}(想定 {exp["楽天"]["gp"]:,})')
    chk(round(rk['margin']) == exp['楽天']['margin'], f'楽天 限界利益 {rk["margin"]:,.0f}(想定 {exp["楽天"]["margin"]:,}) 率 {rk["rate"]:.1%}')
    chk(round(am['sales']) == exp['Amazon']['sales'], f'Amazon 売上 {am["sales"]:,}')
    chk(round(am['gp']) == exp['Amazon']['gp'], f'Amazon 商品別粗利 {am["gp"]:,.0f}(想定 {exp["Amazon"]["gp"]:,})')
    chk(round(am['margin']) == exp['Amazon']['margin'], f'Amazon 商品別限界利益 {am["margin"]:,.0f}(想定 {exp["Amazon"]["margin"]:,})')
    chk(am['ch_fee'] == exp['Amazon']['ch_fee'], f'Amazon チャネル実額手数料 {am["ch_fee"]:,}')
    chk(round(am['ch_gp']) == exp['Amazon']['ch_gp'], f'Amazon チャネル粗利 {am["ch_gp"]:,.0f}(想定 {exp["Amazon"]["ch_gp"]:,})')
    chk(round(am['ch_margin']) == exp['Amazon']['ch_margin'], f'Amazon チャネル限界利益 {am["ch_margin"]:,.0f}(想定 {exp["Amazon"]["ch_margin"]:,}) 率 {am["ch_rate"]:.1%}')
    print(f'   2チャネル合計: 売上 {rk["sales"] + am["sales"]:,} / 限界利益(楽天15%推定＋Amazonチャネル実額) {rk["margin"] + am["ch_margin"]:,.0f}')
    # 例外一覧の承認(本番のみファイルを書く。copy は複製へ)
    exc_path = P['exc']
    if a.target == 'production':
        print('   バックアップ:', backup(exc_path, '承認前'))
    rows = list(csv.DictReader(open(exc_path, encoding='utf-8-sig')))
    n_app = 0
    for d in rows:
        if d['楽天商品管理番号'] in ('b0842r98vl', 'b0dffzzxqw') and d['区分'] == '共有' and not d['承認日']:
            d['承認者'] = '英樹(確認リスト 8月最終確認 F-15/F-16「共有」)'
            d['承認日'] = DAY
            d['根拠'] = d['根拠'].replace('英樹の確認待ち', '英樹確認済み(共有)')
            n_app += 1
    with open(exc_path, 'w', encoding='utf-8-sig', newline='') as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0].keys())); w.writeheader(); w.writerows(rows)
    chk(n_app == 2, f'共有在庫の例外一覧: 承認 {n_app} 件(PCA所得税・ドルツ)')
    # 記録
    if a.target == 'production':
        lp = f'{OUT}/8月最終確認_反映記録_{date.today():%Y%m%d}.csv'
        with open(lp, 'w', encoding='utf-8-sig', newline='') as f:
            w = csv.writer(f); w.writerow(['区分', 'チャネル', '対象', 'セル', '変更前', '変更後', '記録'])
            w.writerows(log)
        print('   記録:', lp)
    print('■ 結果:', 'PASS' if ok else 'FAIL(想定値と不一致。停止)')
    sys.exit(0 if ok else 1)


if __name__ == '__main__':
    main()
