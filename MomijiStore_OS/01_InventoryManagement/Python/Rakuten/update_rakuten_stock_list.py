# -*- coding: utf-8 -*-
"""update_rakuten_stock_list.py — RMSの在庫CSV(dl-normal-item_*.csv)から 楽天在庫リスト_import.xlsx を作る(コピー先へ)

使い方:
    python3 update_rakuten_stock_list.py plan  <RMS CSV...>                 差分だけ表示(何も書かない)
    python3 update_rakuten_stock_list.py build <RMS CSV...> --out <出力xlsx>  コピー先へ新しい在庫リストを書く

例:
    python3 update_rakuten_stock_list.py plan  ../../import/Rakuten/dl-normal-item_20260916081129-*.csv
    python3 update_rakuten_stock_list.py build ../../import/Rakuten/dl-normal-item_20260916081129-*.csv \\
        --out ../../SourceData/V3_stock_test/Import/楽天在庫リスト_import.xlsx

RMS CSV の形(2026-09-16 確認):
    列: 商品管理番号（商品URL）, 商品番号, 在庫表示, SKU管理番号, システム連携用SKU番号, 在庫数, 在庫戻しフラグ, 在庫切れ時の注文受付
    ・親行 = SKU管理番号が空。商品番号はここにだけある
    ・SKU行 = SKU管理番号と在庫数がある。商品番号は空(親行から引く)
    ・**ダウンロードは複数ファイルに分割される**(…-1.csv, -2.csv …)。全部渡すこと。1つだけだと「消えた」が大量に出る
    ・Shift_JIS。Excelで開いて保存すると先頭0・カンマ・スラッシュが壊れるので、CSVのまま渡す

やること(2026-08-03 に確立した流れ・引継ぎ/記憶「在庫更新パイプライン」):
    1. 既存の「在庫」シートを (商品管理番号×SKU管理番号) ペアで照合し、在庫数と商品番号を更新する
    2. 新規ペアは追加(商品名・販売価格・単価は空欄 → 単価は黄色。JANは商品番号の先頭13桁だけから補完)
    3. CSVに無いペアは**消さない**。「CSVに無い」列に印を付けて残し、英樹が削除方針を決める
    4. 識別子(JAN・管理番号・商品番号・SKU)は **文字列のまま**書く(数値化しない)
    5. 差分(在庫数変更・新規・消えた・商品番号変更)をCSVに残す

本番の Import/楽天在庫リスト_import.xlsx は**このスクリプトでは書き換えない**(--out はコピー先)。
"""
import argparse
import csv
import glob
import io
import os
import re
import sys
import warnings
from collections import OrderedDict
from datetime import date, datetime

warnings.filterwarnings('ignore')
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

BASE = '/Users/hide0726/Desktop/Claude Code/MomijiStore_OS'
CUR_LIST = BASE + '/01_InventoryManagement/SourceData/Import/楽天在庫リスト_import.xlsx'
YELLOW = PatternFill('solid', fgColor='FFFF00')
PINK = PatternFill('solid', fgColor='FFC7CE')
# 在庫シートの列(1始まり): JAN, 楽天商品管理番号（ASIN), 商品番号, SKU管理番号, 商品名, 販売価格, 単価, 在庫数, 合計金額, JAN検証結果, AmazonのJAN
C_JAN, C_CTRL, C_PN, C_SKU, C_NAME, C_PRICE, C_COST, C_STOCK, C_TOTAL, C_JANCHK, C_AJAN = range(1, 12)
C_MISSING = 12          # 追加列: CSVに無い(印)


def norm(v):
    return str(v).strip().upper() if v is not None and str(v).strip() else ''


def read_rms(paths):
    """複数パートを読み、親行(管理番号→商品番号)とSKU行(管理番号,SKU→在庫数)を返す。"""
    parents, skus, stamp = {}, OrderedDict(), None
    for p in paths:
        raw = open(p, 'rb').read()
        txt = raw.decode('cp932') if not raw.startswith(b'\xef\xbb\xbf') else raw.decode('utf-8-sig')
        rows = list(csv.reader(io.StringIO(txt)))
        hdr = rows[0]
        need = ['商品管理番号（商品URL）', '商品番号', 'SKU管理番号', '在庫数']
        for n in need:
            if n not in hdr:
                sys.exit(f'❌ {os.path.basename(p)} に列「{n}」がありません: {hdr}')
        i = {n: hdr.index(n) for n in need}
        m = re.search(r'(\d{8})(\d{6})', os.path.basename(p))
        if m:
            stamp = f'{m.group(1)[:4]}-{m.group(1)[4:6]}-{m.group(1)[6:]} {m.group(2)[:2]}:{m.group(2)[2:4]}:{m.group(2)[4:]}'
        for r in rows[1:]:
            if not r or len(r) <= max(i.values()):
                continue
            ctrl, pn, sku, stock = (r[i[n]].strip() for n in need)
            if sku == '':
                parents[ctrl] = pn                     # 親行
            else:
                skus[(ctrl, sku)] = stock              # SKU行(文字列のまま)
    return parents, skus, stamp


def load_current():
    wb = load_workbook(CUR_LIST, read_only=True, data_only=True)
    ws = wb['在庫']
    rows = [list(r[:11]) for r in ws.iter_rows(min_row=2, values_only=True)
            if (r[3] not in (None, '')) or (r[1] not in (None, ''))]
    wb.close()
    return rows


def diff(parents, skus, cur):
    cur_by = {(norm(r[1]), norm(r[3])): r for r in cur}
    new_by = {(norm(c), norm(s)): (c, s, st) for (c, s), st in skus.items()}
    out = {'更新(在庫数変更)': [], '更新(在庫数同じ)': [], '商品番号変更': [], '新規': [], 'CSVに無い': [], '親行なし': []}
    for k, (c, s, st) in new_by.items():
        pn = parents.get(c)
        if pn is None:
            out['親行なし'].append((c, s, st))
        r = cur_by.get(k)
        if r is None:
            out['新規'].append((c, s, st, pn)); continue
        old_st = r[7] if r[7] is not None else ''
        if str(old_st) != str(int(st) if st.isdigit() else st):
            out['更新(在庫数変更)'].append((c, s, old_st, st))
        else:
            out['更新(在庫数同じ)'].append((c, s, st))
        old_pn = str(r[2]).strip() if r[2] is not None else ''
        if pn not in (None, '') and old_pn != pn:
            out['商品番号変更'].append((c, s, old_pn, pn))
    for k, r in cur_by.items():
        if k not in new_by:
            out['CSVに無い'].append((r[1], r[3], r[7], str(r[4] or '')[:30]))
    return out


def build(parents, skus, cur, out_path, stamp):
    """既存の在庫シートの構造を保ったまま、コピー先へ新しいリストを書く。"""
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    wb = load_workbook(CUR_LIST)                        # 書式・2シート構成をそのまま使う
    ws = wb['在庫']
    # 既存行の位置
    pos = {}
    for r in range(2, ws.max_row + 1):
        c, s = ws.cell(r, C_CTRL).value, ws.cell(r, C_SKU).value
        if c in (None, '') and s in (None, ''):
            continue
        pos[(norm(c), norm(s))] = r
    ws.cell(1, C_MISSING).value = f'CSVに無い({stamp or "取得日時不明"})'
    last = max(pos.values()) if pos else 1
    n_upd = n_new = 0
    for (c, s), st in skus.items():
        k = (norm(c), norm(s))
        pn = parents.get(c, '')
        stock = int(st) if st.isdigit() else st
        if k in pos:
            r = pos[k]
            ws.cell(r, C_STOCK).value = stock
            if pn:
                ws.cell(r, C_PN).value = pn
            n_upd += 1
        else:
            last += 1
            r = last
            jan = pn[:13] if re.match(r'\d{13}', pn or '') else None
            vals = {C_JAN: jan, C_CTRL: c, C_PN: pn or None, C_SKU: s, C_NAME: None, C_PRICE: None,
                    C_COST: None, C_STOCK: stock, C_TOTAL: None, C_JANCHK: None, C_AJAN: None}
            for col, v in vals.items():
                ws.cell(r, col).value = v
            ws.cell(r, C_COST).fill = YELLOW          # 単価は人が入れる
            n_new += 1
        # 識別子は文字列のまま(数値化しない)
        for col in (C_JAN, C_CTRL, C_PN, C_SKU):
            ws.cell(r, col).number_format = '@'
            v = ws.cell(r, col).value
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                ws.cell(r, col).value = str(int(v)) if float(v).is_integer() else str(v)
        ws.cell(r, C_MISSING).value = None
    # CSVに無いペア: 消さずに印
    n_missing = 0
    new_keys = {(norm(c), norm(s)) for (c, s) in skus}
    for k, r in pos.items():
        if k not in new_keys:
            ws.cell(r, C_MISSING).value = 'CSVに無い'
            ws.cell(r, C_MISSING).fill = PINK
            n_missing += 1
    wb.save(out_path)
    return n_upd, n_new, n_missing


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('mode', choices=['plan', 'build'])
    ap.add_argument('csv', nargs='+')
    ap.add_argument('--out')
    a = ap.parse_args()
    paths = sorted(p for pat in a.csv for p in glob.glob(pat))
    if not paths:
        sys.exit('❌ CSV が見つかりません')
    parents, skus, stamp = read_rms(paths)
    cur = load_current()
    d = diff(parents, skus, cur)
    print(f'■ RMS CSV {len(paths)}ファイル(取得 {stamp or "?"}): 親行 {len(parents)} / SKU行 {len(skus)}')
    for p in paths:
        print('   ', os.path.basename(p))
    print(f'■ 既存リスト {len(cur)}行 との差')
    for k, v in d.items():
        print(f'   {k:12} {len(v):4}')
    if len(paths) == 1 and re.search(r'-1\.csv$', paths[0]) and len(d['CSVに無い']) > 20:
        print('   ⚠️ ファイルが1つ(…-1.csv)だけです。RMSは大きな一覧を複数ファイルに分けます。'
              '続きの -2.csv 等が無いか確認してください(「CSVに無い」が多いのはそのためかもしれません)')
    print('   商品番号変更:', d['商品番号変更'][:5])
    print('   CSVに無い(在庫>0):', sum(1 for x in d['CSVに無い'] if isinstance(x[2], (int, float)) and x[2] > 0))
    if a.mode == 'build':
        if not a.out or os.path.abspath(a.out) == os.path.abspath(CUR_LIST):
            sys.exit('❌ --out はコピー先を指定してください(本番の Import は書き換えません)')
        n_upd, n_new, n_missing = build(parents, skus, cur, a.out, stamp)
        print(f'✅ 書出: {a.out}  更新 {n_upd} / 新規 {n_new} / CSVに無い(印) {n_missing}')
        log = os.path.splitext(a.out)[0] + f'_差分_{datetime.now():%Y%m%d_%H%M%S}.csv'
        with open(log, 'w', encoding='utf-8-sig', newline='') as f:
            w = csv.writer(f)
            w.writerow(['区分', '管理番号', 'SKU', '値1', '値2'])
            for k, v in d.items():
                for x in v:
                    w.writerow([k] + [str(y) for y in x])
        print(f'   差分: {log}')


if __name__ == '__main__':
    main()
