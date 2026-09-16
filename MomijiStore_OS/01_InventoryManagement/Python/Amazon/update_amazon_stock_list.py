# -*- coding: utf-8 -*-
"""update_amazon_stock_list.py — Amazonの在庫CSV(自己発送・FBA)から Amazon在庫リスト_import.xlsx を作る(コピー先へ)

使い方:
    python3 update_amazon_stock_list.py plan  <自己発送CSV> <FBA CSV>
    python3 update_amazon_stock_list.py build <自己発送CSV> <FBA CSV> --out <出力xlsx>

CSVの形(2026-09-16 確認・Shift_JIS・値は ="..." 形式):
    SKU, ASIN, number(在庫数), price(販売価格), point, cost(単価), condition, title

やること(2026-08-03 に確立した流れ):
    ・「在庫」シートの ▼FBA / ▼自己発送 のセクション構成と列構成を維持する
    ・SKUで照合して 在庫数・販売価格・単価 を更新。合計金額は数式でなく値(在庫数×単価)
    ・SKUが無く同じASINが1つだけあれば「SKU変更」として JAN/商品名を引き継ぐ(候補として印を付ける。自動確定しない)
    ・新規SKUはセクション末尾に追加(JANは空欄)
    ・CSVに無い行は**消さない・0にしない・繰り越さない**。旧数量・旧単価・旧基準日ごと「未掲載_未確認」シートへ分ける(別枠管理)
    ・識別子(JAN・ASIN・SKU)は文字列のまま
本番の Amazon在庫リスト_import.xlsx は書き換えない(--out はコピー先)。
"""
import argparse
import csv
import io
import os
import sys
import warnings
from datetime import datetime

warnings.filterwarnings('ignore')
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

BASE = '/Users/hide0726/Desktop/Claude Code/MomijiStore_OS'
CUR_LIST = BASE + '/01_InventoryManagement/SourceData/Amazon在庫リスト_import.xlsx'
PINK = PatternFill('solid', fgColor='FFC7CE')
YELLOW = PatternFill('solid', fgColor='FFFF00')
C_JAN, C_ASIN, C_PN, C_SKU, C_NAME, C_PRICE, C_COST, C_STOCK, C_TOTAL = range(1, 10)
C_MISSING = 10


def norm(v):
    return str(v).strip().upper() if v is not None and str(v).strip() else ''


def read_csv(path):
    raw = open(path, 'rb').read()
    txt = raw.decode('utf-8-sig') if raw.startswith(b'\xef\xbb\xbf') else raw.decode('cp932')
    rows = list(csv.reader(io.StringIO(txt)))
    hdr = rows[0]
    need = ['SKU', 'ASIN', 'number', 'price', 'cost', 'title']
    for n in need:
        if n not in hdr:
            sys.exit(f'❌ {os.path.basename(path)} に列「{n}」がありません: {hdr}')
    i = {n: hdr.index(n) for n in need}
    clean = lambda s: s[2:-1] if s.startswith('="') and s.endswith('"') else s.strip()
    out = {}
    for r in rows[1:]:
        if not r or len(r) <= max(i.values()):
            continue
        rec = {n: clean(r[i[n]]) for n in need}
        out[norm(rec['SKU'])] = rec
    return out


def to_num(s):
    try:
        return int(s) if str(s).strip().lstrip('-').isdigit() else float(s)
    except (TypeError, ValueError):
        return None


def sections(ws):
    """(セクション名, 開始行, 終了行) を返す。"""
    secs, cur = [], None
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if isinstance(v, str) and v.startswith('▼'):
            if cur:
                secs.append((cur[0], cur[1], r - 1))
            cur = (v, r + 1)
    if cur:
        secs.append((cur[0], cur[1], ws.max_row))
    return secs


def run(mode, self_csv, fba_csv, out_path=None):
    src = {'FBA': read_csv(fba_csv), '自己発送': read_csv(self_csv)}
    wb = load_workbook(CUR_LIST)
    ws = wb['在庫']
    secs = sections(ws)
    stat = {'更新': 0, '在庫数変更': 0, '新規': 0, 'CSVに無い': 0, 'SKU変更候補': 0}
    diffs = []
    ws.cell(1, C_MISSING).value = f'CSVに無い({datetime.now():%Y-%m-%d %H:%M})'
    insert_at = {}                                   # セクションごとの追加位置(末尾)
    for name, s0, s1 in secs:
        key = 'FBA' if 'FBA' in name else '自己発送'
        rows = src[key]
        seen = set()
        asin_rows = {}
        for r in range(s0, s1 + 1):
            if ws.cell(r, C_ASIN).value in (None, ''):
                continue
            asin_rows.setdefault(norm(ws.cell(r, C_ASIN).value), []).append(r)
        for r in range(s0, s1 + 1):
            sku = norm(ws.cell(r, C_SKU).value)
            if not sku and ws.cell(r, C_ASIN).value in (None, ''):
                continue
            rec = rows.get(sku)
            if rec is None:
                ws.cell(r, C_MISSING).value = 'CSVに無い'
                ws.cell(r, C_MISSING).fill = PINK
                stat['CSVに無い'] += 1
                diffs.append(('CSVに無い', key, ws.cell(r, C_SKU).value, ws.cell(r, C_ASIN).value, ws.cell(r, C_STOCK).value, ''))
                continue
            seen.add(sku)
            old_stock = ws.cell(r, C_STOCK).value
            n = to_num(rec['number']); p = to_num(rec['price']); c = to_num(rec['cost'])
            if str(old_stock) != str(n):
                stat['在庫数変更'] += 1
                diffs.append(('在庫数変更', key, rec['SKU'], rec['ASIN'], old_stock, n))
            ws.cell(r, C_STOCK).value = n
            if p is not None:
                ws.cell(r, C_PRICE).value = p
            if c is not None and c > 0:
                ws.cell(r, C_COST).value = c
            cost_v = ws.cell(r, C_COST).value
            ws.cell(r, C_TOTAL).value = (n or 0) * cost_v if isinstance(cost_v, (int, float)) else None
            ws.cell(r, C_MISSING).value = None
            stat['更新'] += 1
        insert_at[key] = s1
        # 新規(CSVにあってシートに無いSKU)
        for sku, rec in rows.items():
            if sku in seen:
                continue
            cand = asin_rows.get(norm(rec['ASIN']), [])
            note = ''
            if len(cand) == 1 and ws.cell(cand[0], C_MISSING).value == 'CSVに無い':
                note = f'SKU変更候補(同じASINの行 {cand[0]} がCSVに無い)'
                stat['SKU変更候補'] += 1
            stat['新規'] += 1
            diffs.append(('新規', key, rec['SKU'], rec['ASIN'], '', to_num(rec['number']), note))
    if mode == 'build':
        # 新規はセクション末尾へ(後ろのセクションから挿入して行ずれを防ぐ)
        for name, s0, s1 in reversed(secs):
            key = 'FBA' if 'FBA' in name else '自己発送'
            existing = {norm(ws.cell(r, C_SKU).value) for r in range(s0, s1 + 1)}
            new = [rec for sku, rec in src[key].items() if sku not in existing]
            if not new:
                continue
            ws.insert_rows(s1 + 1, len(new))
            for k, rec in enumerate(new):
                r = s1 + 1 + k
                n = to_num(rec['number']); p = to_num(rec['price']); c = to_num(rec['cost'])
                vals = {C_JAN: None, C_ASIN: rec['ASIN'], C_PN: None, C_SKU: rec['SKU'], C_NAME: rec['title'],
                        C_PRICE: p, C_COST: (c if c and c > 0 else None), C_STOCK: n,
                        C_TOTAL: ((n or 0) * c if c and c > 0 else None)}
                for col, v in vals.items():
                    ws.cell(r, col).value = v
                if not (c and c > 0):
                    ws.cell(r, C_COST).fill = YELLOW
                for col in (C_JAN, C_ASIN, C_SKU):
                    ws.cell(r, col).number_format = '@'
        # 別枠管理(ChatGPT 2026-09-16): CSVに無い行は本体から外し「未掲載_未確認」へ(旧数量・旧単価・旧基準日を保持)
        old_stamp = datetime.fromtimestamp(os.path.getmtime(CUR_LIST)).strftime('%Y-%m-%d %H:%M')
        wu = wb.create_sheet('未掲載_未確認')
        for c, h in enumerate(['JAN', 'ASIN', '商品番号', 'SKU管理番号', '商品名', '販売価格', '単価(旧)', '在庫数(旧)', '合計金額(旧・参考)',
                               '旧基準日', '状態', 'セクション', '同じASINの新規SKU(候補・統合しない)'], 1):
            wu.cell(1, c).value = h
        drop = []
        sec_of = {}
        for name, s0, s1 in sections(ws):
            for r in range(s0, s1 + 1):
                sec_of[r] = 'FBA' if 'FBA' in name else '自己発送'
        new_by_asin = {}
        for key, rows in src.items():
            for sku, rec in rows.items():
                new_by_asin.setdefault(norm(rec['ASIN']), []).append(rec['SKU'])
        for r in range(2, ws.max_row + 1):
            if ws.cell(r, C_MISSING).value == 'CSVに無い':
                vals = [ws.cell(r, c).value for c in range(1, 10)]
                qty, cost = vals[7], vals[6]
                cands = [x for x in new_by_asin.get(norm(vals[1]), []) if norm(x) != norm(vals[3])]
                rowv = vals[:6] + [cost, qty, (qty * cost if isinstance(qty, (int, float)) and isinstance(cost, (int, float)) else None),
                                   old_stamp, '最新CSVに未掲載・現在在庫未確認', sec_of.get(r, ''), ', '.join(cands)]
                rr = wu.max_row + 1
                for c, v in enumerate(rowv, 1):
                    wu.cell(rr, c).value = v
                for c in (1, 2, 4):
                    wu.cell(rr, c).number_format = '@'
                drop.append(r)
        for r in sorted(drop, reverse=True):
            ws.delete_rows(r, 1)
        os.makedirs(os.path.dirname(out_path), exist_ok=True)
        wb.save(out_path)
        log = os.path.splitext(out_path)[0] + f'_差分_{datetime.now():%Y%m%d_%H%M%S}.csv'
        with open(log, 'w', encoding='utf-8-sig', newline='') as f:
            w = csv.writer(f)
            w.writerow(['区分', 'セクション', 'SKU', 'ASIN', '旧', '新', '備考'])
            for d in diffs:
                w.writerow(list(d) + [''] * (7 - len(d)))
        print(f'✅ 書出: {out_path}\n   差分: {log}')
    return stat, diffs


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('mode', choices=['plan', 'build'])
    ap.add_argument('self_csv'); ap.add_argument('fba_csv')
    ap.add_argument('--out')
    a = ap.parse_args()
    if a.mode == 'build' and (not a.out or os.path.abspath(a.out) == os.path.abspath(CUR_LIST)):
        sys.exit('❌ --out はコピー先を指定してください(本番は書き換えません)')
    stat, diffs = run(a.mode, a.self_csv, a.fba_csv, a.out)
    print('■ Amazon在庫リスト との差:', stat)
    print('   取得(ファイル時刻):', {os.path.basename(p): datetime.fromtimestamp(os.path.getmtime(p)).strftime('%Y-%m-%d %H:%M') for p in (a.self_csv, a.fba_csv)})


if __name__ == '__main__':
    main()
