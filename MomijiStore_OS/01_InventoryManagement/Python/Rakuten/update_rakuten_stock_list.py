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
    3. CSVに無いペアは本体(ツールが読む「在庫」)から外し、最新CSV掲載分だけにする。外した行の扱いは --missing で選ぶ:
         sheet … 旧数量・旧単価・旧基準日を保持したままシート「未掲載_未確認」へ分ける(別枠管理)
         log   … シートには残さず、同じ内容を 変更ログCSV(…_削除記録_<日時>.csv)に残す(英樹 Q14「消す」・ChatGPT 2026-09-16 採用)
       どちらも「在庫0」「損失確定」の意味ではない(今回のCSVに無いだけ・現在在庫は未確認)
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

# 共通モジュール(02_Analytics/Python/momiji_paths.py)の場所は __file__ からの相対位置で求める
sys.path.insert(0, os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..', '..', '02_Analytics', 'Python')))
import momiji_paths as MP  # noqa: E402

BASE = MP.OS_ROOT_STR
CUR_LIST = BASE + '/01_InventoryManagement/SourceData/Import/楽天在庫リスト_import.xlsx'
YELLOW = PatternFill('solid', fgColor='FFFF00')
PINK = PatternFill('solid', fgColor='FFC7CE')
# 在庫シートの列(1始まり): JAN, 楽天商品管理番号（ASIN), 商品番号, SKU管理番号, 商品名, 販売価格, 単価, 在庫数, 合計金額, JAN検証結果, AmazonのJAN
C_JAN, C_CTRL, C_PN, C_SKU, C_NAME, C_PRICE, C_COST, C_STOCK, C_TOTAL, C_JANCHK, C_AJAN = range(1, 12)
C_MISSING = 12          # 追加列: CSVに無い(印)


def norm(v):
    return str(v).strip().upper() if v is not None and str(v).strip() else ''


# CSV出力後にRMS側で直った商品番号(英樹の申告)。CSVの値が「旧」と一致するときだけ「新」に置き換え、備考に根拠を残す。
# 生CSVは変えない。適用するのは根拠が揃ったものだけ(ChatGPT 2026-09-16: 根拠不足なら自動適用しない)。
#
# 取り下げ(2026-09-16): 'b0dktchdgk': '4987176260659-2' → '4987176260635-2'
#   一度は「英樹 2026-09-16 現在はRMS上では修正済み」を根拠に置換したが、根拠不足として外した。
#   ・最新CSV(08:11)の親行は …659-2 のまま。CSVで変わったのは SKU(b0dfyjnwjt → b0dktchdgk)だけ
#   ・旧リストの b0dktchdgk 行は JAN 4987176260659・商品名「ホワイトピーチ&カモミール」。…635 は別商品(ホワイトリリー・管理番号 b0dfyjnwjt)のJAN
#   ・2026-09-11 の回答「b0dfyjnwjt は 4987176260635-2 が正」は SKU b0dfyjnwjt について。管理番号 b0dktchdgk の商品番号を指すかは未確認
#   → 確認リスト Q15 で英樹に確認。回答が出るまで派生データはCSVの値(…659-2)のまま
OVERRIDES = {}


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
                if ctrl in OVERRIDES and pn == OVERRIDES[ctrl][0]:
                    pn = OVERRIDES[ctrl][1]           # 申告済みの訂正を反映(根拠は OVERRIDES)
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


def build(parents, skus, cur, out_path, stamp, missing='sheet'):
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
    drop_rows = []
    for k, r in pos.items():
        if k not in new_keys:
            ws.cell(r, C_MISSING).value = 'CSVに無い'
            ws.cell(r, C_MISSING).fill = PINK
            n_missing += 1
            drop_rows.append(r)
    # 外した行の記録(旧数量・旧単価・旧基準日・SKU変更候補)。sheet=別シートへ / log=変更ログCSVへ
    old_stamp = datetime.fromtimestamp(os.path.getmtime(CUR_LIST)).strftime('%Y-%m-%d %H:%M')
    head = ['JAN', '楽天商品管理番号（ASIN)', '商品番号', 'SKU管理番号', '商品名', '販売価格', '単価(旧)', '在庫数(旧)', '合計金額(旧・参考)',
            '旧基準日', '状態', '最新CSV(取得)', 'SKU変更候補(同じ管理番号の新規SKU)']
    # 同じ管理番号で「消えたSKU」と「新規SKU」がある場合は候補として書くだけ(統合はしない)
    new_by_ctrl = {}
    for (c, s_) in skus:
        if (norm(c), norm(s_)) not in pos:
            new_by_ctrl.setdefault(norm(c), []).append(s_)
    records = []
    for r in sorted(drop_rows):
        vals = [ws.cell(r, c).value for c in range(1, 12)]
        vals = [(str(int(v)) if float(v).is_integer() else str(v)) if isinstance(v, (int, float)) and not isinstance(v, bool) and i < 4 else v
                for i, v in enumerate(vals)]                # 旧リストで数値になっていた識別子も文字列に揃える
        qty, cost = vals[7], vals[6]
        records.append(vals[:6] + [cost, qty, (qty * cost if isinstance(qty, (int, float)) and isinstance(cost, (int, float)) else None),
                                   old_stamp, '最新CSVに未掲載・現在在庫未確認', stamp or '', ', '.join(new_by_ctrl.get(norm(vals[1]), []))])
    if missing == 'sheet':
        wu = wb['未掲載_未確認'] if '未掲載_未確認' in wb.sheetnames else wb.create_sheet('未掲載_未確認')
        if wu.max_row <= 1 and wu.cell(1, 1).value is None:
            for c, h in enumerate(head, 1):
                wu.cell(1, c).value = h
        for rowv in records:
            rr = wu.max_row + 1
            for c, v in enumerate(rowv, 1):
                wu.cell(rr, c).value = v
            for c in (1, 2, 3, 4):
                wu.cell(rr, c).number_format = '@'
    else:
        if '未掲載_未確認' in wb.sheetnames:                  # 元リストに残っていても運用中の一覧には持ち込まない
            del wb['未掲載_未確認']
        log = os.path.splitext(out_path)[0] + f'_削除記録_{datetime.now():%Y%m%d_%H%M%S}.csv'
        with open(log, 'w', encoding='utf-8-sig', newline='') as f:
            w = csv.writer(f)
            w.writerow(['削除日時', '元リスト'] + head)
            for rowv in records:
                w.writerow([f'{datetime.now():%Y-%m-%d %H:%M:%S}', os.path.basename(CUR_LIST)] + rowv)
        print(f'   外した{len(records)}行の記録: {log}')
    for r in sorted(drop_rows, reverse=True):            # 本体からは外す(記録は別シートまたは変更ログに保持済み)
        ws.delete_rows(r, 1)
    wb.save(out_path)
    return n_upd, n_new, n_missing


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('mode', choices=['plan', 'build'])
    ap.add_argument('csv', nargs='+')
    ap.add_argument('--out')
    ap.add_argument('--missing', choices=['sheet', 'log'], default='sheet',
                    help='CSVに無い行: sheet=シート「未掲載_未確認」へ分ける / log=変更ログCSVだけに残す(Q14「消す」)')

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
        n_upd, n_new, n_missing = build(parents, skus, cur, a.out, stamp, a.missing)
        print(f'✅ 書出: {a.out}  更新 {n_upd} / 新規 {n_new} / CSVに無い(本体から外した) {n_missing} [{a.missing}]')
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
