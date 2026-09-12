# -*- coding: utf-8 -*-
"""restore_product_numbers.py — 楽天CSV取込シートの商品番号を、取込前の文字列へ復元する

使い方:
    python3 restore_product_numbers.py plan  <候補CSV>    読み取り専用。適用できる行と保留する行を表示
    python3 restore_product_numbers.py apply <候補CSV>    バックアップ → Excelで文字列として書込 → 再読込で検証 → ログ
    python3 restore_product_numbers.py plan-ids           JAN/管理番号/SKU の先頭0消失の候補を表示(§7-3・読み取り専用)
    python3 restore_product_numbers.py apply-ids          同上を適用(バックアップ・再読込検証・ログ)

例:
    python3 restore_product_numbers.py plan  ../../01_InventoryManagement/SourceData/Output/商品番号_復元候補_20260911.csv

なぜ必要か:
    在庫ツールの取込VBAは商品番号を `.Value` 代入しており、Excelが手入力と同じ解釈をする。
    先頭0が落ちる・カンマが桁区切りとして消える・`N/N` が分数として評価される。
    復元の根拠は **同じ取込に使った元xlsx**(Import/楽天在庫リスト_import.xlsx)だけ。
    小数からの逆算や桁数による先頭0の補完はしない(2026-09-11 ChatGPT承認 V-3)。

適用の条件(2026-09-12 承認):
    1. 候補CSVの差の理由が「取込で変わった4種」か「英樹確認の誤記」であること。それ以外の行は触らない
    2. 管理番号＋SKUでシート上の行を**一意に**照合できること(重複・該当なしはスキップ)
    3. 照合した行が候補作成時の行と一致し、**現在値が候補作成時から変わっていない**こと
    4. 元xlsxを今もう一度読み、同じ位置の商品番号・商品名が候補と一致すること(根拠の再確認)
    5. 文字列として保存し、元値・新値・根拠・更新結果をログへ残す

書込は Excel(AppleScript)で行う。xlsm は openpyxl で保存できない(drawing1.xml が壊れる)。
書込の前に NumberFormat="@" を設定する。設定しないと同じ変換がもう一度起きる。

⚠️ この修正では、元xlsxの時点で既に数値だった24件は復元できない(上流の確認が要る)。
"""
import csv
import os
import shutil
import subprocess
import sys
import warnings
from datetime import date

warnings.filterwarnings('ignore')
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from openpyxl import load_workbook

import build_todo_lists as T
import sync_cost_master as S

STOCK = T.RAKUTEN_STOCK
IMPORT_XLSX = os.path.dirname(STOCK) + '/Import/楽天在庫リスト_import.xlsx'
SHEET = '楽天CSV取込'
COL_ITEM_NO = 3                                   # C列 = 商品番号(VBA TC_ITEM_NO)
OUT = os.path.dirname(STOCK) + '/Output'

# 候補CSVの「差の理由」→ 適用区分。ここに無い理由の行は**変更しない**
KIND = {
    '文字列→整数(値は同じ・型のみ)': '型修正',
    '文字列→整数(先頭0消失)':       '文字列復元',
    '文字列→整数(カンマ消失)':      '文字列復元',
    '文字列→小数(分数評価)':        '文字列復元',
    '英樹確認の誤記':               '上流誤記訂正',
}


def canon(v):
    """セル値を比較用の文字列にする。int は桁そのまま、float は repr、str はそのまま。"""
    if v is None:
        return ''
    if isinstance(v, bool):
        return str(v)
    if isinstance(v, int):
        return str(v)
    if isinstance(v, float):
        return str(int(v)) if v.is_integer() else repr(v)
    return str(v)


def key_of(ctrl, sku):
    return (S.norm(ctrl), S.norm(sku))


def load_sheet():
    """取込先シートを {行番号: (管理番号, SKU, 商品番号, 商品名)} と key→[行番号] で返す。"""
    wb = load_workbook(STOCK, read_only=True, data_only=True)
    ws = wb[SHEET]
    rows, by_key = {}, {}
    for rno, r in enumerate(ws.iter_rows(min_row=3, values_only=True), 3):
        if not r or (r[1] in (None, '') and r[3] in (None, '')):
            continue
        rows[rno] = (r[1], r[3], r[2], r[4])
        by_key.setdefault(key_of(r[1], r[3]), []).append(rno)
    wb.close()
    return rows, by_key


def load_source_aligned():
    """元xlsxを取込VBAと同じ順序で読み、取込先の行番号に並べる(VBAは空SKU行を飛ばすだけ)。"""
    wb = load_workbook(IMPORT_XLSX, read_only=True, data_only=True)
    ws = wb.worksheets[0]                          # VBA: wbSrc.Sheets(1)
    out, rno = {}, 3
    for r in ws.iter_rows(min_row=2, values_only=True):
        sku = str(r[3]).strip() if r[3] is not None else ''
        mg = str(r[1]).strip() if r[1] is not None else ''
        if not (sku or mg):
            continue
        out[rno] = (r[1], r[3], r[2], r[4])
        rno += 1
    wb.close()
    return out


def plan(cand_path):
    """候補CSVの各行を条件1〜4で判定し、適用行と保留行に分ける(何も書かない)。"""
    cands = list(csv.DictReader(open(cand_path, encoding='utf-8-sig')))
    rows, by_key = load_sheet()
    src = load_source_aligned()
    plan_rows = []
    for c in cands:
        rec = dict(c)
        rec['区分'] = KIND.get(c['差の理由'], '')
        rec['更新結果'] = ''
        rec['新値'] = ''
        if not rec['区分']:
            rec['更新結果'] = '対象外(承認範囲外の理由。変更しない)'
            plan_rows.append(rec)
            continue

        # 条件2: 管理番号＋SKUで一意
        k = key_of(c['楽天商品管理番号'], c['SKU管理番号'])
        hit = by_key.get(k, [])
        if len(hit) != 1:
            rec['更新結果'] = f'スキップ(管理番号+SKUの照合が一意でない: {len(hit)}件)'
            plan_rows.append(rec)
            continue
        rno = hit[0]
        # 条件3: 行と現在値が候補作成時のまま
        if str(rno) != str(c['取込先行']).strip():
            rec['更新結果'] = f'スキップ(行が候補作成時と違う: {c["取込先行"]}→{rno})'
            plan_rows.append(rec)
            continue
        cur = rows[rno][2]
        # 候補CSVの現在値は 数値行=そのまま / 文字列行=repr(引用符付き) で書かれている(2026-09-11版)。
        # **型も含めて**一致を見る。既に文字列へ直された行を「値が同じ」で通さない
        expect = repr(cur) if isinstance(cur, str) else canon(cur)
        if c['現在値'].strip() != expect:
            rec['更新結果'] = f'スキップ(現在値が候補作成時から変わっている: {canon(cur)!r})'
            plan_rows.append(rec)
            continue
        # 条件4: 元xlsxを今読み直して根拠を再確認(同じ位置の商品番号・商品名)
        s = src.get(rno)
        if s is None:
            rec['更新結果'] = '保留(元xlsxに同じ位置の行が無い)'
            plan_rows.append(rec)
            continue
        if str(s[3] or '').strip() != str(rows[rno][3] or '').strip():
            rec['更新結果'] = '保留(元xlsxと商品名が一致しない。根拠に矛盾)'
            plan_rows.append(rec)
            continue
        src_pn = s[2]
        if rec['区分'] == '上流誤記訂正':
            # 取込では変わっていない(元xlsx=現在値)ことを確認し、英樹確認の値を新値にする
            if canon(src_pn) != canon(cur):
                rec['更新結果'] = '保留(上流誤記のはずが元xlsxと現在値が違う。根拠に矛盾)'
                plan_rows.append(rec)
                continue
            new = c['復元候補'].strip()
            if not new or '英樹' not in c['確認根拠']:
                rec['更新結果'] = '保留(英樹確認の値が候補に無い)'
                plan_rows.append(rec)
                continue
        else:
            if not isinstance(src_pn, str):
                rec['更新結果'] = f'保留(元xlsxの値が文字列ではない: {type(src_pn).__name__})'
                plan_rows.append(rec)
                continue
            if repr(src_pn) != c['元xlsx値'].strip() or src_pn != c['復元候補'].strip():
                rec['更新結果'] = f'保留(元xlsxの値が候補と一致しない: {src_pn!r})'
                plan_rows.append(rec)
                continue
            new = src_pn
        rec['新値'] = new
        rec['_row'] = rno
        rec['_cur'] = cur
        rec['更新結果'] = '適用可'
        plan_rows.append(rec)
    return plan_rows


def _as_lit(v):
    """AppleScriptのリテラルへ。数値はそのまま、文字列は引用符で囲む。"""
    if isinstance(v, bool):
        return 'true' if v else 'false'
    if isinstance(v, int):
        return str(v)
    if isinstance(v, float):
        return repr(v)
    s = str(v).replace('\\', '\\\\').replace('"', '\\"')
    return f'"{s}"'


def apply_excel(targets):
    """Excelで書き込む。各セルは**書く直前にもう一度現在値を照合**し、違えば書かない。

    targets: [(row, current_value, new_str)]  戻り値: {row: 'written' | 'mismatch:<値>'}
    """
    lines = []
    for t in targets:
        rno, cur, new = t[0], t[-2], t[-1]
        col = t[1] if len(t) == 4 else 'C'            # (row, col, cur, new) または (row, cur, new)=商品番号C列
        # 型も見る: 数値のはずのセルが既に文字列なら書かない(値が同じでも「変わっている」)
        chk = (f'(class of (value of c)) is real and (value of c) is equal to {_as_lit(cur)}'
               if isinstance(cur, (int, float))
               else f'(class of (value of c)) is not real and ((value of c) as string) is equal to {_as_lit(str(cur))}')
        lines.append(f'''
    set c to cell "{col}{rno}" of ws
    if {chk} then
        set number format of c to "@"
        set value of c to {_as_lit(new)}
        set end of res to "{col}{rno}:written"
    else
        set end of res to "{col}{rno}:mismatch:" & ((value of c) as string)
    end if''')
    script = f'''
set p to POSIX file "{STOCK}"
set res to {{}}
with timeout of 600 seconds
tell application "Microsoft Excel"
    open p
    delay 1
    set wb to active workbook
    set ws to worksheet "{SHEET}" of wb
    {''.join(lines)}
    save wb
    close wb saving no
end tell
end timeout
set AppleScript's text item delimiters to linefeed
return res as string
'''
    r = subprocess.run(['osascript', '-e', script], capture_output=True, text=True, timeout=660)
    if r.returncode != 0:
        raise RuntimeError(f'Excelでの書込に失敗: {r.stderr}')
    out = {}
    for line in r.stdout.strip().split('\n'):
        if ':' in line:
            addr, _, rest = line.partition(':')
            key = int(addr[1:]) if addr.startswith('C') and len(targets) and len(targets[0]) == 3 else addr
            out[key] = rest
    return out


ID_COLS = {'JAN': 'A', '管理番号': 'B', 'SKU': 'D'}      # 商品番号(C)は plan/apply が扱う
SRC_IDX = {'JAN': 0, '管理番号': 1, '商品番号': 2, 'SKU': 3}


def _digits_equal(a, b):
    """先頭0の有無・型の違いを無視して同じ数字列か(それ以外は文字列として同じか)。"""
    ca, cb = canon(a), canon(b)
    if ca.isdigit() and cb.isdigit():
        return ca.lstrip('0') == cb.lstrip('0')
    return ca == cb


def plan_ids():
    """JAN/管理番号/SKU の先頭0消失を、**行の対応を複数項目で一意に確認してから**候補にする(§7-3)。

    対応の根拠: 商品名(完全一致)＋管理番号・SKU・商品番号(先頭0と型を無視して一致)＋販売価格＋在庫数。
    この組が 元xlsx側でも取込先側でも1行だけ に絞れたときだけ候補にする。行番号だけでは決めない。
    """
    rows, _ = load_sheet()
    src = load_source_aligned()
    wb = load_workbook(STOCK, read_only=True, data_only=True)
    full_t = {rno: r[:11] for rno, r in enumerate(wb[SHEET].iter_rows(min_row=3, values_only=True), 3) if rno in rows}
    wb.close()
    wb = load_workbook(IMPORT_XLSX, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    full_s, rno = {}, 3
    for r in ws.iter_rows(min_row=2, values_only=True):
        sku = str(r[3]).strip() if r[3] is not None else ''
        mg = str(r[1]).strip() if r[1] is not None else ''
        if not (sku or mg):
            continue
        full_s[rno] = r[:11]
        rno += 1
    wb.close()

    def blank0(v):
        # 取込VBAは元xlsxの空欄(販売価格・在庫数)を 0 で書く。行の対応を見る目的では 空欄=0 と扱う
        # (OS原則「0と空欄は別」に反しない: ここでは値を使わず、同じ行かどうかを見ているだけ)
        c = canon(v)
        return '' if c in ('', '0') else c

    def sig(r):
        return (str(r[4] or '').strip(), canon(r[1]).lstrip('0'), canon(r[3]).lstrip('0'),
                canon(r[2]).lstrip('0'), blank0(r[5]), blank0(r[7]))
    from collections import Counter
    cs, ct = Counter(sig(r) for r in full_s.values()), Counter(sig(r) for r in full_t.values())

    out = []
    for rno, t in full_t.items():
        s_ = full_s.get(rno)
        for name, ci in SRC_IDX.items():
            if name == '商品番号':
                continue
            sv, tv = (s_[ci] if s_ else None), t[ci]
            if not (isinstance(sv, str) and isinstance(tv, (int, float)) and sv.startswith('0')):
                continue                                   # 先頭0消失だけが対象(§7-3)
            rec = {'取込先行': rno, '列': name, '楽天商品管理番号': canon(t[1]), 'SKU管理番号': canon(t[3]),
                   '現在値': canon(tv), '元xlsx値': repr(sv), '復元候補': sv, '差の理由': '文字列→整数(先頭0消失)',
                   '対応の根拠': '', '更新結果': '', '商品名': str(t[4] or '')[:40]}
            if sig(t) != sig(s_):
                rec['更新結果'] = '保留(同じ位置の元xlsx行と項目が一致しない)'
            elif cs[sig(t)] != 1 or ct[sig(t)] != 1:
                rec['更新結果'] = f'保留(対応が曖昧: 元xlsx {cs[sig(t)]}行 / 取込先 {ct[sig(t)]}行 が同じ項目)'
            elif not _digits_equal(sv, tv):
                rec['更新結果'] = '保留(先頭0を除いても数字が一致しない)'
            else:
                rec['対応の根拠'] = '商品名・管理番号・SKU・商品番号・販売価格・在庫数が元xlsxと1対1で一致'
                rec['更新結果'] = '適用可'
                rec['_col'] = ID_COLS[name]; rec['_cur'] = tv
            out.append(rec)
    return out


def apply_ids(plan_rows, mode):
    from collections import Counter
    ok = [r for r in plan_rows if r['更新結果'] == '適用可']
    print(f'■ 識別子(先頭0消失)候補 {len(plan_rows)}セル → 適用可 {len(ok)}セル')
    for k, v in Counter(r['列'] for r in ok).items():
        print(f'   {k:6} {v:3}セル')
    for k, v in Counter(r['更新結果'] for r in plan_rows if r['更新結果'] != '適用可').items():
        print(f'   保留 {v:3}  {k}')
    if mode == 'plan-ids':
        for r in ok[:12]:
            print(f"   行{r['取込先行']:<4} {r['列']:6} {r['現在値']:>16} → {r['復元候補']!r}  {r['商品名'][:20]}")
        if len(ok) > 12:
            print(f'   … 他 {len(ok) - 12}セル')
        return
    bdir = os.path.dirname(STOCK) + '/Backup'
    os.makedirs(bdir, exist_ok=True)
    base = os.path.splitext(os.path.basename(STOCK))[0]
    bpath = f'{bdir}/{base}_backup_{date.today():%Y%m%d}_識別子先頭0復元前.xlsm'
    seq = 2
    while os.path.exists(bpath):
        bpath = f'{bdir}/{base}_backup_{date.today():%Y%m%d}_識別子先頭0復元前_{seq}.xlsm'; seq += 1
    shutil.copy2(STOCK, bpath)
    print(f'\n✅ バックアップ: {bpath}')
    res = apply_excel([(r['取込先行'], r['_col'], r['_cur'], r['復元候補']) for r in ok])
    wb = load_workbook(STOCK, read_only=True, data_only=True)
    after = {rno: r[:11] for rno, r in enumerate(wb[SHEET].iter_rows(min_row=3, values_only=True), 3)}
    wb.close()
    n_ok = 0
    for r in ok:
        st = res.get(f"{r['_col']}{r['取込先行']}", 'no-result')
        v = after[r['取込先行']][SRC_IDX[r['列']]]
        if st == 'written' and isinstance(v, str) and v == r['復元候補']:
            r['更新結果'] = '更新済(文字列・再読込で一致)'; n_ok += 1
        elif st == 'written':
            r['更新結果'] = f'⚠️ 書込後の再読込が一致しない: {v!r}'
        else:
            r['更新結果'] = f'スキップ(書込直前の照合で不一致: {st})'
    print(f'✅ 更新 {n_ok}/{len(ok)}セル')
    log = f'{OUT}/識別子_先頭0復元ログ_{date.today():%Y%m%d}.csv'
    head = ['取込先行', '列', '楽天商品管理番号', 'SKU管理番号', '元値', '新値', '差の理由', '対応の根拠', '更新結果', '商品名']
    with open(log, 'w', encoding='utf-8-sig', newline='') as f:
        w = csv.writer(f); w.writerow(head)
        for r in plan_rows:
            w.writerow([r['取込先行'], r['列'], r['楽天商品管理番号'], r['SKU管理番号'], r['現在値'], r['復元候補'],
                        r['差の理由'], r['対応の根拠'], r['更新結果'], r['商品名']])
    print(f'   → {log}')


def main():
    if len(sys.argv) >= 2 and sys.argv[1] in ('plan-ids', 'apply-ids'):
        apply_ids(plan_ids(), sys.argv[1])
        return
    if len(sys.argv) < 3 or sys.argv[1] not in ('plan', 'apply'):
        print(__doc__)
        sys.exit(1)
    mode, cand_path = sys.argv[1], sys.argv[2]
    rows = plan(cand_path)

    from collections import Counter
    ok = [r for r in rows if r['更新結果'] == '適用可']
    print(f'■ 候補 {len(rows)}行 → 適用可 {len(ok)}行')
    for k, v in Counter(r['区分'] or '—' for r in ok).items():
        print(f'   {k:10} {v:3}行')
    print('■ 適用しない行')
    for k, v in Counter(r['更新結果'] for r in rows if r['更新結果'] != '適用可').items():
        print(f'   {v:3}行  {k}')

    if mode == 'plan':
        for r in ok:
            print(f"   行{r['_row']:<4} {r['区分']:8} {canon(r['_cur']):>22} → {r['新値']!r}")
        return

    # ── apply ──
    bdir = os.path.dirname(STOCK) + '/Backup'
    os.makedirs(bdir, exist_ok=True)
    base = os.path.splitext(os.path.basename(STOCK))[0]
    # 同日に2回目以降の適用があっても**毎回**直前の状態を残す(既存名なら連番を付ける)
    stem = os.path.splitext(os.path.basename(cand_path))[0]
    bpath = f'{bdir}/{base}_backup_{date.today():%Y%m%d}_V3商品番号復元前.xlsm'
    seq = 2
    while os.path.exists(bpath):
        bpath = f'{bdir}/{base}_backup_{date.today():%Y%m%d}_V3商品番号復元前_{seq}.xlsm'
        seq += 1
    shutil.copy2(STOCK, bpath)
    print(f'\n✅ バックアップ: {bpath}')

    res = apply_excel([(r['_row'], r['_cur'], r['新値']) for r in ok])

    # 再読込で検証: 文字列型で新値と一致しているか
    wb = load_workbook(STOCK, read_only=True, data_only=True)
    ws = wb[SHEET]
    after = {rno: r[COL_ITEM_NO - 1] for rno, r in enumerate(ws.iter_rows(min_row=3, values_only=True), 3)}
    wb.close()
    n_ok = 0
    for r in ok:
        rno = r['_row']
        st = res.get(rno, 'no-result')
        v = after.get(rno)
        if st == 'written' and isinstance(v, str) and v == r['新値']:
            r['更新結果'] = '更新済(文字列・再読込で一致)'
            n_ok += 1
        elif st == 'written':
            r['更新結果'] = f'⚠️ 書込後の再読込が一致しない: {v!r}({type(v).__name__})'
        else:
            r['更新結果'] = f'スキップ(書込直前の照合で不一致: {st})'
    print(f'✅ 更新 {n_ok}/{len(ok)}行')

    # ログは候補ファイルごとに分ける(同名で上書きすると前回のログが消える。2026-09-12 に一度消した)
    log = f'{OUT}/商品番号_復元ログ_{date.today():%Y%m%d}_{stem}.csv'
    head = ['取込先行', '楽天商品管理番号', 'SKU管理番号', '区分', '元値', '新値', '差の理由',
            '根拠', '更新結果', '商品名']
    with open(log, 'w', encoding='utf-8-sig', newline='') as f:
        w = csv.writer(f)
        w.writerow(head)
        for r in rows:
            w.writerow([r['取込先行'], r['楽天商品管理番号'], r['SKU管理番号'], r['区分'],
                        r['現在値'], r['新値'], r['差の理由'], r['確認根拠'], r['更新結果'],
                        r['商品名']])
    print(f'   → {log}')


if __name__ == '__main__':
    main()
