# -*- coding: utf-8 -*-
"""build_todo_lists.py — 月次の「要対応一覧」を生成する

使い方:
    python3 build_todo_lists.py <月シート名> [Transaction.csv]

例:
    python3 build_todo_lists.py 8月 ../SourceData/2026AugMonthlyTransaction.csv

出力:
    01_InventoryManagement/SourceData/Output/要対応一覧_<月>_<YYYYMMDD>.xlsx
    00_サマリー / 01_新商品 / 02_原価確認 / 03_商品情報不足
    04_データ不整合 / 05_Amazon手数料未反映 / 06_廃番整理候補
    ★ シートは「修正する順番」に並べる

方針:
- **修正が必要な行だけを出す。正常な行は出さない。**
- **読み取り専用。** 商品マスター・KPIシート・在庫テーブルは一切変更しない。
- 一覧はそのまま修正作業に使えること。判断に必要な材料と対応案を各行へ添える。
- 分析の成果物ではなく **OSを改善するためのTODOリスト** として出力する。

⚠️ AIは候補を提示するまで。反映は人の承認後に行う(DEC-MST-05)。
"""
import csv
import os
import sys
from collections import Counter, defaultdict
from datetime import date

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import sync_cost_master as S          # noqa: E402  突合ロジックを再利用する

BASE = S.BASE
INV_FILE = BASE + '/01_InventoryManagement/SourceData/在庫管理テーブル_v1.1.xlsm'
OUTPUT_DIR = BASE + '/01_InventoryManagement/SourceData/Output'

HEAD_FILL = PatternFill('solid', fgColor='DDDDDD')
S_FILL = PatternFill('solid', fgColor='FFC7CE')   # 優先度 最優先
A_FILL = PatternFill('solid', fgColor='FFEB9C')   # 優先度 高
CARRY_FILL = PatternFill('solid', fgColor='D9E1F2')  # 前月からの継続
BOLD = Font(bold=True)


# --- 読み込み(すべて読み取り専用)-------------------------------------------
def load_master_full():
    wb = load_workbook(S.MASTER_FILE, read_only=True, data_only=True)
    pm, lt = {}, []
    for r in wb[S.SH_MASTER].iter_rows(min_row=2, values_only=True):
        if r[0]:
            pm[str(r[0])] = {
                'jan': r[1], 'name': r[2] or '', 'type': r[3], 'cost': r[4],
                'channel': r[5], 'status': r[6], 'note': r[9],
                'ship_size': r[10] if len(r) > 10 else None,
                'weight': r[11] if len(r) > 11 else None,
            }
    for r in wb[S.SH_LISTING].iter_rows(min_row=2, values_only=True):
        if r[0]:
            lt.append({'lid': r[0], 'pid': r[1], 'ch': r[2], 'rctrl': r[3],
                       'rsku': r[4], 'asin': r[5], 'asku': r[6]})
    wb.close()
    return pm, lt


def load_stock():
    """内部管理ID → 在庫数。在庫行が無い商品は含まれない(それ自体が情報)"""
    wb = load_workbook(INV_FILE, read_only=True, data_only=True)
    stock = {}
    for r in wb['在庫管理テーブル'].iter_rows(min_row=2, values_only=True):
        if r[1]:
            stock[str(r[1])] = r[5]
    wb.close()
    return stock


def load_sales(month_sheet, pair2pid, ctrl2pid, asin2pid, asku2pid):
    """内部管理ID → その月の売上個数。楽天とAmazonを合算する"""
    sold = Counter()
    try:
        wb = load_workbook(S.KPI_FILE, read_only=True, data_only=True)
        if month_sheet in wb.sheetnames:
            ws = wb[month_sheet]
            for row in ws.iter_rows(min_row=8, values_only=True):
                if not row or row[2] is None:
                    break
                pid = (pair2pid.get((S.norm(row[2]), S.norm(row[4])))
                       or ctrl2pid.get(S.norm(row[2])))
                if pid and isinstance(row[7], (int, float)):
                    sold[pid] += row[7]
        wb.close()
    except Exception as e:
        print(f'  ※ 楽天KPIを読めません({e})')
    try:
        wb = load_workbook(S.AMZ_KPI_FILE, read_only=True, data_only=True)
        if month_sheet in wb.sheetnames:
            ws = wb[month_sheet]
            for row in ws.iter_rows(min_row=8, values_only=True):
                if not row or row[2] is None:
                    break
                pid = asin2pid.get(S.norm(row[2])) or asku2pid.get(S.norm(row[1]))
                if pid and isinstance(row[7], (int, float)):
                    sold[pid] += row[7]
        wb.close()
    except Exception as e:
        print(f'  ※ Amazon KPIを読めません({e})')
    return sold


# --- ① 新商品 ---------------------------------------------------------------
def build_new_products(month_sheet):
    rows = []
    cost_by_pid, pair2pid, ctrl2pid = S.load_master()
    _, _, missing = S.classify(S.load_kpi_month(month_sheet),
                               cost_by_pid, pair2pid, ctrl2pid)
    for r in missing:
        rows.append(['楽天', r['name'], r['ctrl'], r['sku'], '', '',
                     S.extract_jan(r['pn']) or '', r['cost'] or '',
                     f'python3 sync_cost_master.py register {month_sheet}'])
    cost_by_pid, asin2pid, asku2pid = S.load_master_amazon()
    _, _, missing_a = S.classify_amazon(S.load_amazon_kpi_month(month_sheet),
                                        cost_by_pid, asin2pid, asku2pid)
    jan_lookup = {}
    try:
        jan_lookup = S.load_amazon_jan_lookup()
    except Exception:
        pass
    for r in missing_a:
        jan = jan_lookup.get(S.norm(r['asin'])) or jan_lookup.get(S.norm(r['sku'])) or ''
        rows.append(['Amazon', r['name'], '', '', r['asin'], r['sku'],
                     jan, r['cost'] or '',
                     f'python3 sync_cost_master.py register-amazon {month_sheet}'])
    return rows


# --- ② 原価確認 -------------------------------------------------------------
def build_cost_conflicts(month_sheet, sold):
    rows = []
    cost_by_pid, pair2pid, ctrl2pid = S.load_master()
    _, conflict, _ = S.classify(S.load_kpi_month(month_sheet),
                                cost_by_pid, pair2pid, ctrl2pid)
    for r, pid, mc in conflict:
        rows.append(_cost_row('楽天', pid, r, mc, sold, month_sheet, 'adopt'))
    cost_by_pid, asin2pid, asku2pid = S.load_master_amazon()
    _, conflict_a, _ = S.classify_amazon(S.load_amazon_kpi_month(month_sheet),
                                         cost_by_pid, asin2pid, asku2pid)
    for r, pid, mc in conflict_a:
        rows.append(_cost_row('Amazon', pid, r, mc, sold, month_sheet, 'adopt-amazon'))
    return rows


def _cost_row(ch, pid, r, mc, sold, month_sheet, mode):
    new = r['cost']
    diff = (float(new) - float(mc)) if (new is not None and mc is not None) else ''
    qty = sold.get(pid, 0)
    return [ch, pid, r['name'], mc if mc is not None else '',
            new if new is not None else '', diff,
            (round(diff / float(mc) * 100, 1) if diff != '' and mc else ''),
            qty, (round(diff * qty) if diff != '' else ''), '',
            f'python3 sync_cost_master.py {mode} {month_sheet}']


# --- ③ 商品情報不足 ---------------------------------------------------------
#  優先順位は「その項目を直すと何が動き出すか」で決める。件数ではなく効果で並べる。
CRITICAL_ITEMS = {'発送サイズ区分', '商品重量',      # ラベル自動化の対象
                  '標準原価', '商品名'}              # 利益計算に影響する


def _priority(qty, stock, item):
    """最優先 > 高 > 中 > 低"""
    if qty > 0 and item in CRITICAL_ITEMS:
        return '最優先'          # 出荷実績あり かつ ラベル自動化/利益計算に効く
    if qty > 0:
        return '高'              # 今月販売実績あり
    if stock not in (None, 0):
        return '中'              # 在庫ありだが販売なし
    return '低'                  # 販売実績なし・在庫なし


def build_shortages(pm, lt, stock, sold):
    by_pid = defaultdict(list)
    for x in lt:
        by_pid[str(x['pid'])].append(x)
    rows = []
    for pid, p in pm.items():
        qty = sold.get(pid, 0)
        st = stock.get(pid)
        ls = by_pid.get(pid, [])
        has_rak = any(x['ch'] == '楽天' for x in ls)
        has_amz = any(x['ch'] == 'Amazon' for x in ls)

        def add(item, cur, how, where):
            rows.append([_priority(qty, st, item), pid, p['name'][:60], item,
                         cur if cur is not None else '', qty,
                         st if st is not None else '(行なし)', how, where])

        if not str(p['name'] or '').strip():
            add('商品名', '', 'ASIN/JANからカタログ照会して補完(自動更新しない)',
                '商品マスター C列')
        if p['cost'] in (None, '', 0):
            add('標準原価', p['cost'], '仕入実績またはKPIシートから確認して入力',
                '商品マスター E列')
        if not p['jan']:
            add('JAN', '', '商品番号/カタログから補完。照合精度に影響', '商品マスター B列')
        if has_amz and not any(x['asin'] for x in ls if x['ch'] == 'Amazon'):
            add('ASIN', '', 'Amazon出品にASINが無い', '出品テーブル F列')
        if has_amz and not any(x['asku'] for x in ls if x['ch'] == 'Amazon'):
            add('AmazonSKU', '', '手数料突合が効かなくなる', '出品テーブル G列')
        if has_rak and not any(x['rsku'] for x in ls if x['ch'] == '楽天'):
            add('楽天SKU', '', '在庫照合に必要', '出品テーブル E列')
        if pid not in stock:
            add('在庫行', '', '在庫0と未登録が区別できない', '在庫管理テーブル')
        if not p['ship_size']:
            add('発送サイズ区分', '', 'PRC-10 の候補提示で埋める(ラベル自動化に必須)',
                '商品マスター K列')
        if p['weight'] in (None, ''):
            add('商品重量', '', 'PRC-10 の候補提示で埋める(監査・ネコポス判定に必要)',
                '商品マスター L列')
    order = {'最優先': 0, '高': 1, '中': 2, '低': 3}
    rows.sort(key=lambda x: (order[x[0]], -x[5], x[1]))
    return rows


# --- ④ データ不整合 ---------------------------------------------------------
def build_inconsistencies(pm, lt):
    rows = []
    by_pid = defaultdict(list)
    for x in lt:
        by_pid[str(x['pid'])].append(x)

    # 同一ASINが複数の内部管理IDに割り当てられている = 重複登録の疑い
    asin_pids = defaultdict(set)
    for x in lt:
        if x['asin']:
            asin_pids[str(x['asin']).strip().upper()].add(str(x['pid']))
    for asin, pids in sorted(asin_pids.items()):
        if len(pids) > 1:
            names = ' / '.join(pm.get(p, {}).get('name', '')[:28] for p in sorted(pids))
            rows.append(['重複登録候補', f'ASIN {asin}', ' , '.join(sorted(pids)), names,
                         '同一ASINに複数の内部管理IDが割り当てられている',
                         'どのIDへ寄せるか人が判断(DEC-MST-01)。実績の紐付けが動く'])

    # 楽天SKUが単独で重複(使い回し)— ペアが同じなら無害、別商品を指すなら危険
    sku_pairs = defaultdict(set)
    for x in lt:
        if x['ch'] == '楽天' and x['rsku']:
            sku_pairs[str(x['rsku']).strip().lower()].add(
                (str(x['rctrl'] or '').strip().lower(), str(x['pid'])))
    for sku, s in sorted(sku_pairs.items()):
        pids = {p for _, p in s}
        if len(pids) > 1:
            rows.append(['楽天SKU重複(別商品)', f'SKU {sku}', ' , '.join(sorted(pids)),
                         ' / '.join(pm.get(p, {}).get('name', '')[:28] for p in sorted(pids)),
                         '同一SKUが別商品に使われている',
                         '🚫 SKU単独で在庫照合しない。管理番号×SKUのペアで扱う'])

    # JAN重複 — セット商品で説明できないものだけを出す
    jan_pids = defaultdict(list)
    for pid, p in pm.items():
        if p['jan']:
            jan_pids[str(p['jan']).strip()].append(pid)
    for jan, pids in sorted(jan_pids.items()):
        if len(pids) > 1 and not any(pm[p]['type'] == 'セット' for p in pids):
            rows.append(['JAN重複(セット外)', f'JAN {jan}', ' , '.join(sorted(pids)),
                         ' / '.join(pm[p]['name'][:28] for p in sorted(pids)),
                         'セットで説明できない同一JAN',
                         '入数違いなら正常。同一商品なら統合候補'])

    # 出品が1件も無い商品 / チャネル表記とのずれ
    for pid, p in pm.items():
        ls = by_pid.get(pid, [])
        if not ls:
            rows.append(['出品なし', pid, pid, p['name'][:28],
                         '出品テーブルに1件も無い', 'モールCSVと突き合わせられない'])
            continue
        chs = {x['ch'] for x in ls}
        ch = p['channel']
        expect = {'楽天': {'楽天'}, 'Amazon': {'Amazon'}, '両方': {'楽天', 'Amazon'}}.get(ch)
        if expect and chs != expect:
            rows.append(['チャネル不一致', pid, pid, p['name'][:28],
                         f'マスター「{ch}」 / 出品テーブル「{"・".join(sorted(chs))}」',
                         '販売チャネル列は出品テーブルから導出できる(Master_Design §3.2)'])
    return rows


# --- ⑤ 廃番・整理候補 -------------------------------------------------------
def build_retire_candidates(pm, stock, sold, sold_prev, month_sheet, prev_sheet):
    rows = []
    for pid, p in pm.items():
        st = stock.get(pid)
        q, qp = sold.get(pid, 0), sold_prev.get(pid, 0)
        if q or qp:
            continue                     # 直近2か月に売れていれば候補にしない
        reasons = [f'{month_sheet}売上0', f'{prev_sheet}売上0']
        conf = '低'
        if st == 0:
            reasons.append('在庫0'); conf = '中'
        elif st is None:
            reasons.append('在庫行なし(在庫0か未登録か不明)')
        else:
            continue                     # 在庫が残っているものは整理候補にしない
        rows.append([pid, p['name'][:60], st if st is not None else '(行なし)',
                     q, qp, ' / '.join(reasons), conf,
                     '販売停止 or 取扱終了の候補。⚠️ 再仕入予定は未確認のため誤検知あり'])
    rows.sort(key=lambda x: (x[6] != '中', x[1]))
    return rows


# --- ⑥ Amazon手数料未反映 ---------------------------------------------------
def build_fee_unresolved(month_sheet, tx_csv):
    if not tx_csv or not os.path.exists(tx_csv):
        return [], '(Transaction.csv 未指定のためスキップ)'
    import apply_amazon_fees as F
    fee_by_sku, _, _, _ = F.read_transactions(tx_csv)
    key_index = F.build_key_index()
    wb = load_workbook(S.AMZ_KPI_FILE, read_only=True, data_only=True)
    if month_sheet not in wb.sheetnames:
        wb.close()
        return [], f'({month_sheet} シートが無い)'
    ws = wb[month_sheet]
    sheet_sku, sheet_asin = set(), set()
    for row in ws.iter_rows(min_row=8, values_only=True):
        if not row or row[2] is None:
            break
        if row[1]:
            sheet_sku.add(str(row[1]).strip().upper())
        sheet_asin.add(str(row[2]).strip().upper())
    wb.close()
    rows = []
    for sku, fee in fee_by_sku.items():
        if sku in sheet_sku:
            continue
        asin = key_index['amazon_sku'].get(sku)
        if asin and asin in sheet_asin:
            continue
        reason = '出品テーブルにAmazonSKUが無い' if not asin else 'シートに該当ASINが無い'
        action = ('旧SKU/終了SKUの可能性。出品テーブルへ登録するか対象外と判断'
                  if not asin else '決済が翌月へずれた注文の可能性')
        rows.append([sku, round(-fee), reason, action])
    rows.sort(key=lambda x: -x[1])
    total = -sum(fee_by_sku.values())
    un = sum(r[1] for r in rows)
    note = (f'総手数料 ¥{total:,.0f} / 未反映 ¥{un:,.0f} '
            f'({un / total * 100:.1f}%) / 反映率 {(total - un) / total * 100:.1f}%')
    return rows, note


# --- 前月比較 ---------------------------------------------------------------
#  修正が済んだ行は今月の一覧から消える。それだけだと「何が片付いたか」が残らないため、
#  前月のファイルを読んで 新規 / 継続 / 解決 を突き合わせる。
#  一覧そのものは「要対応のみ」を保ち、履歴は 07_前月比較 と月次ファイルの蓄積で残す。
def find_prev_file(prev_sheet):
    if not os.path.isdir(OUTPUT_DIR):
        return None
    cands = sorted(f for f in os.listdir(OUTPUT_DIR)
                   if f.startswith(f'要対応一覧_{prev_sheet}_') and f.endswith('.xlsx'))
    return os.path.join(OUTPUT_DIR, cands[-1]) if cands else None


def load_prev_keys(path, sheet_title, key_cols):
    """前月ファイルの該当シートから突合キーの集合を作る"""
    if not path or not os.path.exists(path):
        return None
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
        if sheet_title not in wb.sheetnames:
            wb.close()
            return set()
        ws = wb[sheet_title]
        keys = set()
        for row in ws.iter_rows(min_row=5, values_only=True):   # 5行目から本文
            if not row or row[0] is None:
                continue
            keys.add(tuple(str(row[c]) if c < len(row) and row[c] is not None else ''
                           for c in key_cols))
        wb.close()
        return keys
    except Exception as e:
        print(f'  ※ 前月ファイルを読めません({e})')
        return None


# --- 出力 -------------------------------------------------------------------
def write_sheet(wb, spec, prev_path):
    """共通の作業列(状態 / 担当 / 完了)を付けて書き出す"""
    title, headers, rows, widths, note, key_cols = (
        spec['title'], spec['headers'], spec['rows'],
        spec['widths'], spec.get('note', ''), spec['key_cols'])

    prev_keys = load_prev_keys(prev_path, title, key_cols)
    def key_of(row):
        return tuple(str(row[c]) if c < len(row) and row[c] is not None else ''
                     for c in key_cols)
    cur_keys = {key_of(r) for r in rows}
    resolved = len(prev_keys - cur_keys) if prev_keys is not None else None
    carried = len(prev_keys & cur_keys) if prev_keys is not None else None

    ws = wb.create_sheet(title)
    ws['A1'] = f'{title}  —  {len(rows)}件'
    ws['A1'].font = Font(bold=True, size=12)
    line = note
    if prev_keys is not None:
        line += (('  |  ' if line else '')
                 + f'前月 {len(prev_keys)}件 → 今月 {len(rows)}件 '
                   f'(新規 {len(cur_keys - prev_keys)} / 継続 {carried} / 解決 {resolved})')
    ws['A2'] = line
    ws['A3'] = '⚠️ 候補である。反映は人の承認後に行う(DEC-MST-05)'

    full = headers + ['状態', '対応方法', '担当', '完了']
    for i, h in enumerate(full, 1):
        c = ws.cell(4, i)
        c.value = h
        c.font = BOLD
        c.fill = HEAD_FILL
        c.alignment = Alignment(vertical='center', wrap_text=True)

    for r, row in enumerate(rows, 5):
        for i, v in enumerate(row, 1):
            ws.cell(r, i).value = v
        n = len(headers)
        state = '新規' if (prev_keys is not None and key_of(row) not in prev_keys) else (
            '継続' if prev_keys is not None else '')
        ws.cell(r, n + 1).value = state
        ws.cell(r, n + 2).value = spec['how']
        ws.cell(r, n + 3).value = '蛯名'
        ws.cell(r, n + 4).value = ''            # 完了チェック欄(人が記入)
        if state == '継続':
            ws.cell(r, n + 1).fill = CARRY_FILL
        if headers[0] == '優先度':
            if row[0] == '最優先':
                ws.cell(r, 1).fill = S_FILL
            elif row[0] == '高':
                ws.cell(r, 1).fill = A_FILL

    for i, w in enumerate(list(widths) + [8, 40, 8, 8], 1):
        ws.column_dimensions[ws.cell(1, i).column_letter].width = w
    ws.freeze_panes = 'A5'
    return {'title': title, 'count': len(rows), 'new': len(cur_keys - (prev_keys or set())),
            'carried': carried, 'resolved': resolved, 'prev': len(prev_keys) if prev_keys is not None else None}


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    month_sheet = sys.argv[1]
    tx_csv = sys.argv[2] if len(sys.argv) > 2 else None
    prev_sheet = f'{int(month_sheet.rstrip("月")) - 1}月'
    prev_path = find_prev_file(prev_sheet)

    print(f'要対応一覧を生成します: {month_sheet}(前月比較: {prev_sheet})')
    if prev_path:
        print(f'  前月ファイル: {os.path.basename(prev_path)}')
    else:
        print(f'  前月ファイルなし — 今回は全件「新規」扱い')

    pm, lt = load_master_full()
    stock = load_stock()
    _, pair2pid, ctrl2pid = S.load_master()
    _, asin2pid, asku2pid = S.load_master_amazon()
    sold = load_sales(month_sheet, pair2pid, ctrl2pid, asin2pid, asku2pid)
    sold_prev = load_sales(prev_sheet, pair2pid, ctrl2pid, asin2pid, asku2pid)

    fees, fee_note = build_fee_unresolved(month_sheet, tx_csv)

    # ★ シートは「修正する順番」で並べる
    specs = [
        {'title': '01_新商品',
         'headers': ['チャネル', '商品名', '楽天商品管理番号', '楽天SKU', 'ASIN',
                     'AmazonSKU', 'JAN(推定)', 'KPI仕入値', '登録コマンド'],
         'rows': build_new_products(month_sheet),
         'widths': (10, 46, 22, 22, 16, 26, 16, 12, 44),
         'key_cols': [0, 2, 3, 4, 5],
         'how': 'register コマンドを実行(承認後)',
         'note': '修正場所: コマンド実行 → 商品マスター・出品テーブルへ自動追加'},
        {'title': '02_原価確認',
         'headers': ['チャネル', '内部管理ID', '商品名', '現在の原価', '新しい原価候補',
                     '差額', '差額率(%)', f'{month_sheet}個数', '影響額', '承認(記入)',
                     '適用コマンド'],
         'rows': build_cost_conflicts(month_sheet, sold),
         'widths': (10, 13, 40, 12, 14, 10, 11, 11, 12, 12, 44),
         'key_cols': [1],
         'how': '承認欄へ記入 → adopt コマンドを実行',
         'note': '修正場所: 商品マスター E列(標準原価)。⚠️ 両方が空の行は adopt では解決しない'},
        {'title': '03_商品情報不足',
         'headers': ['優先度', '内部管理ID', '商品名', '不足項目', '現在値',
                     f'{month_sheet}出荷', '在庫数', '対応方法(詳細)', '修正場所'],
         'rows': build_shortages(pm, lt, stock, sold),
         'widths': (9, 13, 44, 16, 12, 11, 10, 46, 22),
         'key_cols': [1, 3],
         'how': '修正場所の列を直接編集',
         'note': ('最優先=出荷実績あり かつ ラベル自動化/利益計算に影響 / '
                  '高=今月販売実績あり / 中=在庫ありだが販売なし / 低=販売実績なし・在庫なし')},
        {'title': '04_データ不整合',
         'headers': ['種別', '対象', '内部管理ID', '商品名', '詳細', '対応方法(詳細)'],
         'rows': build_inconsistencies(pm, lt),
         'widths': (20, 26, 26, 40, 38, 46),
         'key_cols': [0, 1],
         'how': '人が統合可否を判断(DEC-MST-01)',
         'note': '修正場所: 商品マスター / 出品テーブル'},
        {'title': '05_Amazon手数料未反映',
         'headers': ['AmazonSKU', '手数料(実額)', '未解決の理由', '想定される対応'],
         'rows': fees,
         'widths': (42, 14, 30, 46),
         'key_cols': [0],
         'how': '出品テーブルへ AmazonSKU を登録',
         'note': (fee_note + ' / 修正場所: 出品テーブル G列(AmazonSKU)')},
        {'title': '06_廃番整理候補',
         'headers': ['内部管理ID', '商品名', '在庫数', f'{month_sheet}売上',
                     f'{prev_sheet}売上', '判定根拠', '確信度', '対応候補'],
         'rows': build_retire_candidates(pm, stock, sold, sold_prev,
                                         month_sheet, prev_sheet),
         'widths': (13, 44, 10, 11, 11, 36, 9, 46),
         'key_cols': [0],
         'how': '販売停止 / 取扱終了を人が判断',
         'note': ('⚠️ 再仕入予定を確認できないため誤検知がある。'
                  '⏸ 販売ステータスの設計は保留中のため、当面は一覧の確認のみ')},
    ]

    wb = Workbook()
    wb.remove(wb.active)

    ws = wb.create_sheet('00_サマリー')
    ws['A1'] = f'要対応一覧 — {month_sheet}'
    ws['A1'].font = Font(bold=True, size=14)
    ws['A2'] = (f'作成 {date.today():%Y-%m-%d} / 商品マスター {len(pm)}件 / '
                f'⚠️ すべて「修正が必要なものだけ」。正常な行は含まない')
    ws['A3'] = '★ シートは「修正する順番」に並んでいる。上から順に対応する'

    stats = []
    for spec in specs:
        stats.append(write_sheet(wb, spec, prev_path))

    head = ['順', '一覧', '今月', '前月', '新規', '継続', '解決', '対応']
    for i, h in enumerate(head, 1):
        c = ws.cell(5, i); c.value = h; c.font = BOLD; c.fill = HEAD_FILL
    for r, (spec, st) in enumerate(zip(specs, stats), 6):
        vals = [r - 5, st['title'], st['count'],
                st['prev'] if st['prev'] is not None else '—',
                st['new'], st['carried'] if st['carried'] is not None else '—',
                st['resolved'] if st['resolved'] is not None else '—', spec['how']]
        for i, v in enumerate(vals, 1):
            ws.cell(r, i).value = v

    sev = Counter(r[0] for r in specs[2]['rows'])
    ws.cell(13, 1).value = '03_商品情報不足の内訳'
    ws.cell(13, 1).font = BOLD
    for i, k in enumerate(['最優先', '高', '中', '低'], 14):
        ws.cell(i, 1).value = k
        ws.cell(i, 2).value = sev.get(k, 0)
    ws.cell(19, 1).value = '月次標準フロー: ①分析 → ②要対応一覧 → ③商品マスター更新 → ④再分析(改善確認)'
    ws.cell(19, 1).font = BOLD
    for col, w in zip('ABCDEFGH', (5, 26, 8, 8, 8, 8, 8, 44)):
        ws.column_dimensions[col].width = w

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    path = f'{OUTPUT_DIR}/要対応一覧_{month_sheet}_{date.today():%Y%m%d}.xlsx'
    wb.save(path)

    print()
    print('═══ 要対応一覧(修正する順番) ═══')
    for i, (spec, st) in enumerate(zip(specs, stats), 1):
        extra = ''
        if st['prev'] is not None:
            extra = f"  [前月{st['prev']} 新規{st['new']} 解決{st['resolved']}]"
        print(f"  {i}. {st['title']:24s} {st['count']:5d}件{extra}")
    print(f"     └ 03の内訳: 最優先{sev.get('最優先',0)} / 高{sev.get('高',0)}"
          f" / 中{sev.get('中',0)} / 低{sev.get('低',0)}")
    print(f'\n  → {path}')


if __name__ == '__main__':
    main()
