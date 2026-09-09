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
import re
import sys
from collections import Counter, defaultdict
from datetime import date

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import sync_cost_master as S          # noqa: E402  突合ロジックを再利用する

BASE = S.BASE
SD = BASE + '/02_Analytics/SourceData'      # 広告分析シートの置き場
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
    """内部管理ID → 在庫数。在庫行が無い商品は含まれない(それ自体が情報)

    🔴 2026-09-09 監査で判明: このテーブルの在庫数は554行すべて0で、
       在庫ステータスは「未棚卸」。**実数が1件も入っていない。**
       つまりここの 0 は「在庫が無い」ではなく「まだ数えていない」である。
       在庫切れの根拠に使ってはならない。

       実在庫は別の場所にある:
         楽天   … 楽天在庫金額集計ツール_v1.0.xlsm「楽天CSV取込」
         Amazon … Amazon在庫リスト_import.xlsx
         統合   … 全体在庫サマリー_v1.0.xlsx
    """
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


# --- 候補生成のための外部データ ---------------------------------------------
#  「問題があります」で終わらせない。**「私はこれだと思います」まで提示する。**
#  人は承認するだけでよい状態にする(AIは候補まで・DEC-MST-05)。
AMZ_INV = BASE + '/01_InventoryManagement/SourceData/Amazon在庫リスト_import.xlsx'
RAK_INV = BASE + '/01_InventoryManagement/SourceData/Import/楽天在庫リスト_import.xlsx'

RE_W = re.compile(r'(\d+(?:\.\d+)?)\s*(kg|ｋｇ|g|ｇ)\b', re.I)
RE_V = re.compile(r'(\d+(?:\.\d+)?)\s*(ml|ｍｌ|L|リットル)', re.I)
RE_N = re.compile(r'[×x]\s*(\d+)\s*(本|個|枚|袋|包|錠|粒|セット)|(\d+)\s*(本|個|枚|袋|包|錠|粒)入')


def load_mall_inventory():
    """モール在庫リストから JAN・商品名の候補を引く索引を作る(読み取り専用)"""
    idx = {'asin': {}, 'asku': {}, 'rctrl': {}, 'rsku': {}}
    for path, kind in ((AMZ_INV, 'amz'), (RAK_INV, 'rak')):
        if not os.path.exists(path):
            continue
        try:
            wb = load_workbook(path, read_only=True, data_only=True)
            for ws in wb.worksheets:
                for r in ws.iter_rows(min_row=2, values_only=True):
                    if not r or len(r) < 5:
                        continue
                    jan, k1, _, k2, name = r[0], r[1], r[2], r[3], r[4]
                    val = {'jan': str(jan).strip() if jan else '',
                           'name': str(name).strip() if name else ''}
                    if k1:
                        idx['asin' if kind == 'amz' else 'rctrl'].setdefault(
                            str(k1).strip().upper(), val)
                    if k2:
                        idx['asku' if kind == 'amz' else 'rsku'].setdefault(
                            str(k2).strip().upper(), val)
            wb.close()
        except Exception as e:
            print(f'  ※ {os.path.basename(path)} を読めません({e})')
    return idx


def _lookup(idx, listings, field):
    """出品テーブルの識別子からモール在庫リストを引く"""
    for x in listings:
        for key, col in (('asin', 'asin'), ('asku', 'asku'),
                         ('rctrl', 'rctrl'), ('rsku', 'rsku')):
            v = x.get(col)
            if v:
                hit = idx[key].get(str(v).strip().upper())
                if hit and hit.get(field):
                    return hit[field], f'{"Amazon" if key in ("asin","asku") else "楽天"}在庫リスト'
    return '', ''


def estimate_weight(name):
    """商品名から商品重量を推定する。

    ⚠️ 商品名の g は**内容量**であって商品重量ではない。包装分を上乗せする。
    確信度は必ず「低」— 単一の情報源からの推定にすぎない。
    """
    if not name:
        return '', '', ''
    m = RE_W.search(name)
    base = unit = None
    if m:
        v = float(m.group(1))
        base = v * 1000 if m.group(2).lower() in ('kg', 'ｋｇ') else v
        unit = f'{m.group(1)}{m.group(2)}'
    else:
        m = RE_V.search(name)
        if m:
            v = float(m.group(1))
            base = v * 1000 if m.group(2) in ('L', 'リットル') else v   # 1ml≒1g
            unit = f'{m.group(1)}{m.group(2)}'
    if base is None:
        return '', '', ''
    n = RE_N.search(name)
    cnt = int(n.group(1) or n.group(3)) if n else 1
    total = base * cnt
    est = int(round(total * 1.15, -1))          # 包装分を15%上乗せ
    why = f'商品名「{unit}' + (f'×{cnt}' if cnt > 1 else '') + '」から内容量'
    why += f'{int(total)}g相当 + 包装15%'
    return est, why, '低'


def estimate_ship_size(weight_est, name):
    """発送サイズ区分の候補。確信度は必ず「低」(単一情報源・厚さが読めない)"""
    if not weight_est:
        return '', '', ''
    w = weight_est
    if w <= 300:
        cand = 'ネコポス'
    elif w <= 1000:
        cand = '宅急便コンパクト'
    elif w <= 3000:
        cand = '60サイズ'
    elif w <= 8000:
        cand = '80サイズ'
    else:
        cand = '100サイズ'
    why = f'推定重量{w}gから。⚠️厚さは商品名から読めないため境界は要確認'
    return cand, why, '低'


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
def _both_empty(r, mc):
    """マスターもKPIも原価が空の行は「不一致」ではなく「未入力」。
    adopt では解決できず、③商品情報不足(標準原価)が同じ商品を扱う。
    ②へ残すと同じ商品が2枚の一覧に並び、どちらを見ればよいか分からなくなる。"""
    return mc in (None, '', 0) and r['cost'] in (None, '', 0)


def _cost_note(rows):
    """②の注記。**影響額が1件も出せないときは「未算定」と書く。0円ではない。**"""
    n = len(rows)
    ok = [x for x in rows if isinstance(x[8], (int, float))]
    unknown = n - len(ok)
    if not ok:
        total = '影響額 合計: **未算定**(単位を揃えられた行が1件も無いため。0円ではない)'
    else:
        total = (f'影響額 合計: ¥{sum(x[8] for x in ok):,.0f}'
                 f'(単位確認済み {len(ok)}件ぶんのみ)'
                 + (f' / 未算定 {unknown}件' if unknown else ''))
    return ('修正場所: 商品マスター E列(標準原価)。'
            'マスターとKPIの値が食い違う行だけを載せる。'
            '両方とも空の行は「未入力」なので ③商品情報不足 側で扱う。'
            f'★ {total}。'
            '★ 個数は**その出品(ASIN/管理番号×SKU)が売れた数**を使う'
            '(内部管理ID全体の個数ではない)。'
            '⚠️ 「単位未確認」は、マスター原価とKPI仕入値が同じものを指しているか'
            '確認できていない行。差額を引き算しても意味がないので出していない。'
            '出品テーブルの L列(販売入数)・M列(原価単位)を根拠付きで埋めると判定できる')


def build_cost_conflicts(month_sheet, sold):
    """② 原価確認 — **単位が揃うと確認できた行だけ差額・影響額を出す**

    (2026-09-10 承認Q-7)。マスターが単品原価でKPIがセット原価なら、
    引き算した数字は差ではない。元のKPI値とマスター値は残すが、
    差額・影響額・合計には出さない。
    """
    pack, _pos = S.load_pack_info()
    rows = []
    cost_by_pid, pair2pid, ctrl2pid = S.load_master()
    _, conflict, _ = S.classify(S.load_kpi_month(month_sheet),
                                cost_by_pid, pair2pid, ctrl2pid)
    for r, pid, mc in conflict:
        if _both_empty(r, mc):
            continue
        rows.append(_cost_row('楽天', pid, r, mc, sold, month_sheet, 'adopt', pack))
    cost_by_pid, asin2pid, asku2pid = S.load_master_amazon()
    _, conflict_a, _ = S.classify_amazon(S.load_amazon_kpi_month(month_sheet),
                                         cost_by_pid, asin2pid, asku2pid)
    for r, pid, mc in conflict_a:
        if _both_empty(r, mc):
            continue
        rows.append(_cost_row('Amazon', pid, r, mc, sold, month_sheet, 'adopt-amazon', pack))
    return rows


def _cost_row(ch, pid, r, mc, sold, month_sheet, mode, pack=None):
    """1行 = 1出品(楽天:管理番号×SKU / Amazon:ASIN×SKU)。

    ⚠️ 影響額の個数は **その出品が売れた数** を使う(2026-09-10 承認R)。
       以前は内部管理ID全体の個数を使っており、同じIDに複数の出品がある商品で
       金額が何倍にも膨らんでいた(納豆菌P000334は5出品あり、26個を3回掛けていた)。

    同じ出品が複数行に分かれ、行ごとに仕入値が違う場合は**曖昧**とし、
    差額・影響額を出さずに「要確認」とする。どの値を正とすべきか決められないため。
    """
    new = r['cost']
    qty = r.get('units', 0)
    ambiguous = len({c for c in r.get('cost_variants', {new}) if c is not None}) > 1
    rows_note = ('/'.join(str(x) for x in r.get('sheet_rows', [])) or '')

    # 単位が揃うと確認できなければ、差額も影響額も出さない(承認Q-7)
    info, amb = S.lookup_pack(pack or {}, ch, S.norm(r['key']), S.norm(r['sku']))
    info = info or {}
    unit_ok = (info.get('unit') in (S.UNIT_SINGLE, S.UNIT_SET)
               and (info.get('unit') == S.UNIT_SET
                    or isinstance(info.get('pack'), (int, float))))
    if not ambiguous and not unit_ok:
        return [ch, pid, r['name'], mc if mc is not None else '',
                new if new is not None else '', '単位未確認', '', qty, '', rows_note,
                (amb or '販売入数・原価単位が未確認。'
                 '出品テーブルへ根拠付きで記入すると差額を出せる'), '']
    if ambiguous:
        return [ch, pid, r['name'], mc if mc is not None else '',
                new if new is not None else '', '要確認', '', qty, '', rows_note,
                '同じ出品が複数行にあり仕入値が食い違う。どれを正とするか人が判断する',
                '']
    diff = (float(new) - float(mc)) if (new is not None and mc is not None) else ''
    return [ch, pid, r['name'], mc if mc is not None else '',
            new if new is not None else '', diff,
            (round(diff / float(mc) * 100, 1) if diff != '' and mc else ''),
            qty, (round(diff * qty) if diff != '' else ''), rows_note, '',
            f'python3 sync_cost_master.py {mode} {month_sheet}']


# --- ③ 商品情報不足 ---------------------------------------------------------
#  「問題があります」で終わらせず、**候補・根拠・確信度まで提示する。**
#  優先順位は「その項目を直すと何が動き出すか」で決める。件数ではなく効果で並べる。
CRITICAL_ITEMS = {'発送サイズ区分', '商品重量',      # ラベル自動化の対象
                  '標準原価', '商品名'}              # 利益計算に影響する

# 業務改善効果の重み。★の算出に使う
ITEM_WEIGHT = {'商品名': 5, '標準原価': 5, '発送サイズ区分': 3, '商品重量': 2,
               'AmazonSKU': 2, 'ASIN': 1, 'JAN': 1, '楽天SKU': 1, '在庫行': 1}


def _priority(qty, stock, item):
    if qty > 0 and item in CRITICAL_ITEMS:
        return '最優先'
    if qty > 0:
        return '高'
    if stock not in (None, 0):
        return '中'
    return '低'


def _stars(qty, item, has_candidate):
    """業務改善効果で ★1〜★5 を付ける。

    効果 = 出荷量 × 項目の重み。候補を提示できる行は**承認だけで済む**ため
    着手コストが低く、同じ効果なら先に手を付けるべきなので加点する。
    """
    score = qty * ITEM_WEIGHT.get(item, 1)
    if has_candidate:
        score *= 1.5
    if score >= 60:
        return '★★★★★', '今日修正すると業務改善効果が非常に大きい'
    if score >= 20:
        return '★★★★', '今月中に修正したい'
    if score >= 6:
        return '★★★', '時間がある時'
    if score >= 1:
        return '★★', '余裕があれば'
    return '★', '放置可能'


def build_shortages(pm, lt, stock, sold, mall_idx):
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
        name = str(p['name'] or '').strip()

        # 候補は先に作る(★の算出に「候補の有無」を使うため)
        w_est, w_why, w_conf = estimate_weight(name)
        s_est, s_why, s_conf = estimate_ship_size(w_est, name)
        jan_c, jan_src = _lookup(mall_idx, ls, 'jan')
        name_c, name_src = _lookup(mall_idx, ls, 'name')

        def add(item, cur, cand, why, conf, where):
            star, star_why = _stars(qty, item, bool(cand))
            rows.append([star, _priority(qty, st, item), pid, name[:52], item,
                         cur if cur is not None else '',
                         cand, why, conf, qty,
                         st if st is not None else '(行なし)', where, star_why])

        if not name:
            add('商品名', '', name_c, (f'{name_src}から取得' if name_c else
                'ASIN/JANでカタログ照会が必要'), ('中' if name_c else ''),
                '商品マスター C列')
        if p['cost'] in (None, '', 0):
            add('標準原価', p['cost'], '', '仕入実績またはKPIシートの仕入値から確認', '',
                '商品マスター E列')
        if not p['jan']:
            add('JAN', '', jan_c, (f'{jan_src}から取得' if jan_c else
                '商品番号から抽出できず。カタログ照会が必要'),
                ('中' if jan_c else ''), '商品マスター B列')
        if has_amz and not any(x['asin'] for x in ls if x['ch'] == 'Amazon'):
            add('ASIN', '', '', 'AmazonSKUからカタログ照会が必要', '', '出品テーブル F列')
        if has_amz and not any(x['asku'] for x in ls if x['ch'] == 'Amazon'):
            add('AmazonSKU', '', '', '手数料突合が効かなくなる。セラーセントラルから取得', '',
                '出品テーブル G列')
        if has_rak and not any(x['rsku'] for x in ls if x['ch'] == '楽天'):
            add('楽天SKU', '', '', 'RMSから取得', '', '出品テーブル E列')
        if pid not in stock:
            add('在庫行', '', '', '在庫0と未登録が区別できない。行を作る', '', '在庫管理テーブル')
        if not p['ship_size']:
            add('発送サイズ区分', '', s_est, s_why or '商品名に物理量が無く推定できない',
                s_conf, '商品マスター K列')
        if p['weight'] in (None, ''):
            add('商品重量', '', (f'{w_est}g' if w_est else ''),
                w_why or '商品名に物理量が無く推定できない', w_conf, '商品マスター L列')
    star_order = {'★★★★★': 0, '★★★★': 1, '★★★': 2, '★★': 3, '★': 4}
    rows.sort(key=lambda x: (star_order[x[0]], -x[9], x[2]))
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


# --- ⑥ 廃番・整理候補 -------------------------------------------------------
#  「消す候補」ではなく **「残す価値があるか」を判断できる一覧**にする。
def load_history(pair2pid, ctrl2pid, asin2pid, asku2pid):
    """全月のKPIから 累計販売数・累計利益・最終販売月 を集める(読み取り専用)"""
    hist = defaultdict(lambda: {'qty': 0, 'profit': 0.0, 'last': ''})
    for path, is_amz in ((S.KPI_FILE, False), (S.AMZ_KPI_FILE, True)):
        try:
            wb = load_workbook(path, read_only=True, data_only=True)
        except Exception:
            continue
        for sheet in wb.sheetnames:
            if not sheet.endswith('月'):
                continue
            for row in wb[sheet].iter_rows(min_row=8, values_only=True):
                if not row or len(row) < 14 or row[2] is None:
                    break
                pid = (asin2pid.get(S.norm(row[2])) or asku2pid.get(S.norm(row[1]))) \
                    if is_amz else (pair2pid.get((S.norm(row[2]), S.norm(row[4])))
                                    or ctrl2pid.get(S.norm(row[2])))
                if not pid:
                    continue
                h = hist[pid]
                if isinstance(row[7], (int, float)) and row[7]:
                    h['qty'] += row[7]
                    h['last'] = max(h['last'], sheet, key=_month_key)
                if isinstance(row[13], (int, float)):
                    h['profit'] += row[13]
        wb.close()
    return hist


def _month_key(m):
    try:
        return int(str(m).rstrip('月'))
    except ValueError:
        return 0


# 粗利率の「要確認」閾値(2026-09-09 ChatGPT承認L / **Ver.1の暫定値**)。
#   ⚠️ これは「異常確定」ではなく「要確認」である。
#      高粗利が実力の商品も、赤字覚悟で売っている商品も実在する。
#      機械は疑うところまでで、正誤は人が決める。
#   3か月程度運用してから閾値を再評価する(次回見直し: 2026-12)
MARGIN_HIGH = 0.50      # これ以上は「仕入値が少なすぎる」疑い(セット品の1個分入力など)
MARGIN_LOW = 0.10       # これ未満は「仕入値が多すぎる/売価が下がった」疑い
MARGIN_SHIFT = 0.15     # 前月からこれ以上動いたら仕入値が変わった疑い


# 棚卸の滞留区分(2026-09-09 ChatGPT承認)。Ver.1の暫定値
STALE_LOW = 3       # 3か月売上0 = 低回転候補
STALE_LONG = 6      # 6か月売上0 = 長期滞留候補

IM = BASE + '/01_InventoryManagement/SourceData'
RAKUTEN_STOCK = IM + '/楽天在庫金額集計ツール_v1.0.xlsm'
AMAZON_STOCK = IM + '/Amazon在庫リスト_import.xlsx'


def load_real_stock():
    """**実在庫**を既存の2ファイルから読む。新しい仕組みは作らない。

    在庫管理テーブルは554行すべて0・「未棚卸」で実数が入っていない(2026-09-09 監査)。
    実在庫はここにある:
      楽天   … 楽天在庫金額集計ツール「楽天CSV取込」(RMS在庫CSVの取込結果)
      Amazon … Amazon在庫リスト_import「在庫」(FBA/自己発送の区分あり)

    ⚠️ 商品番号が「-a」で終わる行はAmazonとの共有在庫。
       Amazon側を正式値とし、楽天側は二重計上を避けるため除外する
       (全体在庫サマリーの運用ルール)。
    """
    out, asof = {}, []
    if os.path.exists(RAKUTEN_STOCK):
        wb = load_workbook(RAKUTEN_STOCK, read_only=True, data_only=True)
        ws = wb['楽天CSV取込']
        hi = None
        for i, row in enumerate(ws.iter_rows(min_row=1, max_row=6, values_only=True), 1):
            if row and row[0] == 'JAN':
                hi = i
                break
        if hi:
            for row in ws.iter_rows(min_row=hi + 1, values_only=True):
                if not row or not row[3]:
                    continue
                sku = str(row[3]).strip()
                if sku.endswith('-a'):        # Amazonとの共有在庫。Amazon側を正とする
                    continue
                q = row[7] if isinstance(row[7], (int, float)) else 0
                if q > 0:
                    out.setdefault(('楽天', S.norm(row[1])), {'sku': sku, 'qty': 0,
                                                             'src': '楽天在庫CSV'})['qty'] += q
        wb.close()
        asof.append(('楽天', date.fromtimestamp(os.path.getmtime(RAKUTEN_STOCK))))
    if os.path.exists(AMAZON_STOCK):
        wb = load_workbook(AMAZON_STOCK, read_only=True, data_only=True)
        ws = wb['在庫']
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or not row[1]:
                continue
            q = row[7] if isinstance(row[7], (int, float)) else 0
            if q > 0:
                out.setdefault(('Amazon', S.norm(row[1])), {'sku': str(row[3] or ''),
                                                            'qty': 0,
                                                            'src': 'Amazon在庫リスト'})['qty'] += q
        wb.close()
        asof.append(('Amazon', date.fromtimestamp(os.path.getmtime(AMAZON_STOCK))))
    return out, asof


def load_monthly_sales(pair2pid, ctrl2pid, asin2pid, asku2pid):
    """月 → {内部管理ID: 個数}。**新しい月が先**の順で返す。

    「何か月売れていないか」を数えるために使う。
    KPIシートに存在する月だけが対象なので、"6か月売上0" は
    正確には "OSが持っている範囲で6か月分の記録が無い" である。
    """
    out = {}
    for path, is_amz in ((S.KPI_FILE, False), (S.AMZ_KPI_FILE, True)):
        if not os.path.exists(path):
            continue
        wb = load_workbook(path, read_only=True, data_only=True)
        for sheet in wb.sheetnames:
            ws = wb[sheet]
            d = out.setdefault(sheet, {})
            for row in ws.iter_rows(min_row=8, values_only=True):
                if not row or len(row) < 8 or row[2] is None:
                    break
                pid = (asin2pid.get(S.norm(row[2])) or asku2pid.get(S.norm(row[1]))) \
                    if is_amz else (pair2pid.get((S.norm(row[2]), S.norm(row[4])))
                                    or ctrl2pid.get(S.norm(row[2])))
                if pid and isinstance(row[7], (int, float)) and row[7]:
                    d[str(pid)] = d.get(str(pid), 0) + row[7]
        wb.close()
    return dict(sorted(out.items(), key=lambda kv: -_month_key(kv[0])))


def build_stocktake(pm, lt, hist_months, month_sheet):
    """⑨ 棚卸対象リスト — **在庫があるものは全部**が対象

    売上上位だけでは足りない。動かない在庫こそ誰も見ないまま金額だけ残るため、
    低回転・長期滞留を優先する。各区分の中は**在庫金額の大きい順**。

    ⚠️ 季節商品があるため、この区分を**自動的な廃棄・処分判定に使わないこと**
       (2026-09-09 ChatGPT承認条件)。あくまで「数えに行く順番」である。
    """
    stock, asof = load_real_stock()
    if not stock:
        return [], []
    ctrl2pid, asin2pid = {}, {}
    for x in lt:
        if x['ch'] == '楽天' and x.get('rctrl'):
            ctrl2pid.setdefault(S.norm(x['rctrl']), str(x['pid']))
        if x['ch'] == 'Amazon' and x.get('asin'):
            asin2pid.setdefault(S.norm(x['asin']), str(x['pid']))
    rows = []
    for (ch, key), d in stock.items():
        pid = (ctrl2pid if ch == '楽天' else asin2pid).get(key)
        p = pm.get(pid) if pid else None
        cost = (p or {}).get('cost')
        amount = (d['qty'] * cost) if isinstance(cost, (int, float)) else None
        # 直近何か月売れていないか
        gap = None
        for i, m in enumerate(hist_months):
            if pid and hist_months[m].get(pid):
                gap = i
                break
        if gap is None:
            gap = len(hist_months)
        if cost is None:
            grp, note = 'E 要調査', '標準原価が無く在庫金額を出せない'
        elif gap >= STALE_LONG:
            grp, note = 'D 長期滞留', f'{STALE_LONG}か月以上売上なし。処分の検討対象'
        elif gap >= STALE_LOW:
            grp, note = 'C 低回転', f'{STALE_LOW}か月以上売上なし。評価損の検討対象'
        elif gap == 0:
            grp, note = 'A 直近販売あり', '動きがあるぶん差異も出やすい'
        else:
            grp, note = 'B 通常', ''
        rows.append([grp, ch, key, (p or {}).get('name', '(マスター未登録)')[:44],
                     d['sku'], d['qty'], cost, amount,
                     f'{gap}か月' if gap < len(hist_months) else f'{len(hist_months)}か月以上',
                     note, d['src'], ''])
    order = {'D 長期滞留': 0, 'C 低回転': 1, 'A 直近販売あり': 2, 'B 通常': 3, 'E 要調査': 4}
    rows.sort(key=lambda x: (order.get(x[0], 9), -(x[7] or 0)))
    return rows, asof


def build_margin_outliers(month_sheet, prev_sheet):
    """利益率 要確認リスト — 仕入値の入力ミスを疑う

    粗利率は「売価」と「仕入値」の2つでしか決まらない。
    売価はモール側の実績なので間違いようがない。
    **粗利率がおかしいときは、ほぼ仕入値が間違っている。**

    よくある間違い方:
      ・セット品なのに仕入値が1個分   → 粗利率が異常に高く出る
      ・値下げしたのに仕入値が古いまま → 粗利率が低く/マイナスに出る
      ・単位違い(ケース単価を個単価に) → どちらにも振れる
    """
    def load(path, kc, ch, sheet):
        if not os.path.exists(path):
            return {}
        wb = load_workbook(path, read_only=True, data_only=True)
        if sheet not in wb.sheetnames:
            wb.close()
            return {}
        ws = wb[sheet]
        out = {}
        for row in ws.iter_rows(min_row=8, values_only=True):
            if not row or row[2] is None:
                break
            key = S.norm(row[kc])
            sales = row[8] if isinstance(row[8], (int, float)) else 0
            if not sales or not key:
                continue
            out[key] = {'name': str(row[0] or '')[:46], 'ch': ch,
                        'price': row[6], 'units': row[7], 'sales': sales,
                        'cost': row[11], 'gp': row[13], 'rate': row[14]}
        wb.close()
        return out

    rows = []
    for path, kc, ch in ((S.KPI_FILE, 2, '楽天'), (S.AMZ_KPI_FILE, 2, 'Amazon')):
        cur = load(path, kc, ch, month_sheet)
        pre = load(path, kc, ch, prev_sheet)
        for key, d in cur.items():
            rate, price, cost = d['rate'], d['price'], d['cost']
            if not isinstance(rate, float):
                continue
            why, pri = [], None
            # 売価より仕入値が高い — 売るほど損。最も分かりやすい矛盾
            if isinstance(price, (int, float)) and isinstance(cost, (int, float)) \
                    and cost >= price:
                why.append(f'仕入値 ¥{cost:,.0f} ≧ 平均単価 ¥{price:,.0f}(売るほど損)')
                pri = '最優先'
            if rate < 0:
                why.append(f'粗利率がマイナス({rate:.1%})')
                pri = '最優先'
            elif rate < MARGIN_LOW:
                why.append(f'粗利率が低すぎる({rate:.1%})。'
                           '値下げ後に仕入値が古いままの可能性')
                pri = pri or '高'
            elif rate >= MARGIN_HIGH:
                why.append(f'粗利率が高すぎる({rate:.1%})。'
                           'セット品なのに仕入値が1個分の可能性')
                pri = pri or '中'
            # 前月から粗利率が大きく動いた = 仕入値が変わった/直された合図
            p = pre.get(key)
            if p and isinstance(p.get('rate'), float):
                sh = rate - p['rate']
                if abs(sh) >= MARGIN_SHIFT:
                    why.append(f'前月から粗利率が {p["rate"]:.1%} → {rate:.1%} '
                               f'({sh:+.1%})へ変動')
                    pri = pri or '中'
            if not why:
                continue
            rows.append([pri, ch, key, d['name'], d['units'], d['sales'],
                         price, cost, rate,
                         (pre.get(key) or {}).get('cost', ''),
                         ' / '.join(why),
                         '商品マスター E列(標準原価)を確認 → 直したら該当月のL列も直す', ''])
    order = {'最優先': 0, '高': 1, '中': 2}
    rows.sort(key=lambda x: (order.get(x[0], 9), -(x[5] or 0)))
    return rows


def build_coverage(pm, sold, month_sheet):
    """配送自動化の進み具合を「出荷個数カバー率(暫定)」で測る(ChatGPT指摘⑦)

    ⚠️ **暫定である理由**: 本来のKPIは
           自動処理可能な発送件数 ÷ 全発送件数
       だが、OSは現在「発送件数」を持っていない(受注データが未連携)。
       そこで当面は出荷個数で代用する。1注文に複数個入ると実態とずれる。
       受注データが入ったら本KPIへ切り替える。

    入力率(商品数ベース)も併記する。売れない商品を1件埋めても入力率は上がるが
    カバー率は上がらない。2つ並べて初めて「効く整備をしているか」が分かる。
    """
    n = len(pm)
    filled = {p for p, v in pm.items() if v.get('ship_size')}
    tot_q = sum(sold.values())
    cov_q = sum(q for p, q in sold.items() if p in filled)
    rows = [
        ['出荷個数カバー率(暫定) ★主KPI',
         f'{cov_q:,} / {tot_q:,}個', f'{(cov_q / tot_q) if tot_q else 0:.1%}',
         '発送サイズ区分が入った商品の出荷個数 ÷ 全出荷個数',
         '⚠️ 本KPIは「自動処理可能な発送件数 ÷ 全発送件数」。'
         '受注データ未連携のため出荷個数で代用中'],
        ['商品マスター入力率(発送サイズ区分) サブKPI',
         f'{len(filled):,} / {n:,}件', f'{(len(filled) / n) if n else 0:.1%}',
         '発送サイズ区分が入った商品数 ÷ 全商品数',
         '売れない商品を埋めても上がる。カバー率と併せて見ること'],
    ]
    w = {p for p, v in pm.items() if v.get('weight')}
    rows.append(['商品マスター入力率(商品重量) サブKPI',
                 f'{len(w):,} / {n:,}件', f'{(len(w) / n) if n else 0:.1%}',
                 '商品重量が入った商品数 ÷ 全商品数', '監査(DQA)の前提'])

    # どこまでやればどこまで届くか。全件を待つ必要がないことを示す
    rank = sorted(sold.items(), key=lambda x: -x[1])
    cum = 0
    marks = {10: None, 20: None, 50: None, 100: None, 200: None, 300: None}
    for i, (pid, q) in enumerate(rank, 1):
        cum += q
        if i in marks:
            marks[i] = cum
    rows.append(['', '', '', '', ''])
    rows.append(['── 出荷上位から整備した場合の到達点 ──', '', '', '', ''])
    for k, v in marks.items():
        if v is None:
            continue
        rows.append([f'上位{k}商品を整備すると', f'{v:,} / {tot_q:,}個',
                     f'{v / tot_q if tot_q else 0:.1%}', '',
                     f'出荷実績のある商品は全{len(rank):,}件'])
    return rows


def build_expense_cells(month_sheet):
    """⑦ 月次経費(黄色セル) — どの値が未入力か / どこから取るか / 誰が入れるか

    **新しい一覧を作らない。** 取得元と入力者の定義は
    apply_monthly_expenses.py が持っているので、それを読んで表にするだけ。
    定義が2箇所にあると必ず片方が古くなる。
    """
    import apply_monthly_expenses as E
    rows = []
    for d in E.survey(month_sheet):
        if d['filled']:
            continue                       # 入力済みは出さない(修正が必要なものだけ)
        rows.append([d['ch'], d['item'], d['cell'], d['src'], d['who'],
                     ('限界利益が確定しない' if d['item'] != '仕入値(商品ごと)'
                      else 'その商品の粗利が出ず、広告の損得も判定できない'), ''])
    return rows


def load_ad_spend(month_sheet, pair2pid, ctrl2pid, asin2pid, asku2pid):
    """広告分析シートから 内部管理ID → (広告費, 広告経由売上) を作る。

    ③重複の解消。廃番候補に挙がった商品へ広告費を払い続けているなら、
    それは2つの別々の問題ではなく**ひとつの判断**である。
    一覧を増やさず、廃番候補の行に広告費を並べて見せる。
    """
    out = {}
    for f, ch, kcol in [(SD + '/楽天RPP広告分析.xlsx', '楽天', 1),
                        (SD + '/AmazonSP広告分析.xlsx', 'Amazon', 2)]:
        if not os.path.exists(f):
            continue
        wb = load_workbook(f, read_only=True, data_only=True)
        if month_sheet not in wb.sheetnames:
            wb.close()
            continue
        ws = wb[month_sheet]
        rows = list(ws.iter_rows(values_only=True))
        wb.close()
        hdr = '商品管理番号' if ch == '楽天' else 'SKU'
        hi = next((i for i, r in enumerate(rows) if r and r[0] == hdr), None)
        if hi is None:
            continue
        sp, sa = (5, 7) if ch == '楽天' else (7, 8)
        for r in rows[hi + 1:]:
            if not r or r[0] in (None, '合計'):
                break
            k = S.norm(r[kcol - 1])
            pid = (ctrl2pid.get(k) or pair2pid.get((k, ''))) if ch == '楽天' \
                else (asin2pid.get(k) or asku2pid.get(k))
            if not pid:
                continue
            spend = r[sp] if isinstance(r[sp], (int, float)) else 0
            sales = r[sa] if isinstance(r[sa], (int, float)) else 0
            cur = out.setdefault(str(pid), [0.0, 0.0])
            cur[0] += spend
            cur[1] += sales
    return out


def build_retire_candidates(pm, lt, stock, sold, sold_prev, hist,
                            month_sheet, prev_sheet, ad_spend=None):
    by_pid = defaultdict(list)
    for x in lt:
        by_pid[str(x['pid'])].append(x)
    rows = []
    for pid, p in pm.items():
        st = stock.get(pid)
        q, qp = sold.get(pid, 0), sold_prev.get(pid, 0)
        if q or qp:
            continue                     # 直近2か月に売れていれば候補にしない
        if st not in (None, 0):
            continue                     # 在庫が残っているものは整理候補にしない
        h = hist.get(pid, {'qty': 0, 'profit': 0.0, 'last': ''})
        ls = by_pid.get(pid, [])
        rak = 'あり' if any(x['ch'] == '楽天' for x in ls) else 'なし'
        amz = 'あり' if any(x['ch'] == 'Amazon' for x in ls) else 'なし'
        reasons = [f'{month_sheet}売上0', f'{prev_sheet}売上0']
        conf = '低'
        # 🔴 在庫数0は「未棚卸」であって在庫切れではない(2026-09-09 監査)。
        #    在庫を根拠に確信度を上げてはいけない
        if st == 0:
            reasons.append('在庫未確認(未棚卸)')
        else:
            reasons.append('在庫行なし')
        # 「残す価値」— 累計実績が大きいものは安易に切らない
        if h['qty'] >= 20 or h['profit'] >= 20000:
            keep = '⚠️ 残す価値あり(累計実績が大きい)'
            conf = '低'
        elif h['qty'] == 0:
            keep = '実績なし — 整理しやすい'
        else:
            keep = '判断が要る'
        # ③ 広告との重複解消。廃番候補へ広告費を払い続けているなら
        #    「やめる」と「広告を止める」は同じ判断である
        ad = (ad_spend or {}).get(pid)
        ad_c, ad_note = ('', '')
        if ad and ad[0]:
            ad_c = round(ad[0])
            ad_note = ('★ 2か月連続で売上0なのに広告費を払っている。まず広告を止める'
                       if not ad[1] else '広告経由では売れている。廃番判断は慎重に')
        rows.append([pid, p['name'][:52], st if st is not None else '(行なし)',
                     h['last'] or '(なし)', '(データなし)',
                     h['qty'], round(h['profit']), rak, amz,
                     ' / '.join(reasons), conf, keep, ad_c, ad_note])
    # 広告費を払っている廃番候補を先頭へ。放置すると毎月お金が出ていく
    rows.sort(key=lambda x: (not x[12], -x[5]))
    return rows


# --- ⑤ Amazon手数料未反映 -------------------------------------------------
def fee_quality(month_sheet):
    """K列(販売手数料)の実額化の品質を測る(ChatGPT指摘H)

    同じ列に「実額」と「暫定15%の数式」が混在する。列を見ても区別できないので、
    **金額ベースと行ベースの両方**を数えて品質差を可視化する。
    金額の大きい商品から実額化されるため、金額88%でも行では34%ということが起きる。
    """
    wb = load_workbook(S.AMZ_KPI_FILE)          # 数式のまま読む
    if month_sheet not in wb.sheetnames:
        return None
    ws = wb[month_sheet]
    prov = act = 0
    r = 8
    while ws.cell(r, 3).value is not None:
        v = ws.cell(r, 11).value                # K列
        if isinstance(v, str) and v.startswith('='):
            prov += 1
        elif v is not None:
            act += 1
        r += 1
    n = prov + act
    wbv = load_workbook(S.AMZ_KPI_FILE, read_only=True, data_only=True)
    wsv = wbv[month_sheet]
    tot = 0.0
    rr = 8
    while wsv.cell(rr, 3).value is not None:
        v = wsv.cell(rr, 11).value
        if isinstance(v, (int, float)):
            tot += v
        rr += 1
    wbv.close()
    return {'act': act, 'prov': prov, 'n': n,
            'row_rate': (act / n) if n else 0, 'total': tot}


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


# 1項目あたりの作業時間目安(分)。(候補あり, 候補なし)
#  候補がある項目は「承認するだけ」なので短い。無いものは調べる時間が要る。
MINUTES = {
    '発送サイズ区分': (0.5, 5), '商品重量': (0.5, 5),
    'JAN': (0.5, 3), '商品名': (0.5, 3), '標準原価': (1, 3),
    'ASIN': (1, 5), 'AmazonSKU': (1, 5), '楽天SKU': (1, 5), '在庫行': (2, 2),
    '標準原価(不一致)': (1, 1),
}
# その項目を直すと何が動き出すか
EFFECT = {
    '発送サイズ区分': 'ラベル自動化の対象になる',
    '商品重量': '監査(DQA)の対象になる',
    '標準原価': '利益計算が正しくなる',
    '標準原価(不一致)': '利益計算が正しくなる',
    '商品名': '商品を識別できるようになる',
    'AmazonSKU': 'Amazon手数料の実額突合が効く',
    'JAN': '照合・名寄せの精度が上がる',
    'ASIN': 'Amazonカタログと突き合わせられる',
    '楽天SKU': '楽天の在庫照合が効く',
    '在庫行': '在庫0と未登録を区別できる',
}


def build_top20(specs, top_n=20):
    """**商品単位**で「今日やること」を組み立てる。

    1商品 = 1行。その商品で直す項目をすべて1か所へまとめる。
    毎朝このシートだけ開けば、その日の作業が分かる状態を目指す。
    """
    by_title = {sp['title']: sp for sp in specs}
    prod = defaultdict(lambda: {'items': [], 'score': 0.0, 'minutes': 0.0,
                                'cands': [], 'whys': [], 'places': set()})

    # 03 商品情報不足 — 属性の欠落を商品ごとにまとめる
    for r in by_title['03_商品情報不足']['rows']:
        star, _, pid, name, item, _cur, cand, why, conf, qty, _st, place, _ = r
        d = prod[pid]
        d['name'] = name
        d['qty'] = qty
        d['items'].append(item)
        d['score'] += qty * ITEM_WEIGHT.get(item, 1) * (1.5 if cand else 1.0)
        d['minutes'] += MINUTES.get(item, (1, 3))[0 if cand else 1]
        if cand:
            d['cands'].append(f'{item}: {cand}' + (f'({conf})' if conf else ''))
            d['whys'].append(f'{item} — {why}')
        d['places'].add(place)

    # 02 原価確認 — 同じ商品なら同じ行へ合流させる(商品単位で1回で終わらせる)
    for r in by_title['02_原価確認']['rows']:
        _ch, pid, name, cur, new, diff, _rate, qty, impact, _rows, _note, _cmd = r
        d = prod[pid]
        d.setdefault('name', str(name)[:52])
        d['qty'] = max(d.get('qty', 0), qty)
        # 単位未補正の差は TOP20 の点数に使わない(セット原価と単品原価の差かもしれない)
        if not isinstance(impact, (int, float)):
            continue
        d['items'].append('標準原価(不一致)')
        d['score'] += abs(impact) / 100
        d['minutes'] += MINUTES['標準原価(不一致)'][0]
        d['cands'].append(f'標準原価: {cur} → {new}')
        d['whys'].append(f'KPIシートの仕入値と不一致(差額{diff} / 影響額{impact})')
        d['places'].add('商品マスター E列')

    rows = []
    for pid, d in prod.items():
        n = len(d['items'])
        score = d['score']
        star = ('★★★★★' if score >= 60 else '★★★★' if score >= 20
                else '★★★' if score >= 6 else '★★' if score >= 1 else '★')
        eff = sorted({EFFECT.get(i, '') for i in d['items']} - {''})
        mins = d['minutes']
        rows.append([star, pid, d.get('name', ''),
                     f'{n}項目\n' + '\n'.join('・' + i for i in d['items']),
                     '\n'.join(d['cands']) if d['cands'] else '(候補なし — 調査が必要)',
                     '\n'.join(d['whys']),
                     '\n'.join(sorted(d['places'])),
                     d.get('qty', 0),
                     f'{mins:.0f}分' if mins >= 1 else '1分未満',
                     '\n'.join(eff), score, n])
    rows.sort(key=lambda x: -x[10])
    # 末尾に項目数を残す — 書き出し側が行の高さを決めるのに使う
    out = [[i] + r[:10] + [r[11]] for i, r in enumerate(rows[:top_n], 1)]

    # 商品に紐付かない作業は末尾へ回す(商品単位の作業を優先する)
    extra = []
    if by_title['01_新商品']['rows']:
        extra.append(['—', '★★★★★', '(一括)', f'新商品 {len(by_title["01_新商品"]["rows"])}件',
                      'マスター未登録', 'register コマンドで一括登録', '',
                      'コマンド実行', '', '数分', '分析の突合が合うようになる'])
    for r in by_title['05_Amazon手数料未反映']['rows'][:3]:
        extra.append(['—', '★★★★', r[0][:28], 'AmazonSKU未登録',
                      f'手数料 ¥{r[1]:,} が未反映', '', r[2],
                      '出品テーブル G列', '', '5分', 'Amazon手数料の実額突合が効く'])
    return out, extra


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
    hist = load_history(pair2pid, ctrl2pid, asin2pid, asku2pid)
    mall_idx = load_mall_inventory()

    fees, fee_note = build_fee_unresolved(month_sheet, tx_csv)
    fq = fee_quality(month_sheet)
    cost_rows = build_cost_conflicts(month_sheet, sold)
    monthly = load_monthly_sales(pair2pid, ctrl2pid, asin2pid, asku2pid)
    stocktake, stock_asof = build_stocktake(pm, lt, monthly, month_sheet)
    # ③ 広告改善一覧との重複を突き合わせるための対応表
    ad_spend = load_ad_spend(month_sheet, pair2pid, ctrl2pid, asin2pid, asku2pid)

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
                     '差額', '差額率(%)', f'{month_sheet}出品別個数', '影響額',
                     'KPIシート行', '注意', '適用コマンド'],
         'rows': cost_rows,
         'widths': (10, 13, 38, 12, 14, 10, 11, 14, 12, 12, 46, 40),
         'key_cols': [1],
         'how': ('単位未確認 → 出品テーブル L/M/N列へ根拠付きで記入する。'
                 '単位確認済みで差がある → push-kpi の候補一覧で判断する'),
         'note': _cost_note(cost_rows)},
        {'title': '03_商品情報不足',
         'headers': ['★', '優先度', '内部管理ID', '商品名', '不足項目', '現在値',
                     '候補(AIの提示)', '根拠', '確信度', f'{month_sheet}出荷',
                     '在庫数', '修正場所', '★の理由'],
         'rows': build_shortages(pm, lt, stock, sold, mall_idx),
         'widths': (11, 9, 13, 40, 15, 11, 22, 44, 8, 10, 10, 20, 34),
         'key_cols': [2, 4],
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
         'note': (fee_note + ' / 修正場所: 出品テーブル G列(AmazonSKU)'
                  + (f' ★K列(販売手数料)の実額化: 行ベース {fq["row_rate"]:.1%}'
                     f'({fq["act"]}行が実額 / {fq["prov"]}行が暫定15%のまま)。'
                     '同じ列に実額と暫定が混在するため、列を見ても区別できない。'
                     '**金額ベースの反映率だけ見ると品質を高く錯覚する**'
                     if fq else ''))},
        {'title': '06_廃番整理候補',
         'headers': ['内部管理ID', '商品名', '現在在庫数', '最後の販売月', '最後の仕入日',
                     '累計販売数', '累計利益', '楽天掲載', 'Amazon掲載',
                     '判定根拠', '確信度', '残す価値', f'{month_sheet}広告費', '広告との関係'],
         'rows': build_retire_candidates(pm, lt, stock, sold, sold_prev, hist,
                                         month_sheet, prev_sheet, ad_spend),
         'widths': (13, 40, 11, 12, 12, 11, 11, 9, 11, 30, 9, 30, 11, 13),
         'key_cols': [0],
         'how': '販売停止 / 取扱終了を人が判断',
         'note': ('🔴 **在庫は根拠に使えない。** 在庫管理テーブルは554行すべて在庫数0・'
                  '「未棚卸」で、実数が入っていない(2026-09-09 監査)。'
                  'この一覧の判定根拠は「2か月連続で売上0」のみ有効。'
                  '実在庫は 楽天在庫金額集計ツール / Amazon在庫リスト_import にある。'
                  '⚠️「最後の仕入日」は発注管理がOSの外にあるため取得できない。'
                  '⚠️ 再仕入予定も未確認のため誤検知がある。'
                  '★ 広告費の列に金額があれば、廃番候補に広告を出し続けている')},
        {'title': '07_棚卸対象リスト',
         'headers': ['滞留区分', 'チャネル', '識別子', '商品名', 'SKU', '在庫数',
                     '標準原価', '在庫金額', '直近の販売', '所見', '在庫の出所',
                     '実棚数(記入)'],
         'rows': stocktake,
         'widths': (14, 9, 20, 40, 24, 9, 10, 12, 12, 40, 16, 12),
         'key_cols': [1, 2],
         'how': '数えた結果は 在庫管理テーブル_v1.1.xlsm「棚卸入力シート」へ記録する',
         'note': (f'**在庫があるものは全部が対象。** 売上上位だけでは足りない。'
                  f'動かない在庫こそ誰も見ないまま金額だけ残るため、'
                  f'D長期滞留({STALE_LONG}か月売上0) → C低回転({STALE_LOW}か月売上0) '
                  f'→ A直近販売あり → B通常 の順に並べ、各区分の中は在庫金額の大きい順。'
                  '⚠️ **季節商品があるため、この区分を自動的な廃棄・処分判定に使わないこと。**'
                  'あくまで「数えに行く順番」である。'
                  + (f' ★在庫データの基準日: '
                     + ' / '.join(f'{c} {d:%Y-%m-%d}' for c, d in stock_asof)
                     + '(現在在庫ではない。取得し直してから数えること)'
                     if stock_asof else '')
                  + f' ⚠️ OSが持っている月次記録は {len(monthly)}か月分'
                  + f'({" / ".join(monthly)})のため、'
                  + (f'D長期滞留({STALE_LONG}か月)の判定はまだ成立しない。'
                     if len(monthly) < STALE_LONG else '')
                  + ' ⚠️「E 要調査」は標準原価が無く在庫金額を計算できない行。'
                  '合計金額はその分だけ過小である')},
        {'title': '08_利益率 要確認',
         'headers': ['優先度', 'チャネル', '識別子', '商品名', f'{month_sheet}個数',
                     f'{month_sheet}売上', '平均単価', '仕入値', '粗利率',
                     f'{prev_sheet}の仕入値', '確認したい理由', '対応方法', '確認結果(記入)'],
         'rows': build_margin_outliers(month_sheet, prev_sheet),
         'widths': (9, 9, 20, 40, 9, 12, 11, 11, 9, 13, 56, 44, 12),
         'key_cols': [1, 2],
         'how': '商品マスター E列(標準原価)を確認する',
         'note': ('粗利率は売価と仕入値だけで決まる。売価はモールの実績なので間違いようがない。'
                  '**粗利率がおかしいときは、ほぼ仕入値が間違っている。**'
                  f'閾値(Ver.1 暫定): {MARGIN_HIGH:.0%}以上 / {MARGIN_LOW:.0%}未満 / '
                  f'前月から±{MARGIN_SHIFT:.0%}以上の変動。'
                  '⚠️ **これは「異常確定」ではなく「要確認」である。** '
                  '高粗利が実力の商品も、赤字覚悟で売っている商品も実在する。'
                  '機械は疑うところまでで、正誤は人が決める。'
                  '3か月運用して閾値を再評価する(次回見直し 2026-12)。'
                  '★ 仕入値を直したら push-kpi で該当月のKPIシートへ反映できる'
                  '(承認した行だけ)')},
        {'title': '09_配送自動化カバー率',
         'headers': ['指標', '実数', '率', '定義', '注記'],
         'rows': build_coverage(pm, sold, month_sheet),
         'widths': (38, 20, 10, 46, 60),
         'key_cols': [0],
         'how': '出荷上位から発送サイズ区分を埋める',
         'note': ('★最優先目標「配送ラベル自動作成」の進み具合を測る唯一の指標。'
                  '入力率だけでは「効く整備をしているか」が分からないため、'
                  '出荷個数カバー率を主KPIとする(2026-09-09 ChatGPT承認)')},
        {'title': '10_月次経費(黄色セル)',
         'headers': ['チャネル', '費目', 'セル', '取得元', '入力者', '埋めないとどうなるか',
                     '入力済(記入)'],
         'rows': build_expense_cells(month_sheet),
         'widths': (10, 20, 12, 44, 40, 40, 12),
         'key_cols': [0, 1],
         'how': 'python3 apply_monthly_expenses.py set <月> <チャネル> <費目> <金額>',
         'note': ('KPIシートの黄色セル。**埋まるまで限界利益は確定しない。**'
                  '広告費はAIが分析から算出できるが、それ以外は人が請求書・RMS・'
                  'セラーセントラルから転記する。'
                  '取得元と入力者の定義は apply_monthly_expenses.py が正本')},
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

    # 一覧の行数だけ下へ送る。**行番号を決め打ちしない**
    # (シートが7枚から10枚に増えたとき、内訳ブロックが一覧に重なっていた)
    base = 6 + len(specs) + 1
    sev = Counter(r[1] for r in specs[2]['rows'])
    ws.cell(base, 1).value = '03_商品情報不足の内訳'
    ws.cell(base, 1).font = BOLD
    for i, k in enumerate(['最優先', '高', '中', '低'], base + 1):
        ws.cell(i, 1).value = k
        ws.cell(i, 2).value = sev.get(k, 0)
    fl = base + 6
    ws.cell(fl, 1).value = ('月次標準フロー: ⓪列構成チェック → ①分析 → ②要対応一覧を確定 → '
                            '③register → ④月次経費 → ⑤広告改善一覧 → ⑥一覧を再生成')
    ws.cell(fl, 1).font = BOLD
    for col, w in zip('ABCDEFGH', (5, 26, 8, 8, 8, 8, 8, 44)):
        ws.column_dimensions[col].width = w

    # ★ 毎日開くのはこのシートだけ。全一覧から「今やるべきこと」を20件だけ抜く
    top, extra = build_top20(specs)
    ws = wb.create_sheet('今日やること_TOP20', 0)
    ws['A1'] = f'今日やること TOP20 — {month_sheet}'
    ws['A1'].font = Font(bold=True, size=14)
    ws['A2'] = ('★ 毎朝このシートだけ開けばよい。**1商品=1行**にまとめてある。'
                '上から順に片付ければ会社が改善する')
    ws['A3'] = f'作成 {date.today():%Y-%m-%d} / ⚠️ 反映は人の承認後(AIは候補まで)'
    head = ['順', '★', '内部管理ID', '商品名', '直す項目(この商品でまとめて)',
            'AI候補', '根拠', '修正場所', '当月出荷', '作業時間目安', '改善効果', '完了']
    for i, h in enumerate(head, 1):
        c = ws.cell(5, i); c.value = h; c.font = BOLD; c.fill = HEAD_FILL
        c.alignment = Alignment(vertical='center', wrap_text=True)
    r = 6
    wrap = Alignment(vertical='top', wrap_text=True)
    for row in top:
        for i, v in enumerate(row[:-1], 1):     # 末尾の項目数は表示しない
            c = ws.cell(r, i); c.value = v; c.alignment = wrap
        if str(row[1]).count('★') >= 5:
            ws.cell(r, 2).fill = S_FILL
        elif str(row[1]).count('★') == 4:
            ws.cell(r, 2).fill = A_FILL
        # セル内改行は行の高さを自動で広げないため、項目数から高さを決める
        ws.row_dimensions[r].height = max(32, 14 * (row[-1] + 1))
        r += 1
    r += 1
    ws.cell(r, 1).value = '── 商品に紐付かない作業(まとめて処理できる)──'
    ws.cell(r, 1).font = BOLD
    r += 1
    for row in extra:
        for i, v in enumerate(row, 1):
            c = ws.cell(r, i); c.value = v; c.alignment = wrap
        r += 1
    total_min = sum(float(str(x[9]).rstrip('分') or 0)
                    for x in top if str(x[9])[:1].isdigit())
    ws.cell(r + 1, 1).value = (f'TOP{len(top)}をすべて片付けると約{total_min:.0f}分。'
                              '上から順に進めればよい')
    ws.cell(r + 1, 1).font = BOLD
    for col, w in zip('ABCDEFGHIJKL', (5, 11, 13, 34, 34, 40, 44, 26, 10, 12, 34, 8)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = 'A6'

    # 月次改善率
    ws = wb.create_sheet('月次改善率')
    ws['A1'] = f'月次改善率 — {prev_sheet} → {month_sheet}'
    ws['A1'].font = Font(bold=True, size=14)
    ws['A2'] = ('毎月これを見れば、会社がどれだけ改善したかが分かる。'
                if prev_path else '⚠️ 前月ファイルが無いため今回は比較できない(次回から有効)')
    head = ['一覧', '先月の問題件数', '今月の問題件数', '改善(解決)件数',
            '改善率(%)', '新たに増えた問題', '継続']
    for i, h in enumerate(head, 1):
        c = ws.cell(4, i); c.value = h; c.font = BOLD; c.fill = HEAD_FILL
    tp = tc = tr = tn = 0
    for r, st in enumerate(stats, 5):
        prev = st['prev']
        rate = (round(st['resolved'] / prev * 100, 1)
                if prev not in (None, 0) and st['resolved'] is not None else '')
        vals = [st['title'], prev if prev is not None else '—', st['count'],
                st['resolved'] if st['resolved'] is not None else '—',
                rate, st['new'], st['carried'] if st['carried'] is not None else '—']
        for i, v in enumerate(vals, 1):
            ws.cell(r, i).value = v
        if prev is not None:
            tp += prev; tc += st['count']; tr += st['resolved']; tn += st['new']
    r = 5 + len(stats) + 1
    ws.cell(r, 1).value = '合計'; ws.cell(r, 1).font = BOLD
    if prev_path:
        for i, v in enumerate([tp, tc, tr,
                               round(tr / tp * 100, 1) if tp else '', tn], 2):
            c = ws.cell(r, i); c.value = v; c.font = BOLD
    ws.cell(r + 2, 1).value = (
        '改善率 = 解決件数 ÷ 先月の問題件数。'
        '「新たに増えた問題」が解決を上回る月は、原因を先に潰す')
    for col, w in zip('ABCDEFG', (26, 15, 15, 15, 12, 16, 10)):
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
    print(f"     └ 03の優先度: 最優先{sev.get('最優先',0)} / 高{sev.get('高',0)}"
          f" / 中{sev.get('中',0)} / 低{sev.get('低',0)}")
    st5 = Counter(r[0] for r in specs[2]['rows'])
    print(f"     └ 03の★    : " + " / ".join(
        f"{k}{st5.get(k,0)}" for k in ('★★★★★', '★★★★', '★★★', '★★', '★')))
    print(f'\n  → {path}')


if __name__ == '__main__':
    main()
