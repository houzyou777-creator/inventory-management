# -*- coding: utf-8 -*-
"""sync_cost_master.py — 原価一元管理の同期ツール

商品マスターを「原価の正」とするための同期スクリプト。
KPI管理シートの月次シートと商品マスターを突き合わせる。

使い方:
    python3 sync_cost_master.py check 7月            # 突合レポートのみ(変更なし)
    python3 sync_cost_master.py register 7月         # 未登録商品をマスターへ追加登録
    python3 sync_cost_master.py adopt 7月            # 原価不一致をKPI側の値でマスター更新
                                                     # (ユーザーが明示承認した場合のみ実行すること)
    python3 sync_cost_master.py push                 # 商品マスター→在庫集計ツール原価マスターへ配信
    python3 sync_cost_master.py push-kpi 8月         # 商品マスター→既存月のKPIシートL列(候補提示のみ)
    python3 sync_cost_master.py push-kpi 8月 --apply <承認済み一覧.xlsx>   # 承認分だけ反映

    python3 sync_cost_master.py check-amazon 7月     # Amazon KPIシートとの突合レポート
    python3 sync_cost_master.py register-amazon 7月  # Amazon側の未登録商品をマスターへ追加登録
    python3 sync_cost_master.py adopt-amazon 7月     # Amazon側の原価不一致をKPI側でマスター更新(要承認)

注意: push は openpyxl ではなく Excel(AppleScript)経由で書き込む。
xlsm を openpyxl で保存するとボタン描画(drawing1.xml)が失われるため。
GS移行時はこの層を Apps Script の直接参照に置き換える。

設計方針(GS移行を見据えて):
- 照合キーは「楽天商品管理番号 × 楽天SKU」のペア(SKU単独照合はNG。使い回しがあるため)
- マスターの既存原価は絶対に上書きしない。不一致は一覧表示してユーザー判断に委ねる
- 追加登録行は備考に出所を残し、後から見分けられるようにする
"""
import os
import re
import shutil
import sys
from datetime import date

from openpyxl import Workbook, load_workbook

BASE = '/Users/hide0726/Desktop/Claude Code/MomijiStore_OS'
MASTER_FILE = BASE + '/01_InventoryManagement/SourceData/商品マスター_単品_v1.0.xlsx'
KPI_FILE = BASE + '/02_Analytics/SourceData/楽天運営 KPI管理シート.xlsx'

SH_MASTER = '商品マスター'
SH_LISTING = '出品テーブル'

# 商品マスター列 (1始まり)
PM_ID, PM_JAN, PM_NAME, PM_TYPE, PM_COST, PM_CH, PM_STATUS, PM_REG, PM_UPD, PM_NOTE = range(1, 11)
# 出品テーブル列
LT_ID, LT_PID, LT_CH, LT_RCTRL, LT_RSKU, LT_ASIN, LT_ASKU, LT_PRICE, LT_STATUS, LT_REG, LT_NOTE = range(1, 12)


def norm(v):
    return str(v).strip().upper() if v is not None and str(v).strip() else ''


def load_master():
    """マスターを読み、(管理番号,SKU)ペア→内部管理ID と ID→標準原価 を返す"""
    wb = load_workbook(MASTER_FILE, data_only=True)
    cost_by_pid = {}
    for r in wb[SH_MASTER].iter_rows(min_row=2, values_only=True):
        if r[0]:
            cost_by_pid[str(r[0])] = r[PM_COST - 1]
    pair2pid, ctrl2pid = {}, {}
    for r in wb[SH_LISTING].iter_rows(min_row=2, values_only=True):
        pid, rctrl, rsku = r[LT_PID - 1], r[LT_RCTRL - 1], r[LT_RSKU - 1]
        if not pid or not rctrl:
            continue
        pair2pid.setdefault((norm(rctrl), norm(rsku)), str(pid))
        ctrl2pid.setdefault(norm(rctrl), str(pid))
    return cost_by_pid, pair2pid, ctrl2pid


def load_jan_by_pid():
    """内部管理ID → 商品マスターのJAN(文字列)。構成違いの検出に使う(2026-09-12 Z-4)。"""
    wb = load_workbook(MASTER_FILE, read_only=True, data_only=True)
    out = {}
    for r in wb[SH_MASTER].iter_rows(min_row=2, values_only=True):
        if r and r[0] and r[PM_JAN - 1] not in (None, ''):
            j = r[PM_JAN - 1]
            out[str(r[0])] = str(int(j)) if isinstance(j, float) and j.is_integer() else str(j).strip()
    wb.close()
    return out


def load_kpi_month(month_sheet):
    """KPI月次シートから 出品(管理番号×SKU)ごとに1件へまとめて返す。

    **個数とシート行番号を持たせる**(2026-09-10 承認R)。
    影響額は「その出品が何個売れたか」で計算しなければならない。
    内部管理ID全体の個数を使うと、同じIDに複数の出品がある商品で
    何倍にも膨らむ(実例: 納豆菌P000334は5出品あり、26個を3回掛けていた)。

    同じ出品が複数行に分かれている場合は個数を合算する。
    ただし仕入値が行によって違うときは**曖昧**として印を付け、金額を出さない。
    """
    wb = load_workbook(KPI_FILE, data_only=True)
    ws = wb[month_sheet]
    out, idx = [], {}
    for row in range(8, ws.max_row + 1):
        ctrl = ws.cell(row, 3).value
        if ctrl is None:          # データ末尾(合計行の手前)まで
            break
        sku = str(ws.cell(row, 5).value).strip() if ws.cell(row, 5).value else ''
        key = (norm(ctrl), norm(sku))
        units = ws.cell(row, 8).value
        units = units if isinstance(units, (int, float)) else 0
        cost = ws.cell(row, 12).value
        if key in idx:
            d = out[idx[key]]
            d['units'] += units
            d['sheet_rows'].append(row)
            d['cost_variants'].add(cost)
            continue
        idx[key] = len(out)
        out.append({
            'ch': '楽天', 'month': month_sheet, 'key': str(ctrl).strip(),
            'ctrl': str(ctrl).strip(), 'sku': sku,
            'name': str(ws.cell(row, 1).value or '').strip(),
            'pn': str(ws.cell(row, 4).value or '').strip(),
            'cost': cost, 'units': units,
            'sheet_rows': [row], 'cost_variants': {cost},
        })
    return out


def classify(rows, cost_by_pid, pair2pid, ctrl2pid):
    """KPI行をマスターと突合して 一致/不一致/未登録 に分類(ペア単位で重複排除)"""
    match, conflict, missing, seen = [], [], [], set()
    for r in rows:
        key = (norm(r['ctrl']), norm(r['sku']))
        if key in seen:
            continue
        seen.add(key)
        pid = pair2pid.get(key) or ctrl2pid.get(key[0])
        if pid is None:
            missing.append(r)
        else:
            mc = cost_by_pid.get(pid)
            if mc is not None and r['cost'] is not None and float(mc) == float(r['cost']):
                match.append(r)
            else:
                conflict.append((r, pid, mc))
    return match, conflict, missing


def extract_jan(pn):
    """商品番号欄から13桁JANらしき先頭数字列を取り出す(なければ空)"""
    m = re.match(r'(\d{13})', pn)
    return m.group(1) if m else ''


def register_missing(month_sheet, missing):
    """未登録商品を商品マスター+出品テーブルへ追加登録する(既存行は変更しない)"""
    wb = load_workbook(MASTER_FILE)
    wm, wl = wb[SH_MASTER], wb[SH_LISTING]
    pids = [str(r[0]) for r in wm.iter_rows(min_row=2, values_only=True) if r[0]]
    cids = [str(r[0]) for r in wl.iter_rows(min_row=2, values_only=True) if r[0]]
    pn_num = max(int(p[1:]) for p in pids)
    cn_num = max(int(c[1:]) for c in cids)
    today = date.today().isoformat()
    note = f'KPI{month_sheet}シートから自動登録'
    mrow, lrow = wm.max_row + 1, wl.max_row + 1
    for r in missing:
        pn_num += 1
        cn_num += 1
        pid, cid = f'P{pn_num:06d}', f'C{cn_num:06d}'
        ptype = 'セット' if 'セット' in r['name'] else '単品'
        vals_m = [pid, extract_jan(r['pn']) or None, r['name'], ptype, r['cost'],
                  '楽天', '販売中', today, today, note]
        for col, v in enumerate(vals_m, start=1):
            wm.cell(mrow, col).value = v
        vals_l = [cid, pid, '楽天', r['ctrl'], r['sku'] or r['ctrl'], None, None,
                  None, '販売中', today, note]
        for col, v in enumerate(vals_l, start=1):
            wl.cell(lrow, col).value = v
        mrow += 1
        lrow += 1
    wb.save(MASTER_FILE)
    return pn_num, cn_num


def adopt_kpi_costs(month_sheet, conflict):
    """不一致分のマスター標準原価をKPI側の値で更新する(ユーザー承認済みの場合のみ)"""
    wb = load_workbook(MASTER_FILE)
    wm = wb[SH_MASTER]
    target = {pid: r['cost'] for r, pid, _ in conflict}
    today = date.today().isoformat()
    updated = 0
    for row in range(2, wm.max_row + 1):
        pid = str(wm.cell(row, PM_ID).value or '')
        if pid in target:
            wm.cell(row, PM_COST).value = target[pid]
            wm.cell(row, PM_UPD).value = today
            old_note = wm.cell(row, PM_NOTE).value
            note = f'原価をKPI{month_sheet}に合わせて更新'
            wm.cell(row, PM_NOTE).value = f'{old_note} / {note}' if old_note else note
            updated += 1
    wb.save(MASTER_FILE)
    return updated


TOOL_FILE = BASE + '/01_InventoryManagement/SourceData/楽天在庫金額集計ツール_v1.0.xlsm'
SH_TOOL_COST = '原価マスター'
AMZ_KPI_FILE = BASE + '/02_Analytics/SourceData/Amazon運営 KPI管理シート.xlsx'
AMZ_INV_FILE = BASE + '/01_InventoryManagement/SourceData/Amazon在庫リスト_import.xlsx'


def load_master_amazon():
    """マスターを読み、ASIN/AmazonSKU→内部管理ID と ID→標準原価 を返す"""
    wb = load_workbook(MASTER_FILE, data_only=True)
    cost_by_pid = {}
    for r in wb[SH_MASTER].iter_rows(min_row=2, values_only=True):
        if r[0]:
            cost_by_pid[str(r[0])] = r[PM_COST - 1]
    asin2pid, asku2pid = {}, {}
    for r in wb[SH_LISTING].iter_rows(min_row=2, values_only=True):
        pid, asin, asku = r[LT_PID - 1], r[LT_ASIN - 1], r[LT_ASKU - 1]
        if not pid:
            continue
        if asin:
            asin2pid.setdefault(norm(asin), str(pid))
        if asku:
            asku2pid.setdefault(norm(asku), str(pid))
    return cost_by_pid, asin2pid, asku2pid


def load_amazon_kpi_month(month_sheet):
    """Amazon KPI月次シートから 出品(ASIN×SKU)ごとに1件へまとめて返す。

    load_kpi_month と同じ理由で個数とシート行番号を持たせる(承認R)。
    """
    wb = load_workbook(AMZ_KPI_FILE, data_only=True)
    ws = wb[month_sheet]
    out, idx = [], {}
    for row in range(8, ws.max_row + 1):
        asin = ws.cell(row, 3).value
        if asin is None:
            break
        sku = str(ws.cell(row, 2).value or '').strip()
        key = (norm(asin), norm(sku))
        units = ws.cell(row, 8).value
        units = units if isinstance(units, (int, float)) else 0
        cost = ws.cell(row, 12).value
        if key in idx:
            d = out[idx[key]]
            d['units'] += units
            d['sheet_rows'].append(row)
            d['cost_variants'].add(cost)
            continue
        idx[key] = len(out)
        out.append({
            'ch': 'Amazon', 'month': month_sheet, 'key': str(asin).strip(),
            'asin': str(asin).strip(), 'sku': sku,
            'name': str(ws.cell(row, 1).value or '').strip(),
            'cost': cost, 'units': units,
            'sheet_rows': [row], 'cost_variants': {cost},
        })
    return out


def classify_amazon(rows, cost_by_pid, asin2pid, asku2pid):
    match, conflict, missing = [], [], []
    for r in rows:
        pid = asin2pid.get(norm(r['asin'])) or asku2pid.get(norm(r['sku']))
        if pid is None:
            missing.append(r)
            continue
        mc = cost_by_pid.get(pid)
        if mc is not None and r['cost'] is not None and float(mc) == float(r['cost']):
            match.append(r)
        else:
            conflict.append((r, pid, mc))
    return match, conflict, missing


def load_amazon_jan_lookup():
    """Amazon在庫リストから ASIN→JAN を引く(マスター登録時のJAN補完用)"""
    wb = load_workbook(AMZ_INV_FILE, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    jan_by_asin = {}
    for r in ws.iter_rows(min_row=2, values_only=True):
        jan, asin = r[0], r[1]
        if asin and jan:
            jan_by_asin.setdefault(norm(asin), str(jan).strip())
    return jan_by_asin


def register_missing_amazon(month_sheet, missing):
    """Amazon側の未登録商品を商品マスター+出品テーブルへ追加登録する"""
    jan_by_asin = load_amazon_jan_lookup()
    wb = load_workbook(MASTER_FILE)
    wm, wl = wb[SH_MASTER], wb[SH_LISTING]
    pn_num = max(int(str(r[0])[1:]) for r in wm.iter_rows(min_row=2, values_only=True) if r[0])
    cn_num = max(int(str(r[0])[1:]) for r in wl.iter_rows(min_row=2, values_only=True) if r[0])
    today = date.today().isoformat()
    note = f'Amazon KPI{month_sheet}シートから自動登録'
    mrow, lrow = wm.max_row + 1, wl.max_row + 1
    for r in missing:
        pn_num += 1
        cn_num += 1
        pid, cid = f'P{pn_num:06d}', f'C{cn_num:06d}'
        ptype = 'セット' if 'セット' in r['name'] else '単品'
        vals_m = [pid, jan_by_asin.get(norm(r['asin'])), r['name'], ptype, r['cost'],
                  'Amazon', '販売中', today, today, note]
        for col, v in enumerate(vals_m, start=1):
            wm.cell(mrow, col).value = v
        vals_l = [cid, pid, 'Amazon', None, None, r['asin'], r['sku'] or None,
                  None, '販売中', today, note]
        for col, v in enumerate(vals_l, start=1):
            wl.cell(lrow, col).value = v
        mrow += 1
        lrow += 1
    wb.save(MASTER_FILE)
    return pn_num, cn_num


def build_push_rows():
    """商品マスター+出品テーブルから集計ツール原価マスター用の行を作る
    (JAN, 楽天商品管理番号, 商品名, 商品種別, 標準原価, 最終更新日, 備考)"""
    wb = load_workbook(MASTER_FILE, data_only=True)
    products = {}
    for r in wb[SH_MASTER].iter_rows(min_row=2, values_only=True):
        if r[0]:
            products[str(r[0])] = r
    rctrl_by_pid = {}
    for r in wb[SH_LISTING].iter_rows(min_row=2, values_only=True):
        pid, ch, rctrl = r[LT_PID - 1], r[LT_CH - 1], r[LT_RCTRL - 1]
        if pid and rctrl and ch and '楽天' in str(ch):
            rctrl_by_pid.setdefault(str(pid), str(rctrl).strip())
    rows, seen = [], set()
    for pid, p in products.items():
        jan = str(p[PM_JAN - 1]).strip() if p[PM_JAN - 1] else ''
        rctrl = rctrl_by_pid.get(pid, '')
        cost = p[PM_COST - 1]
        if cost is None or (jan == '' and rctrl == ''):
            continue
        key = (jan, rctrl)
        if key in seen:
            continue
        seen.add(key)
        upd = p[PM_UPD - 1]
        upd = upd.isoformat()[:10] if hasattr(upd, 'isoformat') else (str(upd)[:10] if upd else '')
        rows.append([jan, rctrl, str(p[PM_NAME - 1] or ''), str(p[PM_TYPE - 1] or '単品'),
                     float(cost), upd, '商品マスターから同期'])
    return rows


def push_to_tool(rows):
    """Excel(AppleScript)経由で集計ツールの原価マスターシートへ書き込む"""
    import subprocess
    import tempfile

    def esc(v):
        if isinstance(v, float):
            return str(int(v)) if v == int(v) else str(v)
        return '"' + str(v).replace('\\', '\\\\').replace('"', '\\"') + '"'

    # 1行が長くなりすぎないよう100行ずつ set value する(AppleScriptの複数行リストは書けないため)
    CHUNK = 100
    stmts = []
    for i in range(0, len(rows), CHUNK):
        chunk = rows[i:i + CHUNK]
        data = ','.join('{' + ','.join(esc(v) for v in row) + '}' for row in chunk)
        r1, r2 = 3 + i, 2 + i + len(chunk)
        stmts.append(f'set value of range ("A{r1}:G{r2}") of ws to {{{data}}}')
    body = '\n    '.join(stmts)
    # 集計シートが10万行あるため、書き込み中は手動計算にして最後に一括再計算する
    script = f'''
set p to POSIX file "{TOOL_FILE}"
with timeout of 900 seconds
tell application "Microsoft Excel"
    set wasRunning to running
    open p
    delay 2
    set wb to active workbook
    set ws to worksheet "{SH_TOOL_COST}" of wb
    set calculation to calculation manual
    set screen updating to false
    clear contents of range "A3:G10000" of ws
    -- JAN(A)・楽天商品管理番号(B)は**代入の前に**文字列書式にする(2026-09-12 §7-2)。
    -- 標準書式のまま set value すると Excel が手入力と同じ解釈をし、先頭0が落ちる(102件が12桁になっていた)
    set number format of range "A3:B10000" of ws to "@"
    {body}
    set calculation to calculation automatic
    calculate
    set screen updating to true
    save wb
    close wb saving no
    if not wasRunning then quit
end tell
end timeout
return "push done: {len(rows)} rows"
'''
    with tempfile.NamedTemporaryFile('w', suffix='.applescript', delete=False) as f:
        f.write(script)
        path = f.name
    r = subprocess.run(['osascript', path], capture_output=True, text=True, timeout=960)
    if r.returncode != 0:
        raise RuntimeError(f'AppleScript failed: {r.stderr}')
    return r.stdout.strip()


def plan_tool_ids():
    """在庫ツール「原価マスター」の JAN/管理番号 を、正本(商品マスター・出品テーブル)の完全な識別子へ戻す候補(§7-2)。

    対応の決め方: 正本から push と同じ手順で作った行と、
      JAN・管理番号(先頭0と型を無視して一致)＋商品名(完全一致)＋標準原価(一致)
    が **1対1** のときだけ。12桁だからという理由で先頭0を足すことはしない(正本にある文字列をそのまま使う)。
    原価金額は変更しない。
    """
    want = build_push_rows()                    # [jan, rctrl, name, type, cost, upd, note]
    def d0(v):
        c = str(int(v)) if isinstance(v, float) and v.is_integer() else str(v if v is not None else '').strip()
        return c.lstrip('0') if c.isdigit() else c
    from collections import Counter
    def sig(jan, rctrl, name, cost):
        return (d0(jan), d0(rctrl), str(name or '').strip(), float(cost) if isinstance(cost, (int, float)) else None)
    idx = {}
    cnt = Counter()
    for w in want:
        k = sig(w[0], w[1], w[2], w[4]); idx.setdefault(k, w); cnt[k] += 1
    wb = load_workbook(TOOL_FILE, read_only=True, data_only=True)
    ws = wb[SH_TOOL_COST]
    out = []
    for rno, r in enumerate(ws.iter_rows(min_row=3, values_only=True), 3):
        if not r or all(v in (None, '') for v in r[:2]):
            continue
        jan, rctrl, name, cost = r[0], r[1], r[2], r[4]
        k = sig(jan, rctrl, name, cost)
        for col, cur, letter, label in ((0, jan, 'A', 'JAN'), (1, rctrl, 'B', '楽天商品管理番号')):
            if cur in (None, ''):
                continue
            rec = {'行': rno, '列': label, '現在値': str(int(cur)) if isinstance(cur, float) and cur.is_integer() else str(cur),
                   '正本の値': '', '商品名': str(name or '')[:40], '標準原価': cost, '更新結果': '', '_col': letter, '_cur': cur}
            w = idx.get(k)
            if w is None:
                rec['更新結果'] = '保留(正本に同じ JAN・管理番号・商品名・原価 の行が無い)'
            elif cnt[k] != 1:
                rec['更新結果'] = f'保留(正本で {cnt[k]} 行が同じ項目。対応が曖昧)'
            else:
                tgt = str(w[col]).strip()
                rec['正本の値'] = tgt
                if isinstance(cur, str) and cur == tgt:
                    continue                                # 既に一致。候補にしない
                if d0(cur) != d0(tgt):
                    rec['更新結果'] = '保留(先頭0を除いても数字が一致しない)'
                elif tgt == '' :
                    rec['更新結果'] = '保留(正本が空)'
                elif not tgt.startswith('0'):
                    # 承認範囲は「先頭0が落ちた識別子」(§7-2)。型だけの違いは次回の push(文字列書式)で揃う
                    rec['更新結果'] = '対象外(型のみ。承認範囲外。次回pushで文字列になる)'
                else:
                    rec['更新結果'] = '適用可'
            out.append(rec)
    wb.close()
    return out


def apply_tool_ids(plan, mode):
    """§7-2 の候補を在庫ツールの原価マスターへ書く(バックアップ・Excelで文字列書込・再読込検証・ログ)。"""
    import subprocess
    from collections import Counter
    from datetime import date
    ok = [r for r in plan if r['更新結果'] == '適用可']
    print(f'■ 原価マスター(在庫ツール) 識別子の候補 {len(plan)}セル → 適用可 {len(ok)}セル')
    for k, v in Counter(r['列'] for r in ok).items():
        print(f'   {k:10} {v:3}セル')
    for k, v in Counter(r['更新結果'] for r in plan if r['更新結果'] != '適用可').items():
        print(f'   保留 {v:3}  {k}')
    lead0 = sum(1 for r in ok if r['正本の値'].startswith('0'))
    print(f'   うち先頭0が戻るもの {lead0}セル / 型だけ {len(ok) - lead0}セル')
    if mode != 'apply':
        for r in ok[:8]:
            print(f"   行{r['行']:<4} {r['列']:10} {r['現在値']:>16} → {r['正本の値']!r}  {r['商品名'][:20]}")
        return
    bdir = os.path.dirname(TOOL_FILE) + '/Backup'
    os.makedirs(bdir, exist_ok=True)
    base = os.path.splitext(os.path.basename(TOOL_FILE))[0]
    bpath = f'{bdir}/{base}_backup_{date.today():%Y%m%d}_原価マスター識別子復元前.xlsm'
    seq = 2
    while os.path.exists(bpath):
        bpath = f'{bdir}/{base}_backup_{date.today():%Y%m%d}_原価マスター識別子復元前_{seq}.xlsm'; seq += 1
    import shutil
    shutil.copy2(TOOL_FILE, bpath)
    print(f'\n✅ バックアップ: {bpath}')

    def lit(v):
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            return str(int(v)) if isinstance(v, float) and v.is_integer() else repr(v)
        return '"' + str(v).replace('\\', '\\\\').replace('"', '\\"') + '"'
    lines = []
    for r in ok:
        chk = (f'(class of (value of c)) is real and (value of c) is equal to {lit(r["_cur"])}'
               if isinstance(r['_cur'], (int, float)) else
               f'(class of (value of c)) is not real and ((value of c) as string) is equal to {lit(str(r["_cur"]))}')
        lines.append(f'''
    set c to cell "{r['_col']}{r['行']}" of ws
    if {chk} then
        set number format of c to "@"
        set value of c to {lit(r['正本の値'])}
        set end of res to "{r['_col']}{r['行']}:written"
    else
        set end of res to "{r['_col']}{r['行']}:mismatch:" & ((value of c) as string)
    end if''')
    script = f'''
set p to POSIX file "{TOOL_FILE}"
set res to {{}}
with timeout of 900 seconds
tell application "Microsoft Excel"
    open p
    delay 1
    set wb to active workbook
    set ws to worksheet "{SH_TOOL_COST}" of wb
    {''.join(lines)}
    save wb
    close wb saving no
end tell
end timeout
set AppleScript's text item delimiters to linefeed
return res as string
'''
    rr = subprocess.run(['osascript', '-e', script], capture_output=True, text=True, timeout=960)
    if rr.returncode != 0:
        raise RuntimeError(f'Excelでの書込に失敗: {rr.stderr}')
    res = {}
    for line in rr.stdout.strip().split('\n'):
        if ':' in line:
            addr, _, rest = line.partition(':'); res[addr] = rest
    wb = load_workbook(TOOL_FILE, read_only=True, data_only=True)
    after = {rno: r[:2] for rno, r in enumerate(wb[SH_TOOL_COST].iter_rows(min_row=3, values_only=True), 3)}
    wb.close()
    n_ok = 0
    for r in ok:
        st = res.get(f"{r['_col']}{r['行']}", 'no-result')
        v = after.get(r['行'], (None, None))[0 if r['_col'] == 'A' else 1]
        if st == 'written' and isinstance(v, str) and v == r['正本の値']:
            r['更新結果'] = '更新済(文字列・再読込で一致)'; n_ok += 1
        elif st == 'written':
            r['更新結果'] = f'⚠️ 書込後の再読込が一致しない: {v!r}'
        else:
            r['更新結果'] = f'スキップ(書込直前の照合で不一致: {st})'
    print(f'✅ 更新 {n_ok}/{len(ok)}セル')
    out_dir = os.path.dirname(TOOL_FILE) + '/Output'
    log = f'{out_dir}/原価マスター_識別子復元ログ_{date.today():%Y%m%d}.csv'
    import csv as _csv
    with open(log, 'w', encoding='utf-8-sig', newline='') as f:
        w = _csv.writer(f)
        w.writerow(['行', '列', '現在値', '正本の値', '商品名', '標準原価(変更なし)', '更新結果'])
        for r in plan:
            w.writerow([r['行'], r['列'], r['現在値'], r['正本の値'], r['商品名'], r['標準原価'], r['更新結果']])
    print(f'   → {log}')


# ═══ push-kpi ═════════════════════════════════════════════════════════════
#  商品マスター → 既存月のKPIシート L列(仕入値)
#
#  ⚠️ **無条件の上書きは禁止**(2026-09-09 ChatGPT承認Kの条件)。
#     過去月のKPIシートはその月の実績記録である。原価を直す理由は2つあり、
#     扱いが正反対になる:
#       ① 仕入価格が変わった   → 過去月は触らない(その月はその原価で売った)
#       ② 元の入力が間違っていた → 過去月も直す(粗利が誤っているため)
#     どちらかを機械が判定することはできない。**人が1件ずつ決める。**
PUSH_KPI_DIR = BASE + '/01_InventoryManagement/SourceData/Output'


# 出品テーブルに「販売入数 / 原価単位 / 確認根拠」があればそこから読む。
# **無ければ推測しない。**(2026-09-10 承認Q)
#   原価の比率や行数から単位を推し量ってはならない。
#   比率が整数に乗るのは単位一致の証拠ではなく、たまたまかもしれない。
PACK_COLS = {'販売入数': None, '原価単位': None, '確認根拠': None}
UNIT_SINGLE = '単品原価'      # マスターの標準原価が単品1個分を指す
UNIT_SET = 'セット原価'       # マスターの標準原価がこの出品と同じ構成のセットを指す


# 確認根拠(N列)の**明示トークン**。自由記述の中に `項目=値` の形で書く(区切りは / または改行)。
#   含有範囲=確認済み(付属品なし)      … 付属品等の含有範囲を確認した   ← RMS商品番号の原価採用に必須
#   含有範囲=未確認 / 含有範囲=資料待ち … 明示的に未確認(採用しない)
#   RMS原価対象月=2026-08〜            … 番号に埋め込まれた原価が適用できる月(〜は以降ずっと)
#   RMS原価対象月=2026-08,2026-09      … 列挙も可
# ※ 過去月の確認済み原価を後の月へ使う判断は、N列のトークンではなく **原価確認記録**(cost_confirmations.py)の
#    適用期間(開始月〜終了月)で持つ(2026-09-12 §7-5: 確認日だけで無期限に引き継がない)
# ⚠️ 「含有範囲」という文字があるだけでは採用しない(2026-09-12 ChatGPT指摘)。
#    値が「確認済」で始まるものだけを確認済みとし、「未確認」「資料待ち」を含めば必ず除外する。
UNCONFIRMED_WORDS = ('未確認', '資料待ち', '未定', '不明')


def parse_proof(text):
    """確認根拠の自由記述から `項目=値` を抜き出す。{項目: 値} を返す(無ければ空)。"""
    import re as _re
    out = {}
    for m in _re.finditer(r'([^\s/／=＝:：]+)\s*[=＝]\s*([^/／\n]+)', str(text or '')):
        out.setdefault(m.group(1).strip(), m.group(2).strip())
    return out


def proof_status(text, item):
    """項目の確認状態を返す: 'confirmed' / 'unconfirmed' / 'absent'。

    確認済みと言えるのは 値が「確認済」で始まり、未確認語を含まないときだけ。
    """
    v = parse_proof(text).get(item)
    if v is None:
        return 'absent', None
    if any(w in v for w in UNCONFIRMED_WORDS):
        return 'unconfirmed', v
    if v.startswith('確認済') or v.startswith('可'):
        return 'confirmed', v
    return 'unconfirmed', v


def proof_months(text, item='RMS原価対象月'):
    """`RMS原価対象月=2026-08〜` / `2026-08,2026-09` を読む。(月の集合, 以降ずっとの開始月)。"""
    import re as _re
    v = parse_proof(text).get(item) or ''
    months = set(_re.findall(r'\d{4}-\d{2}', v))
    open_from = None
    m = _re.search(r'(\d{4}-\d{2})\s*[〜~～]', v)
    if m:
        open_from = m.group(1)
    return months, open_from


def month_covered(text, ym, item='RMS原価対象月'):
    """対象月 ym('2026-08') が確認根拠の対象月に含まれるか。**記載が無ければ False。**"""
    months, open_from = proof_months(text, item)
    if ym in months:
        return True
    return bool(open_from) and ym >= open_from


def load_pack_info():
    """出品(ASIN/SKU)単位の 販売入数・原価単位・確認根拠 を読む。

    列がまだ無ければ全件「未確認」を返す。**それが正しい状態である。**
    確認していない入数を使って金額を出すより、出さない方がよい。
    """
    wb = load_workbook(MASTER_FILE, read_only=True, data_only=True)
    ws = wb[SH_LISTING]
    header = [str(c or '').strip() for c in
              next(ws.iter_rows(min_row=1, max_row=1, values_only=True))]
    pos = {k: (header.index(k) if k in header else None) for k in PACK_COLS}
    out = {}
    if all(v is None for v in pos.values()):
        wb.close()
        return out, pos            # 列そのものが無い
    for r in ws.iter_rows(min_row=2, values_only=True):
        if not r or not r[1]:
            continue
        ch = r[2]
        key = norm(r[5]) if ch == 'Amazon' else norm(r[3])
        sku = norm(r[6]) if ch == 'Amazon' else norm(r[4])
        g = lambda k: (r[pos[k]] if pos[k] is not None and pos[k] < len(r) else None)
        out.setdefault((ch, key), []).append(
            {'sku': sku, 'pid': str(r[1]), 'pack': g('販売入数'),
             'unit': g('原価単位'), 'proof': g('確認根拠')})
    wb.close()
    return out, pos


def resolve_unit_cost(pack, ch, key, sku, master_cost):
    """出品1件あたりの原価を求める(2026-09-10 承認W-4)。

    ・単品原価          → 販売入数を掛ける
    ・同じ構成のセット原価 → **そのまま使う**(二重に掛けない)
    ・単位/構成/商品対応が未確認 → **None**(呼び出し側は空欄にする)

    ⚠️ 「確認済みなら一律に掛ける」ではない。セット原価に入数を掛けると
       構成の数だけ二重に膨らむ。原価単位で分岐する。
    """
    if not isinstance(master_cost, (int, float)):
        return None, 'マスターの標準原価が無い'
    info, amb = lookup_pack(pack, ch, key, sku)
    if amb:
        return None, amb
    info = info or {}
    unit, n = info.get('unit'), info.get('pack')
    if unit == UNIT_SET:
        return master_cost, ''                    # 掛けない
    if unit == UNIT_SINGLE:
        if not isinstance(n, (int, float)) or n <= 0:
            return None, '単品原価だが販売入数が未確認'
        return master_cost * n, ''
    return None, ('原価単位が未確認' if unit in (None, '') else f'原価単位が不正: {unit}')


def lookup_pack(pack, ch, key, sku):
    """出品の入数情報を引く。**引けない/一意でないなら None を返す。**

    ⚠️ Amazonは2026年8月のレポートからSKU列が消えたため、KPI側にSKUが無い。
       ASIN単独で引くしかないが、同じASINに入数の違う出品が複数ぶら下がっていたら
       どれか分からない。**そのときは推測せず「曖昧」として扱う。**
    """
    cands = pack.get((ch, key), [])
    if not cands:
        return None, '出品テーブルに該当の出品が無い'
    if sku:
        hit = [c for c in cands if c['sku'] == sku]
        if len(hit) == 1:
            return hit[0], ''
        if len(hit) > 1:
            return None, 'SKUで引いても複数該当する'
    if len(cands) == 1:
        return cands[0], ''
    # SKUで絞れないときは、**入数だけでなく 内部管理ID・原価単位 の一致まで見る**
    # (2026-09-10 承認Q-6)。どれか1つでも食い違えば、どの出品の話か決められない
    for f, label in (('pid', '内部管理ID'), ('pack', '販売入数'), ('unit', '原価単位')):
        vals = {c.get(f) for c in cands}
        if len(vals) > 1:
            return None, (f'同じ識別子に{len(cands)}件の出品があり{label}が一致しない'
                          '(SKUが無く特定できない)')
        if f in ('pack', 'unit') and None in vals:
            return None, (f'同じ識別子に{len(cands)}件の出品があり、{label}が未確認'
                          '(SKUが無く特定できない)')
    return cands[0], ''


def _ratio_note(cur, mc):
    """整数倍かどうかは**参考情報**としてだけ出す。判定には使わない。"""
    if not (isinstance(cur, (int, float)) and isinstance(mc, (int, float)) and mc):
        return ''
    r = cur / mc
    n = round(r)
    if n >= 1 and abs(r - n) <= 0.01:
        return f'参考: KPI ÷ マスター = {r:.2f}(約{n}倍)'
    return f'参考: KPI ÷ マスター = {r:.2f}'


def collect_push_kpi(month_sheet):
    """KPIシートのL列と商品マスターの標準原価を、**単位を揃えてから**比べる。

    単位を揃えられない行は金額を出さない(2026-09-10 承認Q)。
      ・販売入数が未確認
      ・原価単位(単品原価/セット原価)が未確認
      ・商品対応(出品→内部管理ID)が付かない
    いずれかに当てはまれば「単位未確認」とし、更新対象から外す。
    元のKPI値とマスター値は表示するが、差額と利益影響額は空欄にする。
    """
    cost_by_pid, pair2pid, ctrl2pid = load_master()
    _, asin2pid, asku2pid = load_master_amazon()
    pack, pos = load_pack_info()
    wbm = load_workbook(MASTER_FILE, read_only=True, data_only=True)
    names, ptype = {}, {}
    for r in wbm[SH_MASTER].iter_rows(min_row=2, values_only=True):
        if r[0]:
            names[str(r[0])] = r[2]
            ptype[str(r[0])] = r[3]
    wbm.close()

    rows = []
    for listing in load_kpi_month(month_sheet) + load_amazon_kpi_month(month_sheet):
        ch = listing['ch']
        cur = listing['cost']
        if not isinstance(cur, (int, float)):
            continue
        key = norm(listing['key'])
        sku = norm(listing['sku'])
        pid = (asin2pid.get(key) or asku2pid.get(sku)) if ch == 'Amazon' \
            else (pair2pid.get((key, sku)) or ctrl2pid.get(key))
        mc = cost_by_pid.get(pid) if pid else None

        info, amb = lookup_pack(pack, ch, key, sku)
        info = info or {}
        p_units, p_unit = info.get('pack'), info.get('unit')
        ratio = _ratio_note(cur, mc)

        # ── 単位を揃えられるか ──────────────────────────────
        if not pid:
            status, why = '単位未確認', '出品→内部管理IDの対応が付かない'
        elif not isinstance(mc, (int, float)):
            status, why = '単位未確認', 'マスターの標準原価が無い'
        elif amb:
            status, why = '単位未確認', amb
        elif p_unit not in (UNIT_SINGLE, UNIT_SET):
            status, why = '単位未確認', ('原価単位が未確認'
                                       + ('(出品テーブルに列が無い)' if pos['原価単位'] is None
                                          else ''))
        elif p_unit == UNIT_SINGLE and not isinstance(p_units, (int, float)):
            status, why = '単位未確認', '単品原価だが販売入数が未確認'
        else:
            status, why = '確認済み', ''

        if status == '単位未確認':
            rows.append([ch, listing['month'], '/'.join(map(str, listing['sheet_rows'])),
                         listing['key'], sku, pid or '', str(names.get(pid) or
                         listing['name'])[:44], listing['units'], cur,
                         mc if isinstance(mc, (int, float)) else '',
                         p_units if p_units is not None else '',
                         p_unit or '', str(ptype.get(pid) or ''),
                         '', '', '',           # 期待値/差額/影響額は空欄
                         status, why, ratio, info.get('proof') or '', '', ''])
            continue

        expected, _why2 = resolve_unit_cost(pack, ch, key, sku, mc)
        diff = cur - expected
        impact = round(-diff * listing['units'])
        rows.append([ch, listing['month'], '/'.join(map(str, listing['sheet_rows'])),
                     listing['key'], sku, pid, str(names.get(pid) or '')[:44],
                     listing['units'], cur, mc, p_units, p_unit,
                     str(ptype.get(pid) or ''), expected, diff, impact,
                     status, ('一致' if diff == 0 else
                              '少額差(低優先)' if abs(diff) <= 10 else '差あり'),
                     ratio, info.get('proof') or '', '', ''])
    # 単位確認済みで差があるものを先頭へ。その中は影響額の大きい順
    rows.sort(key=lambda x: (x[16] != '確認済み',
                             0 if isinstance(x[14], (int, float)) and x[14] else 1,
                             -abs(x[15]) if isinstance(x[15], (int, float)) else 0))
    return rows, pos


PUSH_HEAD = ['チャネル', '対象月', 'KPIシート行', '識別子', 'SKU', '内部管理ID', '商品名',
             '出品別個数', 'KPIの現在値', 'マスター標準原価', '販売入数', '原価単位',
             'マスター商品種別', '単位を揃えた期待値', '差額', '利益への影響額',
             '単位判定', '所見', '参考(倍率)', '確認根拠',
             '更新理由(記入: 入力ミス / 価格変更)', '承認(1を記入)']


def write_push_kpi_list(month_sheet, rows, pos):
    """承認欄付きの候補一覧。**単位未確認の行は差額も影響額も空欄で出す。**"""
    from openpyxl.styles import Alignment, Font, PatternFill
    os.makedirs(PUSH_KPI_DIR, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = 'KPI仕入値更新候補'
    ws['A1'] = f'KPI仕入値 更新候補 — {month_sheet}'
    ws['A1'].font = Font(bold=True, size=13)
    miss = [k for k, v in pos.items() if v is None]
    ws['A2'] = (
        '商品マスターの標準原価とKPIシートのL列(仕入値)を、**単位を揃えてから**比べた結果。\n'
        '⚠️ 単位を揃えられない行は「単位未確認」とし、差額・利益への影響額を空欄にしている。\n'
        '   比率が整数に乗ることは単位一致の証拠ではない。倍率は参考情報にすぎない。\n'
        + (f'🔴 出品テーブルに {" / ".join(miss)} の列が無いため、全行が単位未確認になる。\n'
           if miss else '')
        + '【反映するには2つとも必要】\n'
          '  ・単位判定が「確認済み」であること(単位未確認は反映側でも弾く)\n'
          '  ・更新理由へ「入力ミス」と記入すること\n'
          '⚠️ **現在の標準原価と違うというだけでは過去月を更新しない。**\n'
          '   仕入価格が変わっただけなら、その月はその原価で売っている。過去月は触らない。\n'
          '記入後: python3 sync_cost_master.py push-kpi <月> --apply <このファイル>')
    ws['A2'].alignment = Alignment(vertical='top', wrap_text=True)
    ws.row_dimensions[2].height = 150
    fill = PatternFill('solid', fgColor='D9E1F2')
    warn = PatternFill('solid', fgColor='FFF2CC')
    for i, h in enumerate(PUSH_HEAD, 1):
        c = ws.cell(4, i)
        c.value, c.font, c.fill = h, Font(bold=True), fill
        c.alignment = Alignment(vertical='top', wrap_text=True)
    for r, row in enumerate(rows, 5):
        for i, v in enumerate(row, 1):
            ws.cell(r, i).value = v
        if row[16] != '確認済み':
            ws.cell(r, 17).fill = warn
    for i, w in enumerate((9, 8, 12, 20, 20, 12, 36, 11, 12, 15, 10, 12, 14,
                           16, 10, 14, 12, 26, 24, 18, 26, 13), 1):
        ws.column_dimensions[ws.cell(4, i).column_letter].width = w
    ws.freeze_panes = 'A5'
    path = f'{PUSH_KPI_DIR}/KPI仕入値更新候補_{month_sheet}_{date.today():%Y%m%d}.xlsx'
    wb.save(path)
    return path


def apply_push_kpi(month_sheet, list_path):
    """承認された行だけを反映する。**反映側でも単位未確認を弾く。**

    候補作成側で除外していても、人が承認欄へ記入すれば通ってしまう。
    正本を書き換える側で二重に止める(2026-09-10 承認Q)。
    """
    wb = load_workbook(list_path, data_only=True)
    ws = wb.active
    head = [str(ws.cell(4, i).value or '') for i in range(1, len(PUSH_HEAD) + 2)]
    col = {h: i + 1 for i, h in enumerate(head) if h}
    need = ['チャネル', 'KPIシート行', 'KPIの現在値', '単位を揃えた期待値',
            '単位判定', '更新理由(記入: 入力ミス / 価格変更)', '承認(1を記入)']
    lack = [n for n in need if n not in col]
    if lack:
        sys.exit(f'❌ 一覧の見出しが想定と違います。足りない列: {lack}')

    approved, rejected = [], []
    r = 5
    while ws.cell(r, col['チャネル']).value:
        g = lambda n: ws.cell(r, col[n]).value
        if g('承認(1を記入)') not in (None, '', 0):
            reason = str(g('更新理由(記入: 入力ミス / 価格変更)') or '').strip()
            if g('単位判定') != '確認済み':
                rejected.append((r, '単位未確認のため反映できない'))
            elif '入力ミス' not in reason:
                rejected.append((r, f'更新理由が「入力ミス」ではない(記入: {reason or "空欄"})。'
                                    '価格変更なら過去月は更新しない'))
            else:
                approved.append({'ch': g('チャネル'), 'rows': str(g('KPIシート行')),
                                 'cur': g('KPIの現在値'), 'new': g('単位を揃えた期待値'),
                                 'key': g('識別子')})
        r += 1

    for rr, why in rejected:
        print(f'  ⛔ 一覧{rr}行目: {why}')
    if not approved:
        sys.exit('反映できる行がありません。'
                 + ('(承認はあるが上記の理由で弾きました)' if rejected else
                    '(承認欄に記入がありません)'))

    touched = {}
    done = 0
    for a in approved:
        path = KPI_FILE if a['ch'] == '楽天' else AMZ_KPI_FILE
        w = touched.get(path) or load_workbook(path)
        touched[path] = w
        sh = w[month_sheet]
        for rn in str(a['rows']).split('/'):
            cell = sh.cell(int(rn), 12)          # L列
            if cell.value != a['cur']:
                print(f'  ⚠️ {a["ch"]} 行{rn}: 現在値が {cell.value} で'
                      f'候補作成時({a["cur"]})と違う — スキップ')
                continue
            cell.value = a['new']
            done += 1
            print(f'  {a["ch"]} 行{rn} {a["key"]}: {a["cur"]} → {a["new"]}')
    for path, w in touched.items():
        bdir = os.path.dirname(path) + '/Backup'
        os.makedirs(bdir, exist_ok=True)
        base = os.path.splitext(os.path.basename(path))[0]
        shutil.copy2(path, f'{bdir}/{base}_backup_{date.today():%Y%m%d}_{month_sheet}仕入値反映前.xlsx')
        w.save(path)
    return done, list(touched)


def main_amazon(mode, month_sheet):
    cost_by_pid, asin2pid, asku2pid = load_master_amazon()
    rows = load_amazon_kpi_month(month_sheet)
    match, conflict, missing = classify_amazon(rows, cost_by_pid, asin2pid, asku2pid)
    print(f'Amazon {month_sheet}: ユニーク{len(rows)}ペア')
    print(f'  マスター一致: {len(match)} / 原価不一致: {len(conflict)} / 未登録: {len(missing)}')
    if conflict:
        if mode == 'adopt-amazon':
            print('--- 原価不一致(これからKPI側の値でマスターを更新します)')
        else:
            print('--- 原価不一致(マスターは変更していません。要ユーザー判断)')
        for r, pid, mc in conflict:
            print(f'  {r["asin"]} ({pid}) マスター:{mc} / KPI:{r["cost"]} — {r["name"][:30]}')
    if mode == 'adopt-amazon':
        if not conflict:
            print('不一致なし — 更新処理スキップ')
            return
        updated = adopt_kpi_costs(f'Amazon {month_sheet}', conflict)
        print(f'マスター原価をKPI側に更新: {updated}件')
    elif mode == 'register-amazon':
        if not missing:
            print('未登録なし — 登録処理スキップ')
            return
        last_p, last_c = register_missing_amazon(month_sheet, missing)
        print(f'登録完了: {len(missing)}件追加 (最終ID P{last_p:06d} / C{last_c:06d})')
    elif missing:
        print(f'--- 未登録 {len(missing)}件(register-amazon モードで追加登録できます)')


def main():
    if len(sys.argv) >= 2 and sys.argv[1] in ('plan-tool-ids', 'apply-tool-ids'):
        apply_tool_ids(plan_tool_ids(), 'apply' if sys.argv[1] == 'apply-tool-ids' else 'plan')
        return
    if len(sys.argv) >= 2 and sys.argv[1] == 'push':
        rows = build_push_rows()
        print(f'配信対象: {len(rows)}行(商品マスター由来)')
        print(push_to_tool(rows))
        return
    if len(sys.argv) >= 3 and sys.argv[1] == 'push-kpi':
        month_sheet = sys.argv[2]
        if '--apply' in sys.argv:
            lp = sys.argv[sys.argv.index('--apply') + 1]
            n, files = apply_push_kpi(month_sheet, lp)
            print(f'\n{n}行を反映しました。Excelで再計算してください:')
            for f in files:
                print(f'  {os.path.basename(f)}')
            return
        rows, pos = collect_push_kpi(month_sheet)
        ok = [x for x in rows if x[16] == '確認済み']
        ng = [x for x in rows if x[16] != '確認済み']
        diff_rows = [x for x in ok if isinstance(x[14], (int, float)) and x[14]]
        print(f'■ {month_sheet}: 出品 {len(rows)}件を単位を揃えて比較')
        print(f'   単位確認済み {len(ok)}件(うち差あり {len(diff_rows)}件) / '
              f'単位未確認 {len(ng)}件')
        miss = [k for k, v in pos.items() if v is None]
        if miss:
            print(f'   🔴 出品テーブルに {" / ".join(miss)} の列が無いため、'
                  '単位を揃えられません')
            print('      → 承認Sの最小構成を入れるまで、更新できる行はありません')
        if diff_rows:
            imp = sum(x[15] for x in diff_rows if isinstance(x[15], (int, float)))
            print(f'   単位補正後の利益影響額(参考): {imp:+,.0f}')
            print(f'\n   {"ch":6} {"識別子":16} {"個数":>4} {"KPI":>8} {"期待値":>8} '
                  f'{"差額":>8} {"影響額":>9}  所見')
            for x in diff_rows[:15]:
                print(f'   {x[0]:6} {str(x[3])[:16]:16} {x[7]:>4.0f} {x[8]:>8,.0f} '
                      f'{x[13]:>8,.0f} {x[14]:>+8,.0f} {x[15]:>+9,.0f}  {x[17]}')
        else:
            print('   単位を揃えたうえで差がある行はありません')
        path = write_push_kpi_list(month_sheet, rows, pos)
        print(f'\n  → {path}')
        print('  ⚠️ 単位未確認の行は反映側でも弾きます。'
              '承認欄に加えて「更新理由=入力ミス」の記入が必要です')
        return
    if len(sys.argv) >= 3 and sys.argv[1] in ('check-amazon', 'register-amazon', 'adopt-amazon'):
        main_amazon(sys.argv[1], sys.argv[2])
        return
    if len(sys.argv) < 3 or sys.argv[1] not in ('check', 'register', 'adopt'):
        print(__doc__)
        sys.exit(1)
    mode, month_sheet = sys.argv[1], sys.argv[2]
    cost_by_pid, pair2pid, ctrl2pid = load_master()
    rows = load_kpi_month(month_sheet)
    match, conflict, missing = classify(rows, cost_by_pid, pair2pid, ctrl2pid)
    print(f'{month_sheet}: {len(rows)}行(ユニーク{len(match)+len(conflict)+len(missing)}ペア)')
    print(f'  マスター一致: {len(match)} / 原価不一致: {len(conflict)} / 未登録: {len(missing)}')
    if conflict:
        if mode == 'adopt':
            print('--- 原価不一致(これからKPI側の値でマスターを更新します)')
        else:
            print('--- 原価不一致(マスターは変更していません。要ユーザー判断)')
        for r, pid, mc in conflict:
            print(f'  {r["ctrl"]} ({pid}) マスター:{mc} / KPI:{r["cost"]} — {r["name"][:30]}')
    if mode == 'adopt':
        if not conflict:
            print('不一致なし — 更新処理スキップ')
            return
        updated = adopt_kpi_costs(month_sheet, conflict)
        print(f'マスター原価をKPI側に更新: {updated}件')
        return
    if mode == 'register':
        if not missing:
            print('未登録なし — 登録処理スキップ')
            return
        last_p, last_c = register_missing(month_sheet, missing)
        print(f'登録完了: {len(missing)}件追加 (最終ID P{last_p:06d} / C{last_c:06d})')
    elif missing:
        print(f'--- 未登録 {len(missing)}件(register モードで追加登録できます)')


if __name__ == '__main__':
    main()
