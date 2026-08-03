# -*- coding: utf-8 -*-
"""sync_cost_master.py — 原価一元管理の同期ツール

商品マスターを「原価の正」とするための同期スクリプト。
KPI管理シートの月次シートと商品マスターを突き合わせる。

使い方:
    python3 sync_cost_master.py check 7月            # 突合レポートのみ(変更なし)
    python3 sync_cost_master.py register 7月         # 未登録商品をマスターへ追加登録
    python3 sync_cost_master.py adopt 7月            # 原価不一致をKPI側の値でマスター更新
                                                     # (ユーザーが明示承認した場合のみ実行すること)

設計方針(GS移行を見据えて):
- 照合キーは「楽天商品管理番号 × 楽天SKU」のペア(SKU単独照合はNG。使い回しがあるため)
- マスターの既存原価は絶対に上書きしない。不一致は一覧表示してユーザー判断に委ねる
- 追加登録行は備考に出所を残し、後から見分けられるようにする
"""
import re
import sys
from datetime import date

from openpyxl import load_workbook

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


def load_kpi_month(month_sheet):
    """KPI月次シートから (管理番号, SKU, 商品名, 商品番号, 仕入値) を行順で返す"""
    wb = load_workbook(KPI_FILE, data_only=True)
    ws = wb[month_sheet]
    out = []
    for row in range(8, ws.max_row + 1):
        ctrl = ws.cell(row, 3).value
        if ctrl is None:          # データ末尾(合計行の手前)まで
            break
        out.append({
            'ctrl': str(ctrl).strip(),
            'sku': str(ws.cell(row, 5).value).strip() if ws.cell(row, 5).value else '',
            'name': str(ws.cell(row, 1).value or '').strip(),
            'pn': str(ws.cell(row, 4).value or '').strip(),
            'cost': ws.cell(row, 12).value,
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


def main():
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
