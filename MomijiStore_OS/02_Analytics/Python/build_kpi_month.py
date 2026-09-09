# -*- coding: utf-8 -*-
"""build_kpi_month.py — KPI管理シートの月次シートをRMSのSKU別売上CSVから自動生成

使い方:
    python3 build_kpi_month.py <SKU_SalesList.csv> [シート名]

例:
    python3 build_kpi_month.py ../SourceData/202608_SKU_SalesList.csv
    # シート名省略時はCSVファイル名の 2026MM から「8月」を自動判定

処理内容:
1. KPI管理シートをBackup/へバックアップ
2. 直近の月シートをテンプレートとしてコピーし、CSVデータを流し込む
3. 仕入値の照合順: 商品マスター(管理番号×SKUペア) → RMS商品番号の「/仕入値」埋め込み
   → 商品番号が6桁以下の数値ならその値 → 過去月シート → 解決不能は黄色塗り空欄
4. Excel(AppleScript)で再計算し、合計値をCSVと突合して検証
5. 月次経費(広告費・送料等)は黄色塗り空欄 — ユーザー入力後に限界利益が確定

生成後は sync_cost_master.py register で新商品をマスターへ還流させること。
"""
import csv
import os
import re
import subprocess
import sys

from openpyxl import load_workbook
from openpyxl.styles import PatternFill

import report_columns as C   # レポートは位置ではなく見出し名で読む

import sync_cost_master as S
from sync_cost_master import KPI_FILE, load_master, norm

YELLOW = PatternFill('solid', fgColor='FFFF00')
BACKUP_DIR = os.path.dirname(KPI_FILE) + '/Backup'


SOURCE = '楽天 SKU別売上CSV'

# KPIシートのA〜J列は**この順番で固定**である。K列以降の数式
# (K=販売手数料 I*0.15 / M=H*L / N=I-K-M / O=N/I)がこの並びに依存している。
# したがって「CSVの並び」ではなく「シートの並び」を正とし、
# CSV側は見出し名で引いて、この順に流し込む。
# こうすればRMSが列を足しても並べ替えても、シートは同じ形のまま保たれる。
SHEET_COLUMNS = [
    ('name',     ['商品名']),          # A
    ('sku_info', ['SKU情報']),         # B
    ('ctrl',     ['商品管理番号']),      # C ← 仕入値の照合キー
    ('pn',       ['商品番号']),         # D ← 「/仕入値」の埋め込み元
    ('sku',      ['SKU管理番号']),      # E ← 仕入値の照合キー
    ('sku_no',   ['SKU番号']),         # F
    ('price',    ['平均単価']),         # G
    ('units',    ['売上個数']),         # H ← 検証対象
    ('sales',    ['売上']),            # I ← 検証対象
    ('orders',   ['売上件数']),         # J ← 検証対象
]


COL = {k: i for i, (k, _) in enumerate(SHEET_COLUMNS)}   # 'ctrl' → 2 など


def read_rms_csv(path, month=''):
    """RMSのSKU別売上CSVを (ヘッダー6行, データ行list) で返す。

    データ行は **SHEET_COLUMNS の順に並べ替えた list** にして返す。
    以降の処理はシートの列順だけを知っていればよい。
    """
    with open(path, encoding='utf-8-sig') as f:
        rows = list(csv.reader(f))
    # 見出しより上にRMSの検索条件(表示期間・端末・消費税など)が入る。
    # 行数を決め打ちせず、「商品名」で始まる行を見出しとして探す
    hi = next((i for i, r in enumerate(rows) if r and r[0].strip() == '商品名'), None)
    if hi is None:
        sys.exit(f'❌ {SOURCE} の見出し行(「商品名」で始まる行)が見つかりません: {path}')
    header = rows[hi]

    C.check_layout(SOURCE, header, month)            # ⓪ 前月と列構成を突き合わせる
    idx = C.resolve(header, dict(SHEET_COLUMNS), source=SOURCE)
    C.report(idx, header, SOURCE)

    if hi != 6:
        print(f'  ※ 見出し行が {hi} 行目にあります(通常は6)。'
              'シート上部へ書き戻すのは先頭6行のみです')
    head6 = rows[:hi][:6]
    data = [[C.get(r, idx, k, '') for k, _ in SHEET_COLUMNS]
            for r in rows[hi + 1:] if any(x.strip() for x in r)]
    return head6, data


def conv(s):
    """CSV文字列をExcel向けに数値変換(数値でなければ文字列のまま)"""
    s = s.strip()
    if s == '':
        return None
    if s.isdigit():
        return int(s)
    try:
        return float(s)
    except ValueError:
        return s


def build_cost_resolver(wb_values):
    """仕入値の解決関数を作る(商品マスター → RMS埋め込み → 過去月シート)

    ⚠️ 商品マスターの標準原価は「単品1個あたり」とは限らない。
       出品がセット販売なら、その出品の原価は 単品原価 × 販売入数 である。
       出品テーブルの 販売入数・原価単位 が**確認済みのときだけ**計算し、
       未確認なら **None を返して黄色セルで人へ回す**(2026-09-10 承認W-4)。
       推測値を入れない。既存の「仕入値未解決」と同じ経路に落とす。
    """
    cost_by_pid, pair2pid, ctrl2pid = load_master()
    pack, pos = S.load_pack_info()
    has_pack_cols = any(v is not None for v in pos.values())

    def rms_cost(ctrl, pn, sku, master_cost):
        """RMS商品番号に埋め込まれた原価を、**条件を満たすときだけ**採用する。

        商品番号は「8507/2283-a」のように 品番/原価 の形で原価を持つ。
        出品ごとの値なので入数の問題は起きないが、
        **数字が抽出できることは確認済みを意味しない**(承認X-1)。

        採用の条件:
          ① チャネル・商品管理番号・SKU の対応が付くこと(出品テーブルに在る)
          ② 1販売あたりの原価であること(出品の販売入数が確認済み)
          ③ 対象時点が分かること(当月のRMS CSVから取得している)
          ④ 付属品等の含有範囲が分かること(原価単位が確認済み)
        ②④は出品テーブルの L/M/N 列で確認する。**未記入なら採用しない。**
        """
        m = None
        if '/' in pn:
            mm = re.match(r'(\d+)', pn.rsplit('/', 1)[1].strip())
            if mm:
                m = int(mm.group(1))
        if m is None:
            q = pn.strip()
            m = int(q) if q.isdigit() and len(q) <= 6 else None
        if m is None:
            return None, 'RMS商品番号に原価が埋め込まれていない'
        info, amb = S.lookup_pack(pack, '楽天', norm(ctrl), norm(sku))
        if amb or not info:
            return None, f'出品の対応が付かない({amb or "出品テーブルに無い"})'
        if not isinstance(info.get('pack'), (int, float)):
            return None, 'RMS原価はあるが販売入数が未確認(1販売あたりか判定できない)'
        if info.get('unit') not in (S.UNIT_SINGLE, S.UNIT_SET):
            return None, 'RMS原価はあるが原価単位が未確認(付属品の含有範囲が不明)'
        return m, ''


    hist_pair = {}
    for name in reversed(wb_values.sheetnames):        # 新しい月を優先
        ws = wb_values[name]
        for r in range(8, ws.max_row + 1):
            c, e, l = ws.cell(r, 3).value, ws.cell(r, 5).value, ws.cell(r, 12).value
            if c is None or not isinstance(l, (int, float)):
                continue
            hist_pair.setdefault((norm(c), norm(e)), l)

    def resolve(ctrl, pn, sku):
        """(原価, 出所, 未確認理由) を返す。**採用できなければ原価は None。**"""
        key = (norm(ctrl), norm(sku))
        pid = pair2pid.get(key) or ctrl2pid.get(key[0])
        mc = cost_by_pid.get(pid) if pid is not None else None
        why = ''
        if isinstance(mc, (int, float)):
            if not has_pack_cols:
                return mc, 'マスター(単位列なし)', ''
            v, why = S.resolve_unit_cost(pack, '楽天', norm(ctrl), norm(sku), mc)
            if v is not None:
                return v, f'マスター×入数(単位確認済み)', ''
        # マスターから解決できないとき、RMS商品番号の原価を**代替候補**として見る。
        # ただし数字が取れるだけでは採用しない(2026-09-10 承認X-1)
        cand, cwhy = rms_cost(ctrl, pn, sku, mc)
        if cand is not None:
            return cand, 'RMS商品番号(条件確認済み)', why
        h = hist_pair.get(key)
        if h is not None:
            return h, '過去月シート', why
        return None, '', (why or cwhy or '原価を確認できる資料が無い')

    return resolve


def build_sheet(kpi_file, sheet_name, head6, data, resolve, master_cost_of=None):
    """テンプレート(直近月シート)をコピーして新しい月シートを構築する"""
    wb = load_workbook(kpi_file)
    if sheet_name in wb.sheetnames:
        raise SystemExit(f'シート「{sheet_name}」は既に存在します。手動で削除してから再実行してください。')
    template = wb.sheetnames[-1]
    ws = wb.copy_worksheet(wb[template])
    ws.title = sheet_name
    ws.freeze_panes = None

    # テンプレートのデータ行数を数え、今月の行数に合わせて増減する
    t_n = 0
    while ws.cell(8 + t_n, 3).value is not None:
        t_n += 1
    n = len(data)
    if n < t_n:
        ws.delete_rows(8, t_n - n)
    elif n > t_n:
        ws.insert_rows(9, n - t_n)
        for r in range(9, 9 + (n - t_n)):
            for col in range(1, 16):
                ws.cell(r, col)._style = ws.cell(8, col)._style

    for i, r in enumerate(head6, start=1):
        ws.cell(i, 1).value = r[0] if r else None
        ws.cell(i, 2).value = r[1] if len(r) > 1 else None
    ws.cell(7, 16).value = '原価の出所(自動記録)'

    master_cost_of = master_cost_of or (lambda *a: None)
    last = 7 + n
    missing = []
    for i, r in enumerate(data):
        row = 8 + i
        for col in range(len(SHEET_COLUMNS)):          # A〜J列。順番はシート側が正
            ws.cell(row, col + 1).value = conv(r[col])
        cost, src, why = resolve(r[COL['ctrl']], r[COL['pn']], r[COL['sku']])
        lc = ws.cell(row, 12)
        lc.value = int(cost) if isinstance(cost, float) and cost == int(cost) else cost
        lc.fill = PatternFill(fill_type=None) if cost is not None else YELLOW
        # 採用した原価の出所を残す。**どこから来た値か後から辿れるようにする**
        # ⚠️ A〜O列はCSVデータと数式が使っているので触らない。**P列(16)へ書く**
        note = ws.cell(row, 16)
        if cost is not None:
            mcv = master_cost_of(r[COL['ctrl']], r[COL['sku']])
            gap = ('' if not isinstance(mcv, (int, float)) or mcv == cost
                   else f' / マスター {mcv:,.0f} と不一致(理由: {why or "未確認"})')
            note.value = f'原価出所: {src}{gap}'
        else:
            note.value = f'原価未確定: {why}'
            missing.append((row, r[COL['ctrl']], r[COL['sku']], why))
        ws.cell(row, 11).value = f'=SUM(I{row}*0.15)'
        # 原価未確定ガード(2026-09-10 承認X-2)
        #   Excelは空セルを0として計算するため、仕入値が空欄だと
        #   仕入0円として粗利が過大に出る。数式エラーも警告も出ない。
        #   ISNUMBER は 空欄・文字列・エラーのいずれもFALSEを返し、**0はTRUE**を返す。
        #   → 「確認済みの0円」と「未入力」を数式の上で区別できる。
        #   対象は販売や返品のある行だけ。売上も個数も0の行は月全体を止めない。
        need = f'OR(H{row}<>0,I{row}<>0)'
        ws.cell(row, 13).value = (f'=IF(AND({need},NOT(ISNUMBER(L{row}))),'
                                  f'"未確定",H{row}*L{row})')
        ws.cell(row, 14).value = f'=IF(ISTEXT(M{row}),"未確定",I{row}-K{row}-M{row})'
        ws.cell(row, 15).value = (f'=IF(ISTEXT(N{row}),"未確定",'
                                  f'IF(I{row}=0,"",N{row}/I{row}))')

    t = last + 1
    # 売上・個数・手数料は原価と独立に確認できるので、そのまま数値で出す
    for col in 'HIJK':
        ws[f'{col}{t}'] = f'=SUM({col}7:{col}{last})'
    # 仕入計と粗利は、対象行に1つでも未確定があれば「未確定」
    und = f'COUNTIF(M7:M{last},"未確定")'
    ws[f'M{t}'] = f'=IF({und}>0,"未確定",SUM(M7:M{last}))'
    ws[f'N{t}'] = f'=IF({und}>0,"未確定",SUM(N7:N{last}))'
    ws[f'O{t}'] = f'=IF(ISTEXT(N{t}),"未確定",IF(I{t}=0,"",N{t}/I{t}))'
    # 原価が判明している行だけの粗利。**月全体の利益ではない**ことを明示する
    ws.cell(t + 1, 13).value = '(参考)原価判明分のみの粗利'
    ws.cell(t + 1, 14).value = f'=SUMIF(N7:N{last},"<>未確定")'
    ws.cell(t + 1, 15).value = f'=IF({und}=0,"全件確定",{und}&"行が未確定")'

    e0 = t + 3
    for j, lab in enumerate(['広告費', 'ポイント費用', 'クーポン利用額', 'クーポン利用手数料']):
        row = e0 + j
        ws.cell(row, 13).value = lab
        nc = ws.cell(row, 14)
        nc.value = None
        nc.fill = YELLOW
        ws.cell(row, 15).value = f'=SUM(N{row}/I{t})' if j < 3 else None
    tot1 = e0 + 4
    ws.cell(tot1, 13).value = '合計'
    ws.cell(tot1, 14).value = f'=SUM(N{e0}:N{e0 + 3})'
    ws.cell(tot1, 15).value = f'=SUM(N{tot1}/I{t})'

    s0 = tot1 + 2
    for j, lab in enumerate(['送料', '梱包資材']):
        row = s0 + j
        ws.cell(row, 13).value = lab
        nc = ws.cell(row, 14)
        nc.value = None
        nc.fill = YELLOW
        ws.cell(row, 15).value = f'=SUM(N{row}/I{t})'
    tot2 = s0 + 2
    ws.cell(tot2, 13).value = '合計'
    ws.cell(tot2, 14).value = f'=SUM(N{s0}:N{s0 + 1})'
    ws.cell(tot2, 15).value = f'=SUM(O{s0}:O{s0 + 1})'

    g = tot2 + 2
    # 必須の経費が1つでも空欄なら「未確定」と出す。
    # Excelは空セルを0として集計するため、ガードが無いと
    # 「送料0円」のもっともらしい数字が限界利益として表示されてしまう
    # (2026-09-09 監査。数式エラーも警告も出ないので人は気づけない)
    guard = f'COUNTBLANK(N{e0}:N{e0 + 3})+COUNTBLANK(N{s0}:N{s0 + 1})'
    # 経費だけでなく**粗利が未確定なら限界利益も未確定**(承認X-2)
    ws.cell(g, 13).value = '限界利益'
    ws.cell(g, 14).value = (f'=IF(OR(ISTEXT(N{t}),{guard}>0),"未確定",'
                            f'N{t}-N{tot1}-N{tot2})')
    ws.cell(g + 1, 13).value = '限界利益率'
    ws.cell(g + 1, 14).value = f'=IF(ISTEXT(N{g}),"未確定",IF(I{t}=0,"",N{g}/I{t}))'

    for row in ws.iter_rows(min_row=g + 2, max_row=ws.max_row):
        for c in row:
            c.value = None

    wb.save(kpi_file)
    return t, missing


def recalc_via_excel(path):
    script = f'''
set p to POSIX file "{path}"
with timeout of 600 seconds
tell application "Microsoft Excel"
    set wasRunning to running
    open p
    delay 2
    calculate
    save active workbook
    close active workbook saving no
    if not wasRunning then quit
end tell
end timeout
return "ok"
'''
    r = subprocess.run(['osascript', '-e', script], capture_output=True, text=True, timeout=660)
    if r.returncode != 0:
        raise RuntimeError(f'Excel再計算に失敗: {r.stderr}')


def verify(kpi_file, sheet_name, data, total_row):
    """再計算後の合計をCSVと突合し、数式エラーを検査する"""
    wb = load_workbook(kpi_file, data_only=True)
    ws = wb[sheet_name]
    exp = [sum(int(r[COL[k]]) for r in data) for k in ('units', 'sales', 'orders')]
    got = [ws.cell(total_row, c).value for c in (8, 9, 10)]
    ok = got == exp
    errs = [c.coordinate for row in ws.iter_rows() for c in row
            if isinstance(c.value, str) and c.value.startswith('#')]
    return ok, exp, got, errs, ws.cell(total_row, 14).value, ws.cell(total_row, 15).value


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    csv_path = sys.argv[1]
    fname = os.path.basename(csv_path)
    if len(sys.argv) >= 3:
        sheet_name = sys.argv[2]
    else:
        m = re.match(r'\d{4}(\d{2})_', fname)
        if not m:
            raise SystemExit('シート名を自動判定できません。第2引数で指定してください(例: 8月)')
        sheet_name = f'{int(m.group(1))}月'

    kpi_file = os.environ.get('KPI_FILE_OVERRIDE', KPI_FILE)

    # ⓪ 列構成チェックは何よりも先。壊れたレポートで作業を始めない
    head6, data = read_rms_csv(csv_path, sheet_name)
    print(f'{sheet_name}: CSVデータ {len(data)}行')

    if kpi_file == KPI_FILE:
        os.makedirs(BACKUP_DIR, exist_ok=True)
        from datetime import date
        base = os.path.splitext(os.path.basename(kpi_file))[0]
        bak = f'{BACKUP_DIR}/{base}_backup_{date.today():%Y%m%d}_{sheet_name}生成前.xlsx'
        import shutil
        shutil.copy2(kpi_file, bak)
        print(f'バックアップ: {os.path.basename(bak)}')

    wb_values = load_workbook(kpi_file, data_only=True)
    resolve = build_cost_resolver(wb_values)
    cost_by_pid, pair2pid, ctrl2pid = load_master()

    def master_cost_of(ctrl, sku):
        pid = pair2pid.get((norm(ctrl), norm(sku))) or ctrl2pid.get(norm(ctrl))
        return cost_by_pid.get(pid) if pid else None

    total_row, missing = build_sheet(kpi_file, sheet_name, head6, data, resolve,
                                     master_cost_of)
    und_sales = 0
    for m_ in missing:
        pass
    print(f'シート生成完了(原価未確定 {len(missing)}行)')
    from collections import Counter
    for why, c in Counter(m_[3] for m_ in missing).most_common():
        print(f'  {c:>4}行  {why}')

    print('Excelで再計算中...')
    recalc_via_excel(kpi_file)

    ok, exp, got, errs, n_total, o_total = verify(kpi_file, sheet_name, data, total_row)
    print(f'検証: 合計{"一致" if ok else "不一致!"} (個数/売上/件数 CSV={exp} シート={got})')
    print(f'数式エラー: {len(errs)}件 {errs[:10] if errs else ""}')
    print(f'粗利: {n_total:,.0f}円 ({o_total * 100:.1f}%)' if n_total else '')
    if not ok or errs:
        sys.exit('*** FAIL — シートを確認してください ***')
    print('PASS。経費(黄色セル)入力後に限界利益が確定します。')
    print('新商品があれば: python3 sync_cost_master.py register ' + sheet_name)


if __name__ == '__main__':
    main()
