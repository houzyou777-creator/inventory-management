# -*- coding: utf-8 -*-
"""run_stock_tool.py — 楽天在庫金額集計ツールの「読込」→「集計」→保存を、人がボタンを押さずに実行する(V-4 以降)

使い方:
    python3 run_stock_tool.py --target copy                       V4_test のコピーで 読込→集計→保存
    python3 run_stock_tool.py --target copy --import-from <xlsx>  先にコピーの Import/楽天在庫リスト_import.xlsx を差し替えてから実行
    python3 run_stock_tool.py --target production --confirm 本番  本番で実行(バックアップ→実行。人が Excel を使っていたら止まる)
    --dry-run を付けると、何をするかを表示するだけで Excel を動かさない

仕組み(2026-09-17):
    ・ツールの VBA(V-4)にある AutoImport / AutoAggregate(MsgBox を出さず結果文字列を返す)を
      excel_bridge.run_macros が AppleScript の run VB macro で呼ぶ
    ・Excel の安全装置は excel_bridge のまま(対象はフルパス・自分が開いたブックだけ保存/閉じる・
      対象や同名ブックが開かれていれば何もしない・quit しない)
    ・本番はさらに厳しく: Excel に何かブックが開いていれば止める(人が作業中の可能性)。実行前に本番ツールと
      Import をバックアップする(一意な名前)

できないこと(制約):
    ・VBA モジュールの入替(VBE の削除→インポート)は Mac Excel に外部から操作する口が無い → 人
    ・Import フォルダへの初回アクセス許可(サンドボックスのダイアログ)は人がクリックする必要がある(1フォルダ1回)
    ・予期しないダイアログが出ると AppleScript は待ち続ける → timeout で止め、ブックは保存しない
"""
import argparse
import hashlib
import os
import shutil
import sys
import warnings
from datetime import datetime

warnings.filterwarnings('ignore')
sys.path.insert(0, '/Users/hide0726/Desktop/Claude Code/MomijiStore_OS/02_Analytics/Python')
import excel_bridge as eb                                   # noqa: E402
from openpyxl import load_workbook                          # noqa: E402

SD = '/Users/hide0726/Desktop/Claude Code/MomijiStore_OS/01_InventoryManagement/SourceData'
TARGETS = {
    'copy': SD + '/V4_test/楽天在庫金額集計ツール_v1.0.xlsm',
    'production': SD + '/楽天在庫金額集計ツール_v1.0.xlsm',
}
TOOL_NAME = '楽天在庫金額集計ツール_v1.0.xlsm'
IMPORT_NAME = '楽天在庫リスト_import.xlsx'


def sha(path):
    return hashlib.sha1(open(path, 'rb').read()).hexdigest()[:10]


def backup(path, label):
    d = os.path.join(os.path.dirname(path), 'Backup')
    os.makedirs(d, exist_ok=True)
    base, ext = os.path.splitext(os.path.basename(path))
    dst = os.path.join(d, f'{base}_backup_{datetime.now():%Y%m%d_%H%M%S}_{label}{ext}')
    if os.path.exists(dst):
        sys.exit(f'❌ バックアップ名が既にあります: {dst}')
    shutil.copy2(path, dst)
    return dst


def summary(path):
    wb = load_workbook(path, read_only=True, data_only=True)
    ag = wb['在庫金額集計']
    tk = wb['楽天CSV取込']
    n = sum(1 for r in tk.iter_rows(min_row=3, values_only=True) if r and (r[1] or r[3]))
    out = dict(rows=n, amount=ag.cell(3, 2).value, qty=ag.cell(4, 2).value, unreg=ag.cell(5, 6).value,
               agg_dt=ag.cell(6, 2).value, import_dt=tk.cell(3, 11).value, size=os.path.getsize(path))
    wb.close()
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--target', choices=list(TARGETS), required=True)
    ap.add_argument('--import-from', help='この xlsx を対象の Import/ へ置いてから実行(元は Import/Backup へ退避)')
    ap.add_argument('--confirm', default='', help='production のときは「本番」と書く')
    ap.add_argument('--dry-run', action='store_true')
    ap.add_argument('--label', default='実行前', help='バックアップ名に付ける印')
    a = ap.parse_args()

    target = TARGETS[a.target]
    imp = os.path.join(os.path.dirname(target), 'Import', IMPORT_NAME)
    print(f'■ 対象: {target}')
    if not os.path.exists(target):
        sys.exit('❌ 対象がありません')
    if a.target == 'production' and a.confirm != '本番':
        sys.exit('❌ 本番で実行するには --confirm 本番 が必要です')

    # 人が Excel を使っていないか
    opened = eb.open_workbooks()
    if any(os.path.basename(p) == TOOL_NAME for p in opened):
        sys.exit(f'⛔ 同名のブックが Excel で開かれています(本番/コピーの取り違え防止で停止): {opened}')
    if a.target == 'production' and opened:
        sys.exit(f'⛔ 本番実行は Excel に他のブックが開いている間は行いません(人が作業中の可能性): {opened}')

    # Import の差し替え(任意)
    if a.import_from:
        src = os.path.abspath(a.import_from)
        if not os.path.exists(src):
            sys.exit(f'❌ --import-from がありません: {src}')
        print(f'   Import 差し替え: {src} (sha {sha(src)}, {datetime.fromtimestamp(os.path.getmtime(src)):%Y-%m-%d %H:%M})')
        print(f'              → {imp}')
        if not a.dry_run:
            if os.path.exists(imp):
                print('   退避:', backup(imp, a.label))
            shutil.copy2(src, imp)
    if not os.path.exists(imp):
        sys.exit(f'❌ Import がありません: {imp}')
    print(f'   Import: {imp} (sha {sha(imp)}, {datetime.fromtimestamp(os.path.getmtime(imp)):%Y-%m-%d %H:%M})')

    before = summary(target)
    print(f'   実行前: 取込 {before["rows"]} 行 / 集計 {before["amount"]} / {before["qty"]} / 未登録 {before["unreg"]} / 集計日時 {before["agg_dt"]} / {before["size"]:,} bytes')
    if a.dry_run:
        print('   (dry-run: Excel は動かしません)')
        return
    print('   バックアップ:', backup(target, a.label))

    started = datetime.now()
    print(f'   実行開始 {started:%Y-%m-%d %H:%M:%S} … AutoImport → AutoAggregate → 保存 → 閉じる')
    try:
        out = eb.run_macros(target, ['AutoImport', 'AutoAggregate'], timeout=900)
    except eb.ExcelAbort as e:
        sys.exit(f'⛔ 何もせず停止: {e}')
    print('   VBA の戻り:', out)
    after = summary(target)
    print(f'   実行後: 取込 {after["rows"]} 行 / 集計 {after["amount"]} / {after["qty"]} / 未登録 {after["unreg"]} / 読込日時 {after["import_dt"]} / 集計日時 {after["agg_dt"]} / {after["size"]:,} bytes')
    log = os.path.join(os.path.dirname(target), 'Backup', f'実行記録_{started:%Y%m%d_%H%M%S}.txt')
    with open(log, 'w', encoding='utf-8') as f:
        f.write(f'対象: {target}\nImport: {imp} sha={sha(imp)}\n実行(取込実行日時): {started:%Y-%m-%d %H:%M:%S}\n戻り: {out}\n実行前: {before}\n実行後: {after}\n')
    print('   記録:', log)


if __name__ == '__main__':
    main()
