# -*- coding: utf-8 -*-
"""run_stock_tool.py — 楽天在庫金額集計ツールの「読込」→「集計」→保存を、人がボタンを押さずに実行する(V-4 以降)

使い方:
    python3 run_stock_tool.py --target copy --expect 15077081,6801,20
    python3 run_stock_tool.py --target copy --import-from <xlsx> --expect 12621422,5911,59
    python3 run_stock_tool.py --target production --confirm 本番 --expect 12621422,5911,59 --source-stamp "楽天 2026-09-16 08:11:29"
    --dry-run を付けると、何をするかを表示するだけで Excel を動かさない

仕組み(2026-09-17/19):
    ・ツールの VBA(V-4)にある AutoImport / AutoAggregate(MsgBox を出さず結果文字列を返す)を
      excel_bridge.run_macros が AppleScript の run VB macro で呼ぶ
    ・Excel の安全装置は excel_bridge のまま(対象はフルパス・自分が開いたブックだけ保存/閉じる・
      対象や同名ブックが開かれていれば何もしない・quit しない)
    ・本番はさらに厳しく: Excel に何かブックが開いていれば止める(人が作業中の可能性)。実行前に本番ツールと
      Import をバックアップする(一意な名前)

ChatGPT 2026-09-19 の条件:
    1. タイムアウトを「処理が停止した」とみなさない。Excel 側の状態(応答・対象ブックが開いたまま・ダイアログ・
       ファイルの更新)を確認して**報告するだけ**。再実行も手動操作への切替もしない
    2. マクロの OK 応答だけで成功にしない。保存後にファイルを読み直し、集計値・識別子・原価未登録件数・
       ファイルサイズ・読込/集計日時を検証する。不一致なら FAIL(終了コード 1)で後続を止める

できないこと(制約):
    ・VBA モジュールの入替(VBE の削除→インポート)は Mac Excel に外部から操作する口が無い → 人
    ・Import フォルダへの初回アクセス許可(サンドボックスのダイアログ)は人がクリックする必要がある(1フォルダ1回)
"""
import argparse
import hashlib
import os
import shutil
import sys
import warnings
import zipfile
from datetime import datetime

warnings.filterwarnings('ignore')
# 共通モジュール(02_Analytics/Python)の場所は __file__ からの相対位置で求める
sys.path.insert(0, os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..', '..', '02_Analytics', 'Python')))
import excel_bridge as eb                                   # noqa: E402
import momiji_paths as MP                                   # noqa: E402
from openpyxl import load_workbook                          # noqa: E402

SD = MP.INV_SD
TARGETS = {
    'copy': SD + '/V4_test/楽天在庫金額集計ツール_v1.0.xlsm',
    'production': SD + '/楽天在庫金額集計ツール_v1.0.xlsm',
}
TOOL_NAME = '楽天在庫金額集計ツール_v1.0.xlsm'
IMPORT_NAME = '楽天在庫リスト_import.xlsx'
SIZE_LIMIT = 1_000_000          # V-3: 6,985,292 / V-4 1回目: 4,714,370 → V-4b 以降は 300KB 台のはず


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


def cstr(v):
    if v is None:
        return None
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return v.strip() if isinstance(v, str) else str(v)


def rows_of(ws, min_row, key_cols, width=13):
    out = []
    for r in ws.iter_rows(min_row=min_row, values_only=True):
        r = tuple(r) + (None,) * (width - len(r))
        if any(r[c] not in (None, '') for c in key_cols):
            out.append(r)
    return out


def summary(path):
    wb = load_workbook(path, read_only=True, data_only=True)
    ag = wb['在庫金額集計']
    tk = wb['楽天CSV取込']
    trows = rows_of(tk, 3, (1, 3))
    out = dict(rows=len(trows), amount=ag.cell(3, 2).value, qty=ag.cell(4, 2).value, unreg=ag.cell(5, 6).value,
               agg_dt=ag.cell(6, 2).value, import_dt=(trows[0][10] if trows else None), size=os.path.getsize(path))
    wb.close()
    return out


def verify_after(target, imp, started, expect):
    """条件2: 保存後の再読込で検証する。全部通れば True。"""
    ok = True

    def chk(cond, msg):
        nonlocal ok
        print(('   ✅ ' if cond else '   ❌ ') + msg)
        ok = ok and bool(cond)

    after = summary(target)
    chk(after['size'] < SIZE_LIMIT, f'ファイルサイズ {after["size"]:,} bytes(< {SIZE_LIMIT:,})')
    z = zipfile.ZipFile(target)
    n4 = z.read('xl/worksheets/sheet4.xml').count(b'<row ')
    chk(n4 < after['rows'] + 300, f'保存された集計シートの行要素 {n4:,}(10万行ではない)')
    # 読込日時・集計日時が実行開始より後
    def dt(v):
        try:
            return datetime.strptime(str(v)[:19].replace('-', '/'), '%Y/%m/%d %H:%M:%S')
        except ValueError:
            return None
    chk(dt(after['import_dt']) and dt(after['import_dt']) >= started.replace(microsecond=0), f'読込日時 {after["import_dt"]} が実行開始 {started:%Y/%m/%d %H:%M:%S} 以後')
    chk(dt(after['agg_dt']) and dt(after['agg_dt']) >= started.replace(microsecond=0), f'集計日時 {after["agg_dt"]} が実行開始以後')
    # 取込 = Import(識別子4列の文字列一致・空欄保持)
    wi = load_workbook(imp, read_only=True, data_only=True)['在庫']
    irows = rows_of(wi, 2, (1, 3))
    wc = load_workbook(target, data_only=True)
    trows = rows_of(wc['楽天CSV取込'], 3, (1, 3))
    chk(len(trows) == len(irows), f'取込 {len(trows)} 行 = Import {len(irows)} 行')
    mism = typebad = blankbad = 0
    for a, b in zip(irows, trows):
        for c in range(4):
            e, g = cstr(a[c]), b[c]
            if e in (None, ''):
                blankbad += g is not None
            else:
                mism += (g != e)
                typebad += not isinstance(g, str)
    chk(mism == 0 and typebad == 0 and blankbad == 0, f'識別子4列: 不一致 {mism}・非文字列 {typebad}・空欄→非空欄 {blankbad}')
    # 集計の識別子(文字列・長さ0なし)と行数
    arows = rows_of(wc['在庫金額集計'], 10, (0, 1, 2))
    chk(len(arows) == len(trows), f'集計 {len(arows)} 行 = 取込 {len(trows)} 行')
    chk(sum(1 for r in arows for c in (0, 1, 2, 11) if isinstance(r[c], (int, float)) or r[c] == '') == 0, '集計の識別子に数値型・長さ0文字列なし')
    crows = rows_of(wc['要確認一覧'], 3, (0, 2))
    unreg_rows = sum(1 for r in crows if r[4] == '原価未登録')
    chk(unreg_rows == after['unreg'], f'要確認一覧の「原価未登録」{unreg_rows} 行 = 集計の未登録 {after["unreg"]}')
    if expect:
        ea, eq, eu = expect
        chk(round(after['amount'] or 0) == ea and after['qty'] == eq and after['unreg'] == eu,
            f'集計値 {after["amount"]:,} / {after["qty"]:,} / 未登録 {after["unreg"]} = 期待 {ea:,} / {eq:,} / {eu}')
    else:
        print(f'   ・集計値 {after["amount"]:,} / {after["qty"]:,} / 未登録 {after["unreg"]}(期待値の指定なし)')
    return ok, after


def report_state(target, why):
    st = eb.excel_state(target)
    print(f'⚠️ {why}')
    print(f'   Excel の状態: 応答={st["responding"]} / 対象ブックが開いたまま={st["target_open"]} / 開いているブック={st["open_workbooks"]}')
    print(f'   ウィンドウ={st["dialogs"]}')
    print(f'   ファイル: 更新 {st["file_mtime"]} / {st["file_size"]:,} bytes')
    print('   → 再実行も手動操作への切替もしません。この状態を報告して判断を待ちます(ChatGPT 2026-09-19 条件1)')
    return st


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--target', choices=list(TARGETS), required=True)
    ap.add_argument('--import-from', help='この xlsx を対象の Import/ へ置いてから実行(元は Import/Backup へ退避)')
    ap.add_argument('--confirm', default='', help='production のときは「本番」と書く')
    ap.add_argument('--expect', help='期待する 総在庫金額,総在庫数量,原価未登録 (例 12621422,5911,59)')
    ap.add_argument('--source-stamp', default='', help='資料の基準日時(記録用。取込実行日時とは別に残す)')
    ap.add_argument('--dry-run', action='store_true')
    ap.add_argument('--label', default='実行前', help='バックアップ名に付ける印')
    a = ap.parse_args()
    expect = tuple(int(x.replace(',', '')) for x in a.expect.split(',')) if a.expect else None

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
    if a.source_stamp:
        print(f'   資料の基準日時: {a.source_stamp}')

    before = summary(target)
    print(f'   実行前: 取込 {before["rows"]} 行 / 集計 {before["amount"]} / {before["qty"]} / 未登録 {before["unreg"]} / 集計日時 {before["agg_dt"]} / {before["size"]:,} bytes')
    if a.dry_run:
        print('   (dry-run: Excel は動かしません)')
        return
    bk = backup(target, a.label)
    print('   バックアップ:', bk)

    started = datetime.now()
    print(f'   取込実行日時 {started:%Y-%m-%d %H:%M:%S} … AutoImport → AutoAggregate → 保存 → 閉じる')
    log = os.path.join(os.path.dirname(target), 'Backup', f'実行記録_{started:%Y%m%d_%H%M%S}.txt')
    rec = [f'対象: {target}', f'Import: {imp} sha={sha(imp)}', f'資料の基準日時: {a.source_stamp or "(指定なし)"}',
           f'取込実行日時: {started:%Y-%m-%d %H:%M:%S}', f'実行前: {before}', f'バックアップ: {bk}']
    result = 'FAIL'
    try:
        out = eb.run_macros(target, ['AutoImport', 'AutoAggregate'], timeout=900)
        print('   VBA の戻り:', out)
        rec.append(f'戻り: {out}')
        print('   保存後の再読込で検証(条件2):')
        ok, after = verify_after(target, imp, started, expect)
        rec.append(f'実行後: {after}')
        result = 'PASS' if ok else 'FAIL(検証不一致。成功扱いにしない・後続を止める)'
    except eb.ExcelAbort as e:
        result = f'ABORT(何もしていない): {e}'
        print('⛔', result)
    except eb.ExcelTimeout as e:
        st = report_state(target, f'タイムアウト: {e}')
        result = f'TIMEOUT(状態確認のみ): {st}'
    except RuntimeError as e:
        # マクロが ERROR を返した等。対象ブックは開いたまま(保存していない)。状態を確認してから、自分が開いた対象だけ閉じる
        st = report_state(target, f'エラー: {e}')
        if st['responding'] and st['target_open'] and not st['dialogs']:
            print('   対象ブックを保存せずに閉じます:', eb.close_without_saving(target))
        result = f'ERROR(保存していない): {e}'
    print(f'■ 結果: {result}')
    rec.append(f'結果: {result}')
    with open(log, 'w', encoding='utf-8') as f:
        f.write('\n'.join(rec) + '\n')
    print('   記録:', log)
    sys.exit(0 if result == 'PASS' else 1)


if __name__ == '__main__':
    main()
