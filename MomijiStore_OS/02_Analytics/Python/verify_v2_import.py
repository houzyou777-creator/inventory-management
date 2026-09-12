# -*- coding: utf-8 -*-
"""verify_v2_import.py — V-2(取込VBAの文字列書式化)をコピーで検証する(読み取り専用)

使い方:
    python3 verify_v2_import.py

前提(英樹の作業が済んでいること):
    SourceData/V2_test/baseline/…xlsm   … 現行VBAのまま「集計実行」だけ押して保存したもの
    SourceData/V2_test/…xlsm            … Module_Rakuten_Tool_V2.bas を取り込み、
                                          「楽天CSV読込」→「集計実行」を押して保存したもの

確認すること(2026-09-12 ChatGPT承認の条件):
    1. 文字列の完全一致 … 取込先の 商品番号/JAN/管理番号/SKU が Import xlsx の文字列と一致する
    2. 保存・再読込後の保持 … 保存済みファイルを読み直して型が文字列のままである
    3. マクロ・既存機能の維持 … VBAの差分が V-2 の追加分だけ／集計結果が baseline と一致する
       (一致しない行は全部列挙する。理由が説明できないものが1つでもあれば FAIL)

何も書き換えない。判定は PASS / FAIL を最後に出す。
"""
import os
import sys
import warnings
from collections import Counter
from datetime import date

warnings.filterwarnings('ignore')
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from openpyxl import load_workbook

import build_todo_lists as T

SD = os.path.dirname(T.RAKUTEN_STOCK)
TEST = SD + '/V2_test/楽天在庫金額集計ツール_v1.0.xlsm'
BASE = SD + '/V2_test/baseline/楽天在庫金額集計ツール_v1.0.xlsm'
IMPORT = SD + '/V2_test/Import/楽天在庫リスト_import.xlsx'
PATCHED_BAS = os.path.dirname(SD) + '/VBA/Module_Rakuten_Tool.bas'
ID_COLS = {0: 'JAN', 1: '管理番号', 2: '商品番号', 3: 'SKU'}

fails = []


def check(cond, msg):
    print(('  ✅ ' if cond else '  ❌ ') + msg)
    if not cond:
        fails.append(msg)


def expected_text(v):
    """Import xlsx のセル値が取込後にどう入るべきか。文字列はそのまま、数値は桁の文字列。"""
    if v is None:
        return None
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip() if not isinstance(v, str) else v


def rows_of(path, sheet, min_row):
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb[sheet]
    out = [r for r in ws.iter_rows(min_row=min_row, values_only=True)]
    wb.close()
    return out


def main():
    for p in (TEST, BASE, IMPORT):
        if not os.path.exists(p):
            sys.exit(f'❌ 無い: {p}')

    # ── 0. そもそもテストが行われたか(手順②③が未実施なら、細かい判定を並べても意味がない) ──
    try:
        from oletools.olevba import VBA_Parser
        vp = VBA_Parser(TEST)
        pre = {n: c for (_, _, n, c) in vp.extract_macros()}
        vp.close()
        if 'SetIdColumnsAsText' not in pre.get('Module_Rakuten_Tool.bas', ''):
            print('⛔ テスト用ファイルに修正版VBAが入っていません(手順② V2-2〜V2-6 が未実施)。')
            print(f'   対象: {TEST}')
            print('   baseline の方ではなく、V2_test 直下のファイルへ Module_Rakuten_Tool_V2.bas を取り込んでください。')
            sys.exit(2)
    except ImportError:
        pass
    import datetime
    wb0 = load_workbook(TEST, read_only=True, data_only=True)
    imp_dt = wb0['楽天CSV取込'].cell(3, 11).value
    agg_dt = wb0['在庫金額集計'].cell(6, 2).value
    wb0.close()
    print(f'■ 0. 取込日時={imp_dt} / 集計日時={agg_dt}')
    if isinstance(imp_dt, datetime.datetime) and imp_dt.date() < date.today():
        print('⛔ 「楽天CSV読込」がまだ実行されていません(手順③ V2-7 未実施)。')
        sys.exit(2)

    v3 = 'IdText' in pre.get('Module_Rakuten_Tool.bas', '')
    print(f'   修正版の世代: {"V3(識別子の型統一＋集計先の文字列書式)" if v3 else "V2(取込4列の文字列書式)"}')

    # ── 1. 文字列の完全一致 ─────────────────────────────
    print('■ 1. 取込先 vs Import xlsx(同じ順序で並べて比較)')
    src = [r[:11] for r in rows_of(IMPORT, '在庫', 2)
           if (r[3] not in (None, '')) or (r[1] not in (None, ''))]
    tgt = [r[:11] for r in rows_of(TEST, '楽天CSV取込', 3)
           if r and (r[3] not in (None, '') or r[1] not in (None, ''))]
    check(len(src) == len(tgt), f'行数 Import {len(src)} / 取込先 {len(tgt)}')
    # 判定の対象は「元xlsxで文字列だったセル」。元xlsxの時点で数値だったセルは、この修正では復元できない
    # (ChatGPT 2026-09-12: 元データで既に失われた文字は戻らない)。数値のまま入るのは想定内として件数だけ出す
    for ci, name in ID_COLS.items():
        exact = num_kept = num_text = mism = 0
        bad = []
        types_from_str = Counter()
        for s, t in zip(src, tgt):
            sv, tv = s[ci], t[ci]
            if sv is None and tv is None:
                exact += 1
                continue
            if isinstance(sv, str):
                types_from_str[type(tv).__name__] += 1
                if tv == sv:
                    exact += 1
                else:
                    mism += 1
                    if len(bad) < 5:
                        bad.append((sv, tv))
            else:
                if tv == expected_text(sv):
                    num_text += 1                       # 数値→桁の文字列(もし文字列化されていれば)
                elif isinstance(tv, (int, float)) and expected_text(tv) == expected_text(sv):
                    num_kept += 1                       # 数値のまま(桁は同じ)
                else:
                    mism += 1
                    if len(bad) < 5:
                        bad.append((sv, tv))
        check(mism == 0, f'{name}: 元が文字列→完全一致 {exact} / 元が数値→数値のまま {num_kept}・文字列化 {num_text} / 不一致 {mism}')
        if v3:   # ③ 識別子の型統一(IdText)が入っていれば、元が数値のセルも桁の文字列になっているはず
            check(num_kept == 0, f'{name}: (V3) 元が数値だったセルの文字列化 {num_text} / 数値のまま {num_kept}(0であること)')
        for b in bad:
            print(f'       不一致例 {b[0]!r} → {b[1]!r}')
        non_str = sum(v for k, v in types_from_str.items() if k != 'str')
        check(non_str == 0, f'{name}: 元が文字列だったセルのうち保存・再読込後に文字列でないもの {non_str}件')

    # ── 2. ファイルの健全性(VBA・図形が残っているか) ─────
    print('■ 2. 保存後のファイル(VBA・図形)')
    import zipfile
    names = zipfile.ZipFile(TEST).namelist()
    check('xl/vbaProject.bin' in names, 'vbaProject.bin が残っている')
    check(any('drawing' in n for n in names), 'drawing(ボタン等)が残っている')

    # ── 3. VBAの差分が V-2 の分だけか ────────────────────
    print('■ 3. VBAソース')
    try:
        from oletools.olevba import VBA_Parser
        vp = VBA_Parser(TEST)
        code = {n: c for (_, _, n, c) in vp.extract_macros()}
        vp.close()
        got = code.get('Module_Rakuten_Tool.bas', '').replace('\r\n', '\n').replace('\r', '\n')
        want = open(PATCHED_BAS, encoding='utf-8').read()
        norm = lambda t: [l.rstrip() for l in t.split('\n') if not l.startswith('Attribute VB_')]
        a, b = norm(got), norm(want)
        import difflib
        d = [x for x in difflib.unified_diff(b, a, lineterm='', n=0) if not x.startswith(('---', '+++', '@@'))]
        check(not d, f'Module_Rakuten_Tool が VBA/Module_Rakuten_Tool.bas(修正版)と一致(差分 {len(d)} 行)')
        for x in d[:10]:
            print('       ', x)
        check('SetIdColumnsAsText' in got, 'SetIdColumnsAsText が含まれる')
        others = [n for n in code if n != 'Module_Rakuten_Tool.bas']
        print(f'     その他のモジュール: {others}')
    except ImportError:
        check(False, 'oletools が無いためVBAを読めない')

    # ── 4. 既存機能: 集計結果を baseline と比較 ──────────
    print('■ 4. 集計結果(baseline=現行VBA vs test=V-2)')
    def agg(path):
        wb = load_workbook(path, read_only=True, data_only=True)
        ws = wb['在庫金額集計']
        head = {ws.cell(3, 1).value: ws.cell(3, 2).value, ws.cell(3, 5).value: ws.cell(3, 6).value,
                ws.cell(4, 1).value: ws.cell(4, 2).value, ws.cell(4, 5).value: ws.cell(4, 6).value,
                ws.cell(5, 1).value: ws.cell(5, 2).value, ws.cell(5, 5).value: ws.cell(5, 6).value}
        det = {}
        for r in ws.iter_rows(min_row=10, values_only=True):
            if not r or r[2] in (None, ''):
                continue
            key = (str(r[1]).strip(), str(r[2]).strip())          # 管理番号×SKU
            det[key] = r[:12]
        chk = [r[:8] for r in wb['要確認一覧'].iter_rows(min_row=3, values_only=True) if r and r[0] not in (None, '')]
        wb.close()
        return head, det, chk
    hb, db, cb = agg(BASE)
    ht, dt, ct = agg(TEST)
    for k in hb:
        same = hb[k] == ht.get(k)
        print(f'     {"✅" if same else "⚠️"} {k}: baseline {hb[k]} / test {ht.get(k)}')
    diffs = []
    for key in set(db) | set(dt):
        b, t = db.get(key), dt.get(key)
        if b is None or t is None:
            diffs.append((key, '片方にしか無い', b, t))
            continue
        for ci, lab in ((4, '在庫数'), (5, '採用原価'), (6, '原価取得元'), (7, '在庫金額'), (8, '販売価格')):
            if b[ci] != t[ci]:
                diffs.append((key, lab, b[ci], t[ci]))
    print(f'     明細の差 {len(diffs)} 件(採用原価・取得元・在庫金額・在庫数・販売価格。商品番号列は文字列化で変わるので除外)')
    for d in diffs[:20]:
        print('       ', d)
    print(f'     要確認一覧: baseline {len(cb)} 行 / test {len(ct)} 行')
    if v3:
        # ② 転記先の識別子が文字列で保たれているか(先頭0のJANが在庫金額集計に残るか)
        wb = load_workbook(TEST, read_only=True, data_only=True)
        ag = [r for r in wb['在庫金額集計'].iter_rows(min_row=10, values_only=True) if r and r[2] not in (None, '')]
        ck = [r for r in wb['要確認一覧'].iter_rows(min_row=3, values_only=True) if r and r[0] not in (None, '')]
        wb.close()
        for label, rows, cols in (('在庫金額集計', ag, (0, 1, 2, 11)), ('要確認一覧', ck, (0, 1, 2))):
            nonstr = sum(1 for r in rows for c in cols if r[c] not in (None, '') and not isinstance(r[c], str))
            lead0 = sum(1 for r in rows for c in cols if isinstance(r[c], str) and r[c].startswith('0'))
            check(nonstr == 0, f'(V3) {label}: 識別子列の文字列以外 {nonstr}件 / 先頭0を保った値 {lead0}件')
        # 金額・数量が変わっていないこと(クリニーク以外)
        print('     (V3) 金額・数量・集計値は上の baseline 比較のとおり')
    # 差は「説明できるもの」だけか … ここでは列挙に留め、判断は人が行う
    diff_keys = {d[0] for d in diffs}
    expected_key = {k for k in diff_keys if k[0].lstrip('0') == '192333006122'}
    check(diff_keys <= expected_key,
          f'明細の差がある出品 {len(diff_keys)}件 — 想定は 管理番号 0192333006122(クリニーク)だけ'
          '(先頭0が保たれ、0の落ちた原価マスターと一致しなくなる。原価マスター側の修正で解消)')

    print('\n' + ('🟢 PASS' if not fails else f'🔴 FAIL {len(fails)}件'))
    for f in fails:
        print('   -', f)


if __name__ == '__main__':
    main()
