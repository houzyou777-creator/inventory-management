# -*- coding: utf-8 -*-
"""build_stock_summary.py — 在庫更新パイプラインの工程4・5をスクリプトにする(2026-09-20。8/3 は手作業だった)

    工程4: 楽天在庫リスト_import.xlsx の「楽天単独在庫」シートを再生成
           = 「在庫」シートのうち 商品番号が -a で終わらない行(Amazon共有を除外)。I列は =SUM(G*H)、最終行に合計
    工程5: 全体在庫サマリー_v1.0.xlsx を更新
           = 楽天単独在庫 + Amazon FBA + Amazon 自己発送(共有在庫は Amazon 側のみ計上)。既存の体裁を保ち、値だけ更新

使い方(順序は 開発 → コピー作成 → コピーで実行 → 再読込検証 → 結果報告 → 承認 → 本番):
    python3 build_stock_summary.py plan  --target copy|production            何が変わるかを表示(書かない)
    python3 build_stock_summary.py make-copy                                 本番3ファイルを Summary_test/ へ複製(本番は読むだけ)
    python3 build_stock_summary.py apply --target copy --rakuten-source "2026-09-16 08:11:29" --amazon-fetch "2026-09-16 08:15"
    python3 build_stock_summary.py apply --target production --confirm 本番 --rakuten-source ... --amazon-fetch ...

基準日時(ChatGPT 2026-09-21 条件①): 固定文字列にしない。
    ・楽天 資料基準日時 … 引数 --rakuten-source(RMS CSV の出力日時)
    ・楽天 取込実行日時 … 在庫ツールの「楽天CSV取込」K列(読込日時)から読む(入力データ)。--rakuten-run で上書き可
    ・Amazon 取得日時   … 引数 --amazon-fetch
    ・Amazon 在庫基準日時 … 既定「未取得」(取得日時を基準日時として扱わない)。分かれば --amazon-basis
在庫数の扱い(条件②): 0 = 在庫数0という値／空欄 = 未入力／非数値 = データ異常。空欄・非数値の行は金額・数量に入れず件数と行を表示。
単価の扱い: 空欄 = 未入力(金額に入れない・件数表示)／0 = 単価0という値(金額0で計上・件数表示)／非数値 = データ異常(件数表示)。
共有在庫の判定(ChatGPT 2026-09-21): 通常の「商品番号の末尾 -a」＋**承認済み共有在庫の例外一覧**(`SourceData/共有在庫_例外一覧.csv`)。
    ・元データ(商品番号)は書き換えない。例外一覧に「誰が・いつ・なぜ」共有と判定したかを残す
    ・除外されるのは 区分=共有 かつ 承認日 が入っている行だけ(候補・未承認・別在庫・無効化済みは除外しない)
    ・キーは 楽天商品管理番号×SKU管理番号(照合ルールと同じ)。列: 楽天商品管理番号, SKU管理番号, 商品番号(判定時), 区分(共有/別在庫/不明),
      判定者, 判定日, 根拠, 承認者, 承認日, 無効化日, 無効化理由
「在庫金額」はリストの単価(G列)だけで計算する(在庫ツールの原価マスター参照とは別物。ツール値は参考として併記)。
共有在庫込み・原価未登録ありの集計なので、シートに「全社の確定在庫数量・確定評価額ではない」と明記する。
"""
import argparse
import os
import shutil
import sys
import warnings
from datetime import datetime

warnings.filterwarnings('ignore')
from openpyxl import load_workbook

# 共通モジュール(02_Analytics/Python/momiji_paths.py)の場所は __file__ からの相対位置で求める
sys.path.insert(0, os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..', '02_Analytics', 'Python')))
import momiji_paths as MP  # noqa: E402

SD = MP.INV_SD
PROD = dict(rakuten=SD + '/Import/楽天在庫リスト_import.xlsx', amazon=SD + '/Amazon在庫リスト_import.xlsx',
            summary=SD + '/全体在庫サマリー_v1.0.xlsx', tool=SD + '/楽天在庫金額集計ツール_v1.0.xlsm')
COPY_DIR = SD + '/Summary_test'
COPY = dict(rakuten=COPY_DIR + '/Import/楽天在庫リスト_import.xlsx', amazon=COPY_DIR + '/Amazon在庫リスト_import.xlsx',
            summary=COPY_DIR + '/全体在庫サマリー_v1.0.xlsx', tool=PROD['tool'])       # ツールは読むだけなので本番を参照
EXCEPTIONS = SD + '/共有在庫_例外一覧.csv'      # 承認済み共有在庫の例外(元データは書き換えない)
EXC_HEAD = ['楽天商品管理番号', 'SKU管理番号', '商品番号(判定時)', '区分', '判定者', '判定日', '根拠', '承認者', '承認日', '無効化日', '無効化理由']
COLS = 11   # JAN, 管理番号, 商品番号, SKU, 商品名, 販売価格, 単価, 在庫数, 合計金額, JAN検証結果, AmazonのJAN
C_PRICE, C_STOCK = 6, 7   # 0始まり: 単価, 在庫数


def backup(path, label):
    d = os.path.join(os.path.dirname(path), 'Backup')
    os.makedirs(d, exist_ok=True)
    base, ext = os.path.splitext(os.path.basename(path))
    dst = os.path.join(d, f'{base}_backup_{datetime.now():%Y%m%d_%H%M%S}_{label}{ext}')
    if os.path.exists(dst):
        sys.exit(f'❌ バックアップ名が既にあります: {dst}')
    shutil.copy2(path, dst)
    return dst


def num(v):
    return v if isinstance(v, (int, float)) and not isinstance(v, bool) else None


def classify_cell(v):
    """セルの状態: 'value'(数値) / 'blank'(空欄=未入力) / 'bad'(非数値=データ異常)"""
    if v is None or (isinstance(v, str) and v.strip() == ''):
        return 'blank'
    if num(v) is not None:
        return 'value'
    if isinstance(v, str):
        try:
            float(v.replace(',', ''))
            return 'text_number'               # "0" や "1,234" のような文字列の数字(型が違う。異常として数える)
        except ValueError:
            return 'bad'
    return 'bad'


def norm(v):
    return str(v).strip().upper() if v not in (None, '') else ''


def load_exceptions(path=EXCEPTIONS):
    """承認済みの共有在庫例外を {(管理番号, SKU): 行dict} で返す。ファイルが無ければ空(それが正しい状態)。
    有効 = 区分が「共有」 かつ 承認日あり かつ 無効化日なし。それ以外(候補・別在庫・不明・無効化)は除外に使わない。"""
    import csv
    approved, others = {}, []
    if not os.path.exists(path):
        return approved, others
    with open(path, encoding='utf-8-sig', newline='') as f:
        for d in csv.DictReader(f):
            if not d.get('楽天商品管理番号'):
                continue
            key = (norm(d['楽天商品管理番号']), norm(d.get('SKU管理番号')))
            ok = (str(d.get('区分') or '').strip() == '共有' and str(d.get('承認日') or '').strip() != ''
                  and str(d.get('無効化日') or '').strip() == '')
            (approved.__setitem__(key, d) if ok else others.append(d))
    return approved, others


def rakuten_rows(path, exceptions=None):
    """(全行, 単独在庫行, 通常-a除外数, 例外で除外した行) を返す。"""
    ws = load_workbook(path, read_only=True, data_only=True)['在庫']
    rows = [tuple(r[:COLS]) + (None,) * (COLS - len(r[:COLS])) for r in ws.iter_rows(min_row=2, values_only=True)
            if r and (r[1] not in (None, '') or r[3] not in (None, ''))]
    exceptions = exceptions or {}
    solo, by_exc, n_a = [], [], 0
    for r in rows:
        if isinstance(r[2], str) and r[2].strip().lower().endswith('-a'):
            n_a += 1
            continue
        if (norm(r[1]), norm(r[3])) in exceptions:
            by_exc.append(r)
            continue
        solo.append(r)
    return rows, solo, n_a, by_exc


def stats(rows, label, id_cols=(1, 3)):
    """集計と品質件数。在庫数が空欄/非数値の行は数量・金額に入れない。
    戻り: dict(sku, qty, amt, price_blank, price_zero, price_bad, stock_blank, stock_bad, rows_excluded)"""
    st = dict(sku=len(rows), qty=0, amt=0.0, price_blank=0, price_zero=0, price_bad=0, stock_blank=0, stock_bad=0, rows_excluded=[])
    for r in rows:
        sc, pc = classify_cell(r[C_STOCK]), classify_cell(r[C_PRICE])
        ident = ' / '.join(str(r[c]) for c in id_cols if r[c] not in (None, ''))
        if sc != 'value':
            st['stock_blank' if sc == 'blank' else 'stock_bad'] += 1
            st['rows_excluded'].append((label, ident, '在庫数 ' + ('空欄' if sc == 'blank' else f'非数値({r[C_STOCK]!r})')))
            continue
        q = r[C_STOCK]
        st['qty'] += q
        if pc == 'value':
            st['amt'] += q * r[C_PRICE]
            if r[C_PRICE] == 0 and q > 0:
                st['price_zero'] += 1
        elif pc == 'blank':
            if q > 0:
                st['price_blank'] += 1
        else:
            st['price_bad'] += 1
            st['rows_excluded'].append((label, ident, f'単価 非数値({r[C_PRICE]!r})'))
    st['amt'] = round(st['amt'])
    return st


def amazon_sections(path):
    ws = load_workbook(path, read_only=True, data_only=True)['在庫']
    secs, cur = {}, None
    for r in ws.iter_rows(min_row=1, values_only=True):
        v = r[0]
        if isinstance(v, str) and v.startswith('▼'):
            cur = 'FBA' if 'FBA' in v else '自己発送'
            secs[cur] = []
            continue
        if cur and r[3] not in (None, ''):
            secs[cur].append(tuple(r[:9]) + (None,) * (9 - len(r[:9])))
    return secs


def tool_figures(path):
    wb = load_workbook(path, read_only=True, data_only=True)
    ag = wb['在庫金額集計']
    tk = wb['楽天CSV取込']
    run_dt = tk.cell(3, 11).value                                # K3 = 読込日時(取込実行日時)
    return dict(amount=ag.cell(3, 2).value, qty=ag.cell(4, 2).value, unreg=ag.cell(5, 6).value, dt=ag.cell(6, 2).value,
                run_dt=(run_dt.strftime('%Y-%m-%d %H:%M:%S') if hasattr(run_dt, 'strftime') else (str(run_dt) if run_dt else '未取得')))


def write_solo_sheet(wb, solo):
    ws_src = wb['在庫']
    if '楽天単独在庫' in wb.sheetnames:
        idx = wb.sheetnames.index('楽天単独在庫')
        del wb['楽天単独在庫']
    else:
        idx = 1
    ws = wb.create_sheet('楽天単独在庫', idx)
    for c in range(1, COLS + 1):
        ws.cell(1, c).value = ws_src.cell(1, c).value
        ws.cell(1, c).font = ws_src.cell(1, c).font.copy()
        ws.cell(1, c).fill = ws_src.cell(1, c).fill.copy()
        ws.column_dimensions[ws.cell(1, c).column_letter].width = ws_src.column_dimensions[ws_src.cell(1, c).column_letter].width
    for i, r in enumerate(solo, 2):
        for c, v in enumerate(r, 1):
            if c == 9:
                ws.cell(i, c).value = f'=SUM(G{i}*H{i})'          # 8/3 と同じ形(単価×在庫数)
            else:
                ws.cell(i, c).value = v
        for c in (1, 2, 3, 4):
            ws.cell(i, c).number_format = '@'
    last = len(solo) + 1
    ws.cell(last + 1, 5).value = '合計'
    ws.cell(last + 1, 8).value = f'=SUM(H2:H{last})'
    ws.cell(last + 1, 9).value = f'=SUM(I2:I{last})'
    return last


def update_summary(path, rk, fba, own, tool, n_shared, stamps, n_exc=0):
    wb = load_workbook(path)
    ws = wb['全体サマリー']
    ws['A2'].value = (f'集計日時: {datetime.now():%Y/%m/%d %H:%M}　（データ元: 楽天在庫リスト_import.xlsx「楽天単独在庫」シート / Amazon在庫リスト_import.xlsx）')
    for row, (label, st, note) in zip((5, 6, 7), (
            ('楽天 単独在庫', rk, f'Amazon共有(-a {n_shared - n_exc}行＋承認済み例外 {n_exc}行)は除外済み。在庫金額はリストの単価×在庫数。単価0(在庫>0) {rk["price_zero"]}行'
                                  + (f'／在庫数 空欄{rk["stock_blank"]}・非数値{rk["stock_bad"]} は除外' if rk['stock_blank'] or rk['stock_bad'] else '')),
            ('Amazon FBA', fba, (f'在庫数 空欄{fba["stock_blank"]}・非数値{fba["stock_bad"]} は除外' if fba['stock_blank'] or fba['stock_bad'] else None)),
            ('Amazon 自己発送', own, '楽天との共有在庫の正式値を含む' + (f'／在庫数 空欄{own["stock_blank"]}・非数値{own["stock_bad"]} は除外' if own['stock_blank'] or own['stock_bad'] else '')))):
        ws.cell(row, 1).value = label
        ws.cell(row, 2).value = st['sku']
        ws.cell(row, 3).value = st['qty']
        ws.cell(row, 4).value = st['amt']
        ws.cell(row, 5).value = st['price_blank']
        ws.cell(row, 6).value = note
    notes = [
        '■ 注記(自動更新)',
        f'・基準日時 — 楽天: 資料基準日時 {stamps["rakuten_source"]} ／ 取込実行日時 {stamps["rakuten_run"]}。'
        f' Amazon: 取得日時 {stamps["amazon_fetch"]} ／ 在庫基準日時 {stamps["amazon_basis"]}(取得日時を基準日時として扱わない)。',
        '・この表は共有在庫込み・原価未登録ありの「掲載分集計」であり、全社の確定在庫数量・確定評価額ではない(ChatGPT 2026-09-19)。',
        f'・楽天の在庫ツール(原価マスター参照・共有在庫込み): 総在庫金額 {tool["amount"]:,} / 総在庫数量 {tool["qty"]:,} / 原価未登録 {tool["unreg"]}(集計 {tool["dt"]})。参考値。',
        '・最新CSVに無い商品(楽天94件・旧469個・参考額1,286,025円は削除記録CSV、Amazon40件・旧186個・参考額246,577円は別シート「未掲載_未確認」)は本表に含めない。現在在庫は未確認。',
        '・単価未入力 = 在庫>0 で単価が空欄の行数(金額に入れない)。単価0 = 0円という値(金額0で計上・備考に件数)。在庫数の空欄=未入力・非数値=データ異常は数量・金額から除外し備考に件数。',
    ]
    start = None
    for r in range(1, ws.max_row + 1):
        if isinstance(ws.cell(r, 1).value, str) and ws.cell(r, 1).value.startswith('■ 注記'):
            start = r
            break
    if start is None:
        start = ws.max_row + 2
    for k in range(start, start + 8):                            # 旧注記(行数が違っても)を消してから書く
        ws.cell(k, 1).value = None
    for k, t in enumerate(notes):
        ws.cell(start + k, 1).value = t
    wb.save(path)


def report(rows, solo, rk, fba, own, tool, stamps, n_a=0, by_exc=(), others=()):
    tot = dict(sku=rk['sku'] + fba['sku'] + own['sku'], qty=rk['qty'] + fba['qty'] + own['qty'], amt=rk['amt'] + fba['amt'] + own['amt'])
    print(f'■ 楽天 在庫 {len(rows)} 行 → 単独在庫 {len(solo)} 行(共有 -a {n_a} 行 ＋ 承認済み例外 {len(by_exc)} 行を除外)')
    for r in by_exc:
        print(f'      例外除外: {r[1]} / {r[3]} / {r[2]} / {str(r[4])[:24]} / 在庫 {r[7]} / 単価 {r[6]}')
    if others:
        print(f'   例外一覧の未承認・別在庫・無効化行 {len(others)} 件は除外に使っていない: ' + ', '.join(f"{d['楽天商品管理番号']}({d.get('区分')}/{'承認あり' if d.get('承認日') else '未承認'})" for d in others))
    for label, st in (('楽天 単独在庫', rk), ('Amazon FBA', fba), ('Amazon 自己発送', own)):
        print(f'   {label}: SKU {st["sku"]} / 在庫数 {st["qty"]:,} / 在庫金額 {st["amt"]:,} / 単価未入力 {st["price_blank"]} / 単価0 {st["price_zero"]} / 単価非数値 {st["price_bad"]}'
              f' / 在庫数 空欄 {st["stock_blank"]} / 在庫数 非数値 {st["stock_bad"]}')
    for st in (rk, fba, own):
        for x in st['rows_excluded']:
            print('      除外行:', x)
    print(f'   合計(共有二重計上なし): SKU {tot["sku"]} / {tot["qty"]:,} 個 / {tot["amt"]:,} 円')
    print(f'   参考 在庫ツール(原価マスター参照・共有込み): {tool["amount"]:,} / {tool["qty"]:,} / 未登録 {tool["unreg"]} ({tool["dt"]})')
    print(f'   基準日時: 楽天 資料 {stamps["rakuten_source"]} / 取込実行 {stamps["rakuten_run"]} ／ Amazon 取得 {stamps["amazon_fetch"]} / 在庫基準 {stamps["amazon_basis"]}')
    return tot


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('mode', choices=['plan', 'make-copy', 'apply'])
    ap.add_argument('--target', choices=['copy', 'production'], default='copy')
    ap.add_argument('--confirm', default='')
    ap.add_argument('--rakuten-source', default='未取得', help='楽天 資料基準日時(RMS CSV の出力日時)')
    ap.add_argument('--rakuten-run', default=None, help='楽天 取込実行日時(省略時は在庫ツールの読込日時から)')
    ap.add_argument('--amazon-fetch', default='未取得', help='Amazon 取得日時(ファイル時刻)')
    ap.add_argument('--amazon-basis', default='未取得', help='Amazon 在庫基準日時(分からなければ未取得のまま)')
    ap.add_argument('--exceptions', default=EXCEPTIONS, help='承認済み共有在庫の例外一覧CSV(既定 SourceData/共有在庫_例外一覧.csv)')
    a = ap.parse_args()

    if a.mode == 'make-copy':
        os.makedirs(COPY_DIR + '/Import', exist_ok=True)
        for k in ('rakuten', 'amazon', 'summary'):
            if os.path.exists(COPY[k]):
                print('   既存のコピーを退避:', backup(COPY[k], '再複製前'))
            shutil.copy2(PROD[k], COPY[k])
            print(f'   複製: {PROD[k]} → {COPY[k]}')
        return

    P = COPY if a.target == 'copy' else PROD
    if a.target == 'production' and a.confirm != '本番':
        sys.exit('❌ 本番に書くには --confirm 本番 が必要です(コピー検証と承認の後)')
    for k in ('rakuten', 'amazon', 'summary'):
        if not os.path.exists(P[k]):
            sys.exit(f'❌ ありません: {P[k]}' + ('(先に make-copy)' if a.target == 'copy' else ''))
    print(f'■ 対象: {a.target} / {P["rakuten"]} / {P["amazon"]} / {P["summary"]}')
    exc, exc_others = load_exceptions(a.exceptions)
    print(f'   共有在庫の例外一覧: {a.exceptions} → 承認済み {len(exc)} 件・その他 {len(exc_others)} 件' if os.path.exists(a.exceptions) else '   共有在庫の例外一覧: なし(通常の -a 判定のみ)')
    rows, solo, n_a, by_exc = rakuten_rows(P['rakuten'], exc)
    n_shared = len(rows) - len(solo)                      # -a ＋ 承認済み例外
    rk = stats(solo, '楽天単独')
    secs = amazon_sections(P['amazon'])
    fba, own = stats(secs.get('FBA', []), 'Amazon FBA', (1, 3)), stats(secs.get('自己発送', []), 'Amazon 自己発送', (1, 3))
    tool = tool_figures(P['tool'])
    stamps = dict(rakuten_source=a.rakuten_source, rakuten_run=a.rakuten_run or tool['run_dt'], amazon_fetch=a.amazon_fetch, amazon_basis=a.amazon_basis)
    tot = report(rows, solo, rk, fba, own, tool, stamps, n_a, by_exc, exc_others)
    if a.mode == 'plan':
        print('(plan: 書いていません)')
        return

    print('バックアップ:', backup(P['rakuten'], '単独在庫再生成前'))
    print('バックアップ:', backup(P['summary'], '更新前'))
    wb = load_workbook(P['rakuten'])
    last = write_solo_sheet(wb, solo)
    wb.save(P['rakuten'])
    update_summary(P['summary'], rk, fba, own, tool, n_shared, stamps, len(by_exc))

    # ── 再読込で検証 ──
    fails = []

    def chk(cond, msg):
        print(('   ✅ ' if cond else '   ❌ ') + msg)
        if not cond:
            fails.append(msg)
    w2 = load_workbook(P['rakuten'])
    s2 = w2['楽天単独在庫']
    got = [tuple(r[:COLS]) for r in s2.iter_rows(min_row=2, max_row=last, values_only=True)]
    chk(w2.sheetnames[:2] == ['在庫', '楽天単独在庫'], f'シート順 {w2.sheetnames}')
    chk(len(got) == len(solo) == rk['sku'], f'楽天単独 SKU数 {len(got)} = 除外後 {len(solo)}')
    chk(all(g[:8] == s[:8] for g, s in zip(got, solo)), '楽天単独 全行の値(8列)が元と一致')
    chk(sum(1 for r in got for c in (0, 1, 2, 3) if isinstance(r[c], (int, float))) == 0, '識別子は文字列(数値型0)')
    chk(not any(isinstance(r[2], str) and r[2].strip().lower().endswith('-a') for r in got), f'-a 行が単独在庫に無い(-a 除外 {n_a} 行)')
    chk(not any((norm(r[1]), norm(r[3])) in exc for r in got), f'承認済み例外 {len(by_exc)} 行が単独在庫に無い')
    chk(all((norm(r[1]), norm(r[3])) not in exc for r in got) and len(got) + n_a + len(by_exc) == len(rows), f'楽天 {len(rows)} = 単独 {len(got)} + -a {n_a} + 例外 {len(by_exc)}')
    chk(s2.cell(last + 1, 8).value == f'=SUM(H2:H{last})' and s2.cell(last + 1, 9).value == f'=SUM(I2:I{last})' and s2.cell(2, 9).value == '=SUM(G2*H2)',
        '合計行と I列の数式')
    a2 = [tuple(r[:COLS]) for r in w2['在庫'].iter_rows(min_row=2, values_only=True) if r and (r[1] or r[3])]
    chk([x[:8] for x in a2] == [x[:8] for x in rows], f'元の「在庫」シート不変({len(rows)}行)')
    sm = load_workbook(P['summary'])['全体サマリー']
    chk([sm.cell(8, c).value for c in (2, 3, 4, 5)] == ['=SUM(B5:B7)', '=SUM(C5:C7)', '=SUM(D5:D7)', '=SUM(E5:E7)'], '全体サマリー 合計行の数式')
    vals = {sm.cell(r, 1).value: (sm.cell(r, 2).value, sm.cell(r, 3).value, sm.cell(r, 4).value, sm.cell(r, 5).value) for r in (5, 6, 7)}
    chk(vals.get('楽天 単独在庫') == (rk['sku'], rk['qty'], rk['amt'], rk['price_blank']), f'サマリー 楽天単独 {vals.get("楽天 単独在庫")}')
    chk(vals.get('Amazon FBA') == (fba['sku'], fba['qty'], fba['amt'], fba['price_blank']), f'サマリー FBA {vals.get("Amazon FBA")}')
    chk(vals.get('Amazon 自己発送') == (own['sku'], own['qty'], own['amt'], own['price_blank']), f'サマリー 自己発送 {vals.get("Amazon 自己発送")}')
    # 二重計上: 楽天単独の管理番号(ASIN)が Amazon 側と重ならない(共有は -a で除外済みなので、重なりは共有印の漏れ)
    amz_asins = {str(r[1]).strip().upper() for sec in secs.values() for r in sec if r[1]}
    overlap = [r[1] for r in solo if str(r[1]).strip().upper() in amz_asins]
    chk(not any(isinstance(r[2], str) and r[2].strip().lower().endswith('-a') for r in got) and rk['sku'] + n_shared == len(rows),
        f'二重計上なし: 共有(-a {n_a}＋例外 {len(by_exc)})は楽天単独に含めず Amazon 側だけで計上(楽天 {len(rows)} = 単独 {rk["sku"]} + 共有 {n_shared})')
    print(f'   参考: 共有印(-a)が無いのに Amazon と同じ管理番号の楽天単独行 {len(overlap)} 件(共有印の漏れか、楽天専用在庫か。人が確認する材料)')
    chk(tot['amt'] == rk['amt'] + fba['amt'] + own['amt'], f'全体在庫金額 {tot["amt"]:,} = 楽天単独 + FBA + 自己発送')
    print('■ 結果:', 'PASS' if not fails else f'FAIL {len(fails)}件')
    sys.exit(0 if not fails else 1)


if __name__ == '__main__':
    main()
