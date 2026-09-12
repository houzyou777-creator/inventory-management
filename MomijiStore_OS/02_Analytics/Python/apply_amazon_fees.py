# -*- coding: utf-8 -*-
"""apply_amazon_fees.py — トランザクションレポートの実額手数料をAmazon KPIシートへ反映

使い方:
    python3 apply_amazon_fees.py <Transaction.csv> <シート名>

例:
    python3 apply_amazon_fees.py ../SourceData/2026JulMonthlyTransaction.csv 7月

処理内容:
1. トランザクションレポート(注文+返金)をSKU別に集計
   手数料 = 販売手数料 + FBA手数料 + トランザクションに関するその他の手数料
2. KPIシートのK列(販売手数料)を、SKUが一致した行だけ実額に置き換える
   一致しないSKU(決済が翌月にずれた注文など)は暫定15%の数式のまま残す
3. 経費ブロックへ自動入力: プロモーション費(プロモ割引+ポイント費用)、
   その他手数料(FBA在庫保管・返送・月額登録料などの注文外費用)、送料(購入配送ラベル実費)
   ※広告費・梱包資材はレポートに含まれないため黄色のまま
4. Excel(AppleScript)で再計算して検証

注意: ビジネスレポート(注文日基準・税込)とトランザクション(決済日基準・税抜)は
集計基準が異なるため、手数料は「その月に決済された実額」として扱う。
"""
import csv
import os
import shutil
import subprocess
import sys
from collections import defaultdict
from datetime import date

from openpyxl import load_workbook
from openpyxl.styles import PatternFill

BASE = '/Users/hide0726/Desktop/Claude Code/MomijiStore_OS'
OUT_FILE = os.environ.get(
    'AMZ_KPI_FILE_OVERRIDE',
    BASE + '/02_Analytics/SourceData/Amazon運営 KPI管理シート.xlsx')
MASTER_FILE = BASE + '/01_InventoryManagement/SourceData/商品マスター_単品_v1.0.xlsx'
OUTPUT_DIR = BASE + '/01_InventoryManagement/SourceData/Output'
NOFILL = PatternFill(fill_type=None)


# --- 識別子の解決 -----------------------------------------------------------
#  ⚠️ 単一キーに依存しないこと。
#  2026-09-08、Amazonのビジネスレポートから「SKU」列が消え、
#  KPIシートのSKU列が全行空になった。手数料はSKUだけで突合していたため
#  281SKU・¥759,756 のうち 0件しか実額化できなかった。
#
#  以後は複数のキーで解決する。1つのキーが欠けても処理全体は止めない。
#  将来 JAN など別のキーを足すときは、この関数と MATCH_ORDER を拡張する。
def build_key_index():
    """出品テーブルから 各識別子 → ASIN の対応表を作る。

    ASIN を突合の共通軸にする。Amazonのレポートで最も欠けにくい identifier のため。
    """
    wb = load_workbook(MASTER_FILE, read_only=True, data_only=True)
    amazon_sku2asin = {}
    for r in wb['出品テーブル'].iter_rows(min_row=2, values_only=True):
        asin, asku = r[5], r[6]
        if asin and asku:
            amazon_sku2asin.setdefault(str(asku).strip().upper(), str(asin).strip().upper())
    wb.close()
    # 将来キーを増やすときはここへ追加する(例: 'jan': jan2asin)
    return {'amazon_sku': amazon_sku2asin}


def num(s):
    s = s.replace(',', '').strip()
    return float(s) if s else 0.0


def read_transactions(path):
    with open(path, encoding='utf-8-sig') as f:
        rows = list(csv.reader(f))
    hi = next(i for i, r in enumerate(rows) if r and r[0].startswith('日付/時間'))
    fee_by_sku = defaultdict(float)
    promo = ship = other = 0.0
    for r in rows[hi + 1:]:
        if len(r) < 28:
            continue
        typ = r[2]
        if typ in ('注文', '返金'):
            fee_by_sku[r[4].strip().upper()] += num(r[23]) + num(r[24]) + num(r[25])
            promo += num(r[20]) + num(r[19])
        elif typ == '配送サービス':
            ship += num(r[27])
        elif typ == '振込み':
            continue
        else:
            other += num(r[27])
    return fee_by_sku, -promo, -ship, -other


def apply(sheet_name, fee_by_sku, promo, ship, other):
    wb = load_workbook(OUT_FILE)
    ws = wb[sheet_name]

    # シート行を SKU と ASIN の両方で引けるようにする。
    # 同一キーが複数行に分かれている場合があるため、行を集めてから売上比で按分する。
    rows_by_sku, rows_by_asin = {}, {}
    row = 8
    while ws.cell(row, 3).value is not None:     # C列(ASIN)が空になるまでがデータ
        sku = str(ws.cell(row, 2).value or '').strip().upper()
        asin = str(ws.cell(row, 3).value or '').strip().upper()
        sales = ws.cell(row, 9).value or 0
        if sku:
            rows_by_sku.setdefault(sku, []).append((row, sales))
        if asin:
            rows_by_asin.setdefault(asin, []).append((row, sales))
        row += 1
    data_rows = row - 8

    # --- 突合 ---------------------------------------------------------------
    #  優先順位: ① SKU一致  ② SKUで一致しない場合のみ AmazonSKU→ASIN 経由
    #  SKUが取れているならSKUを必ず優先する(最も細かい粒度のため)。
    key_index = build_key_index()
    groups = {}                      # 突合先キー -> [(行, 売上)]
    fee_by_group = defaultdict(float)   # 同一ASINへ複数SKU(旧SKU等)が集まる場合を合算
    via_count = defaultdict(int)
    unresolved = []                  # (SKU, 手数料, 理由)

    for sku, fee in fee_by_sku.items():
        if sku in rows_by_sku:                                   # ① SKU一致
            key, rlist, via = ('SKU', sku), rows_by_sku[sku], 'SKU一致'
        else:
            asin = key_index['amazon_sku'].get(sku)              # ② ASIN経由
            if asin and asin in rows_by_asin:
                key, rlist, via = ('ASIN', asin), rows_by_asin[asin], 'ASIN経由'
            else:
                reason = '出品テーブルにAmazonSKUが無い' if not asin else 'シートに該当ASINが無い'
                unresolved.append((sku, fee, reason))
                continue
        groups[key] = rlist
        fee_by_group[key] += fee
        via_count[via] += 1

    matched_rows = set()
    matched_fee = 0.0
    for key, rlist in groups.items():
        total_fee = round(-fee_by_group[key], 2)
        total_sales = sum(s for _, s in rlist)
        assigned = 0.0
        for i, (r, s) in enumerate(rlist):
            if i == len(rlist) - 1:
                v = round(total_fee - assigned, 2)   # 端数は最終行で調整
            else:
                share = s / total_sales if total_sales else 1 / len(rlist)
                v = round(total_fee * share, 2)
                assigned += v
            ws.cell(r, 11).value = int(v) if v == int(v) else v
            matched_rows.add(r)
        matched_fee += total_fee

    matched = len(matched_rows)
    unmatched = data_rows - matched               # 暫定15%の数式のまま残る行

    # 経費ブロックはM列のラベルで行を特定する(行番号のハードコード回避)
    label_row = {}
    for r in range(row, ws.max_row + 1):
        lab = ws.cell(r, 13).value
        if lab:
            label_row.setdefault(str(lab), r)
    for lab, val in [('プロモーション費', promo), ('その他手数料', other), ('送料', ship)]:
        r = label_row.get(lab)
        if r is None:
            raise SystemExit(f'経費ラベル「{lab}」が見つかりません')
        c = ws.cell(r, 14)
        c.value = round(val)
        c.fill = NOFILL

    # ヘッダーの注記を実額版に更新
    ws.cell(5, 2).value = '実額(トランザクションレポートより)。未マッチSKUのみ暫定15%'

    wb.save(OUT_FILE)
    return matched, unmatched, matched_fee, via_count, unresolved


def write_shortage_list(sheet_name, unresolved, total_fee):
    """未反映SKUを「商品情報不足一覧」として Output へ出す。

    🚫 エラーとして落とさない。**商品マスター・出品テーブルの整備対象**として扱う。
    キーが1つ欠けただけで処理全体を止めないための出口である。
    """
    if not unresolved:
        return None
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    path = f'{OUTPUT_DIR}/商品情報不足_AmazonSKU未解決_{sheet_name}_{date.today():%Y%m%d}.xlsx'
    from openpyxl import Workbook
    from openpyxl.styles import Font
    wb = Workbook()
    ws = wb.active
    ws.title = 'AmazonSKU未解決'
    ws['A1'] = f'商品情報不足一覧 — Amazon手数料が突合できなかったSKU({sheet_name})'
    ws['A1'].font = Font(bold=True, size=12)
    ws['A2'] = (f'作成日 {date.today():%Y-%m-%d} / 未反映 {len(unresolved)}件 / '
                f'未反映金額 ¥{-sum(f for _, f, _ in unresolved):,.0f} '
                f'(トランザクション総額 ¥{-total_fee:,.0f} に対して '
                f'{sum(f for _, f, _ in unresolved) / total_fee * 100:.1f}%)')
    ws['A3'] = '※ エラーではない。商品マスター・出品テーブルの整備対象として扱う。'
    head = ['AmazonSKU', '手数料(実額)', '未解決の理由', '想定される対応']
    for i, h in enumerate(head, 1):
        c = ws.cell(5, i); c.value = h; c.font = Font(bold=True)
    ACTION = {
        '出品テーブルにAmazonSKUが無い':
            '旧SKU/終了SKUの可能性。出品テーブルへ登録するか、対象外と判断する',
        'シートに該当ASINが無い':
            '当月の販売が無いASIN。決済が翌月へずれた注文の可能性',
    }
    for i, (sku, fee, reason) in enumerate(
            sorted(unresolved, key=lambda x: x[1]), 6):
        ws.cell(i, 1).value = sku
        ws.cell(i, 2).value = round(-fee)
        ws.cell(i, 3).value = reason
        ws.cell(i, 4).value = ACTION.get(reason, '')
    for col, w in zip('ABCD', (42, 14, 30, 52)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = 'A6'
    wb.save(path)
    return path


def recalc_via_excel(path):
    script = f'''
set p to POSIX file "{path}"
with timeout of 600 seconds
tell application "Microsoft Excel"
    set wasRunning to running
    open p
    delay 2
    calculate
    -- 「active workbook」は使わない。人が別のブックを開いていると、そちらを保存・閉じてしまう(2026-09-12 発生)
    set wb to workbook (name of (info for p))
    save wb
    close wb saving no
    if not wasRunning then quit
end tell
end timeout
return "ok"
'''
    r = subprocess.run(['osascript', '-e', script], capture_output=True, text=True, timeout=660)
    if r.returncode != 0:
        raise RuntimeError(f'Excel再計算に失敗: {r.stderr}')


def main():
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)
    csv_path, sheet_name = sys.argv[1], sys.argv[2]

    bdir = os.path.dirname(OUT_FILE) + '/Backup'
    os.makedirs(bdir, exist_ok=True)
    base = os.path.splitext(os.path.basename(OUT_FILE))[0]
    shutil.copy2(OUT_FILE, f'{bdir}/{base}_backup_{date.today():%Y%m%d}_{sheet_name}手数料実額化前.xlsx')
    print('バックアップ取得済み')

    fee_by_sku, promo, ship, other = read_transactions(csv_path)
    print(f'トランザクション集計: SKU {len(fee_by_sku)}件 / プロモ費 ¥{promo:,.0f} / 送料 ¥{ship:,.0f} / その他 ¥{other:,.0f}')

    matched, unmatched, matched_fee, via_count, unresolved = apply(
        sheet_name, fee_by_sku, promo, ship, other)
    print(f'K列置換: 実額 {matched}行 (¥{matched_fee:,.0f}) / 暫定15%のまま {unmatched}行')

    # --- 突合の検証レポート -------------------------------------------------
    total_fee = sum(fee_by_sku.values())
    un_fee = sum(f for _, f, _ in unresolved)
    rate = (total_fee - un_fee) / total_fee * 100 if total_fee else 0.0
    print()
    print('═══ 手数料突合の検証 ═══')
    print(f'  Transactionレポート総手数料 : ¥{-total_fee:,.0f} ({len(fee_by_sku)} SKU)')
    print(f'  実額反映できた金額          : ¥{matched_fee:,.0f}')
    # 金額ベースと行ベースを**両方**出す。金額だけ見ると品質を高く錯覚する。
    # 金額の大きい商品から実額化されるため、金額88%でも行では34%ということが起きる
    # (2026-09-09 ChatGPT指摘H)
    row_rate = matched / (matched + unmatched) * 100 if (matched + unmatched) else 0.0
    print(f'  実額化率(金額ベース)        : {rate:.1f}%')
    print(f'  実額化率(行ベース)          : {row_rate:.1f}%'
          f'  ← {matched}行 / {matched + unmatched}行')
    if rate - row_rate >= 10:
        print(f'  ⚠️ 金額と行で {rate - row_rate:.0f}pt の差がある。'
              '金額の大きい商品から実額化されているため、'
              '**行数で見ると多くが暫定15%のまま**である')
    print(f'  未反映件数                  : {len(unresolved)} SKU')
    print(f'  未反映金額                  : ¥{-un_fee:,.0f}')
    print('  突合の内訳                  : ' +
          ' / '.join(f'{k} {v}SKU' for k, v in sorted(via_count.items())) or '(なし)')
    diff = round(-total_fee - matched_fee - (-un_fee))
    print(f'  差額(総額 − 反映 − 未反映)  : ¥{diff:,.0f}' +
          ('  ✅ 完全一致' if abs(diff) < 1 else '  ❌ 不一致 — 要調査'))
    if unresolved:
        print(f'\n  未反映SKU一覧(金額の大きい順・上位10件 / 全{len(unresolved)}件):')
        for sku, fee, reason in sorted(unresolved, key=lambda x: x[1])[:10]:
            print(f'    ¥{-fee:>9,.0f}  {sku:<40s} {reason}')
        path = write_shortage_list(sheet_name, unresolved, total_fee)
        print(f'\n  → 商品情報不足一覧を出力: {path}')

    print('Excelで再計算中...')
    recalc_via_excel(OUT_FILE)

    wb = load_workbook(OUT_FILE, data_only=True)
    ws = wb[sheet_name]
    errs = [c.coordinate for row in ws.iter_rows() for c in row
            if isinstance(c.value, str) and c.value.startswith('#')]
    t = 8
    while ws.cell(t, 3).value is not None:
        t += 1
    print(f'数式エラー: {len(errs)}件')
    print(f'手数料合計 K{t}: ¥{ws.cell(t, 11).value:,.0f} / 粗利 N{t}: ¥{ws.cell(t, 14).value:,.0f} ({ws.cell(t, 15).value * 100:.1f}%)')
    gr = next((r for r in range(t, ws.max_row + 1) if ws.cell(r, 13).value == '限界利益'), None)
    if gr:
        v = ws.cell(gr, 14).value
        vr = ws.cell(gr + 1, 14).value
        print(f'限界利益: ¥{v:,.0f} ({vr * 100:.1f}%)' if v is not None else '')
    if errs:
        sys.exit('*** FAIL ***')
    print('PASS')


if __name__ == '__main__':
    main()
