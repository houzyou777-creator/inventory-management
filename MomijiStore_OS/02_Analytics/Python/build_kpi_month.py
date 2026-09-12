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


def build_cost_resolver(wb_values, ym=''):
    """仕入値の解決関数を作る(商品マスター → RMS商品番号 → 過去月シート)

    ⚠️ 商品マスターの標準原価は「単品1個あたり」とは限らない。
       出品がセット販売なら、その出品の原価は 単品原価 × 販売入数 である。
       出品テーブルの 販売入数・原価単位 が**確認済みのときだけ**計算し、
       未確認なら **None を返して黄色セルで人へ回す**(2026-09-10 承認W-4)。
       推測値を入れない。既存の「仕入値未解決」と同じ経路に落とす。

    ym: 生成する月 '2026-08'。RMS原価の対象月判定に使う(2026-09-12 Z-4)
    """
    cost_by_pid, pair2pid, ctrl2pid = load_master()
    pack, pos = S.load_pack_info()
    has_pack_cols = any(v is not None for v in pos.values())
    jan_by_pid = S.load_jan_by_pid() if hasattr(S, 'load_jan_by_pid') else {}

    # RMS商品番号の原価を使うか。**既定はOFF**(2026-09-12 時点で既定ONは保留)。
    #   RMS_COST=1 を付けたときだけ有効。本番生成でこの経路を通す前に承認を得る
    use_rms = os.environ.get('RMS_COST', '') not in ('', '0')
    import classify_product_numbers as CP

    def listing_of(ctrl, sku):
        """出品テーブルの1行を引く。重複・該当なしなら (None, 理由)。"""
        info, amb = S.lookup_pack(pack, '楽天', norm(ctrl), norm(sku))
        if amb:
            return None, amb
        return info, ''

    def composition_conflict(pn, info):
        """商品番号から読める構成と、出品テーブルの記録が食い違っていないか。食い違えば理由を返す。"""
        kind, cost, csrc, n, nsrc, disc, shared, comps, note = CP.parse(pn, {})
        rec_n = (info or {}).get('pack')
        if n is not None and isinstance(rec_n, (int, float)) and int(rec_n) != int(n):
            return f'構成違い: 商品番号の入数{n} ≠ 出品テーブルの販売入数{int(rec_n)}'
        pid = (info or {}).get('pid')
        full = [c for c in comps if len(c) == 13]
        mj = jan_by_pid.get(pid)
        if full and mj and mj not in full:
            return f'構成違い: 商品番号のJAN{full} ≠ マスターのJAN{mj}'
        return ''

    def rms_cost(ctrl, pn, sku, master_cost):
        """RMS商品番号に埋め込まれた原価を、**条件を満たすときだけ**採用する。

        2026-09-12 の条件(Z-4 修正版):
          ① 出品を一意に特定できる(管理番号×SKU→出品テーブル→内部管理ID)。重複なら不可
          ② 原価の意味が形式確認済み(明記型 = 1販売分の原価が直接書いてある)
             積み上げ型(構成ごとの原価)は 数量が未確認なら算定しない
          ③ 対象時点 … 確認根拠に `RMS原価対象月=…` が**明示**され、生成月を含むこと
             (「当月CSVにある＝当月原価として正しい」とはしない)
          ④ 含有範囲 … 確認根拠に `含有範囲=確認済み…` が**明示**されていること。
             「含有範囲」の文字があるだけでは不可。未確認・資料待ちは必ず除外
          ⑤ 商品番号から読める構成と出品テーブルの記録が食い違わないこと
        **構成単品が商品マスターに無いことだけでは除外しない。**
        **数字が抽出できることは確認済みを意味しない。**
        """
        if not use_rms:
            return None, 'RMS原価は未承認(RMS_COST=1 で検証時のみ有効)'
        kind, cost, csrc, n, nsrc, disc, shared, comps, note = CP.parse(pn, {})
        if cost is None:
            if kind.startswith('形式確認済み(積み上げ型'):
                return None, '積み上げ型。構成の数量が未確認のため算定しない'
            return None, f'番号に1販売分の原価が無い({kind})'
        key = (norm(ctrl), norm(sku))
        pid = pair2pid.get(key) or ctrl2pid.get(key[0])
        if not pid:
            return None, 'RMS原価はあるが出品→内部管理IDの対応が付かない'
        info, amb = listing_of(ctrl, sku)
        if amb:
            return None, f'RMS原価はあるが出品を一意に特定できない({amb})'
        proof = (info or {}).get('proof') or ''
        st, v = S.proof_status(proof, '含有範囲')
        if st != 'confirmed':
            return None, ('RMS原価はあるが含有範囲が未確認' +
                          (f'(記載: {v})' if v else '(確認根拠に 含有範囲=確認済み… の記載なし)'))
        if not ym or not S.month_covered(proof, ym):
            rec = S.parse_proof(proof).get('RMS原価対象月')
            return None, (f'RMS原価はあるが対象月 {ym or "?"} への適用が未確認'
                          + (f'(記載: RMS原価対象月={rec} は {ym} を含まない)' if rec
                             else '(確認根拠に RMS原価対象月=… の記載なし)'))
        conflict = composition_conflict(pn, info)
        if conflict:
            return None, f'RMS原価はあるが{conflict}'
        return cost, ''

    # 過去月シート: 新しい月から順に、同じ出品(管理番号×SKU)の L列・商品番号・P列(出所メモ)を持つ
    hist = {}
    for name in reversed(wb_values.sheetnames):
        ws = wb_values[name]
        for r in range(8, ws.max_row + 1):
            c, e, l = ws.cell(r, 3).value, ws.cell(r, 5).value, ws.cell(r, 12).value
            if c is None or not isinstance(l, (int, float)):
                continue
            hist.setdefault((norm(c), norm(e)), []).append(
                {'month': name, 'cost': l, 'pn': str(ws.cell(r, 4).value or '').strip(),
                 'note': str(ws.cell(r, 16).value or '')})

    # 「人が確認した1販売分の原価」の記録(§7-6)。KPIの数値があるだけでは確認済みにしない。
    import cost_confirmations as K
    records = K.load('楽天')

    def same_composition(pn_now, pn_past):
        """販売構成が同じと言えるか。商品番号が同一か、番号から読める構成(JAN/構成JAN下4桁・入数)が一致する。"""
        if pn_now and pn_now == pn_past:
            return True, '商品番号同一'
        a = CP.parse(pn_now, {}) if pn_now else None
        b = CP.parse(pn_past, {}) if pn_past else None
        if a and b and a[7] and a[7] == b[7] and a[3] == b[3]:
            return True, '番号から読める構成・入数が一致'
        return False, f'販売構成が同じと確認できない(商品番号 {pn_past!r} → {pn_now!r})'

    def confirmed_cost(ctrl, pn, sku):
        """人が確認した原価を、**記録の条件を満たすときだけ**採用する(2026-09-12 §7-5/§7-6)。

          a. 同じ出品(管理番号×SKU)で、出品テーブルで一意に特定できる
          b. 記録の適用期間(開始月〜終了月)が生成月を含む。確認日だけで無期限に引き継がない
          c. 記録の販売構成(確認時の商品番号)が現在の商品番号と同じ。違えば無効
          d. 記録に参照月があれば、その月のシートのL列が記録の原価と一致する。違えば無効(原価値が変わった)
          e. 出品テーブルに販売入数があれば記録の入数と一致する。違えば無効
        「L列は1販売分だから単位同一」という解釈は採らない。記録が構成・入数・含有範囲を明示する。
        満たさなければ (None, 理由, 参考候補, 出所) を返す。参考候補はP列に残すだけで、L列は埋めない。
        """
        key = (norm(ctrl), norm(sku))
        cands = hist.get(key) or []
        ref = f'{cands[0]["month"]} L={cands[0]["cost"]:,.0f}' if cands else ''
        pid = pair2pid.get(key) or ctrl2pid.get(key[0])
        if not pid:
            return None, ('過去月に値はあるが出品→内部管理IDの対応が付かない' if cands else ''), ref, ''
        info, amb = listing_of(ctrl, sku)
        if amb:
            return None, f'出品を一意に特定できない({amb})', ref, ''
        recs = K.find(records, '楽天', ctrl, sku, ym) if ym else []
        if not recs:
            return None, ('過去月に値はあるが人が確認した記録が無い(原価確認記録に該当なし)' if cands
                          else '人が確認した記録が無い'), ref, ''
        why_all = []
        for rec in recs:
            rid = rec.get('記録ID')
            ok, why_c = same_composition(pn, str(rec.get('販売構成(確認時の商品番号)') or '').strip())
            if not ok:
                why_all.append(f'記録{rid}: {why_c}'); continue
            rcost = rec.get('1販売分の原価')
            if not isinstance(rcost, (int, float)):
                why_all.append(f'記録{rid}: 原価が数値でない'); continue
            refm = str(rec.get('参照月') or '').strip()
            if refm:
                past = next((h for h in cands if h['month'] == refm), None)
                if past is None:
                    why_all.append(f'記録{rid}: 参照月 {refm} のシートに同じ出品の値が無い'); continue
                if float(past['cost']) != float(rcost):
                    why_all.append(f'記録{rid}: 参照月 {refm} のL列 {past["cost"]:,.0f} と記録の原価 {rcost:,.0f} が違う(原価値が変わった→無効)'); continue
            rec_n = rec.get('販売入数')
            tab_n = (info or {}).get('pack')
            if isinstance(rec_n, (int, float)) and isinstance(tab_n, (int, float)) and int(rec_n) != int(tab_n):
                why_all.append(f'記録{rid}: 入数 {int(rec_n)} が出品テーブルの販売入数 {int(tab_n)} と違う(→無効)'); continue
            scope = str(rec.get('含有範囲') or '')
            if not scope or any(w in scope for w in S.UNCONFIRMED_WORDS):
                why_all.append(f'記録{rid}: 含有範囲が未確認({scope or "空"})'); continue
            src = (f'人が確認({rid} {rec.get("確認者")} {str(rec.get("確認日"))[:10]}'
                   f' 適用{K._ym(rec.get("適用開始月"))}〜{K._ym(rec.get("適用終了月")) or K._ym(rec.get("適用開始月"))})')
            return float(rcost), '', ref, src
        return None, ' / '.join(why_all), ref, ''

    def resolve(ctrl, pn, sku):
        """(原価, 出所, 未確認理由, 参考候補) を返す。**採用できなければ原価は None。**"""
        key = (norm(ctrl), norm(sku))
        pid = pair2pid.get(key) or ctrl2pid.get(key[0])
        mc = cost_by_pid.get(pid) if pid is not None else None
        why = ''
        if isinstance(mc, (int, float)):
            if not has_pack_cols:
                return mc, 'マスター(単位列なし)', '', ''
            v, why = S.resolve_unit_cost(pack, '楽天', norm(ctrl), norm(sku), mc)
            if v is not None:
                return v, 'マスター×入数(単位確認済み)', '', ''
        # マスターから解決できないとき、RMS商品番号の原価を**代替候補**として見る。
        # ただし数字が取れるだけでは採用しない(2026-09-10 承認X-1 / 2026-09-12 Z-4)
        cand, cwhy = rms_cost(ctrl, pn, sku, mc)
        if cand is not None:
            return cand, 'RMS商品番号(明記型・出品特定・含有範囲/対象月 確認済み)', why, ''
        h, hwhy, ref, hsrc = confirmed_cost(ctrl, pn, sku)
        if h is not None:
            return h, hsrc, why, ''
        reasons = [x for x in (why, cwhy, hwhy) if x]
        return None, '', (' / '.join(reasons) or '原価を確認できる資料が無い'), ref

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
        cost, src, why, ref = resolve(r[COL['ctrl']], r[COL['pn']], r[COL['sku']])
        lc = ws.cell(row, 12)
        lc.value = int(cost) if isinstance(cost, float) and cost == int(cost) else cost
        lc.fill = PatternFill(fill_type=None) if cost is not None else YELLOW
        # 採用した原価の出所を残す。**どこから来た値か後から辿れるようにする**
        # ⚠️ A〜O列はCSVデータと数式が使っているので触らない。**P列(16)へ書く**
        note = ws.cell(row, 16)
        if cost is not None:
            mcv = master_cost_of(r[COL['ctrl']], r[COL['sku']])
            # マスター×入数 で解決した値は「単品原価×入数」なので単品の標準原価とは一致しなくて当然。
            # 不一致メモを付けるのは マスター以外の出所 のときだけ
            gap = ('' if src.startswith('マスター') or not isinstance(mcv, (int, float)) or mcv == cost
                   else f' / マスター(単品) {mcv:,.0f} と不一致(理由: {why or "単位未確認"})')
            note.value = f'原価出所: {src}{gap}'
        else:
            # 参考候補(条件不足の過去月値)は**P列に残すだけ**。L列は埋めない(2026-09-12 Y-1)
            note.value = f'原価未確定: {why}' + (f' / 参考候補: {ref}' if ref else '')
            missing.append((row, r[COL['ctrl']], r[COL['sku']], why, r[COL['sales']]))
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
    # 経費ブロックの位置を先に決め、**テンプレートの残骸を消してから**書く。
    #   (参考)行を1行挟んだことで新旧の行がずれ、区切り行に前月の「広告費 76,791」
    #   「送料 248,314」が残っていた(2026-09-12 コピー検証で発見)。
    e0 = t + 3                       # 広告費〜クーポン利用手数料
    tot1 = e0 + 4                    # 合計
    s0 = tot1 + 2                    # 送料・梱包資材
    tot2 = s0 + 2                    # 合計
    g = tot2 + 2                     # 限界利益・限界利益率
    for row in range(t + 1, g + 2):
        for col in (13, 14, 15):
            ws.cell(row, col).value = None
            ws.cell(row, col).fill = PatternFill(fill_type=None)

    # 原価が判明している行だけの粗利。**月全体の利益ではない**ことを明示する
    ws.cell(t + 1, 13).value = '(参考)原価判明分のみの粗利'
    ws.cell(t + 1, 14).value = f'=SUMIF(N7:N{last},"<>未確定")'
    ws.cell(t + 1, 15).value = f'=IF({und}=0,"全件確定",{und}&"行が未確定")'

    for j, lab in enumerate(['広告費', 'ポイント費用', 'クーポン利用額', 'クーポン利用手数料']):
        row = e0 + j
        ws.cell(row, 13).value = lab
        nc = ws.cell(row, 14)
        nc.value = None
        nc.fill = YELLOW
        ws.cell(row, 15).value = f'=SUM(N{row}/I{t})' if j < 3 else None
    ws.cell(tot1, 13).value = '合計'
    ws.cell(tot1, 14).value = f'=SUM(N{e0}:N{e0 + 3})'
    ws.cell(tot1, 15).value = f'=SUM(N{tot1}/I{t})'

    for j, lab in enumerate(['送料', '梱包資材']):
        row = s0 + j
        ws.cell(row, 13).value = lab
        nc = ws.cell(row, 14)
        nc.value = None
        nc.fill = YELLOW
        ws.cell(row, 15).value = f'=SUM(N{row}/I{t})'
    ws.cell(tot2, 13).value = '合計'
    ws.cell(tot2, 14).value = f'=SUM(N{s0}:N{s0 + 1})'
    ws.cell(tot2, 15).value = f'=SUM(O{s0}:O{s0 + 1})'

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
    m = re.match(r'(\d{4})(\d{2})_', fname)
    ym = f'{m.group(1)}-{m.group(2)}' if m else os.environ.get('KPI_YM', '')   # RMS原価の対象月判定に使う
    if len(sys.argv) >= 3:
        sheet_name = sys.argv[2]
    else:
        if not m:
            raise SystemExit('シート名を自動判定できません。第2引数で指定してください(例: 8月)')
        sheet_name = f'{int(m.group(2))}月'

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
    resolve = build_cost_resolver(wb_values, ym)
    cost_by_pid, pair2pid, ctrl2pid = load_master()

    def master_cost_of(ctrl, sku):
        pid = pair2pid.get((norm(ctrl), norm(sku))) or ctrl2pid.get(norm(ctrl))
        return cost_by_pid.get(pid) if pid else None

    total_row, missing = build_sheet(kpi_file, sheet_name, head6, data, resolve,
                                     master_cost_of)
    # 採用できた件数 / 未確定件数 / 対象売上額 を出所ごとに報告する(2026-09-12 Z-4)
    from collections import Counter
    by_src = Counter(); sales_src = Counter()
    for r in data:
        cost, src, why, ref = resolve(r[COL['ctrl']], r[COL['pn']], r[COL['sku']])
        lab = src if cost is not None else '未確定'
        by_src[lab] += 1
        sales_src[lab] += float(conv(r[COL['sales']]) or 0)
    print(f'シート生成完了(原価未確定 {len(missing)}行)')
    print('  出所別:')
    for lab, c in by_src.most_common():
        print(f'    {c:>4}行  売上 ¥{sales_src[lab]:>12,.0f}  {lab}')
    print('  未確定の理由:')
    for why, c in Counter(m_[3] for m_ in missing).most_common():
        print(f'    {c:>4}行  {why}')

    print('Excelで再計算中...')
    recalc_via_excel(kpi_file)

    ok, exp, got, errs, n_total, o_total = verify(kpi_file, sheet_name, data, total_row)
    print(f'検証: 合計{"一致" if ok else "不一致!"} (個数/売上/件数 CSV={exp} シート={got})')
    print(f'数式エラー: {len(errs)}件 {errs[:10] if errs else ""}')
    if isinstance(n_total, (int, float)) and isinstance(o_total, (int, float)):
        print(f'粗利: {n_total:,.0f}円 ({o_total * 100:.1f}%)')
    else:
        print(f'粗利: {n_total}(原価未確定の行があるため月全体の粗利は出ない)')
    if not ok or errs:
        sys.exit('*** FAIL — シートを確認してください ***')
    print('PASS。経費(黄色セル)入力後に限界利益が確定します。')
    print('新商品があれば: python3 sync_cost_master.py register ' + sheet_name)


if __name__ == '__main__':
    main()
