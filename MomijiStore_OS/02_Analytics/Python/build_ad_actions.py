# -*- coding: utf-8 -*-
"""build_ad_actions.py — 広告分析から「利益を増やすための行動一覧」を生成

使い方:
    python3 build_ad_actions.py <月> [前月]

例:
    python3 build_ad_actions.py 8月 7月

処理内容:
    楽天RPP広告分析.xlsx / AmazonSP広告分析.xlsx の当月・前月シートを読み、
    SourceData/Output/ へ広告改善一覧を出力する(読み取り専用。何も書き換えない)。

────────────────────────────────────────────────────────
このファイルの中心にある一つの数字
────────────────────────────────────────────────────────
    粗利率      = 商品月間粗利 ÷ 商品月間売上
    広告経由粗利 = 広告経由売上 × 粗利率
    **純利益貢献 = 広告経由粗利 − 広告費**

純利益貢献がマイナスの広告は、出せば出すほど利益を減らしている。
止めればその絶対値だけ利益が増える。これを「改善額」と呼ぶ。
売上ゼロの広告は広告経由粗利が0なので、改善額はそのまま広告費になる。

ROASだけでは判断できない。ROAS 400%でも粗利率20%なら
広告経由粗利は売上の20% = 広告費の80%しかなく、赤字である。
**利益で見るからこの一覧は「止める・増やす」を言い切れる。**

⚠️ 改善額は**上限値**である。広告を止めても、その売上の一部は
   自然検索で拾える可能性がある(カニバリゼーション)。
   逆に言えば「これ以上は改善しない」という天井を示している。
"""
import os
import sys
from datetime import date

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

BASE = '/Users/hide0726/Desktop/Claude Code/MomijiStore_OS'
SD = BASE + '/02_Analytics/SourceData'
RPP_FILE = SD + '/楽天RPP広告分析.xlsx'
SP_FILE = SD + '/AmazonSP広告分析.xlsx'
OUT_DIR = BASE + '/01_InventoryManagement/SourceData/Output'

BOLD = Font(bold=True)
HEAD_FILL = PatternFill('solid', fgColor='D9E1F2')
S_FILL = PatternFill('solid', fgColor='FFC7CE')      # 止めるべき
G_FILL = PatternFill('solid', fgColor='C6EFCE')      # 増やせる
TOP = Alignment(vertical='top', wrap_text=True)

# 両ファイルは列の意味は同じだが位置が違う。**位置ではなく意味で扱う**
LAYOUT = {
    '楽天': {'hdr': '商品管理番号', 'key': 1, 'name': 2, 'clicks': 4, 'spend': 6,
             'ad_sales': 8, 'ad_orders': 9, 'roas': 11, 'm_sales': 13,
             'dep': 14, 'm_profit': 15, 'attr': '720時間(30日)'},
    'Amazon': {'hdr': 'SKU', 'key': 2, 'name': 3, 'clicks': 5, 'spend': 8,
               'ad_sales': 9, 'ad_orders': 10, 'roas': 12, 'm_sales': 13,
               'dep': 14, 'm_profit': 15, 'attr': '7日'},
}


KPI_RAKUTEN = SD + '/楽天運営 KPI管理シート.xlsx'
KPI_AMAZON = SD + '/Amazon運営 KPI管理シート.xlsx'


def load_margin_now(month):
    """現在の限界利益(楽天+Amazon)。改善率の分母に使う。"""
    total = 0.0
    for f in (KPI_RAKUTEN, KPI_AMAZON):
        if not os.path.exists(f):
            continue
        wb = load_workbook(f, data_only=True)
        if month not in wb.sheetnames:
            continue
        ws = wb[month]
        for r in range(8, ws.max_row + 1):
            if ws.cell(r, 13).value == '限界利益':
                v = ws.cell(r, 14).value
                if isinstance(v, (int, float)):
                    total += v
                break
    return total


def load_no_cost(month):
    """仕入値(KPIシートL列)が未入力の商品キーを集める。

    仕入値が空だと粗利 = 売上 − 手数料 となり、**粗利率が実態より大幅に高く出る**。
    その状態で「広告を増やせば儲かる」と判断すると損をする。
    黄色セルが埋まるまでは、この商品群を金額の集計から外す。
    """
    bad = set()
    for f, ch, kc in [(KPI_RAKUTEN, '楽天', 3), (KPI_AMAZON, 'Amazon', 3)]:
        if not os.path.exists(f):
            continue
        wb = load_workbook(f, data_only=True)
        if month not in wb.sheetnames:
            continue
        ws = wb[month]
        for r in range(8, ws.max_row + 1):
            k = ws.cell(r, kc).value
            if k is None:
                break
            if not ws.cell(r, 12).value:            # L列 = 仕入値
                bad.add((ch, str(k).strip().lower()))
    return bad


def load_ads(path, sheet, ch, no_cost=frozenset()):
    """広告分析シートを1商品1辞書で読む。無ければ空リストを返す。"""
    if not os.path.exists(path):
        return []
    wb = load_workbook(path, data_only=True)
    if sheet not in wb.sheetnames:
        return []
    ws, L = wb[sheet], LAYOUT[ch]
    hr = next((r for r in range(1, 25) if ws.cell(r, 1).value == L['hdr']), None)
    if hr is None:
        sys.exit(f'❌ {path} の {sheet} に見出し行「{L["hdr"]}」がありません')

    def v(r, k):
        x = ws.cell(r, L[k]).value
        return x if isinstance(x, (int, float)) else None

    out, r = [], hr + 1
    while ws.cell(r, 1).value not in (None, '合計'):     # A列(識別子)で走査する
        spend = v(r, 'spend') or 0
        if spend > 0:                                    # 出稿していない行は対象外
            m_sales, m_profit = v(r, 'm_sales'), v(r, 'm_profit')
            # 粗利率はKPIシートと結合できた商品しか出せない。
            # 結合できない = その月に売れていない商品(KPIシートに行が無い)
            rate = (m_profit / m_sales) if (m_sales and m_profit is not None) else None
            ad_sales = v(r, 'ad_sales') or 0
            ad_gp = (ad_sales * rate) if rate is not None else None
            out.append({
                'ch': ch, 'key': str(ws.cell(r, L['key']).value or '').strip(),
                'name': str(ws.cell(r, L['name']).value or '').strip(),
                'clicks': v(r, 'clicks') or 0, 'spend': spend,
                'ad_sales': ad_sales, 'ad_orders': v(r, 'ad_orders') or 0,
                'roas': v(r, 'roas'), 'm_sales': m_sales, 'm_profit': m_profit,
                'dep': v(r, 'dep'), 'rate': rate, 'ad_gp': ad_gp,
                # 純利益貢献。粗利率が出せない商品は None(判定できない)
                'net': (ad_gp - spend) if ad_gp is not None else None,
                # 仕入値未入力 = 粗利率が過大。金額の判断には使えない
                'no_cost': (ch, str(ws.cell(r, L['key']).value or '').strip().lower())
                           in no_cost,
            })
        r += 1
    return out


def label(d):
    return f'{d["ch"]} / {d["name"][:44] or d["key"]}'


def gain(d):
    """止めた場合の改善額(月間)。プラスの純利益貢献なら止める理由が無いので0。"""
    if d['ad_sales'] == 0:
        return d['spend']                 # 売上ゼロ — 広告費が丸ごと戻る
    if d['net'] is None:
        return None                       # 粗利率不明 — 判定できない
    return max(0.0, -d['net'])


# ═══ 各一覧 ═══════════════════════════════════════════════════════════════
def sheet_zero_sales(cur):
    """② 売上0円広告費ランキング"""
    rows = sorted((d for d in cur if d['ad_sales'] == 0), key=lambda d: -d['spend'])
    return [[i, d['ch'], d['key'], d['name'][:52], round(d['spend']), d['clicks'],
             ('クリックはあるが1件も売れていない' if d['clicks']
              else '表示のみでクリックすら無い'),
             round(d['spend']), '止める(または入札を下げる)']
            for i, d in enumerate(rows, 1)]


def sheet_over_gp(cur):
    """③ 広告費＞粗利ランキング — 広告が生んだ粗利より広告費が大きい"""
    rows = [d for d in cur if d['ad_sales'] > 0 and d['net'] is not None and d['net'] < 0]
    rows.sort(key=lambda d: d['net'])
    return [[i, d['ch'], d['key'], d['name'][:52], round(d['spend']),
             round(d['ad_sales']), f'{d["rate"]:.1%}', round(d['ad_gp']),
             round(d['net']), f'{d["roas"]:.0%}' if d['roas'] else '',
             round(-d['net']), '⚠️ 仕入値未入力' if d['no_cost'] else '',
             '止める / 入札を下げる / 原価を見直す']
            for i, d in enumerate(rows, 1)]


def sheet_roas_drop(cur, prev, month, prev_month):
    """④ ROAS急落ランキング(前月比較)

    ⚠️ 「ROASの下がり幅」だけで並べてはいけない。
       前月に極小額だけ出稿してたまたま売れた商品はROASが数万%になり、
       翌月まともに出稿すると必ず「急落」になる。売上は増えているのに、である。
       (実例: 前月 広告費¥9 で ROAS 25683% → 当月 ¥3,042 で 1584%、売上は20倍)

    そこで次の3つを満たすものだけを「急落」とする:
      1. 両月とも意味のある額(¥300以上)を使っている
      2. 広告経由売上が**実際に減っている**(増えているなら問題ではない)
      3. ROASが下がっている
    並び順は仮定を置かず、**実際に減った広告経由粗利**にする。
    """
    MIN_SPEND = 300
    pv = {(d['ch'], d['key'].lower()): d for d in prev}
    rows = []
    for d in cur:
        p = pv.get((d['ch'], d['key'].lower()))
        if not p or not p['roas'] or not d['roas']:
            continue
        if d['spend'] < MIN_SPEND or p['spend'] < MIN_SPEND:
            continue
        if d['roas'] >= p['roas'] or d['ad_sales'] >= p['ad_sales']:
            continue
        # 仮定を置かず、実際に減った粗利で測る
        lost = ((p['ad_gp'] or 0) - (d['ad_gp'] or 0))
        # 粗利が減った主因を切り分ける。出稿を絞ったなら効率の問題ではない
        cut = d['spend'] / p['spend'] - 1
        cause = ('出稿を絞った結果(広告費 {:.0%})。効率も落ちているが主因は出稿量'
                 .format(cut) if cut <= -0.2 else
                 '出稿量は同程度なのに効率が落ちた' if cut < 0.2 else
                 '出稿を増やしたのに効率が落ちた')
        rows.append([d['ch'], d['key'], d['name'][:48], round(p['spend']),
                     round(d['spend']), f'{p["roas"]:.0%}', f'{d["roas"]:.0%}',
                     round(p['ad_sales']), round(d['ad_sales']),
                     round(lost), cause, '仕入値未入力' if d['no_cost'] else '',
                     '入札・キーワード・在庫切れ・価格改定を確認'])
    rows.sort(key=lambda x: -x[9])
    return [[i] + r for i, r in enumerate(rows, 1)]


MIN_SPEND, MIN_ORDERS = 300, 3      # これ未満は「効率が良い/悪い」と言い切れない


def classify(d):
    """1商品を **停止 / 維持 / テスト / 増額** のどれかに振り分け、理由を返す。

    判断の軸は純利益貢献(= 広告経由粗利 − 広告費)。ROASではない。
    ただし**根拠が薄いものを言い切らない**。実績が少ない商品、
    粗利率が出せない商品、仕入値が未入力の商品は「テスト」へ回す。
    分からないものは分からないと書く。
    """
    thin = d['spend'] < MIN_SPEND or d['ad_orders'] < MIN_ORDERS
    eff = (d['net'] / d['spend']) if (d['net'] is not None and d['spend']) else None

    # ① 売上ゼロ — 粗利率に関係なく判定できる。最も確実
    if d['ad_sales'] == 0:
        return ('停止', d['spend'], '高',
                f'広告費 ¥{d["spend"]:,.0f} を使って広告経由の売上が1円も無い'
                + (f'(クリック{d["clicks"]:.0f}回)' if d['clicks'] else '(クリックも無い)'))

    # ② 粗利率が出せない — KPIシートに行が無い = その月に売れていない
    if d['net'] is None:
        return ('テスト', 0, '—',
                '商品の粗利率が出せず損得を判定できない(KPIシートに行が無い)。'
                '少額で様子を見るか、まず商品マスターを整える')

    # ③ 仕入値未入力 — 粗利率が過大に出るので言い切れない
    if d['no_cost']:
        return ('テスト', 0, '低',
                f'⚠️ 仕入値が未入力で粗利率 {d["rate"]:.0%} が過大に出ている。'
                '先に原価を入れれば正しく判定できる')

    # ④ 赤字 — 広告が生んだ粗利より広告費が大きい
    #    ⚠️ ここで実績量(thin)を理由に確信度を下げてはいけない。
    #    赤字は**今月すでに発生した実測値**であり、外挿ではない。
    #    実績量の下限は「増額しても同じ効率が続くか」を推し量るときの話であって、
    #    起きてしまった損失を疑う理由にはならない
    if d['net'] < 0:
        return ('停止', -d['net'], '中',
                f'広告が生んだ粗利 ¥{d["ad_gp"]:,.0f} < 広告費 ¥{d["spend"]:,.0f}'
                f'(粗利率 {d["rate"]:.1%} / ROAS {d["roas"]:.0%})。売るほど利益が減る'
                + (f'。ただし{d["ad_orders"]:.0f}件のみの実績なので'
                   '来月も同じとは限らない' if thin else ''))

    # ⑤ 黒字だが実績が薄い — 増額の根拠にはならない
    if thin:
        return ('テスト', 0, '低',
                f'黒字だが実績が少ない(広告費 ¥{d["spend"]:,.0f} / {d["ad_orders"]:.0f}件)。'
                '効率の推定が不確かなので、小さく増やして確かめる')

    dep = d['dep']
    # ⑥ 黒字で実績も十分 — あとは伸びしろがあるか
    if eff >= 0.5 and (dep is None or dep < 0.8):
        return ('増額', d['spend'] * 0.5 * eff, '中',
                f'広告費1円が純利益 {eff:.1f}円 を生む'
                + (f'。広告依存度 {dep:.0%} でまだ伸びしろがある' if dep is not None
                   and dep < 0.5 else
                   f'。広告依存度 {dep:.0%} なので小さく増やして様子を見る'
                   if dep is not None else ''))

    return ('維持', 0, '中',
            f'黒字(1円あたり純利益 {eff:.1f}円)だが'
            + (f'広告依存度が {dep:.0%} と高く、増やしても伸びしろは小さい'
               if dep is not None and dep >= 0.8 else '増額の妙味は小さい')
            + '。今のまま続ける')


def sheet_actions(cur):
    """⑤ 広告アクション判定 — **出稿している全商品**を4分類する

    旧「停止候補」「増額候補」を1枚にまとめたもの。
    2枚に分けると同じ商品を2箇所で探すことになり、
    どちらにも載らない商品(維持)が見えなくなる。
    """
    ORDER = {'停止': 0, '増額': 1, 'テスト': 2, '維持': 3}
    rows = []
    for d in cur:
        act, gain_v, conf, why = classify(d)
        eff = (d['net'] / d['spend']) if (d['net'] is not None and d['spend']) else None
        rows.append([act, d['ch'], d['key'], d['name'][:46], round(d['spend']),
                     round(d['ad_sales']),
                     f'{d["roas"]:.0%}' if d['roas'] else '',
                     f'{d["rate"]:.1%}' if d['rate'] is not None else '不明',
                     round(d['net']) if d['net'] is not None else '',
                     round(eff, 2) if eff is not None else '',
                     f'{d["dep"]:.0%}' if d['dep'] is not None else '',
                     round(gain_v), conf, why,
                     ('RMS広告センター' if d['ch'] == '楽天' else 'Amazon広告コンソール'),
                     ''])
    rows.sort(key=lambda x: (ORDER[x[0]], -x[11]))
    return [[i] + r for i, r in enumerate(rows, 1)]


def sheet_forecast(cur, acts, budget, margin_now):
    """⑦ 改善した場合の利益増加予測(月間・年間)+ 改善率

    確信度の高いものから積み上げる。**確信度の低いものを合計に混ぜない。**
    """
    stop = [r for r in acts if r[1] == '停止']
    zero = [r for r in stop if r[6] == 0]                  # 広告経由売上 0
    over = [r for r in stop if r[6] > 0]
    over_solid = [r for r in over if r[13] != '低']
    boost_rk = [r for r in acts if r[1] == '増額' and r[2] == '楽天']
    boost_am = [r for r in acts if r[1] == '増額' and r[2] == 'Amazon']
    test = [r for r in acts if r[1] == 'テスト']

    z = sum(r[12] for r in zero)
    o = sum(r[12] for r in over_solid)
    o_thin = sum(r[12] for r in over) - o
    # 増額は楽天の予算余力の範囲でしか置けない。
    # Amazonはレポートに予算上限が無いため、金額を置かず候補提示にとどめる
    eff_rk = (sum(r[9] for r in boost_rk) / len(boost_rk)) if boost_rk else 0
    b = min(budget * eff_rk, sum(r[12] for r in boost_rk))

    ad_total = sum(d['spend'] for d in cur)

    def row(name, amt, conf, note):
        return [name, round(amt), round(amt * 12),
                (amt / ad_total) if ad_total else '',
                (amt / margin_now) if margin_now else '',
                conf, note]

    rows = [
        row('① 売上ゼロの広告を止める', z, '高',
            f'{len(zero)}商品。広告経由の売上が1円も無い。止めれば広告費がそのまま'
            '利益になる。粗利率に依存しないので最も確実'),
        row('② 粗利を超える広告を止める', o, '中',
            f'{len(over_solid)}商品。広告が生む粗利より広告費が大きい。'
            + (f'別に実績の薄いものが{len(over) - len(over_solid)}商品'
               f'(¥{o_thin:,.0f})あるが、確度が低いので含めない' if o_thin else '')),
        row('③ 楽天の予算余力を効率の良い広告へ回す', b, '低',
            f'余力 ¥{budget:,.0f} × 楽天の増額候補{len(boost_rk)}商品の平均効率'
            f' {eff_rk:.1f}円/円。**今の効率が続くと仮定した参考値**。'
            '広告は増やすほど効率が落ちるため実際はこれより小さい'),
        row('合計(①+②) — 止めるだけ', z + o, '高〜中',
            '増額の不確実性を含まない。**まずここだけ実行するのが安全**'),
        row('合計(①+②+③)', z + o + b, '中', '③を含めた場合'),
    ]
    notes = [
        '',
        '【改善率の分母】',
        f'  ・広告費に対する改善率 … 8月の広告費 ¥{ad_total:,.0f} に対する割合',
        f'  ・限界利益に対する改善率 … 現在の限界利益 ¥{margin_now:,.0f} に対する割合'
        '(楽天+Amazon。経費が未入力のため暫定値)',
        '',
        '【この数字の読み方】',
        '  ・改善額は**上限値**。広告を止めても売上の一部は自然検索で拾える可能性がある',
        '    (カニバリゼーション)。実際の改善はこれより小さくなる',
        '  ・①は広告経由売上が0なのでカニバリの影響を受けない。最も確度が高い',
        '  ・年間は月間×12。**季節商材が多い月は過大になる**'
        '(8月は冷感・虫よけ等が動く)。傾向値として見る',
        '  ・③は「増やせば同じ効率で returns が続く」という仮定。実際は逓減する',
        '',
        '【合計に含めていないもの】',
        f'  ・「テスト」に分類した {len(test)}商品(広告費 ¥{sum(r[5] for r in test):,.0f})。'
        '実績が薄い・粗利率が出せない・仕入値が未入力で、損得を言い切れない',
        f'  ・Amazonの増額余地({len(boost_am)}商品)。レポートに予算上限が無いため'
        '金額を置けない。⑤の一覧で個別に判断する',
        '',
        '  ・最終判断は人が行う。AIは候補を出すところまで(AI Constitution 第1条)',
    ]
    return rows, notes


def sheet_top20(acts, drop_rows, top_n=20):
    """① 利益改善インパクトTOP20 — 商品単位で、改善額が大きい順に1行"""
    drop = {(r[1], r[2]): r for r in drop_rows}
    rows = [r for r in acts if r[12] > 0]
    rows.sort(key=lambda r: -r[12])
    out = []
    for i, r in enumerate(rows[:top_n], 1):
        why = r[14]
        d = drop.get((r[2], r[3]))
        if d:
            why += f'\n加えて ROASが前月 {d[6]} → {d[7]} へ悪化(粗利 ¥{d[10]:,} 減)'
        out.append([i, r[1], r[2], r[3], r[4], r[5], r[12], round(r[12] * 12),
                    r[13], why, r[15]])
    return out


# ═══ 出力 ═════════════════════════════════════════════════════════════════
def write(ws, title, subtitle, headers, rows, widths, fill_col=None, fill=None):
    ws['A1'] = title
    ws['A1'].font = Font(bold=True, size=13)
    ws['A2'] = subtitle
    ws['A2'].alignment = TOP
    if '\n' in subtitle:
        ws.row_dimensions[2].height = 14 * (subtitle.count('\n') + 1)
    ws['A3'] = f'作成 {date.today():%Y-%m-%d} / ⚠️ 実行は人の判断(AIは候補まで)'
    for i, h in enumerate(headers, 1):
        c = ws.cell(5, i)
        c.value, c.font, c.fill, c.alignment = h, BOLD, HEAD_FILL, TOP
    for r, row in enumerate(rows, 6):
        for i, v in enumerate(row, 1):
            c = ws.cell(r, i)
            c.value = v
            c.alignment = TOP
            if isinstance(v, (int, float)) and headers[i - 1].endswith('額'):
                c.number_format = '#,##0'
        if fill_col:
            c = ws.cell(r, fill_col)
            v = str(c.value)
            c.fill = (fill or (S_FILL if v == '停止' else G_FILL if v == '増額'
                               else PatternFill(fill_type=None)))
        n = max((str(v).count('\n') for v in row if isinstance(v, str)), default=0)
        if n:
            ws.row_dimensions[r].height = 14 * (n + 1)
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[ws.cell(5, i).column_letter].width = w
    if not rows:
        ws.cell(6, 1).value = '該当なし'
    ws.freeze_panes = 'A6'


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    month = sys.argv[1]
    prev = sys.argv[2] if len(sys.argv) > 2 else f'{int(month.rstrip("月")) - 1}月'

    nc_cur, nc_pre = load_no_cost(month), load_no_cost(prev)
    cur = (load_ads(RPP_FILE, month, '楽天', nc_cur)
           + load_ads(SP_FILE, month, 'Amazon', nc_cur))
    pre = (load_ads(RPP_FILE, prev, '楽天', nc_pre)
           + load_ads(SP_FILE, prev, 'Amazon', nc_pre))
    if not cur:
        sys.exit(f'❌ {month} の広告分析シートがありません。先に広告分析を実行してください')
    print(f'広告改善一覧を生成します: {month}(前月比較: {prev})')
    print(f'  出稿商品 {len(cur)}件 / 前月 {len(pre)}件 / '
          f'広告費計 ¥{sum(d["spend"] for d in cur):,.0f}')
    n_nc = sum(1 for d in cur if d['no_cost'])
    if n_nc:
        print(f'  ⚠️ 仕入値(黄色セル)未入力 {n_nc}件 — 粗利率が過大に出るため'
              '金額の集計から除外する')

    # 楽天の予算余力。サマリー(有効予算・消化率)はRPPシート上部にある
    budget = 0.0
    if os.path.exists(RPP_FILE):
        wb = load_workbook(RPP_FILE, data_only=True)
        if month in wb.sheetnames:
            ws = wb[month]
            b = {ws.cell(r, 1).value: ws.cell(r, 2).value for r in range(3, 11)}
            if isinstance(b.get('有効予算'), (int, float)):
                spent = sum(d['spend'] for d in cur if d['ch'] == '楽天')
                budget = max(0.0, b['有効予算'] - spent)

    margin_now = load_margin_now(month)
    zero = sheet_zero_sales(cur)
    over = sheet_over_gp(cur)
    drop = sheet_roas_drop(cur, pre, month, prev)
    acts = sheet_actions(cur)
    fc, notes = sheet_forecast(cur, acts, budget, margin_now)
    top = sheet_top20(acts, drop)

    wb = Workbook()
    wb.remove(wb.active)

    write(wb.create_sheet('01_利益改善TOP20'),
          f'利益改善インパクト TOP20 — {month}',
          '★ 1商品=1行。改善額が大きい順。'
          '改善額 = 広告経由粗利 − 広告費(ROASではなく利益で判断する)\n'
          '⚠️ 「止める」は確実に戻る額。「増やす」は今の効率が続くと仮定した外挿値で、'
          '実際は逓減する。**確実に取りたいなら「停止」の行から手を付ける**',
          ['順', 'する事', 'チャネル', '識別子', '商品名', '広告費', '広告経由売上',
           '月間改善額', '年間改善額', '確信度', '理由', '実行場所'],
          top, (5, 8, 9, 15, 32, 10, 12, 12, 12, 9, 52, 22), 8, S_FILL)

    write(wb.create_sheet('02_売上0円広告費'),
          f'売上0円 広告費ランキング — {month}',
          '広告費を使ったのに広告経由の売上が1円も無かった商品。'
          '広告経由粗利が0なので、**止めれば広告費がそのまま利益になる**(最も確度が高い)',
          ['順', 'チャネル', '識別子', '商品名', '広告費', 'クリック数', '状況',
           '止めた場合の改善額', '対応'],
          zero, (5, 9, 15, 40, 10, 10, 30, 15, 24))

    write(wb.create_sheet('03_広告費が粗利超過'),
          f'広告費 > 広告が生んだ粗利 — {month}',
          '売れてはいるが、広告が生んだ粗利より広告費の方が大きい商品。'
          '**売るほど赤字**。ROASが高く見えても粗利率が低ければここに入る',
          ['順', 'チャネル', '識別子', '商品名', '広告費', '広告経由売上', '粗利率',
           '広告経由粗利', '純利益貢献', 'ROAS', '止めた場合の改善額', '注意', '対応'],
          over, (5, 9, 15, 36, 10, 12, 8, 12, 11, 8, 15, 14, 30))

    write(wb.create_sheet('04_ROAS急落'),
          f'ROAS急落ランキング — {month}(前月 {prev} 比較)',
          '前月より効率が落ち、かつ**広告経由売上が実際に減った**商品(両月とも¥300以上出稿)。'
          '仮定を置かず「実際に減った粗利額」で並べる。'
          '⚠️ ROASの下がり幅だけで並べると、前月に極小額でたまたま売れた商品が上位に来てしまう',
          ['順', 'チャネル', '識別子', '商品名', f'{prev} 広告費', f'{month} 広告費',
           f'{prev} ROAS', f'{month} ROAS', f'{prev} 広告売上', f'{month} 広告売上',
           '減った粗利額', '主因', '注意', '確認すること'],
          drop, (5, 9, 15, 34, 11, 11, 10, 10, 13, 13, 12, 34, 12, 30))

    write(wb.create_sheet('05_広告アクション判定'),
          f'広告アクション判定 — {month}(出稿している全 {len(acts)} 商品)',
          '**停止 / 増額 / テスト / 維持** の4分類。旧「停止候補」「増額候補」を1枚にした。'
          '2枚に分けると同じ商品を2箇所で探すことになり、'
          'どちらにも載らない「維持」の商品が見えなくなるため\n'
          '⚠️ AIは候補を出すところまで。決めるのは人(AI Constitution 第1条)',
          ['順', 'する事', 'チャネル', '識別子', '商品名', '広告費', '広告経由売上',
           'ROAS', '粗利率', '純利益貢献', '1円あたり純利益', '広告依存度',
           '月間改善額', '確信度', '理由', '実行場所', '判断(記入)'],
          acts, (5, 8, 9, 15, 34, 10, 12, 8, 8, 11, 13, 10, 12, 8, 54, 20, 12),
          2, None)

    ws6 = wb.create_sheet('06_利益増加予測')
    write(ws6, f'改善した場合の利益増加予測 — {month}',
          '上の一覧をすべて実行した場合。**確信度の高いものから積み上げている**',
          ['施策', '月間の利益増加額', '年間の利益増加額', '広告費に対する改善率',
           '限界利益に対する改善率', '確信度', '内訳・前提'],
          fc, (32, 16, 16, 17, 19, 10, 74))
    for r in range(6, 6 + len(fc)):
        for c in (4, 5):
            ws6.cell(r, c).number_format = '0.0%'
    r = 6 + len(fc) + 1
    for line in notes:
        ws6.cell(r, 1).value = line
        ws6.cell(r, 1).font = BOLD if line.startswith('【') else Font()
        r += 1

    os.makedirs(OUT_DIR, exist_ok=True)
    out = f'{OUT_DIR}/広告改善一覧_{month}_{date.today():%Y%m%d}.xlsx'
    wb.save(out)

    from collections import Counter
    cnt = Counter(r[1] for r in acts)
    print(f'\n═══ 広告改善一覧 ═══')
    for name, rows_ in [('01_利益改善TOP20', top), ('02_売上0円広告費', zero),
                        ('03_広告費が粗利超過', over), ('04_ROAS急落', drop),
                        ('05_広告アクション判定', acts)]:
        print(f'  {name:22} {len(rows_):4}件')
    print('     └ 4分類: ' + ' / '.join(f'{k}{cnt[k]}'
                                       for k in ('停止', '増額', 'テスト', '維持')))
    print(f'\n  止めるだけの改善額  月間 ¥{fc[3][1]:,} / 年間 ¥{fc[3][2]:,}'
          f' (広告費の{fc[3][3]:.1%} / 限界利益の{fc[3][4]:.1%})')
    print(f'  増額を含めた場合    月間 ¥{fc[4][1]:,} / 年間 ¥{fc[4][2]:,}'
          f' (広告費の{fc[4][3]:.1%} / 限界利益の{fc[4][4]:.1%})')
    print(f'\n  → {out}')


if __name__ == '__main__':
    main()
