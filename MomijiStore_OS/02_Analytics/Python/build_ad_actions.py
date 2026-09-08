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


def sheet_stop(cur):
    """⑤ 広告停止候補 — ②③をまとめ、止める理由と回収額を1行に"""
    rows = []
    for d in cur:
        g = gain(d)
        if not g:
            continue
        if d['ad_sales'] == 0:
            why, conf = ('広告経由の売上が1円も無い', '高')
        else:
            why = (f'広告が生んだ粗利 ¥{d["ad_gp"]:,.0f} < 広告費 ¥{d["spend"]:,.0f}'
                   f'(粗利率 {d["rate"]:.1%})')
            conf = '低' if d['no_cost'] else '中'
            if d['no_cost']:
                why += ' ⚠️ 仕入値未入力のため粗利率が過大。先に原価を入れる'
        rows.append([d['ch'], d['key'], d['name'][:48], round(d['spend']),
                     round(d['ad_sales']), round(g), why, conf,
                     ('RMS広告センター' if d['ch'] == '楽天' else 'Amazon広告コンソール'),
                     ''])
    rows.sort(key=lambda x: -x[5])
    return [[i] + r for i, r in enumerate(rows, 1)]


def sheet_boost(cur):
    """⑥ 広告増額候補 — **次の1円が最も稼ぐ商品**から並べる

    絶対額(純利益貢献)で並べると、既に大きく出稿していて広告依存度が
    9割に達している商品が上位に来る。そこへ足しても伸びしろは小さい。
    増額の判断に効くのは「広告費1円あたりいくら純利益が出るか」である。

    ⚠️ 仕入値が未入力の商品は粗利率が過大に出るため、確信度を下げて末尾へ回す。
    """
    MIN_SPEND, MIN_ORDERS = 300, 3
    rows = []
    for d in cur:
        if d['net'] is None or d['net'] <= 0 or d['spend'] <= 0:
            continue
        eff = d['net'] / d['spend']            # 広告費1円が生む純利益
        if eff < 0.5:                          # 1円で0.5円未満なら増額の妙味は薄い
            continue
        dep = d['dep']
        # 広告費¥12で1件売れただけの商品は「1円あたり175円」になるが、
        # そこから「増額すれば儲かる」とは言えない。実績量の下限を設ける
        thin = d['spend'] < MIN_SPEND or d['ad_orders'] < MIN_ORDERS
        if d['no_cost']:
            note, conf = '⚠️ 仕入値が未入力。粗利率が過大に出ている。先に原価を入れる', '低'
        elif thin:
            note = (f'⚠️ 実績が少ない(広告費¥{d["spend"]:,.0f} / {d["ad_orders"]:.0f}件)。'
                    '効率の推定が不確か。まず小さく試す')
            conf = '低'
        elif dep is not None and dep < 0.5:
            note, conf = '広告依存度が低く、伸びしろがある', '中'
        elif dep is not None and dep < 0.8:
            note, conf = 'そこそこ広告に頼っている。小さく増やして様子を見る', '中'
        else:
            note, conf = '既に広告依存が高い。増額しても伸びしろは小さい', '低'
        rows.append([d['ch'], d['key'], d['name'][:46], round(d['spend']),
                     round(d['ad_sales']), f'{d["roas"]:.0%}' if d['roas'] else '',
                     f'{d["rate"]:.1%}' if d['rate'] else '', round(d['net']),
                     round(eff, 2), f'{dep:.0%}' if dep is not None else '',
                     conf, note, round(d['spend'] * 0.5 * eff)])
    # 確信度の低いもの(仕入値未入力・実績が薄い)は後ろへ。その中では効率順
    rows.sort(key=lambda x: (x[10] == '低', -x[8]))
    return [[i] + r for i, r in enumerate(rows, 1)]


def sheet_forecast(cur, stop_rows, boost_rows, budget):
    """⑦ 改善した場合の利益増加予測(月間・年間)

    確信度の高いものから積み上げる。**確信度の低いものを合計に混ぜない。**
    """
    zero = [r for r in stop_rows if r[5] == 0]
    over = [r for r in stop_rows if r[5] > 0]
    zero_total = sum(r[6] for r in zero)
    over_total = sum(r[6] for r in over)
    # 仕入値未入力の商品は粗利率が過大なので、②の金額としては信用できない
    over_nc = [r for r in over if r[8] == '低']
    over_solid = over_total - sum(r[6] for r in over_nc)

    # 増額は楽天の予算余力の範囲でしか置けない。
    # Amazonはレポートに予算上限が無いため、金額を置かず候補提示にとどめる
    rk = [r for r in boost_rows if r[1] == '楽天' and r[11] != '低']
    eff_rk = (sum(r[8] for r in rk) / sum(r[4] for r in rk)) if rk else 0
    boost_rk = budget * eff_rk

    unknown = [d for d in cur if d['ad_sales'] > 0 and d['net'] is None]
    am = [r for r in boost_rows if r[1] == 'Amazon' and r[11] != '低']

    rows = [
        ['① 売上ゼロの広告を止める', round(zero_total), round(zero_total * 12), '高',
         f'{len(zero)}商品。広告経由の売上が1円も無い。'
         '止めれば広告費がそのまま利益になる。粗利率に依存しないので最も確実'],
        ['② 粗利を超える広告を止める', round(over_solid), round(over_solid * 12), '中',
         f'{len(over) - len(over_nc)}商品。広告が生む粗利より広告費が大きい。'
         + (f'別に仕入値未入力が{len(over_nc)}商品(¥{sum(r[6] for r in over_nc):,.0f})'
            'あるが、粗利率が過大なため金額には含めない' if over_nc else '')],
        ['③ 楽天の予算余力を効率の良い広告へ回す', round(boost_rk), round(boost_rk * 12),
         '低', f'余力 ¥{budget:,.0f} × 楽天の優良広告の平均効率 {eff_rk:.2f}円/円。'
         '**今の効率が続くと仮定した参考値**。広告は増やすほど効率が落ちるため、実際はこれより小さい'],
        ['合計(①+②) — 止めるだけ', round(zero_total + over_solid),
         round((zero_total + over_solid) * 12), '高〜中',
         '増額の不確実性を含まない。**まずここだけ実行するのが安全**'],
        ['合計(①+②+③)', round(zero_total + over_solid + boost_rk),
         round((zero_total + over_solid + boost_rk) * 12), '中', '③を含めた場合'],
    ]
    notes = [
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
        f'  ・粗利率が出せず判定できない商品 {len(unknown)}件'
        f'(広告費 ¥{sum(d["spend"] for d in unknown):,.0f})。'
        'KPIシートに行が無い = その月に売れていない商品',
        f'  ・仕入値(黄色セル)未入力のため粗利率が過大な商品。'
        f'②で{len(over_nc)}件を除外。**黄色セルを埋めれば精度が上がる**',
        f'  ・Amazonの増額余地({len(am)}商品)。レポートに予算上限が無いため金額を置けない。'
        '⑥の一覧で個別に判断する',
        '',
        '  ・最終判断は人が行う。AIは候補を出すところまで(AI Constitution 第1条)',
    ]
    return rows, notes


def sheet_top20(stop_rows, boost_rows, drop_rows, top_n=20):
    """① 利益改善インパクトTOP20 — 商品単位で、改善額が大きい順に1行"""
    acts = {}

    def add(ch, key, name, act, gainv, why, where, conf):
        k = (ch, key)
        a = acts.setdefault(k, {'ch': ch, 'key': key, 'name': name, 'acts': [],
                                'gain': 0.0, 'whys': [], 'where': set(), 'conf': []})
        a['acts'].append(act)
        a['gain'] += gainv
        a['whys'].append(why)
        a['where'].add(where)
        a['conf'].append(conf)

    for r in stop_rows:
        add(r[1], r[2], r[3], '広告を止める', r[6], r[7], r[9], r[8])
    for r in boost_rows:
        if r[11] == '低':      # 粗利率が怪しい・実績が薄いものは載せない
            continue
        add(r[1], r[2], r[3], '広告を増やす', r[13],
            f'広告費1円が純利益 {r[9]}円 を生む({r[12]})',
            ('RMS広告センター' if r[1] == '楽天' else 'Amazon広告コンソール'), r[11])
    for r in drop_rows:
        k = (r[1], r[2])
        if k in acts:                       # 既に止める/増やす対象なら理由だけ足す
            acts[k]['whys'].append(f'ROASが前月 {r[6]} → {r[7]} へ悪化')
        else:
            add(r[1], r[2], r[3], '原因を調べる', 0,
                f'ROASが前月 {r[6]} → {r[7]} へ悪化(粗利 ¥{r[10]:,} 減)',
                '入札・キーワード・在庫・価格', '—')

    rows = sorted(acts.values(), key=lambda a: -a['gain'])[:top_n]
    return [[i, a['ch'], a['key'], a['name'],
             '\n'.join('・' + x for x in dict.fromkeys(a['acts'])),
             round(a['gain']), round(a['gain'] * 12),
             '\n'.join(a['whys']), '/'.join(sorted(set(a['conf']) - {'—'})) or '—',
             '\n'.join(sorted(a['where']))]
            for i, a in enumerate(rows, 1)]


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
        if fill_col and fill:
            ws.cell(r, fill_col).fill = fill
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

    zero = sheet_zero_sales(cur)
    over = sheet_over_gp(cur)
    drop = sheet_roas_drop(cur, pre, month, prev)
    stop = sheet_stop(cur)
    boost = sheet_boost(cur)
    fc, notes = sheet_forecast(cur, stop, boost, budget)
    top = sheet_top20(stop, boost, drop)

    wb = Workbook()
    wb.remove(wb.active)

    write(wb.create_sheet('01_利益改善TOP20'),
          f'利益改善インパクト TOP20 — {month}',
          '★ 1商品=1行。ご指定どおり改善額が大きい順。'
          '改善額 = 広告経由粗利 − 広告費(ROASではなく利益で判断する)\n'
          '⚠️ 「止める」は確実に戻る額。「増やす」は今の効率が続くと仮定した外挿値で、'
          '実際は逓減する。**確信度の列と合わせて読むこと。'
          '確実に取りたいなら「止める」行から手を付ける**',
          ['順', 'チャネル', '識別子', '商品名', 'すること', '月間改善額', '年間改善額',
           '根拠', '確信度', '実行場所'],
          top, (5, 9, 15, 34, 20, 12, 12, 46, 9, 22), 6, S_FILL)

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

    write(wb.create_sheet('05_広告停止候補'),
          f'広告停止候補 — {month}',
          '②と③をまとめたもの。**AIは候補を出すところまで。止めるかどうかは人が決める**'
          '(AI Constitution 第1条)。⚠️ 改善額は上限値 — 止めても一部は自然検索で拾える',
          ['順', 'チャネル', '識別子', '商品名', '広告費', '広告経由売上',
           '止めた場合の改善額', '止める理由', '確信度', '実行場所', '判断(記入)'],
          stop, (5, 9, 15, 38, 10, 12, 15, 46, 8, 20, 12), 7, S_FILL)

    write(wb.create_sheet('06_広告増額候補'),
          f'広告増額候補 — {month}',
          '**次の1円が最も稼ぐ商品**から並べる(絶対額ではなく効率順)。'
          f'⚠️ 楽天の予算余力は ¥{budget:,.0f}。'
          '**広告は増やすほど効率が落ちる**ため、予測額は上限の参考値として見る',
          ['順', 'チャネル', '識別子', '商品名', '広告費', '広告経由売上', 'ROAS',
           '粗利率', '純利益貢献', '1円あたり純利益', '広告依存度', '確信度', '所見',
           '50%増額時の追加利益額'],
          boost, (5, 9, 15, 34, 10, 12, 8, 8, 11, 13, 10, 8, 40, 16), 10, G_FILL)

    ws7 = wb.create_sheet('07_利益増加予測')
    write(ws7, f'改善した場合の利益増加予測 — {month}',
          '上の一覧をすべて実行した場合。**確信度の高いものから積み上げている**',
          ['施策', '月間の利益増加額', '年間の利益増加額', '確信度', '内訳・前提'],
          fc, (30, 16, 16, 10, 76))
    r = 6 + len(fc) + 1
    for line in notes:
        ws7.cell(r, 1).value = line
        ws7.cell(r, 1).font = BOLD if line.startswith('【') else Font()
        r += 1

    os.makedirs(OUT_DIR, exist_ok=True)
    out = f'{OUT_DIR}/広告改善一覧_{month}_{date.today():%Y%m%d}.xlsx'
    wb.save(out)

    print(f'\n═══ 広告改善一覧 ═══')
    for name, rows_ in [('01_利益改善TOP20', top), ('02_売上0円広告費', zero),
                        ('03_広告費が粗利超過', over), ('04_ROAS急落', drop),
                        ('05_広告停止候補', stop), ('06_広告増額候補', boost)]:
        print(f'  {name:22} {len(rows_):4}件')
    print(f'\n  止めるだけの改善額  月間 ¥{fc[3][1]:,} / 年間 ¥{fc[3][2]:,}')
    print(f'  増額を含めた場合    月間 ¥{fc[4][1]:,} / 年間 ¥{fc[4][2]:,}')
    print(f'\n  → {out}')


if __name__ == '__main__':
    main()
