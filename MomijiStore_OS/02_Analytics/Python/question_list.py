# -*- coding: utf-8 -*-
"""question_list.py — 英樹への「やること・質問」をExcelの記入用リストにする／回答を読む

使い方:
    python3 question_list.py build            Output/英樹への確認リスト_<日付>.xlsx を作る(前回の回答を引き継ぐ)
    python3 question_list.py read <xlsx>      黄色の回答欄を読んで一覧表示する(何も書き換えない)

なぜ作るか(2026-09-12 英樹の要望 / ChatGPT承認で標準形式):
    チャットの文章では「どのファイルをどう直すか」が分からない。
    1行1操作の「やること」と、選択肢つきの「質問」を**黄色のマスだけ埋めればよい**形にする。
    英樹はこのファイルだけ開けばよく、マスター等の正本を直接編集しなくてよい。
    回答を受けて正本へ書くのは Claude Code(バックアップつき・承認後)。

ルール(2026-09-12 ChatGPT):
    ・黄色のセル以外は触らなくてよい
    ・選択肢があるものはプルダウン。分からなければ「不明」。**推測で埋めない**
    ・**事実の回答と変更の承認は別。** 事実を答えただけで正本更新を承認した扱いにしない(列「種別」)
    ・チャットで回答済みの内容は「既知の回答」シートに根拠つきで載せ、同じ質問を繰り返さない
    ・再生成時は 質問ID＋対象 で前回の回答を引き継ぐ。ただし金額・対象月・構成が変わった質問は
      「再確認」とし、古い回答を流用しない(質問キー=質問文と選択肢のハッシュで判定)
    ・「済」の記入ではなく、保存されたファイルの検証で完了判定する(やること)
"""
import glob
import hashlib
import os
import sys
import warnings
from datetime import date

warnings.filterwarnings('ignore')
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

import build_todo_lists as T
import sync_cost_master as S

OUT_DIR = os.path.dirname(T.RAKUTEN_STOCK) + '/Output'
YELLOW = PatternFill('solid', fgColor='FFFF00')
HEAD = PatternFill('solid', fgColor='D9E1F2')
GRAY = PatternFill('solid', fgColor='EEEEEE')
BOLD = Font(bold=True)
WRAP = Alignment(vertical='top', wrap_text=True)

V2 = 'MomijiStore_OS/01_InventoryManagement/SourceData/V2_test'
TARGET_YM = '2026-09'                      # Q4 の対象月。変われば質問キーが変わり再確認になる

FACT, APPROVE, WORK = '事実確認', '変更承認', '作業'


def qkey(*parts):
    return hashlib.md5('|'.join(str(p) for p in parts).encode('utf-8')).hexdigest()[:8]


# ──────────────────────────────────────────────────────────────
# やること(1行1操作)。最終操作は V2-9「保存して閉じる」。完了判定は verify_v2_import.py
# ──────────────────────────────────────────────────────────────
TASKS = [
    ('V2-1', 'V-2テスト ① baseline', f'{V2}/baseline/楽天在庫金額集計ツール_v1.0.xlsm',
     '—', '「集計実行」を押して保存', '✅ 2026-09-12 20:28 実施済み(ファイルで確認)', '済'),
    ('V2-2', 'V-2テスト ② 修正版VBAを入れる', f'{V2}/楽天在庫金額集計ツール_v1.0.xlsm(baselineフォルダの**外**)',
     'Excelで開く', 'マクロを有効にして開く', 'シート「楽天CSV取込」が見える', ''),
    ('V2-3', '', '同上', 'メニュー', 'ツール → マクロ → Visual Basic Editor', 'VBEの画面が開く', ''),
    ('V2-4', '', '同上', 'VBE 左の一覧', '「Module_Rakuten_Tool」を右クリック →「Module_Rakuten_Toolの解放」→ エクスポートしますか？は「いいえ」',
     '一覧から Module_Rakuten_Tool が消える', ''),
    ('V2-5', '', '同上', 'VBE メニュー', 'ファイル → ファイルのインポート → ' + f'{V2}/Module_Rakuten_Tool_V2.bas',
     '一覧に「Module_Rakuten_Tool」が**1つだけ**戻る(Module_Rakuten_Tool1 になっていたらV2-4をやり直す)', ''),
    ('V2-6', '', '同上', 'VBE', 'VBEのウィンドウを閉じる(Excelは閉じない)', 'Excelのシートに戻る', ''),
    ('V2-7', 'V-2テスト ③ 取込と集計', '同上', 'シート「楽天CSV取込」', '「楽天CSV読込」ボタン → 上書き確認は「はい」→ 完了は「OK」',
     'K列「読込日時」が今日の日時になる', ''),
    ('V2-8', '', '同上', 'ボタン', '「集計実行」→ 完了は「OK」', 'シート「在庫金額集計」の集計日時が今日になる', ''),
    ('V2-9', '', '同上', '保存', '⌘S で保存 → ウィンドウを閉じる(**ここが最終操作**)', 'Finderでファイルの更新日時が今になる', ''),
]

# ──────────────────────────────────────────────────────────────
# 質問の材料
# ──────────────────────────────────────────────────────────────
UP24 = [  # 元xlsx時点で数値だった24行を 管理番号×数字 でまとめる(同じ数字でも管理番号が違えば別問)
    ('5946',  '3264680015946', ['3264680015946'], 'Nuxe プロディジュー フローラル オイル 100mL'),
    ('272',   '6941057452425', ['6941057452425'], 'INTEX 子供用 浮き輪 オーシャンリーフ'),
    ('9785',  'b000o7lk1c', ['b000o7lk1c'], 'Nuxe プロディジュー ゴールドオイル 50mL'),
    ('9761',  'b06y13f2sg', ['b06y13f2sg'], 'Nuxe プロディジューオイル 50mL'),
    ('1965',  'b085p2f2v8', ['b085p2f2v8'], 'ソフィーナiP ベースケア セラム'),
    ('4382',  'b0866fcr7t', ['b0866fcr7t'], 'Nuxe プロディジュー フローラル オイル 50mL'),
    ('7480',  'b09b2rlplv', ['b09b2pf8s4', 'b09b2rlplv', 'b09b2t3qbn'], 'Echo Show 5 第3世代'),
    ('15980', 'b0bl6fdw1f', ['b0bl5qt2d1', 'b0bl66dwc3', 'b0bl6fdw1f'], 'Amazon Fire HD 10 キッズモデル'),
    ('5538',  'b0bthtbcrn', ['b0bthtbcrn'], '資生堂 HAKU メラノフォーカスEV 45g'),
    ('7480',  'b0c2s4k41g', ['b0c2rzwv9b', 'b0c2s4d67b', 'b0c2s4k41g'], 'Echo Spot(2024年発売)'),
    ('577',   'b0chs4xy6d', ['b0chs4xy6d'], 'BOSCH スバル車用エアコンフィルター'),
    ('5980',  'b0cjlfprmh', ['b0cjlfprmh'], 'Amazon Fire TV Stick 4K'),
    ('2176',  'b0ckhk5b77', ['b0ckhk5b77'], 'さらさ 柔軟剤 詰め替え 1350mL×2'),
    ('32780', 'b0ds2b3p2b', ['b0ds2b3p2b'], 'DJI Osmo Action 4'),
    ('1980',  'b0f3wyss1s', ['b0f3wyss1s'], 'ラックス バスグロウ セット'),
    ('2000',  'b0gsqcp1mx-a', ['b0gsq2q1ry', 'b0gsqc8jp3', 'b0gsqcp1mx'], 'クーリア 福袋 3点セット'),
]
LISTINGS = [  # 出品テーブルに販売入数が入っている12出品(内部管理ID, チャネル, 識別子, 商品名)
    ('P000049', 'Amazon', 'B0DJCX5CCW', 'ニチドウ 毛玉クリーン 60g'),
    ('P000049', 'Amazon', 'B0DJCXKM4G', 'ニチドウ 毛玉クリーン 60g'),
    ('P000056', 'Amazon', 'B0DKSGDMCZ', 'ドクターズチョイス ゴートミルク 120g'),
    ('P000056', 'Amazon', 'B0DKSVB9B5', 'ドクターズチョイス ゴートミルク 120g'),
    ('P000225', '楽天', 'b0c656jxkd', '肌ラボ 白潤プレミアム 乳液 140ml'),
    ('P000225', 'Amazon', 'B08WS9L34F', '肌ラボ 白潤プレミアム 乳液 140ml'),
    ('P000327', 'Amazon', 'B0DGQJXT78', 'ニチドウ Dr.PRO. ベビーミルク'),
    ('P000334', 'Amazon', 'B0DKSXGXNZ', 'ドクターズチョイス 納豆菌 ふりかけ 80g'),
    ('P000334', 'Amazon', 'B0DKSPG7VK', 'ドクターズチョイス 納豆菌 ふりかけ 80g'),
    ('P000334', 'Amazon', 'B0DKS7KSRB', 'ドクターズチョイス 納豆菌 ふりかけ 80g'),
    ('P000655', '楽天', 'b0fcfk55zk-4', '【4個セット】アリエール 液体洗剤'),
    ('P000846', 'Amazon', 'B0H6L99WP3', 'トイレその後に 280ml×2本'),
]

# チャットで既に回答済みの事実。**同じ質問を繰り返さない**ためにリストへ載せる(根拠=回答日)
KNOWN = [
    ('2026-09-10', '4987176309099-4-3516-a = JAN／入数4／1販売分の原価3,516。SET-8668-4304-a = 構成JAN下4桁 各1個。SET-12-242 = 入数12×単価242。h=廃盤。末尾a=Amazon共有在庫'),
    ('2026-09-10', '7378/1642-a = JAN下4桁／1販売分の原価。4901301451217-6-a = 入数6・原価なし'),
    ('2026-09-10', 'リポソームショット B0FZ3Y7QNM の単価は175円(事実)。マスター登録の承認は別 → Q7'),
    ('2026-09-10', '肌ラボ P000225: 単品在庫から3本取り出して発送。単品原価761円(事実)。楽天=入数3／Amazon=入数1'),
    ('2026-09-10', 'b07bk6696f(HDD P000059)は仕入値が変わった(事実)。標準原価17,980円への更新承認は別 → Q6'),
    ('2026-09-11', 'A-1〜A-9: N,N/N=構成JAN下4桁/1販売分原価、N,N/N,N-a=構成ごとの原価、N,N,N-a=全部JAN下4桁、AST=アソート、BR=バリエーション、13桁JAN-a=入数1、N単独=基本JAN下4桁だが原価もありうる、SET-0014/6668.6637=/以降構成JAN'),
    ('2026-09-11', 'b087lzdd59 は 5036/1273 が正。b0dfyjnwjt は 4987176260635-2 が正。b0fcfk55zk-4 の -a は楽天管理でAmazon共有を示すため。b0g2tqzmy1 は先頭0あり(0840414684973)が正'),
    ('2026-09-12', '14993499s1-3 = 型番＋数量3。4901133863318-2-w の -w = 訳あり品。M-8991/730 = カタログ無しで不明(再確認しない)'),
    ('2026-09-12', 'SET-4916-5050-a 等7件 = SET／構成JAN下4桁／a=Amazon共有(7件を個別登録。数量は別確認)'),
    ('2026-09-12', 'M-0089/1599-a = Mは意味なし／JAN下4桁 0089／1販売分の原価1,599／Amazon共有(この番号のみ)'),
]


def load_cost_facts():
    """Q4 の表示用: 出品ごとの 8月KPIのL列(1販売分) と マスター単品×入数 を集める(読み取り専用)。"""
    cost_by_pid, pair2pid, ctrl2pid = S.load_master()
    pack, pos = S.load_pack_info()
    rk = {(r['ctrl'].lower(), r['sku'].lower()): r for r in S.load_kpi_month('8月')}
    ak = {r['key'].upper(): r for r in S.load_amazon_kpi_month('8月')}
    out = {}
    for pid, ch, key, name in LISTINGS:
        mc = cost_by_pid.get(pid)
        info, amb = S.lookup_pack(pack, ch, S.norm(key), S.norm(key) if ch == '楽天' else '')
        n, unit = (info or {}).get('pack'), (info or {}).get('unit')
        k = rk.get((key.lower(), key.lower())) if ch == '楽天' else ak.get(key.upper())
        out[(pid, ch, key)] = {'master': mc, 'pack': n, 'unit': unit,
                               'kpi': (k or {}).get('cost'), 'pn': (k or {}).get('pn') or ''}
    return out


def build_questions():
    Q = []   # (番号, 種別, 質問, 対象, 選択肢, 出典)
    Q.append(('Q1', FACT, '8/3に楽天RMSからダウンロードした在庫CSV(Excelに変換する前の元ファイル)はPCに残っていますか？ '
              '残っていれば MomijiStore_OS/01_InventoryManagement/SourceData/Import/ にコピーしてください',
              '上流24行の復元', '残っている(Importへ置いた)／残っていない／探している', '上流24件'))
    Q.append(('Q2', WORK, 'RMSの商品番号が誤っている2件を、RMSの商品ページで直しますか？ 直さないと次回の取込で誤りが戻ります(2026-09-11 に正しい値は回答済み)',
              'b087lzdd59 → 5036/1273 ／ b0dfyjnwjt → 4987176260635-2', '直した／後で直す／直さない', '上流誤記2件'))
    for i, (num, ctrl, skus, name) in enumerate(UP24, 1):
        Q.append((f'Q3-{i}', FACT,
                  f'商品番号「{num}」は何の数字ですか？ 元のファイルで数字になっていたため、先頭の0やカンマが消えた可能性があります'
                  f'(元が「0{num}」のような形なら「補足」に書いてください)。対象の行はシート「Q3対応表」',
                  f'管理番号 {ctrl} / SKU {len(skus)}件 / {name}', 'JAN下4桁／原価(円)／不明', '上流24件(A-7)'))
    facts = load_cost_facts()
    for i, (pid, ch, key, name) in enumerate(LISTINGS, 1):
        f = facts[(pid, ch, key)]
        n = f['pack']
        a = f['kpi']
        b = (f['master'] * n) if isinstance(f['master'], (int, float)) and isinstance(n, (int, float)) else None
        src_a = f'{ch}運営KPI管理シート 8月 L列' + (f'(商品番号 {f["pn"]})' if f['pn'] else '')
        text = (f'【{TARGET_YM} のKPI】この出品の「1販売分の原価」として {TARGET_YM} に適用してよい金額はどれですか？'
                f'　A: ¥{a:,.0f}(参照元: {src_a})' if a is not None else
                f'【{TARGET_YM} のKPI】この出品の「1販売分の原価」として {TARGET_YM} に適用してよい金額はどれですか？　A: 8月KPIに値なし')
        text += (f'　B: 商品マスター単品原価 ¥{f["master"]:,.0f} × 販売入数{n} = ¥{b:,.0f}(参照元: 商品マスター E列 × 出品テーブル L列)'
                 if b is not None else '　B: マスター単品×入数は算出不能(単品原価か入数が無い)')
        if a is not None and b is not None and a == b:
            text += '　※AとBは同額'
        Q.append((f'Q4-{i}', FACT, text,
                  f'{pid} {ch} {key} 入数{n} 原価単位={f["unit"] or "未記録"} / {name}',
                  'A(8月KPIの値)／B(マスター単品×入数)／AもBも違う(補足に金額)／不明', f'Y-1 原価継続({TARGET_YM})'))
    for i, (pid, ch, key, name) in enumerate(LISTINGS, 1):
        Q.append((f'Q5a-{i}', FACT, 'この出品に付属品・おまけ(おしぼり、サンプル等)は付いていますか？',
                  f'{pid} {ch} {key} / {name}', 'なし／あり／不明', 'Z-4 含有範囲'))
        Q.append((f'Q5b-{i}', FACT, '(Q5aが「あり」のとき)その付属品の費用は、上の原価に含まれていますか？ それとも別の費目で計上していますか？ Q5aが「なし」なら空欄でよい',
                  f'{pid} {ch} {key} / {name}', '原価に含まれる／別費目で計上／不明', 'Z-4 含有範囲'))
    Q.append(('Q6', APPROVE, '【承認】HDD(P000059)の標準原価を 17,980円 に更新してよいですか？ 8月KPIは変えません。今後の標準原価としての更新です(仕入値が変わった事実は2026-09-10 回答済み)',
              'P000059 商品マスター E列', '更新してよい／まだ／不明', '保留T-1'))
    Q.append(('Q7', APPROVE, '【承認】リポソームショット(B0FZ3Y7QNM / JAN 04571509302842)の標準原価 175円 を商品マスターへ登録してよいですか？(単価175円の事実は2026-09-10 回答済み)',
              'B0FZ3Y7QNM 商品マスター', '登録してよい／まだ／不明', '保留'))
    return Q


def previous_answers(exclude):
    """前回の確認リストから 質問ID＋対象＋質問キー → (回答, 補足) を読む。やることは 番号＋操作 → 済。"""
    files = sorted(f for f in glob.glob(f'{OUT_DIR}/英樹への確認リスト_*.xlsx') if os.path.abspath(f) != os.path.abspath(exclude))
    if not files:
        return {}, {}, None
    wb = load_workbook(files[-1], data_only=True)
    qa, ta = {}, {}
    if '質問' in wb.sheetnames:
        ws = wb['質問']
        hdr = [str(c.value or '') for c in ws[1]]
        col = {h: i + 1 for i, h in enumerate(hdr)}
        for r in range(2, ws.max_row + 1):
            no = ws.cell(r, col.get('番号', 1)).value
            if not no:
                continue
            key = (str(no), str(ws.cell(r, col.get('対象', 3)).value or ''),
                   str(ws.cell(r, col['質問キー']).value or '') if '質問キー' in col else '')
            qa[key] = (ws.cell(r, col.get('回答', 5)).value, ws.cell(r, col.get('補足(自由に)', 6)).value)
    if 'やること' in wb.sheetnames:
        ws = wb['やること']
        for r in range(2, ws.max_row + 1):
            no = ws.cell(r, 1).value
            if no:
                ta[(str(no), str(ws.cell(r, 5).value or ''))] = ws.cell(r, 7).value
    return qa, ta, files[-1]


def build():
    path = f'{OUT_DIR}/英樹への確認リスト_{date.today():%Y%m%d}.xlsx'
    prev_q, prev_t, prev_file = previous_answers(path)
    if os.path.exists(path):                       # 同日の再生成: 今日のファイルの回答も引き継ぐ
        q2, t2, _ = previous_answers('__none__')
        prev_q = {**prev_q, **{k: v for k, v in q2.items() if v[0]}}
        prev_t = {**prev_t, **{k: v for k, v in t2.items() if v}}

    wb = Workbook()
    ws = wb.active
    ws.title = '使い方'
    lines = ['英樹への確認リスト', '',
             '1. 黄色のマスだけ埋めてください。それ以外のセルは触らなくて大丈夫です。',
             '2. 選択肢があるマスはプルダウンです。分からなければ「不明」を選んでください(推測で埋めない)。',
             '3. 「やること」は1行1操作です。最終操作は V2-9「保存して閉じる」。完了かどうかは Claude Code が保存されたファイルを検証して判定します(「済」は目安)。',
             '4. 「質問」の列「種別」: 事実確認=事実を答えるだけ(正本の更新を承認したことにはなりません) / 変更承認=正本を更新してよいかの判断 / 作業=英樹の側で行う作業。',
             '5. 「既知の回答」シートは、チャットで既に答えてもらった内容です。同じことは聞きません。違っていれば「質問」シートの補足欄に書いてください。',
             '6. 終わったら保存して、チャットで「回答した」と一言ください。',
             '7. 正本(商品マスター・KPIシート・在庫ツール)への書き込みは、回答を受けて変更案を出し、承認後に Claude Code がバックアップつきで行います。',
             '', f'作成 {date.today():%Y-%m-%d} / Claude Code' + (f' / 前回の回答を引き継ぎ: {os.path.basename(prev_file)}' if prev_file else '')]
    for i, l in enumerate(lines, 1):
        ws.cell(i, 1).value = l
    ws.cell(1, 1).font = Font(bold=True, size=14)
    ws.column_dimensions['A'].width = 110

    # ── やること ──
    wt = wb.create_sheet('やること')
    head = ['番号', '作業', '開くファイル', 'どこで', '操作(1行1操作)', '確認できること', '済', 'メモ']
    for c, h in enumerate(head, 1):
        wt.cell(1, c).value = h; wt.cell(1, c).font = BOLD; wt.cell(1, c).fill = HEAD
    dv = DataValidation(type='list', formula1='"済,未,できない"', allow_blank=True)
    wt.add_data_validation(dv)
    for r, t in enumerate(TASKS, 2):
        for c, v in enumerate(t[:6], 1):
            wt.cell(r, c).value = v; wt.cell(r, c).alignment = WRAP
        done = wt.cell(r, 7)
        done.value = t[6] or prev_t.get((t[0], t[4])) or None
        done.fill = YELLOW
        wt.cell(r, 8).fill = YELLOW
        dv.add(done)
    for col, w in zip('ABCDEFGH', (7, 22, 44, 16, 56, 44, 8, 30)):
        wt.column_dimensions[col].width = w
    wt.freeze_panes = 'A2'

    # ── 質問 ──
    wq = wb.create_sheet('質問')
    head = ['番号', '種別', '質問', '対象', '選択肢', '回答', '補足(自由に)', '出典', '状態', '質問キー', '前回の回答']
    for c, h in enumerate(head, 1):
        wq.cell(1, c).value = h; wq.cell(1, c).font = BOLD; wq.cell(1, c).fill = HEAD
    Q = build_questions()
    carried = reask = 0
    for r, (no, kind, text, target, opts, src) in enumerate(Q, 2):
        k = qkey(text, opts)
        for c, v in enumerate((no, kind, text, target, opts), 1):
            wq.cell(r, c).value = v; wq.cell(r, c).alignment = WRAP
        ans, memo = wq.cell(r, 6), wq.cell(r, 7)
        ans.fill = YELLOW; memo.fill = YELLOW
        wq.cell(r, 8).value = src
        wq.cell(r, 10).value = k
        wq.cell(r, 10).font = Font(color='999999')
        # 前回の回答: 質問ID＋対象＋質問キーが一致すれば引き継ぐ。IDと対象が同じでキーが違えば「再確認」
        same = prev_q.get((no, target, k))
        changed = [v for (pn, pt, pk), v in prev_q.items() if pn == no and pt == target and pk != k and v[0]]
        if same and same[0]:
            ans.value, memo.value = same[0], same[1]
            carried += 1
        elif changed:
            wq.cell(r, 11).value = f'{changed[0][0]}(条件が変わったため再確認)'
            wq.cell(r, 11).fill = GRAY
            reask += 1
        wq.cell(r, 9).value = f'=IF(F{r}<>"","回答済",IF(K{r}<>"","再確認","未回答"))'
        if opts:
            d = DataValidation(type='list', formula1='"' + opts.replace('／', ',') + '"', allow_blank=True)
            wq.add_data_validation(d); d.add(ans)
    for col, w in zip('ABCDEFGHIJK', (7, 9, 78, 44, 36, 24, 30, 16, 8, 10, 24)):
        wq.column_dimensions[col].width = w
    wq.freeze_panes = 'A2'

    # ── Q3対応表(24行を1行ずつ) ──
    wm = wb.create_sheet('Q3対応表')
    head = ['質問番号', '取込先行', '楽天商品管理番号', 'SKU管理番号', '現在の商品番号', '商品名']
    for c, h in enumerate(head, 1):
        wm.cell(1, c).value = h; wm.cell(1, c).font = BOLD; wm.cell(1, c).fill = HEAD
    import csv
    logs = sorted(glob.glob(f'{OUT_DIR}/商品番号_復元ログ_*.csv'))
    r = 2
    if logs:
        rows = [x for x in csv.DictReader(open(logs[-1], encoding='utf-8-sig')) if x['差の理由'] == '元xlsx時点で既に数値型']
        for i, (num, ctrl, skus, name) in enumerate(UP24, 1):
            for x in rows:
                if x['楽天商品管理番号'] == ctrl and x['元値'] == num:
                    for c, v in enumerate((f'Q3-{i}', x['取込先行'], x['楽天商品管理番号'], x['SKU管理番号'], x['元値'], x['商品名']), 1):
                        wm.cell(r, c).value = v
                    r += 1
    n_map = r - 2
    for col, w in zip('ABCDEF', (9, 9, 20, 20, 14, 70)):
        wm.column_dimensions[col].width = w

    # ── 既知の回答 ──
    wk = wb.create_sheet('既知の回答')
    for c, h in enumerate(['回答日(チャット)', '内容'], 1):
        wk.cell(1, c).value = h; wk.cell(1, c).font = BOLD; wk.cell(1, c).fill = HEAD
    for r, (d, txt) in enumerate(KNOWN, 2):
        wk.cell(r, 1).value = d; wk.cell(r, 2).value = txt; wk.cell(r, 2).alignment = WRAP
    wk.column_dimensions['A'].width = 16; wk.column_dimensions['B'].width = 120

    os.makedirs(OUT_DIR, exist_ok=True)
    wb.save(path)
    print(f'→ {path}')
    print(f'   やること {len(TASKS)}行 / 質問 {len(Q)}件(前回から引き継ぎ {carried}・再確認 {reask}) / Q3対応表 {n_map}行 / 既知の回答 {len(KNOWN)}件')
    return path


def read(path):
    wb = load_workbook(path, data_only=True)
    print('■ やること')
    wt = wb['やること']
    for r in range(2, wt.max_row + 1):
        no, done, memo = wt.cell(r, 1).value, wt.cell(r, 7).value, wt.cell(r, 8).value
        if no:
            print(f'  {no:6} {done or "—":4} {memo or ""}')
    print('■ 質問')
    wq = wb['質問']
    hdr = [str(c.value or '') for c in wq[1]]
    col = {h: i + 1 for i, h in enumerate(hdr)}
    n_ans = n = 0
    for r in range(2, wq.max_row + 1):
        no = wq.cell(r, col['番号']).value
        if not no:
            continue
        n += 1
        ans, memo, kind = (wq.cell(r, col['回答']).value, wq.cell(r, col['補足(自由に)']).value,
                           wq.cell(r, col['種別']).value)
        if ans:
            n_ans += 1
        print(f'  {no:7} [{kind}] {ans or "—":24} {memo or ""}')
    print(f'  回答 {n_ans} / {n}')


if __name__ == '__main__':
    if len(sys.argv) >= 2 and sys.argv[1] == 'build':
        build()
    elif len(sys.argv) >= 3 and sys.argv[1] == 'read':
        read(sys.argv[2])
    else:
        print(__doc__)
