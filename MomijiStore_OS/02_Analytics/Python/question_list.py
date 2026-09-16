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
IMPORT_DIR = 'MomijiStore_OS/01_InventoryManagement/SourceData/Import'
LIST_NAME = '英樹への確認リスト_<日付>.xlsx(このファイル)'

# いま英樹がやる操作だけを載せる(1行1操作)。終わったものは DONE(シート「完了」)へ移す。
TASKS = [
    # Q10 は ChatGPT が「新しい在庫リストを取る」と判断(2026-09-16)。8/3資料の本番再処理(旧T-4)は行わない
    # Q14 は ChatGPT が「最新CSVに未掲載・現在在庫未確認として別枠管理」と判断(2026-09-16)。コピーはその方式で作成済み
    ('T-3d-2', '最新在庫への更新④-2: 在庫ツールの**コピー**で取込と集計(2ボタン・1回だけ)', 'MomijiStore_OS/01_InventoryManagement/SourceData/V3_stock_test/楽天在庫金額集計ツール_v1.0.xlsm(**コピー**)',
     'ボタン「楽天在庫ファイル読込」「集計実行」', '**着手OK(Claude Code 2026-09-16 13:00)**。コピーの Import には別枠管理版(本体814行・未掲載94行は別シート「未掲載_未確認」)を置いてあります。'
     '開く → 「楽天在庫ファイル読込」→ はい →(アクセス許可は「アクセス権を付与」)→ OK → 「集計実行」→ OK → ⌘S → 閉じる(本番ツールは触らない)',
     'コピーの読込日時・集計日時が今になる → チャットで「コピーで集計した」。Claude Code が実測と予測(数量5,911・評価額12,621,422・未登録59)を照合する', ''),
    ('T-3d-3', '最新在庫への更新④-3: b0dktchdgk の商品番号を確認する', LIST_NAME, 'シート「質問」の Q15',
     'RMSの商品ページ(管理番号 b0dktchdgk・ホワイトピーチ&カモミール)で商品番号が 4987176260659-2 か 4987176260635-2 かを見て Q15 に答える(RMSは直さなくてよい。見るだけ)',
     'Q15 の「状態」が「回答済」になる', ''),
    ('T-3e', '最新在庫への更新⑤: 差分を見て本番へ進めるか決める', LIST_NAME, 'シート「質問」に追加される Q(コピー実測の後)',
     'Claude Code がコピーの実測と予測の差・未確認項目(Q15・未掲載94件・原価未登録59件)を出したら、実施する／しない を選ぶ', '本番の取込日時・集計日時が更新される(実施のとき)', ''),
]
DONE = [
    ('V2', 'V-2 取込4列の文字列書式', 'コピー検証PASS → 本番差し替え 2026-09-12 22:37 検証PASS'),
    ('V3', 'V-3 識別子の型統一＋集計先の文字列書式', 'コピー検証2回目PASS(09-16 06:16) → 本番差し替え検証PASS(着手OK前に差し替えられていた: 引継ぎ102)'),
    ('Q1〜Q9', '質問 63件', '2026-09-16 までに全件回答。原価確認記録 K00001〜12 登録、商品マスター7件更新に使用'),
    ('T-2/Q11', 'トイレその後に 453円の単位', '「2本セット分」→ P000846 453→556(セット原価のまま) 2026-09-16 更新'),
    ('T-3a〜c', '最新在庫の資料取得', '楽天 dl-normal-item_20260916081129-1.csv(取得 08:11・1ファイルで全部・b087lzdd59 は 5036/1273 に訂正済。b0dktchdgk は SKU が b0dfyjnwjt→b0dktchdgk に変わり商品番号は …659-2 のまま → Q15)／Amazon inventoryList 2本(取得 08:15・在庫の基準日時は不明)'),
    ('T-3d-1/Q14', '楽天CSVに無い94件の扱い', 'ChatGPT判断(2026-09-16)「最新CSVに未掲載・現在在庫未確認として別枠管理」。コピーの Import を別枠管理版で作成(本体814行／未掲載_未確認 94行・旧数量469・参考額1,286,025)'),
    ('T-5/Q12・Q13', 'ファイル整理', 'Q12 移動52件を実行。Q13 削除候補31件は ChatGPT の推奨で**削除せず** Archive/削除候補_20260916/ へ退避(対応表あり)'),
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
        n, unit, mc, a = f['pack'], f['unit'], f['master'], f['kpi']
        src_a = f'{ch}運営KPI管理シート 8月 L列' + (f'(商品番号 {f["pn"]})' if f['pn'] else '')
        head = f'【{TARGET_YM} のKPI】この出品の「1販売分の原価」として {TARGET_YM} に適用してよい金額はどれですか？'
        text = head + (f'　A: ¥{a:,.0f}(参照元: {src_a})' if a is not None else '　A: 8月KPIに値なし')
        # B は原価単位で分岐する(ChatGPT 2026-09-12)。未確認なら参考値であって確定値ではない
        b = None; b_kind = ''
        if isinstance(mc, (int, float)):
            if unit == S.UNIT_SINGLE and isinstance(n, (int, float)):
                b = mc * n; b_kind = '確定'
                text += f'　B: 商品マスター単品原価 ¥{mc:,.0f} × 販売入数{n} = ¥{b:,.0f}(原価単位=単品原価・参照元: 商品マスター E列 × 出品テーブル L列)'
            elif unit == S.UNIT_SET:
                b = mc; b_kind = '確定'
                text += f'　B: 商品マスターのセット原価 ¥{mc:,.0f} をそのまま(原価単位=セット原価・参照元: 商品マスター E列)'
            else:
                b_kind = '参考'
                ref = f'単品なら ¥{mc:,.0f}×{n}=¥{mc * n:,.0f} ／ セットなら ¥{mc:,.0f}' if isinstance(n, (int, float)) else f'マスター ¥{mc:,.0f}(入数未記録)'
                text += f'　B: **原価単位が未確認のため参考値**({ref}。参照元: 商品マスター E列。確定値ではありません)'
        else:
            text += '　B: 商品マスターに標準原価が無い'
        if a is not None and b is not None and a == b:
            text += '　※AとBは同額'
        text += '　※この回答は事実確認です。記録(原価確認記録)の作成は別に承認をとります'
        opts = ('A(8月KPIの値)／B(マスター由来)／AもBも違う(補足に金額)／不明' if b_kind == '確定'
                else 'A(8月KPIの値)／AもBも違う(補足に金額)／不明')
        Q.append((f'Q4-{i}', FACT, text,
                  f'{pid} {ch} {key} 入数{n} 原価単位={unit or "未記録"} 構成={f["pn"] or "—"} / {name}',
                  opts, f'原価確認記録({TARGET_YM})'))
    for i, (pid, ch, key, name) in enumerate(LISTINGS, 1):
        Q.append((f'Q5a-{i}', FACT, 'この出品に付属品・おまけ(おしぼり、サンプル等)は付いていますか？',
                  f'{pid} {ch} {key} / {name}', 'なし／あり／不明', 'Z-4 含有範囲'))
        Q.append((f'Q5b-{i}', FACT, '(Q5aが「あり」のとき)その付属品の費用は、上の原価に含まれていますか？ それとも別の費目で計上していますか？ '
                  'Q5aが「なし」なら自動で「該当なし」になります。「別費目で計上」なら**計上先(費目名)を補足に**書いてください',
                  f'{pid} {ch} {key} / {name}', '原価に含まれる／別費目で計上／該当なし／不明', 'Z-4 含有範囲'))
    Q.append(('Q6', APPROVE, '【承認】HDD(P000059)の標準原価を 17,980円 に更新してよいですか？ 8月KPIは変えません。今後の標準原価としての更新です(仕入値が変わった事実は2026-09-10 回答済み)',
              'P000059 商品マスター E列', '更新してよい／まだ／不明', '保留T-1'))
    Q.append(('Q7', APPROVE, '【承認】リポソームショット(B0FZ3Y7QNM / JAN 04571509302842)の標準原価 175円 を商品マスターへ登録してよいですか？(単価175円の事実は2026-09-10 回答済み)',
              'B0FZ3Y7QNM 商品マスター', '登録してよい／まだ／不明', '保留'))
    # ── 商品マスター標準原価の更新(ChatGPT 2026-09-16 条件付き保留解除)。9月KPIの記録とは別の承認 ──
    #   条件: 対象ID・変更前後の金額・原価単位が明確／**今後の標準原価**への採用承認／販売構成・含有範囲が確認済み／
    #         単位を変える場合は利用先(KPI生成・在庫評価)が対応済み
    MASTER_UPD = [  # (ID, 商品, 変更前, 変更後, 単位, 根拠, 単位変更の有無)
        ('P000056', 'ドクターズチョイス ゴートミルク 120g', 1130, 1270, '単品原価(種別=単品)', 'Q4-3/4 回答「単品原価1270円」・Q5a あり(原価に含まれる)', 'なし'),
        ('P000327', 'ニチドウ Dr.PRO. ベビーミルク', 1254, 1227, '単品原価(種別=単品)', 'Q4-7 回答「単品原価1227円」・Q5a なし', 'なし'),
        ('P000334', 'ドクターズチョイス 納豆菌 ふりかけ 80g', 695, 694, '単品原価(種別=単品)', 'Q4-8/9/10 回答「単品原価694円」・Q5a なし', 'なし'),
        ('P000655', '【4個セット】アリエール 液体洗剤', 3940, 3516, 'セット原価(4個・種別=セット)のまま', 'Q4-11 回答「単品879円」×4=3,516・Q5a なし。ChatGPT: 3,940は4個セット分', 'なし(セットのまま金額だけ)'),
        ('P000846', 'トイレその後に 280ml×2本セット', 453, 556, 'セット原価(2本・種別=セット)のまま', 'Q4-12 回答「単品278円」×2=556・Q5a あり(原価に含まれる)。現在値453の単位は未記録', 'なし(セットのまま金額だけ。453が何の単位かは要確認)'),
    ]
    for i, (pid, name, before, after, unit, basis, uchg) in enumerate(MASTER_UPD, 1):
        Q.append((f'Q8-{i}', APPROVE,
                  f'【承認・標準原価】{pid} の商品マスター標準原価を {before:,}円 → {after:,}円({unit})へ更新してよいですか？'
                  f' これは**今後の標準原価**としての更新で、9月の原価確認記録(登録済み)とは別です。単位変更: {uchg}。根拠: {basis}',
                  f'{pid} 商品マスター E列 / {name}', '更新してよい／まだ／不明', '標準原価更新(ChatGPT 2026-09-16 条件)'))
    # Q6/Q7 は「更新してよい」と回答済みだが、含有範囲(付属品)の確認が条件に含まれるため事実確認を追加
    Q.append(('Q9-1', FACT, 'HDD(P000059 バッファロー 外付けHDD 4TB)の原価 17,980円 に付属品・おまけの費用は含まれていますか？(ChatGPT条件: 含有範囲の確認)',
              'P000059 楽天 b07bk6696f', '付属品なし／付属品あり・原価に含まれる／付属品あり・別費目で計上／不明', '標準原価更新(Q6の条件)'))
    Q.append(('Q10', APPROVE,
              '【判断】本番の在庫ツール(楽天在庫金額集計ツール)へ、2026-08-03 15:18 の在庫リスト(Import/楽天在庫リスト_import.xlsx)を再取込しますか？ '
              'これは8月3日資料の再処理で、現在在庫の更新ではありません。見込み: 数量 6,801(不変)／評価額 14,986,097→15,077,081(+90,984)／原価未登録 24→20。'
              '⚠️ この資料では上流誤記2件(b087lzdd59・b0dfyjnwjt)が古い値に戻るため、「再取込する」なら直後に Claude Code が2行訂正を再適用します。'
              '「新しい在庫リストを取る」なら T-3 の手順でCSVを置いてください(現在在庫の更新という別の作業になります)',
              '本番 楽天在庫金額集計ツール_v1.0.xlsm', '再取込する(直後に2行訂正を再適用)／新しい在庫リストを取る(T-3)／しない／不明', '本番再取込(引継ぎ91)'))
    Q.append(('Q11', FACT,
              'トイレその後に フレッシュグリーン 280ml×2本セット(P000846)の商品マスター標準原価 453円 は、何の金額ですか？ '
              '(英樹回答の単品278円×2本=556円と比べるために、453円が2本セット分か1本分かを確認します。Q8-5「更新してよい」は記録済みで、この回答が入れば更新します)',
              'P000846 商品マスター E列', '2本セット分／1本分／不明', '標準原価更新(P000846 保留解除の条件)'))
    Q.append(('Q12', APPROVE,
              '【承認・ファイル整理】Output と Backup が増えたので整理します。案は Output/整理案_20260916.txt。'
              '「移動(退避)52件」は Archive/<年月>/ へ動かすだけ(名前は変えず、消しません)。実行してよいですか？',
              'Output/整理案_20260916.txt（tidy_files.py plan の出力）', '実行してよい／一部だけ(補足に書く)／まだ', 'ファイル整理'))
    Q.append(('Q13', APPROVE,
              '【承認・削除】同じ整理案の「削除候補31件・2.9MB」(確認リストの途中バックアップ12本、8/3作業中の在庫リスト連続バックアップ13本、7月の作業中バックアップ等)を消してよいですか？ '
              '消すのは英樹の承認後に Claude Code が候補どおりに行い、消したファイル名を記録します。',
              'Output/整理案_20260916.txt の「削除候補」', '全部消してよい／一部だけ(補足に書く)／消さない', 'ファイル整理'))
    Q.append(('Q14', APPROVE,
              '【判断・在庫リスト】2026-09-16 08:11 の楽天在庫CSV(1ファイルで全部)に、いまの在庫リストの94件がありません(在庫あり84件・旧数量469個・旧単価での参考額 1,286,025円。クーリアのセット品・福袋など)。'
              '「RMSから消えた」とは確認できていません(今回のCSVに無い、というだけ)。在庫リストでどう扱いますか？'
              '(ChatGPT は「最新CSVに未掲載・現在在庫未確認として別枠管理」= 消さない・0にしない・繰り越さない、と判断。コピーはこの方式で作ってあります)',
              '楽天在庫リスト_import.xlsx の94行(コピーのシート「未掲載_未確認」)', '最新CSVに未掲載・現在在庫未確認として別枠管理／消す(RMSに合わせる)／残す(在庫0にする)／残す(そのまま)／ChatGPTに確認する', '最新在庫への更新 T-3d(別枠管理)'))
    Q.append(('Q15', FACT,
              '【事実確認・商品番号】管理番号 b0dktchdgk(レノア リセット セラム ホワイトピーチ&カモミール つめかえ 特大 750mL×2個セット)の商品番号は、9/16 08:11 のRMS CSVでも 4987176260659-2 のままです。'
              'CSVで変わったのは SKU(b0dfyjnwjt → b0dktchdgk)だけでした。9/11 の回答「b0dfyjnwjt は 4987176260635-2 が正」を根拠に、本番ツールの楽天CSV取込シート(行583)はこの行の商品番号を …635-2 に直してあります。'
              'ただし …635 はホワイトリリー(管理番号 b0dfyjnwjt・別の行)のJANで、この行のJAN欄は …659 です。b0dktchdgk(ホワイトピーチ&カモミール)の正しい商品番号はどれですか？',
              'RMS 管理番号 b0dktchdgk の商品番号／本番ツール 楽天CSV取込 行583(…635-2 に訂正済み)', '4987176260659-2(CSVのまま・ホワイトピーチのJAN)／4987176260635-2(9/11回答どおり)／不明',
              '上流誤記(b0dktchdgk 商品番号 659-2/635-2)'))
    Q.append(('Q9-2', FACT, 'リポソームショット(P000897 / B0FZ3Y7QNM)の原価 175円 に付属品・おまけの費用は含まれていますか？ また商品マスターのJANが空欄です。JANは 4571509302842(13桁)でよいですか？(補足に書いてください)',
              'P000897 Amazon B0FZ3Y7QNM 現在の標準原価=0', '付属品なし／付属品あり・原価に含まれる／付属品あり・別費目で計上／不明', '標準原価更新(Q7の条件)'))
    return Q


# ──────────────────────────────────────────────────────────────
# 原価確認記録の追加案(2026-09-15 ChatGPT指示で修正)
#   ・1販売分 = 英樹回答の単品原価 × 確認済み販売入数。適用期間は 2026-09 のみ(質問範囲を超えて延ばさない)
#   ・マスターとの比較は「比較単位」で分ける: 出品テーブルM列が セット原価 なら1販売分と、単品原価 なら単品と比べる。
#     M列が空欄なら両方を出し「どちらの単位か未確認」と書く(アリエールは ChatGPT 2026-09-15 の情報で4個セット分)
#   ・8月KPI L との差は「比較情報」として履歴に残す(参照月にはしない。8月は変更しない。差の理由は未確認)
#   ・含有範囲(Q5a)が「なし」のときだけ Q5b=該当なし。「あり」なら Q5b を人が答え、別費目なら計上先。不明は記録しない
#   ・「通常」も金額一致だけで承認済みにしない。承認は確認リストの黄色セル
# ──────────────────────────────────────────────────────────────
PROPOSAL_YM = '2026-09'

# ChatGPT が下した判断(記録)。「ChatGPTへの確認」シートと質問の「前回の回答」欄に表示する
CHATGPT_DECISIONS = {
    'Q10': 'ChatGPT判断(2026-09-16): 「新しい在庫リストを取る」。8月3日資料の本番再処理は行わない。T-3a〜e で進める',
    'Q14': 'ChatGPT判断(2026-09-16): 「最新CSVに未掲載・現在在庫未確認として別枠管理」。消す／0にする／繰り越す を自動でしない。'
           '旧数量・旧単価・旧基準日を保持し、最新掲載分(814行)と分けて集計・表示。1,286,025円は旧単価での参考額(損失や確定在庫ではない)',
}

# 「ChatGPTへの確認」シートに添える Claude Code の整理(判断材料)
CHATGPT_NOTES = {
    'Q10': '2択: (a) 8/3資料の再処理 = 数量不変・評価額+90,984(原価マスター更新分)・未登録24→20。ただし上流誤記2件が戻るため直後に2行を再訂正する。'
           '(b) 最新在庫への更新 = T-3a〜e(RMS誤記訂正→最新CSV→Amazonも近い時刻→コピー検証→差分提示→英樹判断)。'
           'Claude Code の見立て: (a)の効用は評価額の再計算だけで小さく、(b)へ進む方が本来の目的(現在在庫)に合う。',
    'Q14': '94件は「今回のCSV(9/16 08:11)に無い」だけで、RMSから消えたかは未確認。旧数量469・旧単価での参考額1,286,025円。'
           '楽天(94件)・Amazon(40件)とも旧数量・旧単価・旧基準日を保持してシート「未掲載_未確認」へ分け、最新掲載分と別に集計する。'
           '同じ管理番号/ASINの新規SKUは候補として書くだけで統合しない。',
}


def _answers(path):
    wb = load_workbook(path, data_only=True)
    ws = wb['質問']
    hdr = [str(c.value or '') for c in ws[1]]
    col = {h: i + 1 for i, h in enumerate(hdr)}
    out = {}
    for r in range(2, ws.max_row + 1):
        no = ws.cell(r, col['番号']).value
        if no:
            out[str(no)] = (ws.cell(r, col['回答']).value, ws.cell(r, col['補足(自由に)']).value)
    return out


def build_cost_proposals(list_path):
    import re
    facts = load_cost_facts()
    ans = _answers(list_path)
    # 出品テーブルM列(原価単位)。マスター金額が何を表すかの記録
    unit_rec = {}
    pack, pos = S.load_pack_info()
    for (ch, key), lst in pack.items():
        for it in lst:
            unit_rec[(ch, key, it['sku'])] = it.get('unit')
    rows = []
    for i, (pid, ch, key, name) in enumerate(LISTINGS, 1):
        f = facts[(pid, ch, key)]
        a, memo = ans.get(f'Q4-{i}', (None, None))
        q5a = (ans.get(f'Q5a-{i}', (None, None))[0] or '')
        q5b, q5b_memo = ans.get(f'Q5b-{i}', (None, None))
        q5b = q5b or ''
        n, kpi, master = f['pack'], f['kpi'], f['master']
        unit_cost = per_sale = None
        basis = ''
        if a == 'A(8月KPIの値)':
            per_sale = kpi; basis = f'英樹確認(確認リストQ4-{i} 回答A: 8月KPI L={kpi})'
        elif a and a.startswith('B'):
            unit_cost = master; per_sale = master * n if isinstance(n, (int, float)) else None
            basis = f'英樹確認(確認リストQ4-{i} 回答B: マスター単品 {master}×入数{n})'
        elif a and a.startswith('AもBも違う'):
            m = re.search(r'(\d[\d,]*)\s*円', str(memo or ''))
            unit_cost = int(m.group(1).replace(',', '')) if m else None
            per_sale = unit_cost * n if unit_cost is not None and isinstance(n, (int, float)) else None
            basis = f'英樹確認(確認リストQ4-{i} 回答「{memo}」×確認済み入数{n})'
        # マスターとの比較(比較単位で分ける)
        # 出品テーブルM列(原価単位)。楽天は管理番号×SKU、Amazonは同じASINの出品で単位が1つに決まるときだけ採る
        if ch == '楽天':
            mu = unit_rec.get((ch, S.norm(key), S.norm(key)))
        else:
            units = {u for (c_, k_, s_), u in unit_rec.items() if c_ == ch and k_ == S.norm(key) and u}
            mu = units.pop() if len(units) == 1 else None
        cmp_notes, exceptions = [], []
        if isinstance(master, (int, float)):
            if mu == S.UNIT_SET or (pid == 'P000655'):
                # セット原価: 1販売分と比べる(アリエール P000655 は ChatGPT 2026-09-15「3,940円は4個セット分」)
                lab = 'セット原価(M列)' if mu == S.UNIT_SET else 'セット原価(M列は空欄。ChatGPT 2026-09-15 情報)'
                cmp_notes.append(f'マスター {master:,}[{lab}] vs 1販売分 {per_sale:,}' if per_sale is not None else f'マスター {master:,}[{lab}]')
                if per_sale is not None and per_sale != master:
                    exceptions.append(f'マスター(セット) {master:,} ≠ 1販売分 {per_sale:,}(差 {per_sale - master:+,})→マスター更新は別承認')
            elif mu == S.UNIT_SINGLE:
                cmp_notes.append(f'マスター {master:,}[単品原価(M列)] vs 単品 {unit_cost}')
                if unit_cost is not None and unit_cost != master:
                    exceptions.append(f'マスター(単品) {master:,} ≠ 回答単品 {unit_cost:,}→マスター更新は別承認')
            else:
                cmp_notes.append(f'マスター {master:,}[単位未記録] 単品比較: {unit_cost} / 1販売分比較: {per_sale} (どちらの単位か未確認)')
                if unit_cost is not None and unit_cost != master and per_sale != master:
                    exceptions.append(f'マスター {master:,} は 単品 {unit_cost:,} とも 1販売分 {per_sale:,} とも一致しない(単位未記録)→単位確認とマスター更新は別承認')
        # 8月KPI との比較(履歴。参照月にはしない)
        if per_sale is not None and kpi is not None and per_sale != kpi:
            cmp_notes.append(f'8月KPI L={kpi:,} / 差 {per_sale - kpi:+,} / 差の理由は未確認(8月は変更しない)')
            exceptions.append(f'8月KPI L={kpi:,} と不一致(差 {per_sale - kpi:+,})')
        elif kpi is not None:
            cmp_notes.append(f'8月KPI L={kpi:,} と一致')
        # 含有範囲(Q5)
        if q5a == 'なし':
            scope, fee = '付属品なし(英樹確認Q5a)', '該当なし'
        elif q5a == 'あり':
            scope = '付属品あり(英樹確認Q5a)'
            fee = {'原価に含まれる': '原価に含まれる', '別費目で計上': f'別費目で計上(計上先: {q5b_memo or "未記入"})'}.get(q5b, '')
        else:
            scope, fee = '', ''
        blockers = []
        if per_sale is None:
            blockers.append('Q4の金額が読めない')
        if not q5a or q5a == '不明':
            blockers.append('含有範囲(Q5a)が未回答/不明')
        elif q5a == 'あり' and (not fee or (q5b == '別費目で計上' and not q5b_memo)):
            blockers.append('Q5b(費用の扱い/計上先)が未回答')
        rows.append({
            '案No': f'K案-{i:02d}', 'チャネル': ch, '楽天商品管理番号': key if ch == '楽天' else '', 'SKU/ASIN': key,
            '内部管理ID': pid, '商品名': name, '適用開始月': PROPOSAL_YM, '適用終了月': PROPOSAL_YM,
            '単品原価(回答)': unit_cost, '販売入数': n, '1販売分の原価(案)': per_sale, '原価単位': '単品原価',
            '含有範囲': scope, '付属品費用の扱い': fee, '販売構成(現在の商品番号)': f['pn'] or '', '参照月': '',
            '根拠': basis, '比較情報': ' / '.join(cmp_notes), '区分': ('例外(ChatGPT判断済: 作成可)' if exceptions else '通常'),
            '例外理由': ' / '.join(exceptions),
            '登録できない理由': ' / '.join(blockers), '確認者': '英樹', '確認日': date.today().isoformat(),
        })
    return rows


PROPOSAL_HEAD = ['案No', 'チャネル', 'SKU/ASIN', '内部管理ID', '商品名', '適用開始月', '適用終了月', '単品原価(回答)', '販売入数',
                 '1販売分の原価(案)', '原価単位', '含有範囲', '付属品費用の扱い', '販売構成(現在の商品番号)', '根拠', '比較情報',
                 '区分', '例外理由', '登録できない理由', '承認', '修正・コメント', '状態']


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
            # 「前回の回答」欄に退避された回答も履歴として持ち続ける(再生成を重ねても消えないように)
            pv = ws.cell(r, col['前回の回答']).value if '前回の回答' in col else None
            if pv and not qa[key][0]:
                import re as _re
                base = _re.sub(r'\((条件が変わったため再確認|処理済.*)\)$', '', str(pv)).strip()
                if base:
                    qa[(str(no), key[1], '__prev__')] = (base, None)
    if 'やること' in wb.sheetnames:
        ws = wb['やること']
        for r in range(2, ws.max_row + 1):
            no = ws.cell(r, 1).value
            if no:
                ta[(str(no), str(ws.cell(r, 5).value or ''))] = ws.cell(r, 7).value
    return qa, ta, files[-1]


def build():
    path = f'{OUT_DIR}/英樹への確認リスト_{date.today():%Y%m%d}.xlsx'
    # 英樹が Excel で開いている最中に上書きしない(2026-09-16 に開いたまま再生成してしまった)
    try:
        import excel_bridge as XB
        opened = [x for x in XB.open_workbooks() if os.path.basename(x) == os.path.basename(path)]
        if opened:
            sys.exit(f'⛔ 確認リストが Excel で開かれています。閉じてから再生成してください: {opened[0]}')
    except SystemExit:
        raise
    except Exception:
        pass
    prev_q, prev_t, prev_file = previous_answers(path)
    if os.path.exists(path):                       # 同日の再生成: 今日のファイルの回答も引き継ぐ
        q2, t2, _ = previous_answers('__none__')
        prev_q = {**prev_q, **{k: v for k, v in q2.items() if v[0]}}
        prev_t = {**prev_t, **{k: v for k, v in t2.items() if v}}

    if os.path.exists(path):                       # 上書き前に必ず残す(2026-09-12 に回答済みの列を作り直して消しかけた)
        bdir = OUT_DIR + '/Backup'
        os.makedirs(bdir, exist_ok=True)
        import shutil
        shutil.copy2(path, f'{bdir}/英樹への確認リスト_{date.today():%Y%m%d}_backup_{__import__("datetime").datetime.now():%H%M%S}.xlsx')

    wb = Workbook()
    ws = wb.active
    ws.title = '使い方'
    lines = ['英樹への確認リスト', '',
             '1. 黄色のマスだけ埋めてください。それ以外のセルは触らなくて大丈夫です。',
             '2. 選択肢があるマスはプルダウンです。分からなければ「不明」を選んでください(推測で埋めない)。',
             '3. 「やること」は、いま英樹がやる操作だけを1行1操作で書いています(終わった作業は「完了」シート)。「済」は目安で、完了かどうかは Claude Code がファイルを検証して判定します。',
             '   「質問」シートは未回答が上に並びます。答え終わった質問は下へ移ります。',
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
    # ⚠️ セルを1つも持たない DataValidation を add_data_validation すると <dataValidations count="0"/> が書かれ、
    #    Excel が「一部の内容に問題」と修復を求める(2026-09-16 承認シートで発生)。セルを足してから登録する
    for r, t in enumerate(TASKS, 2):
        for c, v in enumerate(t[:6], 1):
            wt.cell(r, c).value = v; wt.cell(r, c).alignment = WRAP
        done = wt.cell(r, 7)
        done.value = t[6] or prev_t.get((t[0], t[4])) or None
        done.fill = YELLOW
        wt.cell(r, 8).fill = YELLOW
        dv.add(done)
    if dv.sqref:
        wt.add_data_validation(dv)
    for col, w in zip('ABCDEFGH', (7, 22, 44, 16, 56, 44, 8, 30)):
        wt.column_dimensions[col].width = w
    wt.freeze_panes = 'A2'

    # ── 質問 ──
    wq = wb.create_sheet('質問')
    head = ['番号', '種別', '質問', '対象', '選択肢', '回答', '補足(自由に)', '出典', '状態', '質問キー', '前回の回答']
    for c, h in enumerate(head, 1):
        wq.cell(1, c).value = h; wq.cell(1, c).font = BOLD; wq.cell(1, c).fill = HEAD
    Q = build_questions()
    # 未回答(前回の回答が無い)質問を上に並べる。答える場所を探さなくてよいように
    def _answered(q):
        no, kind, text, target, opts, src = q
        k = qkey(text, opts, target, TARGET_YM)
        same = prev_q.get((no, target, k))
        return bool(same and same[0]) or any(pn == no and pt == target and v[0] for (pn, pt, pk), v in prev_q.items())
    Q = sorted(Q, key=lambda q: (1 if _answered(q) else 0))
    carried = reask = 0
    q5a_row = {}
    # 原価確認記録へ登録済みの出品(Q4/Q5)は、その後マスター金額が変わって質問文の数字が変わっても
    # 再確認にしない(回答は記録として確定済み)。処理済みと表示する
    try:
        import cost_confirmations as K
        reg_by_sku = {}
        for rec in K.load():
            if not rec.get('無効化日'):
                reg_by_sku[(str(rec.get('チャネル') or ''), S.norm(rec.get('SKU/ASIN')))] = rec.get('記録ID')
    except Exception:
        reg_by_sku = {}
    done_idx = {i: reg_by_sku[(ch, S.norm(key))] for i, (pid, ch, key, name) in enumerate(LISTINGS, 1)
                if (ch, S.norm(key)) in reg_by_sku}
    for r, (no, kind, text, target, opts, src) in enumerate(Q, 2):
        # 質問キー = 質問文(金額を含む)＋選択肢＋対象(出品・入数・原価単位・構成)＋対象月。どれかが変われば再確認
        k = qkey(text, opts, target, TARGET_YM)
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
        m_idx = None
        import re as _re
        mm = _re.match(r'Q(?:4|5a|5b)-(\d+)$', no)
        if mm:
            m_idx = int(mm.group(1))
        if same and same[0]:
            ans.value, memo.value = same[0], same[1]
            carried += 1
        elif changed and m_idx in done_idx:
            ans.value, memo.value = changed[0]           # 回答は記録として確定済み。質問文の数字だけが変わった
            ans.fill = GRAY
            wq.cell(r, 11).value = f'処理済(原価確認記録 {done_idx[m_idx]} 登録済み。マスター金額の変更で質問文の数字は更新)'
            carried += 1
        elif changed:
            wq.cell(r, 11).value = f'{changed[0][0]}(条件が変わったため再確認)'
            wq.cell(r, 11).fill = GRAY
            reask += 1
        if no in CHATGPT_DECISIONS and not wq.cell(r, 11).value:
            wq.cell(r, 11).value = CHATGPT_DECISIONS[no]; wq.cell(r, 11).fill = GRAY
        wq.cell(r, 9).value = f'=IF(F{r}<>"","回答済",IF(K{r}<>"","再確認","未回答"))'
        if no.startswith('Q5a-'):
            q5a_row[no.split('-', 1)[1]] = r
        if no.startswith('Q5b-') and not ans.value:
            ra = q5a_row.get(no.split('-', 1)[1])
            if ra:
                ans.value = f'=IF(F{ra}="なし","該当なし","")'     # Q5aが「なし」なら自動で該当なし。プルダウンで上書き可
        if opts:
            opts_all = opts + ('／ChatGPTに確認する' if kind == APPROVE and 'ChatGPTに確認する' not in opts else '')
            d = DataValidation(type='list', formula1='"' + opts_all.replace('／', ',') + '"', allow_blank=True)
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
    logs = sorted(glob.glob(f'{OUT_DIR}/商品番号_復元ログ_*.csv') + glob.glob(f'{OUT_DIR}/Archive/*/*/商品番号_復元ログ_*.csv'))   # 退避後も読める
    r = 2
    if logs:
        rows, seen = [], set()
        for lg in logs:
            for x in csv.DictReader(open(lg, encoding='utf-8-sig')):
                if x['差の理由'] == '元xlsx時点で既に数値型' and x['取込先行'] not in seen:
                    rows.append(x); seen.add(x['取込先行'])
        for i, (num, ctrl, skus, name) in enumerate(UP24, 1):
            for x in rows:
                if x['楽天商品管理番号'] == ctrl and x['元値'] == num:
                    for c, v in enumerate((f'Q3-{i}', x['取込先行'], x['楽天商品管理番号'], x['SKU管理番号'], x['元値'], x['商品名']), 1):
                        wm.cell(r, c).value = v
                    r += 1
    n_map = r - 2
    for col, w in zip('ABCDEF', (9, 9, 20, 20, 14, 70)):
        wm.column_dimensions[col].width = w

    # ── 原価記録の承認(追加案を黄色セルで承認する。CSVは内部処理用) ──
    try:
        src_list = prev_file or path
        props = build_cost_proposals(src_list) if os.path.exists(src_list) else []
    except Exception as e:                        # 案が作れない状態なら空シートにして理由を書く
        props, prop_err = [], str(e)
    else:
        prop_err = ''
    wp = wb.create_sheet('原価記録の承認')
    for c, h in enumerate(PROPOSAL_HEAD, 1):
        wp.cell(1, c).value = h; wp.cell(1, c).font = BOLD; wp.cell(1, c).fill = HEAD
    prev_appr = {}
    if prev_file and os.path.exists(prev_file):
        pw = load_workbook(prev_file, data_only=True)
        if '原価記録の承認' in pw.sheetnames:
            pws = pw['原価記録の承認']
            ph = [str(c.value or '') for c in pws[1]]; pc = {h: i + 1 for i, h in enumerate(ph)}
            for r in range(2, pws.max_row + 1):
                k = (pws.cell(r, pc['案No']).value, pws.cell(r, pc['SKU/ASIN']).value, pws.cell(r, pc['1販売分の原価(案)']).value,
                     pws.cell(r, pc['適用開始月']).value, pws.cell(r, pc.get('含有範囲', 12)).value)
                if k[0]:
                    prev_appr[k] = (pws.cell(r, pc['承認']).value, pws.cell(r, pc['修正・コメント']).value)
    dva = DataValidation(type='list', formula1='"承認,修正あり,保留"', allow_blank=True)
    # 既に原価確認記録へ登録済みの案は「登録済(記録ID)」として固定し、承認欄を黄色にしない
    try:
        import cost_confirmations as K
        registered = {}
        for rec in K.load():
            if rec.get('無効化日'):
                continue
            key = (str(rec.get('チャネル') or ''), S.norm(rec.get('SKU/ASIN')), K._ym(rec.get('適用開始月')),
                   float(rec.get('1販売分の原価') or 0))
            registered[key] = rec.get('記録ID')
    except Exception:
        registered = {}
    for r, pr in enumerate(props, 2):
        for c, h in enumerate(PROPOSAL_HEAD, 1):
            if h in ('承認', '修正・コメント', '状態'):
                continue
            wp.cell(r, c).value = pr.get(h); wp.cell(r, c).alignment = WRAP
        ca, cm = wp.cell(r, PROPOSAL_HEAD.index('承認') + 1), wp.cell(r, PROPOSAL_HEAD.index('修正・コメント') + 1)
        rid = registered.get((pr['チャネル'], S.norm(pr['SKU/ASIN']), pr['適用開始月'],
                              float(pr['1販売分の原価(案)'] or 0)))
        if rid:
            ca.value = f'登録済({rid})'; ca.fill = GRAY
            wp.cell(r, PROPOSAL_HEAD.index('状態') + 1).value = (
                f'登録済({rid})・反映未対応(Amazonの生成処理は記録を参照しない)' if pr['チャネル'] == 'Amazon'
                else f'登録済({rid})・次の新月生成で反映(本番の新月生成は保留)')
            continue
        ca.fill = YELLOW; cm.fill = YELLOW; dva.add(ca)
        # 金額・対象月・含有範囲が同じ案だけ前回の承認を引き継ぐ(変わっていれば再確認)
        k = (pr['案No'], pr['SKU/ASIN'], pr['1販売分の原価(案)'], pr['適用開始月'], pr['含有範囲'])
        if k in prev_appr and prev_appr[k][0]:
            ca.value, cm.value = prev_appr[k]
        st = wp.cell(r, PROPOSAL_HEAD.index('状態') + 1)
        col_ok, col_ap, col_blk = (PROPOSAL_HEAD.index(x) + 1 for x in ('承認', '承認', '登録できない理由'))
        L = lambda i: wp.cell(1, i).column_letter
        st.value = (f'=IF({L(col_blk)}{r}<>"","登録不可("&{L(col_blk)}{r}&")",'
                    f'IF({L(col_ap)}{r}="承認","登録待ち(Claude Code)",IF({L(col_ap)}{r}<>"",{L(col_ap)}{r},"未承認")))')
    if dva.sqref:                                  # セルがあるときだけ登録(空の dataValidations を書かない)
        wp.add_data_validation(dva)
    if prop_err:
        wp.cell(2, 1).value = f'(追加案を作れませんでした: {prop_err})'
    for col, w in zip('ABCDEFGHIJKLMNOPQRSTUV', (8, 8, 14, 10, 28, 9, 9, 10, 7, 12, 9, 18, 18, 18, 40, 46, 16, 40, 30, 10, 26, 22)):
        wp.column_dimensions[col].width = w
    wp.freeze_panes = 'A2'
    # 内部処理用CSV(承認の入力はここに求めない)
    if props:
        import csv as _csv
        with open(f'{OUT_DIR}/原価確認記録_追加案_{date.today():%Y%m%d}.csv', 'w', encoding='utf-8-sig', newline='') as f:
            w = _csv.DictWriter(f, fieldnames=list(props[0].keys())); w.writeheader(); w.writerows(props)

    # ── ChatGPTへの確認(英樹が判断できないもの・ChatGPTの判断待ち) ──
    wc = wb.create_sheet('ChatGPTへの確認')
    for c, h in enumerate(['番号', '質問', '英樹の回答', '英樹の補足', 'Claude Code の整理', 'ChatGPTの判断(貼り付け)'], 1):
        wc.cell(1, c).value = h; wc.cell(1, c).font = BOLD; wc.cell(1, c).fill = HEAD
    rr = 2
    for r in range(2, wq.max_row + 1):
        no, kind, ans_v, memo_v = wq.cell(r, 1).value, wq.cell(r, 2).value, wq.cell(r, 6).value, wq.cell(r, 7).value
        # ChatGPT が(英樹の回答前に)判断を出したものも載せる(Q14)。判断は記録であって英樹の回答を消さない
        if ans_v == 'ChatGPTに確認する' or (kind == APPROVE and ans_v == '不明') or str(no) in CHATGPT_DECISIONS:
            wc.cell(rr, 1).value = no; wc.cell(rr, 2).value = wq.cell(r, 3).value; wc.cell(rr, 3).value = ans_v
            wc.cell(rr, 4).value = memo_v
            wc.cell(rr, 5).value = CHATGPT_NOTES.get(str(no), '')
            wc.cell(rr, 6).value = CHATGPT_DECISIONS.get(str(no))
            wc.cell(rr, 6).fill = YELLOW if not CHATGPT_DECISIONS.get(str(no)) else GRAY
            for c in (2, 4, 5, 6):
                wc.cell(rr, c).alignment = WRAP
            rr += 1
    for col, w in zip('ABCDEF', (8, 70, 16, 30, 50, 40)):
        wc.column_dimensions[col].width = w
    wc.freeze_panes = 'A2'

    # ── 完了(終わった作業。やることには載せない) ──
    wd = wb.create_sheet('完了')
    for c, h in enumerate(['番号', '作業', '結果'], 1):
        wd.cell(1, c).value = h; wd.cell(1, c).font = BOLD; wd.cell(1, c).fill = HEAD
    for r, (no, name, res) in enumerate(DONE, 2):
        wd.cell(r, 1).value = no; wd.cell(r, 2).value = name; wd.cell(r, 3).value = res; wd.cell(r, 3).alignment = WRAP
    for col, w in zip('ABC', (10, 40, 90)):
        wd.column_dimensions[col].width = w

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
    print(f'   やること {len(TASKS)}行 / 質問 {len(Q)}件(前回から引き継ぎ {carried}・再確認 {reask}) / Q3対応表 {n_map}行 / 原価記録の承認 {len(props)}件 / 既知の回答 {len(KNOWN)}件')
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
