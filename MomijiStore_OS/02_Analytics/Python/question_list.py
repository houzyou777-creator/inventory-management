# -*- coding: utf-8 -*-
"""question_list.py — 英樹への「やること・質問」をExcelの記入用リストにする／回答を読む

使い方:
    python3 question_list.py build            Output/英樹への確認リスト_<日付>.xlsx を作る
    python3 question_list.py read <xlsx>      黄色の回答欄を読んで一覧表示する(何も書き換えない)

なぜ作るか(2026-09-12 英樹の要望):
    チャットの文章では「どのファイルをどう直すか」が分からない。
    1行1操作の「やること」と、選択肢つきの「質問」を**黄色のマスだけ埋めればよい**形にする。
    英樹はこのファイルだけ開けばよく、マスター等の正本を直接編集しなくてよい。
    回答を受けて正本へ書くのは Claude Code(バックアップつき)。

ルール:
    ・黄色のセル以外は触らなくてよい(数式やIDは動かさない)
    ・選択肢があるものはプルダウン。分からなければ「不明」を選ぶ。**推測で埋めない**
    ・「済」を付けるのは、その行に書いてある「確認できること」が実際に見えたときだけ
"""
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

OUT_DIR = os.path.dirname(T.RAKUTEN_STOCK) + '/Output'
YELLOW = PatternFill('solid', fgColor='FFFF00')
HEAD = PatternFill('solid', fgColor='D9E1F2')
BOLD = Font(bold=True)
WRAP = Alignment(vertical='top', wrap_text=True)

V2 = 'MomijiStore_OS/01_InventoryManagement/SourceData/V2_test'

# ──────────────────────────────────────────────────────────────
# やること(1行1操作)。列: 番号 / 作業 / 開くファイル / どこで / 操作 / 確認できること
# ──────────────────────────────────────────────────────────────
TASKS = [
    ('V2-1', 'V-2テスト ① baseline', f'{V2}/baseline/楽天在庫金額集計ツール_v1.0.xlsm',
     '—', '「集計実行」を押して保存', '✅ 2026-09-12 20:28 実施済み(Claude Code確認)', '済'),
    ('V2-2', 'V-2テスト ② 修正版VBAを入れる', f'{V2}/楽天在庫金額集計ツール_v1.0.xlsm(baselineフォルダの**外**の方)',
     'Excelで開く', 'マクロを有効にして開く', 'シート「楽天CSV取込」が見える', ''),
    ('V2-3', '', '同上', 'メニュー', 'ツール → マクロ → Visual Basic Editor', 'VBEの画面が開く', ''),
    ('V2-4', '', '同上', 'VBE 左の一覧', '「Module_Rakuten_Tool」を右クリック →「Module_Rakuten_Toolの解放」→ エクスポートしますか？は「いいえ」',
     '一覧から Module_Rakuten_Tool が消える', ''),
    ('V2-5', '', '同上', 'VBE メニュー', 'ファイル → ファイルのインポート → ' + f'{V2}/Module_Rakuten_Tool_V2.bas を選ぶ',
     '一覧に「Module_Rakuten_Tool」が**1つだけ**戻る(Module_Rakuten_Tool1 になっていたらV2-4をやり直す)', ''),
    ('V2-6', '', '同上', 'VBE', 'VBEのウィンドウを閉じる(Excelは閉じない)', 'Excelのシートに戻る', ''),
    ('V2-7', 'V-2テスト ③ 取込と集計', '同上', 'シート「楽天CSV取込」', '「楽天CSV読込」ボタン → 上書き確認は「はい」→ 完了は「OK」',
     'K列「読込日時」が今日の日時になる', ''),
    ('V2-8', '', '同上', 'ボタン', '「集計実行」→ 完了は「OK」', 'シート「在庫金額集計」の集計日時が今日になる', ''),
    ('V2-9', '', '同上', '保存', '⌘S で保存して閉じる', 'ファイルの更新日時が今になる', ''),
    ('V2-10', '', 'チャット', '—', '「V2-9まで済」と送る', 'Claude Code が verify_v2_import.py で判定する', ''),
]

# ──────────────────────────────────────────────────────────────
# 質問(選択肢つき)。列: 番号 / 質問 / 対象 / 選択肢 / 回答 / 補足 / 出典
# ──────────────────────────────────────────────────────────────
UP24 = [  # 元xlsx時点で数値だった24行(商品ごとにまとめる)
    ('272', 'INTEX 子供用 浮き輪'), ('577', 'BOSCH スバル車用エアコンフィルター'),
    ('5946', 'Nuxe プロディジュー フローラル オイル(行10)'), ('9785', 'Nuxe プロディジュー ゴールドオイル'),
    ('9761', 'Nuxe プロディジューオイル 50mL'), ('4382', 'Nuxe プロディジュー フローラル オイル(行212)'),
    ('1965', 'ソフィーナiP ベースケア セラム'), ('7480', 'Echo Show 5 第3世代(3行)／Echo Spot 2024(3行)'),
    ('15980', 'Amazon Fire HD 10 キッズモデル(3行)'), ('5538', '資生堂 HAKU メラノフォーカスEV'),
    ('5980', 'Fire TV Stick 4K'), ('2176', 'さらさ 柔軟剤 詰め替え 1350mL×2'),
    ('32780', 'DJI Osmo Action 4'), ('1980', 'ラックス バスグロウ セット'),
    ('2000', 'クーリア 福袋 3点セット(3行)'),
]
LISTINGS = [  # 出品テーブルに販売入数が入っている12出品
    ('P000049', 'Amazon', 'B0DJCX5CCW', 3, 'ニチドウ 毛玉クリーン 60g'),
    ('P000049', 'Amazon', 'B0DJCXKM4G', 2, 'ニチドウ 毛玉クリーン 60g'),
    ('P000056', 'Amazon', 'B0DKSGDMCZ', 2, 'ドクターズチョイス ゴートミルク 120g'),
    ('P000056', 'Amazon', 'B0DKSVB9B5', 4, 'ドクターズチョイス ゴートミルク 120g'),
    ('P000225', '楽天', 'b0c656jxkd', 3, '肌ラボ 白潤プレミアム 乳液 140ml'),
    ('P000225', 'Amazon', 'B08WS9L34F', 1, '肌ラボ 白潤プレミアム 乳液 140ml'),
    ('P000327', 'Amazon', 'B0DGQJXT78', 3, 'ニチドウ Dr.PRO. ベビーミルク'),
    ('P000334', 'Amazon', 'B0DKSXGXNZ', 2, 'ドクターズチョイス 納豆菌 ふりかけ 80g'),
    ('P000334', 'Amazon', 'B0DKSPG7VK', 5, 'ドクターズチョイス 納豆菌 ふりかけ 80g'),
    ('P000334', 'Amazon', 'B0DKS7KSRB', 4, 'ドクターズチョイス 納豆菌 ふりかけ 80g'),
    ('P000655', '楽天', 'b0fcfk55zk-4', 4, '【4個セット】アリエール 液体洗剤'),
    ('P000846', 'Amazon', 'B0H6L99WP3', 2, 'トイレその後に 280ml×2本'),
]

QUESTIONS = []
QUESTIONS.append(('Q1', '8/3に楽天RMSからダウンロードした在庫CSV(Excelに変換する前の元ファイル)はPCに残っていますか？ 残っていれば MomijiStore_OS/01_InventoryManagement/SourceData/Import/ にコピーしてください',
                  '上流24件の復元', '残っている(Importへ置いた)／残っていない／探している', '上流24件'))
QUESTIONS.append(('Q2', 'RMSの商品番号が誤っている2件を、RMSの商品ページで直しますか？ 直さないと次回の取込で誤りが戻ります',
                  'b087lzdd59 → 5036/1273 ／ b0dfyjnwjt → 4987176260635-2', '直した／後で直す／直さない', '上流誤記2件'))
for i, (num, name) in enumerate(UP24, 1):
    QUESTIONS.append((f'Q3-{i}', f'商品番号「{num}」は何の数字ですか？(元のファイルで数字になっていたため先頭の0やカンマが消えた可能性があります。元が「0{num}」のような形なら「補足」に書いてください)',
                      f'{name}', 'JAN下4桁／原価(円)／不明', '上流24件(A-7)'))
for i, (pid, ch, key, n, name) in enumerate(LISTINGS, 1):
    QUESTIONS.append((f'Q4-{i}', f'この出品の原価(1販売分)は、来月以降も同じと考えてよいですか？(「可」ならKPI生成で過去月の確認済み原価を引き継ぎます)',
                      f'{pid} {ch} {key} 入数{n} {name}', '可／不可(仕入値が変わる)／不明', 'Y-1 原価継続'))
for i, (pid, ch, key, n, name) in enumerate(LISTINGS, 1):
    QUESTIONS.append((f'Q5-{i}', f'この出品の原価に付属品やおまけは含まれていませんか？(含有範囲。「なし」なら RMS商品番号の原価を使う条件が1つ揃います)',
                      f'{pid} {ch} {key} 入数{n} {name}', '付属品なし(確認済み)／付属品あり／不明', 'Z-4 含有範囲'))
QUESTIONS.append(('Q6', 'HDD(P000059)の標準原価を 17,980円 に更新してよいですか？(8月KPIは変えません。今後の標準原価として)',
                  'P000059', '更新してよい／まだ／不明', '保留T-1'))
QUESTIONS.append(('Q7', 'リポソームショット(B0FZ3Y7QNM / JAN 04571509302842)の標準原価を 175円 としてマスターへ登録してよいですか？',
                  'B0FZ3Y7QNM', '登録してよい／まだ／不明', '保留'))


def build():
    wb = Workbook()
    ws = wb.active
    ws.title = '使い方'
    lines = ['英樹への確認リスト', '',
             '1. 黄色のマスだけ埋めてください。それ以外のセルは触らなくて大丈夫です。',
             '2. 選択肢があるマスはプルダウンです。分からなければ「不明」を選んでください(推測で埋めない)。',
             '3. 「やること」シートは1行1操作です。右の「確認できること」が実際に見えたら「済」にしてください。',
             '4. 終わったら保存して、チャットで「回答した」と一言ください。',
             '5. 正本(商品マスター・KPIシート・在庫ツール)への書き込みは Claude Code がバックアップつきで行います。',
             '', f'作成 {date.today():%Y-%m-%d} / Claude Code']
    for i, l in enumerate(lines, 1):
        ws.cell(i, 1).value = l
    ws.cell(1, 1).font = Font(bold=True, size=14)
    ws.column_dimensions['A'].width = 100

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
        done = wt.cell(r, 7); done.value = t[6] or None; done.fill = YELLOW
        wt.cell(r, 8).fill = YELLOW
        dv.add(done)
    for col, w in zip('ABCDEFGH', (7, 22, 44, 16, 56, 44, 8, 30)):
        wt.column_dimensions[col].width = w
    wt.freeze_panes = 'A2'

    # ── 質問 ──
    wq = wb.create_sheet('質問')
    head = ['番号', '質問', '対象', '選択肢', '回答', '補足(自由に)', '出典', '状態']
    for c, h in enumerate(head, 1):
        wq.cell(1, c).value = h; wq.cell(1, c).font = BOLD; wq.cell(1, c).fill = HEAD
    for r, q in enumerate(QUESTIONS, 2):
        no, text, target, opts, src = q
        for c, v in enumerate((no, text, target, opts), 1):
            wq.cell(r, c).value = v; wq.cell(r, c).alignment = WRAP
        ans = wq.cell(r, 5); ans.fill = YELLOW
        wq.cell(r, 6).fill = YELLOW
        wq.cell(r, 7).value = src
        wq.cell(r, 8).value = f'=IF(E{r}<>"","回答済","未回答")'
        if opts:
            d = DataValidation(type='list', formula1='"' + opts.replace('／', ',') + '"', allow_blank=True)
            wq.add_data_validation(d); d.add(ans)
    for col, w in zip('ABCDEFGH', (7, 70, 40, 34, 24, 30, 14, 8)):
        wq.column_dimensions[col].width = w
    wq.freeze_panes = 'A2'

    os.makedirs(OUT_DIR, exist_ok=True)
    path = f'{OUT_DIR}/英樹への確認リスト_{date.today():%Y%m%d}.xlsx'
    wb.save(path)
    print(f'→ {path}  やること {len(TASKS)}行 / 質問 {len(QUESTIONS)}件')
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
    n_ans = 0
    for r in range(2, wq.max_row + 1):
        no, ans, memo = wq.cell(r, 1).value, wq.cell(r, 5).value, wq.cell(r, 6).value
        if not no:
            continue
        if ans:
            n_ans += 1
        print(f'  {no:6} {ans or "—":24} {memo or ""}')
    print(f'  回答 {n_ans} / {wq.max_row - 1}')


if __name__ == '__main__':
    if len(sys.argv) >= 2 and sys.argv[1] == 'build':
        build()
    elif len(sys.argv) >= 3 and sys.argv[1] == 'read':
        read(sys.argv[2])
    else:
        print(__doc__)
