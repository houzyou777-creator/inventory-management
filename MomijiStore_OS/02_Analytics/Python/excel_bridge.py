# -*- coding: utf-8 -*-
"""excel_bridge.py — Excel(AppleScript)へ書くときの共通の安全装置(2026-09-15 ChatGPT指示)

2026-09-12 22:45 の事故:
    英樹が本番の在庫ツールを開いている最中に、KPIコピーの再計算スクリプトが
    「active workbook」を保存して閉じた(事故後の保存内容はバックアップと一致。
    直前の未保存状態まで影響なしとは断定しない)。

この層を通さずに Excel を操作するスクリプトを書かない。守ること:
    1. 対象は**フルパス**で特定する(ブック名だけでは本番とコピーを取り違える)
    2. 開いた直後だけでなく、**書込み・保存・閉じる直前にも**対象を再確認する
    3. スクリプト開始前から開いていたブックは、保存も閉じもしない
    4. その実行で自分が開いたブックだけを保存・閉じる。エラー時も同じ
    5. 対象ブックがすでに(人によって)開かれていたら、処理を止めて知らせる。
       他のブックが開いているだけでは止めない。ただし**同名の別ブック**が開いていたら止める
       (Excel は同名のブックを2つ開けず、取り違えのもと)
    6. Excel の終了(quit)や、全ブックを対象にする保存・終了は使わない

使い方:
    from excel_bridge import run_on_workbook, recalc
    out = run_on_workbook(path, body, timeout=600)   # body は wb / ws を使う AppleScript 断片
    recalc(path)                                     # 開く→再計算→保存→閉じる
body の中では `wb`(対象ブック)が使える。`active workbook` は書かない。
戻り値はスクリプトの return 文字列。'ABORT:' で始まれば何もしていない。
"""
import os
import subprocess
import tempfile


class ExcelAbort(RuntimeError):
    """対象ブックが人に開かれている等、安全に処理できないため**何もせず**止めたことを表す。"""


def _lit(s):
    return '"' + str(s).replace('\\', '\\\\').replace('"', '\\"') + '"'


def build_script(path, body, result_expr='"ok"', timeout=600, close_on_error=True):
    """安全装置つきの AppleScript を組み立てる。body は tell ブロック内に挿入される。

    Excel は `full name of every workbook`(リスト)は返すが、`repeat with w in workbooks` の
    要素に対する `full name of w` はエラー(-50)になるため、名前とフルパスのリストを添字で回す。
    /tmp は /private/tmp のシンボリックリンクで、Excel は開いたときの表記で返すため両方を対象にする。
    """
    path = os.path.abspath(path)
    real = os.path.realpath(path)
    name = os.path.basename(path)
    # macOS の /tmp,/var,/etc は /private/... へのシンボリックリンク。Excel は開いたときの表記で返すので両表記を対象にする
    variants = set()
    for x in (path, real):
        variants.add(x)
        if x.startswith('/private/'):
            variants.add(x[len('/private'):])
        elif x.startswith(('/tmp/', '/var/', '/etc/')):
            variants.add('/private' + x)
    paths = '{' + ', '.join(_lit(x) for x in sorted(variants)) + '}'
    return f'''
set targetPaths to {paths}
set targetName to {_lit(name)}
set p to POSIX file {_lit(path)}
with timeout of {int(timeout)} seconds
tell application "Microsoft Excel"
    -- 【事前】対象(フルパス)や同名のブックが開いていないか。開いていたら何もしない
    set fns to {{}}
    set nms to {{}}
    try
        set fns to full name of every workbook
        set nms to name of every workbook
    end try
    repeat with i from 1 to (count of nms)
        set fn to (item i of fns) as string
        set nm to (item i of nms) as string
        if targetPaths contains fn then return "ABORT:already-open:" & fn
        if nm is targetName then return "ABORT:same-name-open:" & fn
    end repeat
    set preCount to count of nms

    open p
    -- 【開いた直後】フルパスで対象を特定する(active workbook は使わない)
    set fns to full name of every workbook
    set nms to name of every workbook
    set wbName to ""
    repeat with i from 1 to (count of nms)
        if targetPaths contains ((item i of fns) as string) then set wbName to (item i of nms) as string
    end repeat
    if wbName is "" then return "ABORT:not-opened"
    set wb to workbook wbName

    try
        {body}
        -- 【保存直前】対象が変わっていないことを再確認してから、自分が開いたブックだけ保存して閉じる
        if not (targetPaths contains ((full name of wb) as string)) then error "target changed before save"
        save wb
        close wb saving no
    on error errMsg number errNum
        -- close_on_error=False のときは閉じない(タイムアウト時は Excel がまだ実行中かもしれない。呼び出し側が状態を確認する)
        if {str(close_on_error).lower()} then
            try
                if targetPaths contains ((full name of wb) as string) then close wb saving no
            end try
        end if
        if errNum is -1712 then return "TIMEOUT:" & errMsg
        return "ERROR:" & errMsg
    end try
    if (count of workbooks) is not preCount then return "WARN:workbook-count-changed:" & (count of workbooks)
end tell
end timeout
return {result_expr}
'''


class ExcelTimeout(RuntimeError):
    """AppleScript の待ち時間を超えた。**処理が止まったとは限らない**(Excel がまだ実行中・ダイアログ表示中の可能性)。
    呼び出し側は excel_state() で状態を確認するまで、再実行も手動操作への切替もしない(ChatGPT 2026-09-19 条件1)。"""


def run_on_workbook(path, body, result_expr='"ok"', timeout=600, close_on_error=True):
    """対象ブックを開き body を実行し、保存して閉じる。ABORT/ERROR は例外にする。
    close_on_error=False: エラー時に対象ブックを閉じない(状態確認を呼び出し側に委ねる)。"""
    if not os.path.exists(path):
        raise FileNotFoundError(path)
    script = build_script(path, body, result_expr, timeout, close_on_error)
    with tempfile.NamedTemporaryFile('w', suffix='.applescript', delete=False, encoding='utf-8') as f:
        f.write(script)
        sp = f.name
    try:
        r = subprocess.run(['osascript', sp], capture_output=True, text=True, timeout=timeout + 120)
    except subprocess.TimeoutExpired:
        raise ExcelTimeout(f'osascript が {timeout + 120} 秒で返らなかった(Excel 側の状態は未確認)')
    finally:
        os.unlink(sp)
    if r.returncode != 0:
        raise RuntimeError(f'AppleScript failed: {r.stderr.strip()}')
    out = r.stdout.strip()
    if out.startswith('TIMEOUT:'):
        raise ExcelTimeout(out[8:])
    if out.startswith('ABORT:already-open:'):
        raise ExcelAbort(f'対象ブックが既に開かれています。閉じてから再実行してください: {out.split(":", 2)[2]}')
    if out.startswith('ABORT:same-name-open:'):
        raise ExcelAbort(f'同名の別ブックが開かれています(本番/コピーの取り違え防止で停止): {out.split(":", 2)[2]}')
    if out.startswith('ABORT:'):
        raise ExcelAbort(out)
    if out.startswith('ERROR:'):
        raise RuntimeError(f'Excel処理でエラー(対象ブックは保存せず閉じた): {out[6:]}')
    return out


def recalc(path, timeout=600):
    """開く → 再計算 → 保存 → 閉じる(対象ブックだけ)。"""
    return run_on_workbook(path, 'delay 1\n        calculate', timeout=timeout)


def open_workbooks():
    """いま開いているブックのフルパス一覧(読み取り専用の確認用)。"""
    r = subprocess.run(['osascript', '-e',
                        'tell application "Microsoft Excel" to if running then get full name of every workbook'],
                       capture_output=True, text=True, timeout=60)
    s = r.stdout.strip()
    if not s or s == 'missing value':
        return []
    return [x.strip() for x in s.split(',')]


def run_macros(path, macros, timeout=900):
    """対象ブックを開き、VBA の Public Function を順に実行し(戻り値が "OK:" で始まらなければ保存せず閉じる)、
    保存して閉じる。戻り値は各マクロの結果を " || " でつないだ文字列。

    在庫ツール V-4 の AutoImport / AutoAggregate(MsgBox を出さず結果を返す)を AppleScript から呼ぶために追加(2026-09-17)。
    マクロ側に画面(MsgBox/InputBox)が残っていると AppleScript が止まるので、呼ぶのは「自動実行用」の関数だけ。
    安全装置(フルパス特定・自分が開いたブックだけ保存・閉じる・既に開かれていれば停止)は build_script のまま。
    """
    lines = []
    for i, m in enumerate(macros, 1):
        lines.append(f'set r{i} to (run VB macro ("\'" & wbName & "\'!" & {_lit(m)})) as text')
        lines.append(f'if r{i} does not start with "OK:" then error {_lit(m + ": ")} & r{i}')
    body = '\n        '.join(lines)
    result = ' & " || " & '.join(f'r{i}' for i in range(1, len(macros) + 1))
    # エラー・タイムアウト時に対象ブックを勝手に閉じない(閉じるのは excel_state で確認した後に close_without_saving で)
    return run_on_workbook(path, body, result_expr=result, timeout=timeout, close_on_error=False)


def excel_state(path, wait=20):
    """Excel 側の実行状態(読み取り専用)。タイムアウト後の確認に使う。
    戻り値: dict(responding, open_workbooks, target_open, dialogs, file_mtime, file_size)
    responding=False は「Excel が Apple Event に応答しない」= マクロ実行中かダイアログ表示中の可能性。"""
    import datetime
    st = dict(responding=None, open_workbooks=[], target_open=None, dialogs=[], file_mtime=None, file_size=None)
    if os.path.exists(path):
        st['file_mtime'] = datetime.datetime.fromtimestamp(os.path.getmtime(path)).strftime('%Y-%m-%d %H:%M:%S')
        st['file_size'] = os.path.getsize(path)
    try:
        r = subprocess.run(['osascript', '-e', f'with timeout of {int(wait)} seconds',
                            '-e', 'tell application "Microsoft Excel" to if running then get full name of every workbook',
                            '-e', 'end timeout'], capture_output=True, text=True, timeout=wait + 15)
        st['responding'] = r.returncode == 0
        v = r.stdout.strip()
        st['open_workbooks'] = [] if (not v or v == 'missing value') else [x.strip() for x in v.split(',')]
        variants = {os.path.abspath(path), os.path.realpath(path)}
        st['target_open'] = any(w in variants for w in st['open_workbooks'])
    except subprocess.TimeoutExpired:
        st['responding'] = False
    try:   # ダイアログ(モーダルウィンドウ)が出ていないか。System Events が使えなければ空
        r = subprocess.run(['osascript', '-e', f'with timeout of {int(wait)} seconds',
                            '-e', 'tell application "System Events" to tell process "Microsoft Excel" to get name of every window',
                            '-e', 'end timeout'], capture_output=True, text=True, timeout=wait + 15)
        if r.returncode == 0:
            st['dialogs'] = [x.strip() for x in r.stdout.strip().split(',') if x.strip()]
    except subprocess.TimeoutExpired:
        pass
    return st


def close_without_saving(path, timeout=60):
    """自分が開いたままになっている対象ブック(フルパス一致)だけを保存せずに閉じる。開いていなければ何もしない。"""
    body = f'''
set targetPaths to {{{_lit(os.path.abspath(path))}, {_lit(os.path.realpath(path))}}}
with timeout of {int(timeout)} seconds
tell application "Microsoft Excel"
    set fns to full name of every workbook
    set nms to name of every workbook
    repeat with i from 1 to (count of nms)
        if targetPaths contains ((item i of fns) as string) then
            close workbook ((item i of nms) as string) saving no
            return "closed"
        end if
    end repeat
end tell
end timeout
return "not-open"
'''
    r = subprocess.run(['osascript', '-e', body], capture_output=True, text=True, timeout=timeout + 30)
    if r.returncode != 0:
        raise RuntimeError(f'close failed: {r.stderr.strip()}')
    return r.stdout.strip()
