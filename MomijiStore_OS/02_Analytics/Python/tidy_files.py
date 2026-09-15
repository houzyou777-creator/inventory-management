# -*- coding: utf-8 -*-
"""tidy_files.py — 成果物・バックアップの整理(移動は承認後・削除は候補提示のみ)

使い方:
    python3 tidy_files.py plan              何をどこへ動かすかを表示する(何もしない)
    python3 tidy_files.py apply             plan の「移動」だけを実行する(削除はしない)
    python3 tidy_files.py delete-candidates 削除候補の一覧を表示する(削除はしない)

方針(CLAUDE.md: 削除・移動・リネームは承認制。古いバックアップの自動削除禁止):
    ・**削除はこのスクリプトでは行わない。** 候補を提示し、英樹が承認したものだけ別途手で消す
    ・「移動」は Archive/<年月>/ への退避だけ。ファイル名は変えない(import名は変更禁止)
    ・いま使うもの(最新の確認リスト・最新の要対応一覧など)は Output に残す
    ・証跡(復元ログ・変更記録・報告書・V-2/V-3の検証コピー)は消さない。退避先へまとめる

なぜ: Output が39ファイル・Backup が44ファイルになり、「どれを見ればいいか」が分からなくなった(英樹 2026-09-16)。
"""
import glob
import os
import re
import shutil
import sys

BASE = '/Users/hide0726/Desktop/Claude Code/MomijiStore_OS'
OUT = BASE + '/01_InventoryManagement/SourceData/Output'
BK = BASE + '/01_InventoryManagement/SourceData/Backup'
BK2 = BASE + '/02_Analytics/SourceData/Backup'
V2T = BASE + '/01_InventoryManagement/SourceData/V2_test'


def _ym_of(name):
    m = re.search(r'(20\d{2})-?(\d{2})-?\d{2}', name)
    return f'{m.group(1)}-{m.group(2)}' if m else '不明'


def _latest_keep(pattern, keep=1):
    """同じ種類の生成物は最新 keep 個を残し、古いものを退避候補にする。"""
    files = sorted(glob.glob(pattern))
    return files[:-keep] if len(files) > keep else []


def plan():
    moves, deletes, keeps = [], [], []
    # ── Output: 生成物は最新だけ残す ──
    for pat, keep, why in (
        (OUT + '/要対応一覧_*.xlsx', 1, '毎月再生成される。最新だけ残す'),
        (OUT + '/広告改善一覧_*.xlsx', 1, '同上'),
        (OUT + '/RMS商品番号_分類_*.csv', 1, '第4版(0912)だけ残す'),
        (OUT + '/原価確認記録_追加案_*.csv', 1, '内部処理用。最新だけ'),
    ):
        for f in _latest_keep(pat, keep):
            moves.append((f, f'{OUT}/Archive/{_ym_of(os.path.basename(f))}/生成物', why))
    # 報告書(md)は証跡として Archive/報告 へ(Output 直下には最新3本だけ)
    reports = sorted(glob.glob(OUT + '/20*.md'))
    for f in reports[:-3]:
        moves.append((f, f'{OUT}/Archive/{_ym_of(os.path.basename(f))}/報告', '証跡。読み返すときは Archive から'))
    # 復元・変更の証跡CSV
    for pat in ('商品番号_復元*', '識別子_先頭0復元ログ_*', '原価マスター_識別子復元ログ_*', '商品マスター_変更記録_*', '原価確認記録_登録ログ_*', '共有在庫_照合表_*'):
        for f in sorted(glob.glob(f'{OUT}/{pat}')):
            moves.append((f, f'{OUT}/Archive/{_ym_of(os.path.basename(f))}/証跡', '証跡。消さない'))
    # 古い単発の成果物
    for name, why in (('KPI仕入値更新候補_8月_20260910.xlsx', '2026-09-10 の候補一覧。push-kpi は保留のまま'),
                      ('商品情報不足_AmazonSKU未解決_8月_20260908.xlsx', '2026-09-08 の一覧。要対応一覧に統合済み'),
                      ('要確認リスト_20260713.xlsx', '7月の旧リスト')):
        f = f'{OUT}/{name}'
        if os.path.exists(f):
            moves.append((f, f'{OUT}/Archive/{_ym_of(name)}/生成物', why))
    # 確認リストのバックアップ: 「旧版」2本は残し、途中の _backup_HHMMSS は削除候補
    for f in sorted(glob.glob(OUT + '/Backup/英樹への確認リスト_*_backup_*.xlsx')):
        deletes.append((f, '再生成前の一時バックアップ。回答は最新リストへ引き継ぎ済み'))
    # ── SourceData/Backup ──
    for f in sorted(glob.glob(BK + '/楽天在庫リスト_backup_20260803_*.xlsx'))[:-1]:
        deletes.append((f, '2026-08-03 の作業中の連続バックアップ。最後の1本(単独再集計前)だけ残す'))
    for f in sorted(glob.glob(BK + '/楽天在庫リスト_backup_20260709_*.xlsx')):
        deletes.append((f, '7月の旧重複版'))
    for name in ('楽天在庫金額集計ツール_backup_20260712_修正前1.xlsm', '楽天在庫金額集計ツール_backup_20260712_修正前2.xlsm',
                 '楽天在庫金額集計ツール_backup_20260712_修正前3.xlsm'):
        f = f'{BK}/{name}'
        if os.path.exists(f):
            deletes.append((f, '7月の作業中の連続バックアップ(商品コード修正前/後 は残す)'))
    for f in sorted(glob.glob(BK + '/*backup_202607*')) + sorted(glob.glob(BK + '/*backup_20260803*')):
        if f not in [d[0] for d in deletes]:
            moves.append((f, f'{BK}/Archive/{_ym_of(os.path.basename(f))}', '7〜8月の節目バックアップ。消さずに退避'))
    # 在庫ツールの V-2/V-3 系(9月)と商品マスターの9月分は Backup 直下に残す
    for f in sorted(glob.glob(BK + '/*202609*')):
        keeps.append((f, '直近(9月)の作業直前バックアップ。残す'))
    # ── 02_Analytics/SourceData/Backup: 同じファイルは最新2本を残す ──
    groups = {}
    for f in sorted(glob.glob(BK2 + '/*.xlsx')):
        key = re.sub(r'_backup_.*', '', os.path.basename(f))
        groups.setdefault(key, []).append(f)
    for key, files in groups.items():
        for f in files[:-2]:
            moves.append((f, f'{BK2}/Archive/{_ym_of(os.path.basename(f))}', f'{key} の古いバックアップ(最新2本は残す)'))
    # ── V2_test: 検証証跡。ChatGPT指示で保持。容量が大きいので退避先だけ提案 ──
    keeps.append((V2T, 'V-2/V-3 の検証証跡(13MB)。当面残す。落ち着いたら Archive/2026-09/V2_V3_検証 へ'))
    return moves, deletes, keeps


def show(moves, deletes, keeps):
    print(f'■ 移動(退避)候補 {len(moves)}件 — apply で実行(ファイル名は変えない)')
    for f, d, why in moves:
        print(f'  {os.path.relpath(f, BASE)}  →  {os.path.relpath(d, BASE)}/   ({why})')
    print(f'\n■ 削除候補 {len(deletes)}件 — **このスクリプトは削除しない**。英樹の承認後に手で消す')
    tot = 0
    for f, why in deletes:
        sz = os.path.getsize(f) if os.path.exists(f) else 0
        tot += sz
        print(f'  {os.path.relpath(f, BASE)}  ({sz/1024:,.0f}KB / {why})')
    print(f'  合計 {tot/1024/1024:.1f}MB')
    print(f'\n■ 残すもの {len(keeps)}件')
    for f, why in keeps:
        print(f'  {os.path.relpath(f, BASE)}  ({why})')


def apply(moves):
    n = 0
    for f, d, why in moves:
        if not os.path.exists(f):
            continue
        os.makedirs(d, exist_ok=True)
        dst = os.path.join(d, os.path.basename(f))
        if os.path.exists(dst):
            print(f'  ⚠️ 既にある(移動しない): {dst}')
            continue
        shutil.move(f, dst)
        n += 1
    print(f'✅ 移動 {n}件(削除は0件)')


if __name__ == '__main__':
    mode = sys.argv[1] if len(sys.argv) > 1 else 'plan'
    moves, deletes, keeps = plan()
    if mode == 'plan':
        show(moves, deletes, keeps)
    elif mode == 'apply':
        apply(moves)
    elif mode == 'delete-candidates':
        show([], deletes, [])
    else:
        print(__doc__)
