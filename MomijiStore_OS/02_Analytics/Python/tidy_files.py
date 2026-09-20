# -*- coding: utf-8 -*-
"""tidy_files.py — 成果物・バックアップの整理(移動は承認後・削除は候補提示のみ)

使い方:
    python3 tidy_files.py plan              何をどこへ動かすかを表示する(何もしない)
    python3 tidy_files.py apply             plan の「移動」だけを実行する(削除はしない)
    python3 tidy_files.py delete-candidates 削除候補の一覧を表示する(削除はしない)
    python3 tidy_files.py apply-quarantine  移動に加え、削除候補を 削除候補_<日付>/ へ退避する(削除はしない)

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
    m = re.search(r'(20\d{2})-?(\d{2})-?\d{2}', name) or re.search(r'(20\d{2})-(\d{2})', name)   # 2026-08_… も拾う
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


# ──────────────────────────────────────────────────────────────
# 2回目(2026-09-21 英樹「ファイルが溜まりすぎてるから整理して」)。8月管理締め後の状態を前提にする。
#   ・移動は Archive/<年月>/<種類>/ への退避だけ(名前不変)。削除はしない(候補として提示)
#   ・スクリプトが参照する基準ファイルは動かさない(verify_v4_copy の BASELINE = Backup/…_20260916_0612時点…xlsm)
#   ・検証用フォルダ(V2_test/V3_stock_test/V4_test/Summary_test)は丸ごと Archive/2026-09/検証/ へ。
#     そのうち「複製の複製」(Summary_test/*/Backup の再複製前・更新前)は削除候補
#   ・楽天94件の削除記録CSV(V4_test/Import_2回目)は証跡なので Archive/2026-09/証跡 へ単独で退避
# ──────────────────────────────────────────────────────────────
KEEP_AS_IS = {
    BK + '/楽天在庫金額集計ツール_v1.0_backup_20260916_0612時点(既にV3差し替え済み).xlsm',   # verify_v4_copy.py の BASELINE
}


def plan2():
    moves, deletes, keeps = [], [], []
    A9 = f'{OUT}/Archive/2026-09'
    # Output: 確認リストは最新だけ、報告は最新(Ver.3)だけ、追加案CSVは最新だけ、証跡CSVは証跡へ
    lists = sorted(glob.glob(OUT + '/英樹への確認リスト_*.xlsx'))
    for f in lists[:-1]:
        moves.append((f, f'{A9}/確認リスト', '古い日付の確認リスト(回答は最新へ引き継ぎ済み)'))
    for f in sorted(glob.glob(OUT + '/Backup/英樹への確認リスト_*旧版*.xlsx')):
        moves.append((f, f'{A9}/確認リスト', '旧版(回答は引き継ぎ済み)。証跡として退避'))
    for f in sorted(glob.glob(OUT + '/Backup/英樹への確認リスト_*_backup_*.xlsx')):
        deletes.append((f, '再生成前の一時バックアップ(回答は最新リスト/証跡へ引き継ぎ済み)'))
    for f in sorted(glob.glob(OUT + '/20*.md')):
        if 'Ver3' not in os.path.basename(f):
            moves.append((f, f'{A9}/報告', '8月締め前の作業報告。Ver.3 に集約済み'))
    for f in _latest_keep(OUT + '/原価確認記録_追加案_*.csv', 1):
        moves.append((f, f'{A9}/生成物', '内部処理用の日次生成物。最新だけ残す'))
    for name, why in (('整理案_20260916.txt', '1回目の整理案(実施済み)'), ('商品番号_訂正取り下げ_20260916.csv', '証跡'),
                      ('8月最終確認_反映記録_20260920.csv', '証跡(8月締めの反映記録)')):
        f = f'{OUT}/{name}'
        if os.path.exists(f):
            moves.append((f, f'{A9}/証跡', why))
    # SourceData/Backup: 商品マスターは最新2本、在庫ツールは 0916基準(固定)・0920 の2本を残す
    pm = sorted(glob.glob(BK + '/商品マスター_単品_v1.0_backup_*.xlsx'))
    for f in pm[:-2]:
        moves.append((f, f'{BK}/Archive/2026-09', '商品マスターの古い作業前バックアップ(最新2本は残す)'))
    for f in sorted(glob.glob(BK + '/楽天在庫金額集計ツール_v1.0_*.xlsm')):
        if f in KEEP_AS_IS or '_backup_20260920_' in f:
            keeps.append((f, '在庫ツール: 検証の基準(0916)と本番反映前(0920)。残す'))
        else:
            moves.append((f, f'{BK}/Archive/2026-09', '在庫ツールの V-2/V-3 作業中バックアップ。証跡として退避'))
    for f in sorted(glob.glob(BK + '/*_backup_20260920_*')) + sorted(glob.glob(BK + '/実行記録_*.txt')):
        keeps.append((f, '8月締め(9/20)の本番反映前バックアップ・実行記録。残す'))
    # 02_Analytics/SourceData/Backup: 同じファイルは最新2本
    groups = {}
    for f in sorted(glob.glob(BK2 + '/*.xlsx')):
        groups.setdefault(re.sub(r'_backup_.*', '', os.path.basename(f)), []).append(f)
    for key, files in groups.items():
        for f in files[:-2]:
            moves.append((f, f'{BK2}/Archive/2026-09', f'{key} の古いバックアップ(最新2本は残す)'))
    # 検証用フォルダ: 丸ごと退避(名前不変)。Summary_test の複製の複製は削除候補
    SDIR = os.path.dirname(BK)
    for name, why in (('V2_test', 'V-2/V-3 検証の証跡(13MB)。8月締め完了につき退避'),
                      ('V3_stock_test', '9/16資料の最初のコピー検証(7MB)。本番反映済み'),
                      ('V4_test', 'V-4 検証(6MB)。本番反映済み。楽天94件の削除記録は別途 証跡へ'),
                      ('Summary_test', '在庫サマリーのコピー検証。本番実行済み(make-copy で再作成できる)')):
        d = f'{SDIR}/{name}'
        if os.path.isdir(d):
            moves.append((d, f'{OUT}/Archive/2026-09/検証', why))
    for f in sorted(glob.glob(f'{SDIR}/Summary_test/Backup/*')) + sorted(glob.glob(f'{SDIR}/Summary_test/Import/Backup/*')):
        deletes.append((f, 'テスト用複製の複製(再複製前/更新前)。本番のバックアップではない'))
    f = f'{SDIR}/V4_test/Import_2回目/楽天在庫リスト_import_削除記録_20260916_133853.csv'
    if os.path.exists(f):
        moves.insert(0, (f, f'{A9}/証跡', '楽天94件(今回のCSVに無い)の変更ログ。フォルダ退避より先に単独で証跡へ'))
    # SourceData 直下の古い成果物、VBA 直下の旧版
    f = f'{SDIR}/照合結果_楽天Amazon_20260709.xlsx'
    if os.path.exists(f):
        moves.append((f, f'{OUT}/Archive/2026-07/生成物', '7月の照合結果。現行の在庫リストに置き換わっている'))
    VBA = os.path.dirname(SDIR) + '/VBA'
    for name, why in (('Module_Rakuten_Tool.bas.backup_20260803_pre_dupkey', '8/3の旧VBA'), ('在庫管理システム_v1.0_VBAコード.bas.frozen_20260907_列不整合', '凍結した旧VBA')):
        f = f'{VBA}/{name}'
        if os.path.exists(f):
            moves.append((f, f'{VBA}/Backup', 'VBA直下から Backup へ(名前不変)'))
    # 残すもの(明示)
    keeps.append((f'{SDIR}/import', 'RMS/Amazon の生CSV(9/16取得)。原資料なので残す'))
    keeps.append((f'{SDIR}/共有在庫_例外一覧.csv', '承認済み共有在庫の例外(運用中)'))
    keeps.append((f'{OUT}/2026-08_経営レポート_Ver3.md', '8月の正式レポート'))
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


def apply(moves, quarantine=None):
    """移動を実行する。削除は**しない**。quarantine に削除候補を渡すと 削除候補_<日付>/ へ退避する
    (ChatGPT 2026-09-16: 今回は削除せず別の退避フォルダへ)。移動元→移動先の対応表を Output/Archive に残す。"""
    import csv
    from datetime import date
    rows, n = [], 0
    # 削除候補の退避を先に行う(検証フォルダごと退避する前に、その中の削除候補を取り出すため)
    jobs = []
    for f, why in (quarantine or []):
        base = OUT if f.startswith(OUT) else (BK2 if f.startswith(BK2) else BK)
        jobs.append((f, f'{base}/Archive/削除候補_{date.today():%Y%m%d}', why, '削除候補(消していない)'))
    jobs += [(f, d, why, '退避') for f, d, why in moves]
    for f, d, why, kind in jobs:
        if not os.path.exists(f):
            rows.append((kind, f, d, '元が無い', why)); continue
        os.makedirs(d, exist_ok=True)
        dst = os.path.join(d, os.path.basename(f))
        if os.path.exists(dst):
            rows.append((kind, f, dst, '同名があるため移動しない', why))
            print(f'  ⚠️ 同名があるため移動しない: {dst}')
            continue
        shutil.move(f, dst)
        rows.append((kind, f, dst, '移動', why))
        n += 1
    os.makedirs(OUT + '/Archive', exist_ok=True)
    log = f'{OUT}/Archive/移動対応表_{date.today():%Y%m%d}.csv'
    new = not os.path.exists(log)
    with open(log, 'a', encoding='utf-8-sig', newline='') as fh:
        w = csv.writer(fh)
        if new:
            w.writerow(['区分', '移動元', '移動先', '結果', '理由'])
        w.writerows(rows)
    print(f'✅ 移動 {n}件(削除は0件) → 対応表 {log}')


if __name__ == '__main__':
    mode = sys.argv[1] if len(sys.argv) > 1 else 'plan'
    round2 = '--round2' in sys.argv
    moves, deletes, keeps = plan2() if round2 else plan()
    if mode == 'plan':
        show(moves, deletes, keeps)
    elif mode == 'apply':
        apply(moves)
    elif mode == 'apply-quarantine':
        apply(moves, quarantine=deletes)
    elif mode == 'delete-candidates':
        show([], deletes, [])
    else:
        print(__doc__)
