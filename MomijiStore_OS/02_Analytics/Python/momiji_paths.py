# -*- coding: utf-8 -*-
"""momiji_paths.py — MomijiStore_OS の保存場所を決める共通モジュール

WHY: 以前は各スクリプトが OS の保存場所(ユーザー名・Desktop を含む絶対パス)を直書きしており、
     フォルダーを移す(Desktop → ~/MomijiStore、Mac の買い替え等)と全スクリプトが壊れた。
     ここで「このファイル自身の位置」から OS ルートを求め、保存場所にもカレントディレクトリにも
     依存しないようにする。

方針:
  ・持つのはディレクトリ構造だけ。具体的なファイル名(商品マスター等)は各スクリプトに置く
    (ここを巨大な設定ファイルにしない)。
  ・通常運用では環境変数は不要。MOMIJI_OS_ROOT はテスト・一時コピー・特殊な実行環境用の上書き。
  ・ルートと判定した場所に MomijiStore_OS 固有のフォルダーが無ければ、処理を続けずに止める
    (誤った場所のファイルを読み書きする事故を防ぐ)。
  ・既存コードは BASE + '/...' の文字列連結なので、値は str で提供する。
"""
import os
import tempfile
from pathlib import Path

# このファイルは MomijiStore_OS/02_Analytics/Python/momiji_paths.py にある → 2つ上が OS ルート
_DEFAULT_ROOT = Path(__file__).resolve().parents[2]

# OS ルートである目印。1つでも無ければ想定外の場所とみなす
_MARKERS = ('01_InventoryManagement', '02_Analytics')


class MomijiPathError(RuntimeError):
    """OS ルートを正しく決められないときのエラー(処理を続けさせない)"""


def _resolve_root():
    override = os.environ.get('MOMIJI_OS_ROOT')
    root = Path(override).expanduser().resolve() if override else _DEFAULT_ROOT
    missing = [m for m in _MARKERS if not (root / m).is_dir()]
    if missing:
        how = f'環境変数 MOMIJI_OS_ROOT={override}' if override else f'momiji_paths.py の位置({__file__})'
        raise MomijiPathError(
            f'MomijiStore_OS のルートとして不正な場所です: {root}\n'
            f'  判定元: {how}\n'
            f'  見つからないフォルダー: {", ".join(missing)}\n'
            f'  → 誤った場所で処理しないよう停止しました')
    return root


OS_ROOT = _resolve_root()
OS_ROOT_STR = str(OS_ROOT)


def _d(*parts):
    return str(OS_ROOT.joinpath(*parts))


# --- 01 在庫管理 ---
INV_SD = _d('01_InventoryManagement', 'SourceData')
INV_OUTPUT = _d('01_InventoryManagement', 'SourceData', 'Output')
INV_BACKUP = _d('01_InventoryManagement', 'SourceData', 'Backup')

# --- 02 分析 ---
ANA_SD = _d('02_Analytics', 'SourceData')
ANA_PY = _d('02_Analytics', 'Python')


def work_dir(name):
    """一時作業領域のパス(作成はしない)。

    MOMIJI_WORK_DIR があればその下、無ければ Python 標準の一時ディレクトリの下。
    本番データの置き場(OS_ROOT 配下)とは分ける。
    """
    base = os.environ.get('MOMIJI_WORK_DIR') or tempfile.gettempdir()
    return os.path.join(base, name)


if __name__ == '__main__':
    # 確認用: python3 momiji_paths.py で判定結果を表示する
    print('OS_ROOT   :', OS_ROOT_STR)
    print('override  :', os.environ.get('MOMIJI_OS_ROOT') or '(なし・通常運用)')
    for k in ('INV_SD', 'INV_OUTPUT', 'INV_BACKUP', 'ANA_SD', 'ANA_PY'):
        print(f'{k:10}:', globals()[k])
    print('work_dir  :', work_dir('example'))
