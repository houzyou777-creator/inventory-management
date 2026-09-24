# -*- coding: utf-8 -*-
"""aikos_intake.py — AIKOS Document Intelligence Ver.1 Phase 1: 原本の取り込みと保全

目的: 「NAS の 00_Inbox に入れた資料を、原本のまま安全に 10_Archive へ保全する」だけを行う。
      OCR・分類・Money Forward 連携などは後続 Phase。ここでは足場(DocID・ハッシュ・台帳・ログ)を作る。

使い方:
  python3 aikos_intake.py check  --target production
  python3 aikos_intake.py init   --target production --confirm 本番
  python3 aikos_intake.py ingest --target production
  python3 aikos_intake.py verify --target production
  (--target sandbox は 35_Documents/99_Sandbox で同じ動作を試す検証用)

安全設計(WHY):
  ・Fail Closed: SMB マウントが外れると /Volumes/MomijiStore は Mac 内蔵ディスクの空フォルダになる。
    ディレクトリの存在ではなく「mount 表で smbfs・想定の NAS 共有であること」と
    「doc root の目印ファイルの中身が一致すること」を書き込み直前に毎回確認し、
    1つでも崩れたら何も書かずに停止する。Mac ローカルへの代替保存は一切しない。
  ・書き込み先は doc root(35_Documents 配下)に限定し、全書き込みを _assert_inside で検査する。
  ・原本は上書き・編集・削除しない。10_Archive へは新しい DocID フォルダを mkdir(既存なら失敗)で
    確保してから書くので、既存の原本を上書きする経路が無い。
  ・Inbox のファイルも削除しない。取り込み後は 00_Inbox/_imported/ へ移すだけ(同じ doc root 内)。
  ・同時実行は常設ロックファイルへの排他ロック(flock)で 1 本に限定する(DocID 連番と台帳追記の衝突を防ぐ)。
    ロックはプロセス終了時に OS が解放するため、ロックファイルは作成後ずっと削除しない。
  ・NAS 上では一切削除しない。UGOS のごみ箱は SMB 経由の削除を拾って 35_Documents 外の #recycle に
    書き込むため(2026-09-25 実 NAS テストで判明)、削除を伴う後片付けを持たない。
  ・複製は 85_Staging で行い、ハッシュ一致を確認してから 10_Archive へ rename で移す。
    失敗した一時ファイルは 85_Staging に残り、verify が「未完了/要確認」として報告する。
"""
import argparse
import csv
import datetime as dt
import errno
import fcntl
import hashlib
import json
import os
import re
import socket
import subprocess
import sys
import time
import unicodedata
import uuid
from pathlib import Path

HERE = Path(__file__).resolve().parent
MODULE_ROOT = HERE.parent  # 06_DocumentIntelligence
DEFAULT_CONFIG = MODULE_ROOT / 'config' / 'aikos_config.json'
# マウント異常は NAS に書けないため、エラー記録だけはローカルに残す(資料の保存はしない)。Git 管理外
DEFAULT_LOCAL_LOG_DIR = MODULE_ROOT / 'logs'

MARKER_NAME = '.aikos_mount_marker'
SUBDIRS = ('00_Inbox', '10_Archive', '20_Catalogs', '30_Extracted', '40_Exports', '80_Ledger', '85_Staging',
           '90_Logs')
STAGING = '85_Staging'  # 複製途中・失敗した一時ファイルの隔離場所。正式原本ではない
INBOX, ARCHIVE, LEDGER_DIR, LOG_DIR = '00_Inbox', '10_Archive', '80_Ledger', '90_Logs'
IMPORTED_DIR = '_imported'              # 00_Inbox 配下。取り込み済みの投入ファイル置き場
DUPLICATE_DIR = '_duplicate_candidates'  # 00_Inbox 配下。重複候補の置き場(人が判断する)
LEDGER_NAME = 'document_ledger.csv'
LOCK_NAME = '.aikos_ingest.lock'
PARTIAL_SUFFIX = '.aikos_partial'

LEDGER_SCHEMA_VERSION = 1
# 台帳 = AIKOS へ正式登録された文書だけの一覧(1 DocID = 1 行)。将来の DB テーブル documents に対応。
# 重複の再投入などの「出来事」は台帳に載せず、ログ(将来の intake_events テーブル)で intake_event_id により追跡する。
# 列の追加は末尾に限る(既存列の変更禁止)
LEDGER_COLUMNS = (
    'schema_version', 'doc_id', 'status', 'ingested_at', 'original_filename',
    'extension', 'file_size', 'sha256', 'archive_path', 'intake_event_id',
    'run_id', 'doc_type', 'note',
)
STATUS_ARCHIVED = 'ARCHIVED'

DOC_ID_RE = re.compile(r'^DOC-(\d{8})-(\d{6})$')
MOUNT_LINE_RE = re.compile(r'^(?P<src>.+?) on (?P<mp>.+) \((?P<opts>[^)]*)\)$')
CHUNK = 1024 * 1024
# Inbox 整理の移動(rename)で一時的に失敗する SMB のエラー。2026-09-25 実 NAS で EDEVERR(83)を観測
TRANSIENT_MOVE_ERRNOS = {getattr(errno, 'EDEVERR', 83), errno.EBUSY, errno.EIO, errno.EAGAIN, errno.ETIMEDOUT}
DEFAULT_MOVE_RETRY_DELAYS = (1, 2, 4)  # config の intake.move_retry_delays_seconds が無い場合の既定値
_sleep = time.sleep  # テストで差し替える


class AikosError(RuntimeError):
    """処理を続けてはいけないエラー"""


class MountError(AikosError):
    """NAS のマウント・目印の確認に失敗(Fail Closed で停止する)"""


def nfc(s):
    return unicodedata.normalize('NFC', s)


def now_local():
    return dt.datetime.now().astimezone()


def load_config(path=DEFAULT_CONFIG):
    with open(path, encoding='utf-8') as f:
        return json.load(f)


# ---------------------------------------------------------------- マウント確認

def read_mount_table():
    """mount コマンドの出力を (source, mount_point, fs_type) のリストで返す"""
    out = subprocess.run(['/sbin/mount'], capture_output=True, text=True, check=True).stdout
    rows = []
    for line in out.splitlines():
        m = MOUNT_LINE_RE.match(line.strip())
        if m:
            rows.append((m['src'], m['mp'], m['opts'].split(',')[0].strip()))
    return rows


class MountGuard:
    """NAS が本当にマウントされているかを判定する。

    テストでは mount_table / ismount を差し替える(CLI からは差し替えられない)。
    """

    def __init__(self, mount_point, fs_type, allowed_sources, mount_table=read_mount_table, ismount=os.path.ismount):
        self.mount_point = mount_point
        self.fs_type = fs_type
        self.allowed_sources = [s.lower() for s in allowed_sources]
        self._mount_table = mount_table
        self._ismount = ismount

    def verify(self):
        mp = self.mount_point
        if not os.path.isdir(mp):
            raise MountError(f'マウントポイントが存在しません: {mp}')
        if not self._ismount(mp):
            # ディレクトリはあるがマウントではない = Mac 内蔵ディスクの空フォルダの可能性
            raise MountError(f'{mp} はマウントされていません(ローカルの通常フォルダです)')
        try:
            table = self._mount_table()
        except Exception as e:  # mount 表を読めない = 判定不能なので止める
            raise MountError(f'mount 情報を取得できません: {e}')
        hits = [r for r in table if r[1] == mp]
        if not hits:
            raise MountError(f'mount 表に {mp} がありません')
        src, _, fstype = hits[-1]
        if fstype != self.fs_type:
            raise MountError(f'{mp} のファイルシステムが想定外です: {fstype}(想定: {self.fs_type})')
        if src.lower() not in self.allowed_sources:
            raise MountError(f'{mp} のマウント元が想定外です: {src}')
        return src


# ---------------------------------------------------------------- doc root

class DocRoot:
    """35_Documents(または sandbox)を表す。書き込みは必ずここを通す"""

    def __init__(self, guard, doc_root_rel, marker_id, requires=None):
        self.guard = guard
        self.path = os.path.join(guard.mount_point, doc_root_rel)
        self.marker_id = marker_id
        self.requires = requires  # sandbox の場合は production の DocRoot

    def p(self, *parts):
        return os.path.join(self.path, *parts)

    def marker_text(self):
        return f'{self.marker_id}\n'

    def verify(self, need_marker=True):
        """マウント → (必要なら親 doc root) → 目印 の順に確認。どれか失敗で MountError"""
        self.guard.verify()
        if self.requires is not None:
            self.requires.verify(need_marker=True)
        real_root = os.path.realpath(self.path)
        real_mp = os.path.realpath(self.guard.mount_point)
        if not real_root.startswith(real_mp + os.sep):
            raise MountError(f'doc root がマウント外を指しています: {real_root}')
        if need_marker:
            mk = self.p(MARKER_NAME)
            try:
                with open(mk, encoding='utf-8') as f:
                    content = f.read()
            except OSError as e:
                raise MountError(f'目印ファイルを読めません: {mk}({e.strerror})')
            if content.strip() != self.marker_id:
                raise MountError(f'目印ファイルの内容が一致しません: {mk}')

    def _assert_inside(self, path):
        real_root = os.path.realpath(self.path)
        real = os.path.realpath(path)
        if real != real_root and not real.startswith(real_root + os.sep):
            raise AikosError(f'doc root 外への書き込みを拒否しました: {path}')


# ---------------------------------------------------------------- ログ

class Logger:
    """90_Logs への JSON Lines。1 イベント 1 行で追記し、毎回 fsync する"""

    def __init__(self, docroot, run_id):
        self.docroot = docroot
        self.run_id = run_id

    def event(self, event, **fields):
        ts = now_local()
        path = self.docroot.p(LOG_DIR, f'aikos_intake_{ts:%Y%m}.jsonl')
        self.docroot._assert_inside(path)
        rec = {'ts': ts.isoformat(timespec='seconds'), 'event_id': new_event_id('LOG', ts),
               'run_id': self.run_id, 'event': event}
        rec.update(fields)
        _append_line(path, json.dumps(rec, ensure_ascii=False))
        return rec


def new_event_id(kind, ts):
    """ログ 1 行ごとの ID(LOG-…)と、Inbox の 1 ファイルの取り込み試行ごとの ID(INT-…)。
    DocID と違い連番にしない(台帳を読まずに発行でき、衝突は乱数部で避ける)"""
    return f'{kind}-{ts:%Y%m%d-%H%M%S}-{uuid.uuid4().hex[:8]}'


def log_local_error(local_log_dir, event, **fields):
    """NAS に書けない状況(マウント異常)の記録。資料そのものは保存しない"""
    os.makedirs(local_log_dir, exist_ok=True)
    rec = {'ts': now_local().isoformat(timespec='seconds'), 'event': event, 'host': socket.gethostname()}
    rec.update(fields)
    _append_line(os.path.join(local_log_dir, 'aikos_local_errors.jsonl'), json.dumps(rec, ensure_ascii=False))


def _append_line(path, line):
    with open(path, 'a', encoding='utf-8') as f:
        f.write(line + '\n')
        f.flush()
        os.fsync(f.fileno())


# ---------------------------------------------------------------- 台帳

def ledger_path(docroot):
    return docroot.p(LEDGER_DIR, LEDGER_NAME)


def read_ledger(docroot):
    path = ledger_path(docroot)
    if not os.path.exists(path):
        return []
    with open(path, encoding='utf-8-sig', newline='') as f:
        return list(csv.DictReader(f))


def append_ledger(docroot, row):
    path = ledger_path(docroot)
    docroot._assert_inside(path)
    new = not os.path.exists(path)
    # 新規作成時だけ BOM 付き(Excel で文字化けしないため)。追記は BOM なし
    with open(path, 'a', encoding='utf-8-sig' if new else 'utf-8', newline='') as f:
        w = csv.DictWriter(f, fieldnames=LEDGER_COLUMNS)
        if new:
            w.writeheader()
        w.writerow({k: row.get(k, '') for k in LEDGER_COLUMNS})
        f.flush()
        os.fsync(f.fileno())


# ---------------------------------------------------------------- DocID

def next_doc_id(docroot, ledger_rows, today):
    """同日の最大連番 + 1。台帳と 10_Archive の両方を見る(片方が欠けても衝突させない)"""
    ymd = f'{today:%Y%m%d}'
    seqs = [0]
    for r in ledger_rows:
        m = DOC_ID_RE.match(r.get('doc_id', ''))
        if m and m[1] == ymd:
            seqs.append(int(m[2]))
    month_dir = archive_month_dir(docroot, today)
    if os.path.isdir(month_dir):
        for name in os.listdir(month_dir):
            m = DOC_ID_RE.match(name)
            if m and m[1] == ymd:
                seqs.append(int(m[2]))
    seq = max(seqs) + 1
    if seq > 999999:
        raise AikosError(f'{ymd} の DocID 連番が上限に達しました')
    return f'DOC-{ymd}-{seq:06d}'


def archive_month_dir(docroot, day):
    return docroot.p(ARCHIVE, f'{day:%Y}', f'{day:%Y-%m}')


# ---------------------------------------------------------------- 取り込み

def sha256_of(path):
    h = hashlib.sha256()
    with open(path, 'rb') as f:
        for chunk in iter(lambda: f.read(CHUNK), b''):
            h.update(chunk)
    return h.hexdigest()


def list_inbox(docroot, min_age_seconds, now_ts):
    """取り込み対象(Inbox 直下の通常ファイル)と、見送ったもの(理由付き)を返す"""
    targets, skipped = [], []
    inbox = docroot.p(INBOX)
    for name in sorted(os.listdir(inbox)):
        full = os.path.join(inbox, name)
        if name.startswith(('.', '_', '~$')) or name.endswith(PARTIAL_SUFFIX):
            continue  # 隠し・管理用・一時ファイル
        if os.path.islink(full) or not os.path.isfile(full):
            skipped.append((name, 'ファイルではない(フォルダ・リンク等)。Ver.1 は Inbox 直下のファイルのみ対象'))
            continue
        age = now_ts - os.stat(full).st_mtime
        if age < min_age_seconds:
            skipped.append((name, f'投入直後のため見送り(書き込み中の可能性・{int(age)}秒)'))
            continue
        targets.append(name)
    return targets, skipped


class MoveFailed(AikosError):
    """Inbox 整理の移動が規定回数で成功しなかった(元ファイルは Inbox に残っている)"""

    def __init__(self, msg, attempts):
        super().__init__(msg)
        self.attempts = attempts


def _move_into(docroot, src, dest_dir, dest_name, delays, log=None, **log_fields):
    """Inbox 内での移動(rename)。上書きはしない。

    SMB 由来の一時的エラーは delays(秒)の順に有限回だけ再試行する。上限に達したら MoveFailed。
    戻り値: (移動先パス, 試行回数)
    """
    os.makedirs(dest_dir, exist_ok=True)
    dest = os.path.join(dest_dir, dest_name)
    docroot._assert_inside(dest)
    attempts = 0
    while True:
        attempts += 1
        if os.path.lexists(dest):
            raise MoveFailed(f'移動先に同名ファイルがあります(上書きしない): {dest}', attempts)
        try:
            os.rename(src, dest)
            return dest, attempts
        except OSError as e:
            if not os.path.lexists(src) and os.path.lexists(dest):
                return dest, attempts  # エラー応答だが移動は完了していた
            if e.errno not in TRANSIENT_MOVE_ERRNOS:
                raise MoveFailed(f'{type(e).__name__}: {e}', attempts)
            if attempts > len(delays):
                raise MoveFailed(f'{attempts} 回試行して失敗: {type(e).__name__}: {e}', attempts)
            delay = delays[attempts - 1]
            if log is not None:
                log.event('MOVE_RETRY', attempt=attempts, next_delay_seconds=delay,
                          error=f'{type(e).__name__}: {e}', **log_fields)
            _sleep(delay)


class StagingError(AikosError):
    """複製・検証に失敗。一時ファイルは削除せず 85_Staging に残す(staging_path で追跡)"""

    def __init__(self, msg, staging_path):
        super().__init__(msg)
        self.staging_path = staging_path


def _stage_copy(docroot, src, intake_id, orig, expected_sha):
    """原本を 85_Staging へ複製し、書き込み後のハッシュ一致まで確認する。
    失敗しても一時ファイルは削除しない(StagingError で場所を返し、ログと verify で追跡する)"""
    staged = docroot.p(STAGING, f'{intake_id}__{orig}{PARTIAL_SUFFIX}')
    docroot._assert_inside(staged)
    try:
        fd = os.open(staged, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o644)
    except OSError as e:
        raise AikosError(f'一時ファイルを作成できません: {e}')
    try:
        h = hashlib.sha256()
        with os.fdopen(fd, 'wb') as out, open(src, 'rb') as inp:
            for chunk in iter(lambda: inp.read(CHUNK), b''):
                h.update(chunk)
                out.write(chunk)
            out.flush()
            os.fsync(out.fileno())
        if h.hexdigest() != expected_sha:
            raise AikosError('複製中に元ファイルが変化しました')
        if sha256_of(staged) != expected_sha:
            raise AikosError('保存後のハッシュが一致しません(書き込み異常)')
    except Exception as e:
        raise StagingError(str(e), os.path.relpath(staged, docroot.path))
    return staged


def _promote(docroot, staged, doc_dir, final_name, expected_sha):
    """検証済みの一時ファイルを DocID フォルダへ rename で移し、正式原本にする"""
    final = os.path.join(doc_dir, final_name)
    docroot._assert_inside(final)
    rel_staged = os.path.relpath(staged, docroot.path)
    try:
        os.mkdir(doc_dir)  # 既存なら失敗 → 既存 DocID フォルダへ書く経路を作らない
        os.rename(staged, final)  # 同一共有内の rename(削除ではないためごみ箱に入らない)
    except OSError as e:
        raise StagingError(f'Archive への確定に失敗: {e}', rel_staged)
    if sha256_of(final) != expected_sha:
        # 起こりえないはずだが、正式登録はしない。原本は動かさず verify で「台帳に記録の無いフォルダ」として検出される
        raise AikosError(f'確定後のハッシュが一致しません: {final}')
    try:
        os.chmod(final, 0o444)  # 付与するが保全手段とはみなさない(SMB では無視されうる)
    except OSError:
        pass
    return final


class RunLock:
    """常設ロックファイルへの排他ロック。

    WHY: 旧方式(作成→終了時に削除)は、削除が UGOS のごみ箱に拾われ #recycle へ書き込みが発生した。
         ファイルは一度作ったら残し、flock で排他する。ロックはプロセス終了(異常終了含む)で解放される。
    """

    def __init__(self, path):
        self.path = path
        self.fd = None

    def acquire(self, info):
        fd = os.open(self.path, os.O_RDWR | os.O_CREAT, 0o644)
        try:
            fcntl.flock(fd, fcntl.LOCK_EX | fcntl.LOCK_NB)
        except OSError as e:
            os.close(fd)
            if e.errno in (errno.EWOULDBLOCK, errno.EAGAIN, errno.EACCES):
                return False
            raise
        self.fd = fd
        # 保持者の情報を上書き(削除はしない)。調査用で、排他の判定には使わない
        os.ftruncate(fd, 0)
        os.lseek(fd, 0, os.SEEK_SET)
        os.write(fd, json.dumps(info, ensure_ascii=False).encode('utf-8'))
        os.fsync(fd)
        return True

    def release(self):
        if self.fd is not None:
            try:
                fcntl.flock(self.fd, fcntl.LOCK_UN)
            finally:
                os.close(self.fd)
                self.fd = None


def _index_ledger(ledger):
    return {r['sha256']: r for r in ledger if r.get('status') == STATUS_ARCHIVED}


def ingest(docroot, local_log_dir, min_age_seconds, now_fn=now_local, move_delays=DEFAULT_MOVE_RETRY_DELAYS):
    """Inbox を 1 回処理する。戻り値は結果の集計 dict

    「証憑の保全状態」(Archive・台帳)と「Inbox の整理状態」(_imported 等への移動)は分けて扱う。
    保全が完了していれば、Inbox の移動だけが失敗しても正式登録は取り消さない(move_pending)。
    """
    run_id = uuid.uuid4().hex[:12]
    summary = {'run_id': run_id, 'archived': [], 'duplicates': [], 'cleaned': [], 'move_pending': [],
               'failed': [], 'skipped': []}
    try:
        docroot.verify()
    except MountError as e:
        log_local_error(local_log_dir, 'MOUNT_ERROR', stage='start', error=str(e), doc_root=docroot.path)
        raise

    log = Logger(docroot, run_id)
    lock_path = docroot.p(LOG_DIR, LOCK_NAME)
    docroot._assert_inside(lock_path)
    lock = RunLock(lock_path)
    if not lock.acquire({'run_id': run_id, 'pid': os.getpid(), 'host': socket.gethostname(),
                         'started': now_fn().isoformat(timespec='seconds')}):
        log.event('RUN_LOCKED', lock=lock_path)
        raise AikosError(f'他の取り込みが実行中です(ロック取得不可): {lock_path}')

    try:
        log.event('RUN_START', doc_root=docroot.path)
        ledger = read_ledger(docroot)
        known = _index_ledger(ledger)
        targets, skipped = list_inbox(docroot, min_age_seconds, now_fn().timestamp())
        for name, reason in skipped:
            log.event('SKIPPED', original_filename=nfc(name), reason=reason)
            summary['skipped'].append(nfc(name))

        for name in targets:
            src = docroot.p(INBOX, name)
            orig = nfc(name)
            intake_id = new_event_id('INT', now_fn())  # この投入 1 件の試行 ID。関連ログはすべてこれで束ねる
            try:
                docroot.verify()  # 1 件ごとに書き込み直前で再確認(途中でマウントが外れた場合に止める)
            except MountError as e:
                log_local_error(local_log_dir, 'MOUNT_ERROR', stage='per_file', error=str(e),
                                doc_root=docroot.path, original_filename=orig, run_id=run_id,
                                intake_event_id=intake_id)
                raise
            try:
                kind, info, pending = _ingest_one(docroot, log, ledger, known, src, orig, run_id, intake_id,
                                                  now_fn(), move_delays)
            except MountError:
                raise
            except Exception as e:
                fields = {}
                if isinstance(e, StagingError):
                    # 復旧・整理の手掛かり: 一時ファイルの場所・期待ハッシュ・投入元(Inbox に残っている)
                    fields = dict(staging_path=e.staging_path, inbox_path=os.path.relpath(src, docroot.path),
                                  recovery='一時ファイルは正式原本ではない。Inbox の元ファイルを再取り込みすれば足りる')
                log.event('INGEST_FAIL', intake_event_id=intake_id, original_filename=orig,
                          error=f'{type(e).__name__}: {e}', **fields)
                summary['failed'].append((orig, str(e)))
                # 台帳登録後に失敗した可能性があるため、メモリ上の台帳を読み直してから次へ進む
                ledger = read_ledger(docroot)
                known = _index_ledger(ledger)
                continue
            if kind == 'archived':
                ledger.append(info)
                known[info['sha256']] = info
                summary['archived'].append((info['doc_id'], orig))
            elif kind == 'cleanup':
                summary['cleaned'].append((info, orig))
            else:
                summary['duplicates'].append((orig, info))
            if pending:
                summary['move_pending'].append((orig, kind, info if kind != 'archived' else info['doc_id']))
        log.event('RUN_END', archived=len(summary['archived']), duplicates=len(summary['duplicates']),
                  cleaned=len(summary['cleaned']), move_pending=len(summary['move_pending']),
                  failed=len(summary['failed']), skipped=len(summary['skipped']))
    finally:
        lock.release()  # ファイルは削除しない。異常終了時もプロセス終了で OS が解放する
    return summary


def _ingest_one(docroot, log, ledger, known, src, orig, run_id, intake_id, now, move_delays):
    """戻り値: (種類, 情報, Inbox 整理が未完了か)
      ('archived', 台帳行)  新規の正式登録
      ('cleanup', DocID)    正式登録済みだが Inbox 整理だけ未完了だった残存ファイルの整理
      ('duplicate', DocID)  既存文書と同一内容の再投入(重複候補)
    """
    size = os.path.getsize(src)
    if size == 0:
        raise AikosError('空ファイル(0 バイト)のため取り込みません')
    sha = sha256_of(src)

    if sha in known:
        row = known[sha]
        existing = row['doc_id']
        imported_dir = docroot.p(INBOX, IMPORTED_DIR, row['ingested_at'][:7])
        imported_name = f'{existing}__{orig}'
        if row['original_filename'] == orig and not os.path.lexists(os.path.join(imported_dir, imported_name)):
            # 登録時の Inbox 移動が失敗して残ったファイル(同じ名前・同じ中身・_imported に未移動)。
            # 重複候補ではなく、前回の整理の続きとして _imported へ移す
            fields = dict(intake_event_id=intake_id, doc_id=existing, original_filename=orig, sha256=sha,
                          registered_intake_event_id=row.get('intake_event_id', ''))
            try:
                moved, attempts = _move_into(docroot, src, imported_dir, imported_name, move_delays, log, **fields)
                log.event('INBOX_CLEANUP', moved_to=os.path.relpath(moved, docroot.path), attempts=attempts,
                          **fields)
                return 'cleanup', existing, False
            except MoveFailed as e:
                log.event('INBOX_MOVE_PENDING', move_pending=True, move_error=str(e), attempts=e.attempts,
                          note='正式登録済み。Inbox 整理のみ未完了。次回実行で再度整理する', **fields)
                return 'cleanup', existing, True

        # 重複候補: 正式登録済みの文書と同一内容。DocID は発行せず、Archive にも保存しない。
        # 投入ファイルは削除せず、既存 DocID と試行 ID が分かる名前で退避し、ログで関係を追跡する
        fields = dict(intake_event_id=intake_id, duplicate_of=existing, original_filename=orig,
                      sha256=sha, file_size=size)
        try:
            moved, attempts = _move_into(docroot, src, docroot.p(INBOX, DUPLICATE_DIR),
                                         f'{existing}__{intake_id}__{orig}', move_delays, log,
                                         intake_event_id=intake_id, original_filename=orig)
            log.event('DUPLICATE_CANDIDATE', moved_to=os.path.relpath(moved, docroot.path), attempts=attempts,
                      move_pending=False, **fields)
            return 'duplicate', existing, False
        except MoveFailed as e:
            # Inbox に残る。次回また重複候補として検出され、移動を再試行する
            log.event('DUPLICATE_CANDIDATE', moved_to='', move_pending=True, move_error=str(e),
                      attempts=e.attempts, **fields)
            return 'duplicate', existing, True

    # 先に 85_Staging で複製・検証し、成功した場合だけ DocID を確定して Archive へ移す
    # (失敗で DocID フォルダが残ったり連番が欠けたりしないように)
    staged = _stage_copy(docroot, src, intake_id, orig, sha)
    doc_id = next_doc_id(docroot, ledger, now)
    doc_dir = os.path.join(archive_month_dir(docroot, now), doc_id)
    docroot._assert_inside(doc_dir)
    os.makedirs(os.path.dirname(doc_dir), exist_ok=True)
    final = _promote(docroot, staged, doc_dir, orig, sha)
    row = {
        'schema_version': LEDGER_SCHEMA_VERSION, 'doc_id': doc_id, 'status': STATUS_ARCHIVED,
        'ingested_at': now.isoformat(timespec='seconds'), 'original_filename': orig,
        'extension': os.path.splitext(orig)[1].lower().lstrip('.'), 'file_size': size,
        'sha256': sha, 'archive_path': os.path.relpath(final, docroot.path),
        'intake_event_id': intake_id, 'run_id': run_id,
    }
    # 保存 → ハッシュ一致 → 台帳 → ログ がすべて成功した後にだけ Inbox から移す。
    # 途中で失敗したら例外で抜け、投入ファイルは Inbox に残る(台帳登録後なら次回 cleanup で整理される)
    append_ledger(docroot, row)
    log.event('INGEST_OK', intake_event_id=intake_id, doc_id=doc_id, original_filename=orig, sha256=sha,
              file_size=size, archive_path=row['archive_path'])
    # ここから先は Inbox の整理。失敗しても正式登録は取り消さない
    try:
        moved, attempts = _move_into(docroot, src, docroot.p(INBOX, IMPORTED_DIR, f'{now:%Y-%m}'),
                                     f'{doc_id}__{orig}', move_delays, log,
                                     intake_event_id=intake_id, doc_id=doc_id)
        log.event('INBOX_MOVED', intake_event_id=intake_id, doc_id=doc_id,
                  moved_to=os.path.relpath(moved, docroot.path), attempts=attempts)
        return 'archived', row, False
    except MoveFailed as e:
        log.event('INBOX_MOVE_PENDING', intake_event_id=intake_id, doc_id=doc_id, original_filename=orig,
                  move_pending=True, move_error=str(e), attempts=e.attempts,
                  note='正式登録済み。Inbox 整理のみ未完了。次回実行で再度整理する')
        return 'archived', row, True


# ---------------------------------------------------------------- 検証・初期化

def verify_archive(docroot):
    """台帳の ARCHIVED 行を再ハッシュし、原本が変わっていないかを確認する(読み取りのみ)。

    戻り値: (正常な原本数, problems, pending)
      problems = 原本の欠落・改ざん・台帳不整合(要対応)
      pending  = 85_Staging に残った未完了の一時ファイル(正式原本ではない・要確認)
    """
    docroot.verify()
    ledger = read_ledger(docroot)
    problems, ok = [], 0
    listed = set()
    seen_ids, seen_sha = set(), set()
    for r in ledger:
        if r['doc_id'] in seen_ids:
            problems.append((r['doc_id'], '台帳で DocID が重複しています', r.get('archive_path', '')))
        if r.get('status') == STATUS_ARCHIVED and r['sha256'] in seen_sha:
            problems.append((r['doc_id'], '同一ハッシュの文書が台帳に複数あります', r.get('archive_path', '')))
        seen_ids.add(r['doc_id'])
        seen_sha.add(r['sha256'])
        if r.get('status') != STATUS_ARCHIVED:
            continue
        path = docroot.p(r['archive_path'])
        listed.add(os.path.dirname(r['archive_path']))
        if not os.path.isfile(path):
            problems.append((r['doc_id'], '原本が見つかりません', r['archive_path']))
        elif sha256_of(path) != r['sha256']:
            problems.append((r['doc_id'], 'ハッシュ不一致(原本が変更された可能性)', r['archive_path']))
        else:
            ok += 1
    # 台帳に無い DocID フォルダ(台帳記録前に異常終了した等)を検出する
    arch = docroot.p(ARCHIVE)
    for dirpath, dirnames, _ in os.walk(arch):
        for d in dirnames:
            if DOC_ID_RE.match(d):
                rel = os.path.relpath(os.path.join(dirpath, d), docroot.path)
                if rel not in listed:
                    problems.append((d, '台帳に記録の無い Archive フォルダ', rel))
    for dirpath, _, files in os.walk(arch):
        for f in files:
            if f.endswith(PARTIAL_SUFFIX):
                rel = os.path.relpath(os.path.join(dirpath, f), docroot.path)
                problems.append(('-', 'Archive 内に一時ファイルがあります', rel))
    pending = []
    staging = docroot.p(STAGING)
    if os.path.isdir(staging):
        for f in sorted(os.listdir(staging)):
            if f.startswith('.'):
                continue
            intake_id = f.split('__', 1)[0] if '__' in f else ''
            pending.append((intake_id, '未完了の一時ファイル(正式原本ではない・要確認)', os.path.join(STAGING, f)))
    return ok, problems, pending


def init_docroot(docroot):
    """doc root・サブフォルダ・目印を作る。既存のものには触らない"""
    docroot.verify(need_marker=False)
    created = []
    marker = docroot.p(MARKER_NAME)
    if os.path.isdir(docroot.path):
        if os.path.exists(marker):
            docroot.verify(need_marker=True)  # 目印が別物なら止める
        elif os.listdir(docroot.path):
            raise AikosError(f'目印の無い既存フォルダに中身があります。想定外のため停止します: {docroot.path}')
    else:
        os.makedirs(docroot.path)
        created.append(docroot.path)
    if not os.path.exists(marker):
        docroot._assert_inside(marker)
        fd = os.open(marker, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o644)
        with os.fdopen(fd, 'w', encoding='utf-8') as f:
            f.write(docroot.marker_text())
        created.append(marker)
    for sub in SUBDIRS:
        path = docroot.p(sub)
        if not os.path.isdir(path):
            docroot._assert_inside(path)
            os.mkdir(path)
            created.append(path)
    docroot.verify(need_marker=True)
    return created


def build_docroot(cfg, target, guard=None):
    nas = cfg['nas']
    guard = guard or MountGuard(nas['mount_point'], nas['fs_type'], nas['allowed_sources'])
    t = cfg['targets'][target]
    requires = None
    if target != 'production':
        p = cfg['targets']['production']
        requires = DocRoot(guard, p['doc_root_rel'], p['marker_id'])
    return DocRoot(guard, t['doc_root_rel'], t['marker_id'], requires)


# ---------------------------------------------------------------- CLI

def main(argv=None):
    ap = argparse.ArgumentParser(description='AIKOS Document Intelligence Phase 1 取り込み')
    ap.add_argument('command', choices=('check', 'init', 'ingest', 'verify'))
    ap.add_argument('--target', required=True, choices=('production', 'sandbox'))
    ap.add_argument('--confirm', default='')
    args = ap.parse_args(argv)

    cfg = load_config()
    docroot = build_docroot(cfg, args.target)
    try:
        if args.command == 'check':
            docroot.verify()
            print(f'OK: NAS マウント・目印を確認しました → {docroot.path}')
        elif args.command == 'init':
            if args.target == 'production' and args.confirm != '本番':
                print('本番の初期化には --confirm 本番 が必要です', file=sys.stderr)
                return 2
            created = init_docroot(docroot)
            print('作成:' if created else '作成なし(既に初期化済み)')
            for c in created:
                print('  ', c)
        elif args.command == 'ingest':
            delays = tuple(cfg['intake'].get('move_retry_delays_seconds', DEFAULT_MOVE_RETRY_DELAYS))
            s = ingest(docroot, DEFAULT_LOCAL_LOG_DIR, cfg['intake']['min_age_seconds'], move_delays=delays)
            print(f"run_id={s['run_id']}  保全 {len(s['archived'])} / 重複候補 {len(s['duplicates'])}"
                  f" / Inbox整理 {len(s['cleaned'])} / 移動保留 {len(s['move_pending'])}"
                  f" / 失敗 {len(s['failed'])} / 見送り {len(s['skipped'])}")
            for d, n in s['archived']:
                print(f'  保全     {d}  {n}')
            for n, dup in s['duplicates']:
                print(f'  重複候補 {n}  (= 既存 {dup}。DocID 発行なし)')
            for d, n in s['cleaned']:
                print(f'  Inbox整理 {d}  {n}  (登録済み文書の残存ファイルを _imported へ)')
            for n, kind, d in s['move_pending']:
                print(f'  移動保留 {n}  ({kind} / {d}。Inbox に残し次回再試行)')
            for n, e in s['failed']:
                print(f'  失敗     {n}  {e}')
            return 1 if s['failed'] else 0
        elif args.command == 'verify':
            ok, problems, pending = verify_archive(docroot)
            print(f'原本 OK {ok} 件 / 問題 {len(problems)} 件 / 未完了・要確認 {len(pending)} 件')
            for p in problems:
                print('  NG    ', *p)
            for p in pending:
                print('  要確認', *p)
            return 1 if problems else (4 if pending else 0)
    except MountError as e:
        if args.command != 'ingest':  # ingest は内部で記録済み
            log_local_error(DEFAULT_LOCAL_LOG_DIR, 'MOUNT_ERROR', stage=args.command, error=str(e),
                            doc_root=docroot.path)
        print(f'停止(Fail Closed): {e}', file=sys.stderr)
        return 3
    except AikosError as e:
        print(f'停止: {e}', file=sys.stderr)
        return 2
    return 0


if __name__ == '__main__':
    sys.exit(main())
