# -*- coding: utf-8 -*-
"""aikos_intake.py のテスト(ダミーファイルのみ・NAS には触れない)

NAS の代わりに一時フォルダを「マウント済み」とみなす MountGuard を差し込む。
D(Fail Closed)では mount 表・ismount を異常な値にして、何も書かれないことを確認する。

実行: python3 -m unittest discover -s MomijiStore_OS/06_DocumentIntelligence/tests -v
"""
import csv
import json
import os
import re
import signal
import stat
import subprocess
import sys
import tempfile
import unicodedata
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'Python'))
import aikos_intake as ai  # noqa: E402

SRC = '//test@nas/MomijiStore'


def make_cfg(mp):
    return {
        'nas': {'mount_point': mp, 'fs_type': 'smbfs', 'allowed_sources': [SRC]},
        'targets': {
            'production': {'doc_root_rel': '35_Documents', 'marker_id': 'TEST-PROD'},
            'sandbox': {'doc_root_rel': '35_Documents/99_Sandbox', 'marker_id': 'TEST-SANDBOX'},
        },
        'intake': {'min_age_seconds': 30},
    }


class Base(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp = self._tmp.name
        self.mp = os.path.join(self.tmp, 'MomijiStore')
        os.mkdir(self.mp)
        self.local_logs = os.path.join(self.tmp, 'local_logs')
        self.table = [(SRC, self.mp, 'smbfs')]
        self.mounted = True
        self.guard = ai.MountGuard(self.mp, 'smbfs', [SRC],
                                   mount_table=lambda: self.table, ismount=lambda p: self.mounted)
        self.cfg = make_cfg(self.mp)
        self.dr = ai.build_docroot(self.cfg, 'production', self.guard)
        ai.init_docroot(self.dr)

    def tearDown(self):
        # テスト内で読み取り専用にしたファイルも片付けられるようにする
        for dirpath, _, files in os.walk(self.tmp):
            for f in files:
                os.chmod(os.path.join(dirpath, f), 0o644)
        self._tmp.cleanup()

    def put(self, name, data, age=120):
        path = self.dr.p('00_Inbox', name)
        with open(path, 'wb') as f:
            f.write(data)
        t = os.stat(path).st_mtime - age
        os.utime(path, (t, t))
        return path

    def run_ingest(self, min_age=30, move_delays=(1, 2, 4)):
        return ai.ingest(self.dr, self.local_logs, min_age, move_delays=move_delays)

    def ledger(self):
        return ai.read_ledger(self.dr)

    def log_events(self):
        out = []
        for name in sorted(os.listdir(self.dr.p('90_Logs'))):
            if name.endswith('.jsonl'):
                with open(self.dr.p('90_Logs', name), encoding='utf-8') as f:
                    out += [json.loads(line) for line in f]
        return out

    def local_events(self):
        path = os.path.join(self.local_logs, 'aikos_local_errors.jsonl')
        if not os.path.exists(path):
            return []
        with open(path, encoding='utf-8') as f:
            return [json.loads(line) for line in f]


class TestInit(Base):
    def test_folders_and_marker(self):
        for sub in ai.SUBDIRS:
            self.assertTrue(os.path.isdir(self.dr.p(sub)), sub)
        with open(self.dr.p(ai.MARKER_NAME), encoding='utf-8') as f:
            self.assertEqual(f.read().strip(), 'TEST-PROD')
        self.assertEqual(ai.init_docroot(self.dr), [])  # 2 回目は何も作らない

    def test_existing_unmarked_folder_is_not_taken_over(self):
        os.makedirs(os.path.join(self.mp, 'other', 'x'))
        cfg = make_cfg(self.mp)
        cfg['targets']['production']['doc_root_rel'] = 'other'
        with self.assertRaises(ai.AikosError):
            ai.init_docroot(ai.build_docroot(cfg, 'production', self.guard))

    def test_sandbox_requires_production_marker(self):
        os.remove(self.dr.p(ai.MARKER_NAME))
        with self.assertRaises(ai.MountError):
            ai.init_docroot(ai.build_docroot(self.cfg, 'sandbox', self.guard))


class TestA_NormalIntake(Base):
    def test_archive_ledger_and_inbox_move(self):
        data = b'%PDF-1.4 dummy invoice'
        self.put('invoice_A.pdf', data)
        s = self.run_ingest()
        self.assertEqual(len(s['archived']), 1)
        doc_id = s['archived'][0][0]
        self.assertRegex(doc_id, r'^DOC-\d{8}-000001$')
        row = self.ledger()[0]
        self.assertEqual(row['status'], 'ARCHIVED')
        self.assertEqual(row['sha256'], ai.hashlib.sha256(data).hexdigest())
        self.assertEqual(row['file_size'], str(len(data)))
        with open(self.dr.p(row['archive_path']), 'rb') as f:
            self.assertEqual(f.read(), data)
        self.assertTrue(re.match(r'^10_Archive/\d{4}/\d{4}-\d{2}/' + doc_id + '/invoice_A.pdf$',
                                 row['archive_path']))
        # Inbox からは消えず _imported へ移っている(削除しない)
        self.assertFalse(os.path.exists(self.dr.p('00_Inbox', 'invoice_A.pdf')))
        imported = [f for _, _, fs in os.walk(self.dr.p('00_Inbox', '_imported')) for f in fs]
        self.assertEqual(imported, [f'{doc_id}__invoice_A.pdf'])

    def test_sequential_doc_ids_across_runs(self):
        self.put('a.pdf', b'a')
        self.put('b.pdf', b'b')
        self.run_ingest()
        self.put('c.pdf', b'c')
        self.run_ingest()
        ids = [r['doc_id'] for r in self.ledger()]
        self.assertEqual([i[-6:] for i in ids], ['000001', '000002', '000003'])
        self.assertEqual(len(set(ids)), 3)

    def test_doc_id_skips_archive_folder_missing_from_ledger(self):
        # 台帳記録前に異常終了して DocID フォルダだけ残った状況でも衝突しない
        self.put('a.pdf', b'a')
        self.run_ingest()
        first = self.ledger()[0]['doc_id']
        orphan = first[:-6] + '000005'
        os.mkdir(os.path.join(os.path.dirname(self.dr.p(self.ledger()[0]['archive_path'])), '..', orphan))
        self.put('b.pdf', b'b')
        self.run_ingest()
        self.assertTrue(self.ledger()[1]['doc_id'].endswith('000006'))

    def test_empty_file_fails_and_stays(self):
        self.put('empty.pdf', b'')
        s = self.run_ingest()
        self.assertEqual(len(s['failed']), 1)
        self.assertTrue(os.path.exists(self.dr.p('00_Inbox', 'empty.pdf')))
        self.assertEqual(self.ledger(), [])

    def test_fresh_file_is_skipped(self):
        self.put('fresh.pdf', b'x', age=0)
        s = self.run_ingest(min_age=30)
        self.assertEqual(s['skipped'], ['fresh.pdf'])
        self.assertTrue(os.path.exists(self.dr.p('00_Inbox', 'fresh.pdf')))

    def test_lock_blocks_concurrent_run(self):
        held = ai.RunLock(self.dr.p('90_Logs', ai.LOCK_NAME))
        self.assertTrue(held.acquire({'holder': 'test'}))
        self.put('a.pdf', b'a')
        try:
            with self.assertRaises(ai.AikosError):
                self.run_ingest()
            self.assertTrue(os.path.exists(self.dr.p('00_Inbox', 'a.pdf')))
            self.assertIn('RUN_LOCKED', [e['event'] for e in self.log_events()])
        finally:
            held.release()
        self.assertEqual(len(self.run_ingest()['archived']), 1)  # 解放後は実行できる

    def test_existing_lock_file_alone_does_not_block(self):
        # 常設ファイルの存在ではなく、ロック保持の有無で判定する
        with open(self.dr.p('90_Logs', ai.LOCK_NAME), 'w') as f:
            f.write('old run info')
        self.put('a.pdf', b'a')
        self.assertEqual(len(self.run_ingest()['archived']), 1)
        self.assertTrue(os.path.exists(self.dr.p('90_Logs', ai.LOCK_NAME)))  # 削除されない

    def test_lock_released_when_holder_killed(self):
        lock = self.dr.p('90_Logs', ai.LOCK_NAME)
        holder = subprocess.Popen(
            [sys.executable, '-c',
             'import sys,time; sys.path.insert(0, sys.argv[1]); import aikos_intake as ai;'
             'l = ai.RunLock(sys.argv[2]); print(l.acquire({"holder": "child"}), flush=True); time.sleep(60)',
             str(Path(ai.__file__).parent), lock], stdout=subprocess.PIPE, text=True)
        try:
            self.assertEqual(holder.stdout.readline().strip(), 'True')
            self.assertFalse(ai.RunLock(lock).acquire({}))  # 保持中は取れない
            holder.send_signal(signal.SIGKILL)  # 異常終了(後片付けなし)
            holder.wait(10)
        finally:
            if holder.poll() is None:
                holder.kill()
            holder.stdout.close()
        again = ai.RunLock(lock)
        self.assertTrue(again.acquire({}))  # OS が解放している
        again.release()


class TestB_Duplicate(Base):
    def test_same_content_is_duplicate_candidate(self):
        self.put('invoice.pdf', b'same-bytes')
        self.run_ingest()
        first = self.ledger()[0]['doc_id']
        self.put('invoice.pdf', b'same-bytes')      # 同名で再投入
        self.put('invoice_copy.pdf', b'same-bytes')  # 別名で再投入
        s = self.run_ingest()
        self.assertEqual(len(s['archived']), 0)
        self.assertEqual(sorted(s['duplicates']), [('invoice.pdf', first), ('invoice_copy.pdf', first)])
        # 新しい DocID・台帳行・Archive 原本は作られない
        self.assertEqual([r['doc_id'] for r in self.ledger()], [first])
        archived_files = [f for _, _, fs in os.walk(self.dr.p('10_Archive')) for f in fs]
        self.assertEqual(len(archived_files), 1)
        # ログに既存 DocID との関係が残り、退避ファイルは削除されていない
        dups = [e for e in self.log_events() if e['event'] == 'DUPLICATE_CANDIDATE']
        self.assertEqual(len(dups), 2)
        for e in dups:
            self.assertEqual(e['duplicate_of'], first)
            self.assertNotIn('doc_id', e)
            self.assertTrue(e['intake_event_id'].startswith('INT-'))
            self.assertTrue(os.path.basename(e['moved_to']).startswith(f"{first}__{e['intake_event_id']}__"))
            self.assertTrue(os.path.exists(self.dr.p(e['moved_to'])))
        # 次の正式登録の連番は重複候補で消費されない
        self.put('next.pdf', b'next')
        self.run_ingest()
        self.assertTrue(self.ledger()[1]['doc_id'].endswith('000002'))

    def test_duplicate_within_same_run(self):
        self.put('x1.pdf', b'dup')
        self.put('x2.pdf', b'dup')
        s = self.run_ingest()
        self.assertEqual((len(s['archived']), len(s['duplicates'])), (1, 1))
        self.assertEqual(len(self.ledger()), 1)


class TestInboxMoveOrder(Base):
    """Inbox から移すのは 保存・ハッシュ一致・台帳・ログ がすべて成功した後だけ"""

    def test_ledger_failure_keeps_inbox_file(self):
        f = self.put('a.pdf', b'a')
        orig = ai.append_ledger

        def boom(*a, **k):
            raise OSError('ledger write failed')
        ai.append_ledger = boom
        try:
            s = self.run_ingest()
        finally:
            ai.append_ledger = orig
        self.assertEqual(len(s['failed']), 1)
        self.assertTrue(os.path.exists(f))
        # 原本フォルダは残る(削除しない)が、verify で台帳未記録として検出できる
        ok, problems, pending = ai.verify_archive(self.dr)
        self.assertEqual(problems[0][1], '台帳に記録の無い Archive フォルダ')

    def test_log_failure_keeps_inbox_file_and_no_double_archive(self):
        f = self.put('a.pdf', b'a')
        self.put('a_same.pdf', b'a')
        orig = ai.Logger.event

        def flaky(self_, event, **fields):
            if event == 'INGEST_OK':
                raise OSError('log write failed')
            return orig(self_, event, **fields)
        ai.Logger.event = flaky
        try:
            s = self.run_ingest()
        finally:
            ai.Logger.event = orig
        self.assertTrue(os.path.exists(f))                      # 移動されていない
        self.assertEqual(len(self.ledger()), 1)                 # 台帳は 1 件
        self.assertEqual(len(s['duplicates']), 1)               # 同内容の 2 件目は二重保存しない
        s2 = self.run_ingest()  # 次回は重複候補ではなく「登録済み文書の Inbox 整理」として _imported へ
        doc_id = self.ledger()[0]['doc_id']
        self.assertEqual(s2['cleaned'], [(doc_id, 'a.pdf')])
        self.assertEqual(s2['duplicates'], [])
        self.assertTrue(os.path.exists(self.dr.p('00_Inbox', '_imported', self.ledger()[0]['ingested_at'][:7],
                                                 f'{doc_id}__a.pdf')))

    def test_hash_mismatch_after_write_keeps_inbox_file(self):
        f = self.put('a.pdf', b'a')
        orig = ai.sha256_of
        calls = {'n': 0}

        def bad(path):
            calls['n'] += 1
            return orig(path) if calls['n'] == 1 else '0' * 64  # 保存後の再計算だけ不一致にする
        ai.sha256_of = bad
        try:
            s = self.run_ingest()
        finally:
            ai.sha256_of = orig
        self.assertEqual(len(s['failed']), 1)
        self.assertTrue(os.path.exists(f))
        self.assertEqual(self.ledger(), [])
        self.assertEqual([f for _, _, fs in os.walk(self.dr.p('10_Archive')) for f in fs], [])
        # 一時ファイルは削除されず 85_Staging に残り、正式原本とは区別される
        fail = [e for e in self.log_events() if e['event'] == 'INGEST_FAIL'][0]
        self.assertTrue(fail['staging_path'].startswith('85_Staging/'))
        self.assertTrue(fail['staging_path'].endswith(ai.PARTIAL_SUFFIX))
        self.assertTrue(os.path.isfile(self.dr.p(fail['staging_path'])))
        self.assertEqual(fail['inbox_path'], '00_Inbox/a.pdf')
        ok, problems, pending = ai.verify_archive(self.dr)
        self.assertEqual((ok, problems), (0, []))
        self.assertEqual(pending, [(fail['intake_event_id'], pending[0][1], fail['staging_path'])])
        # 再取り込みは成功し、DocID は 000001(失敗で連番を消費しない)。staging の残骸はそのまま
        s2 = self.run_ingest()
        self.assertTrue(s2['archived'][0][0].endswith('000001'))
        self.assertTrue(os.path.isfile(self.dr.p(fail['staging_path'])))

    def test_no_delete_calls_on_docroot(self):
        # UGOS のごみ箱が削除を拾って #recycle に書くため、成功・重複・失敗・ロックのどの経路でも削除しない
        self.put('a.pdf', b'a')
        self.put('dup.pdf', b'a')
        self.put('empty.pdf', b'')
        orig = (os.remove, os.unlink, os.rmdir)

        def forbidden(*a, **k):
            raise AssertionError(f'削除が呼ばれました: {a}')
        os.remove = os.unlink = os.rmdir = forbidden
        try:
            s = self.run_ingest()
        finally:
            os.remove, os.unlink, os.rmdir = orig
        self.assertEqual((len(s['archived']), len(s['duplicates']), len(s['failed'])), (1, 1, 1))


class TestC_JapaneseName(Base):
    def test_nfd_name_is_recorded_as_nfc(self):
        nfc_name = '請求書_株式会社ガイド_2026年10月.pdf'
        nfd_name = unicodedata.normalize('NFD', nfc_name)
        self.assertNotEqual(nfc_name, nfd_name)
        self.put(nfd_name, 'ダミー'.encode())
        s = self.run_ingest()
        self.assertEqual(len(s['archived']), 1)
        row = self.ledger()[0]
        self.assertEqual(row['original_filename'], nfc_name)
        self.assertTrue(row['archive_path'].endswith('/' + nfc_name))
        self.assertTrue(os.path.isfile(self.dr.p(row['archive_path'])))
        # NFC で再投入しても重複と判定される(名前ではなく中身で照合)
        self.put(nfc_name, 'ダミー'.encode())
        s2 = self.run_ingest()
        self.assertEqual(len(s2['duplicates']), 1)


class TestD_FailClosed(Base):
    def assert_nothing_written(self, inbox_file):
        self.assertTrue(os.path.exists(inbox_file))
        self.assertEqual(self.ledger(), [])
        self.assertEqual([f for _, _, fs in os.walk(self.dr.p('10_Archive')) for f in fs], [])
        self.assertEqual(self.log_events(), [])
        ev = self.local_events()
        self.assertEqual(ev[-1]['event'], 'MOUNT_ERROR')

    def test_not_a_mount_point(self):
        f = self.put('a.pdf', b'a')
        self.mounted = False  # /Volumes/MomijiStore が単なるローカルフォルダになった状態
        with self.assertRaises(ai.MountError):
            self.run_ingest()
        self.assert_nothing_written(f)

    def test_not_in_mount_table(self):
        f = self.put('a.pdf', b'a')
        self.table = []
        with self.assertRaises(ai.MountError):
            self.run_ingest()
        self.assert_nothing_written(f)

    def test_wrong_fs_type(self):
        f = self.put('a.pdf', b'a')
        self.table = [(SRC, self.mp, 'apfs')]
        with self.assertRaises(ai.MountError):
            self.run_ingest()
        self.assert_nothing_written(f)

    def test_wrong_source(self):
        f = self.put('a.pdf', b'a')
        self.table = [('//someone@other-nas/Share', self.mp, 'smbfs')]
        with self.assertRaises(ai.MountError):
            self.run_ingest()
        self.assert_nothing_written(f)

    def test_marker_mismatch(self):
        f = self.put('a.pdf', b'a')
        os.chmod(self.dr.p(ai.MARKER_NAME), 0o644)
        with open(self.dr.p(ai.MARKER_NAME), 'w', encoding='utf-8') as fh:
            fh.write('SOMETHING-ELSE\n')
        with self.assertRaises(ai.MountError):
            self.run_ingest()
        self.assert_nothing_written(f)

    def test_real_guard_with_missing_mount_point(self):
        # 差し替え無しの本物の判定でも、存在しないマウントポイントでは止まる
        guard = ai.MountGuard(os.path.join(self.tmp, 'NoSuchVolume'), 'smbfs', [SRC])
        dr = ai.build_docroot(make_cfg(guard.mount_point), 'production', guard)
        with self.assertRaises(ai.MountError):
            ai.ingest(dr, self.local_logs, 0)
        with self.assertRaises(ai.MountError):
            ai.init_docroot(dr)
        self.assertFalse(os.path.exists(guard.mount_point))  # 代替フォルダを作っていない

    def test_real_guard_rejects_plain_local_folder(self):
        # 本物の ismount・mount 表で、ローカルの通常フォルダは NAS と認めない
        guard = ai.MountGuard(self.mp, 'smbfs', [SRC])
        with self.assertRaises(ai.MountError):
            guard.verify()

    def test_mount_lost_mid_run(self):
        self.put('a.pdf', b'a')
        f2 = self.put('b.pdf', b'b')
        calls = {'n': 0}

        def flaky():
            calls['n'] += 1
            return self.table if calls['n'] <= 2 else []  # 開始時・1件目は OK、2件目の直前で喪失
        self.guard._mount_table = flaky
        with self.assertRaises(ai.MountError):
            self.run_ingest()
        self.assertEqual(len(self.ledger()), 1)  # 1件目だけ保全
        self.assertTrue(os.path.exists(f2))       # 2件目は Inbox に残る
        self.assertEqual(self.local_events()[-1]['stage'], 'per_file')
        # 何も Mac ローカル(マウント外)へ保存していない: tmp 直下にあるのは NAS 役とローカルログだけ
        self.assertEqual(sorted(os.listdir(self.tmp)), ['MomijiStore', 'local_logs'])

    def test_write_outside_docroot_is_refused(self):
        with self.assertRaises(ai.AikosError):
            self.dr._assert_inside(os.path.join(self.mp, 'data', 'products', 'x.xlsx'))
        with self.assertRaises(ai.AikosError):
            self.dr._assert_inside(self.dr.p('..', '20_Backup', 'x'))


class TestE_ArchiveImmutable(Base):
    def test_original_unchanged_by_later_runs(self):
        self.put('a.pdf', b'original')
        self.run_ingest()
        row = self.ledger()[0]
        path = self.dr.p(row['archive_path'])
        before = (ai.sha256_of(path), os.stat(path).st_mtime_ns, os.stat(path).st_ino)
        self.assertFalse(os.stat(path).st_mode & (stat.S_IWUSR | stat.S_IWGRP | stat.S_IWOTH))  # 読み取り専用
        self.put('a.pdf', b'original')   # 重複
        self.put('a2.pdf', b'different')  # 別ファイル(同名別内容でも別 DocID)
        self.put('a.pdf ', b'x')
        self.run_ingest()
        after = (ai.sha256_of(path), os.stat(path).st_mtime_ns, os.stat(path).st_ino)
        self.assertEqual(before, after)
        ok, problems, pending = ai.verify_archive(self.dr)
        self.assertEqual((ok, problems, pending), (3, [], []))

    def test_same_name_different_content_gets_new_doc_id(self):
        self.put('scan.pdf', b'v1')
        self.run_ingest()
        self.put('scan.pdf', b'v2')
        self.run_ingest()
        rows = self.ledger()
        self.assertEqual([r['status'] for r in rows], ['ARCHIVED', 'ARCHIVED'])
        self.assertNotEqual(rows[0]['archive_path'], rows[1]['archive_path'])
        with open(self.dr.p(rows[0]['archive_path']), 'rb') as f:
            self.assertEqual(f.read(), b'v1')

    def test_verify_detects_tampering(self):
        self.put('a.pdf', b'original')
        self.run_ingest()
        path = self.dr.p(self.ledger()[0]['archive_path'])
        os.chmod(path, 0o644)
        with open(path, 'wb') as f:
            f.write(b'tampered')
        ok, problems, pending = ai.verify_archive(self.dr)
        self.assertEqual(ok, 0)
        self.assertIn('ハッシュ不一致', problems[0][1])


class TestF_LedgerAndLogs(Base):
    def test_ledger_columns_and_events(self):
        self.put('ok.pdf', b'1')
        self.put('empty.pdf', b'')
        self.put('fresh.pdf', b'2', age=0)
        self.run_ingest()
        self.put('ok_again.pdf', b'1')
        self.run_ingest()

        with open(ai.ledger_path(self.dr), 'rb') as f:
            self.assertTrue(f.read(3) == b'\xef\xbb\xbf')  # Excel 用 BOM
        with open(ai.ledger_path(self.dr), encoding='utf-8-sig', newline='') as f:
            header = next(csv.reader(f))
        self.assertEqual(tuple(header), ai.LEDGER_COLUMNS)
        rows = self.ledger()
        self.assertEqual([r['status'] for r in rows], ['ARCHIVED'])  # 重複候補は台帳に載らない
        for r in rows:
            for col in ('doc_id', 'ingested_at', 'original_filename', 'sha256', 'run_id', 'file_size',
                        'archive_path', 'intake_event_id'):
                self.assertTrue(r[col], col)

        events = [e['event'] for e in self.log_events()]
        for need in ('RUN_START', 'INGEST_OK', 'INBOX_MOVED', 'INGEST_FAIL', 'SKIPPED', 'DUPLICATE_CANDIDATE', 'RUN_END'):
            self.assertIn(need, events)
        self.assertEqual(len({e['event_id'] for e in self.log_events()}), len(events))  # 全行に一意の ID
        ok_ev = [e for e in self.log_events() if e['event'] == 'INGEST_OK'][0]
        self.assertEqual(ok_ev['intake_event_id'], rows[0]['intake_event_id'])  # 台帳とログを突き合わせられる
        fail = [e for e in self.log_events() if e['event'] == 'INGEST_FAIL'][0]
        self.assertEqual(fail['original_filename'], 'empty.pdf')
        dup = [e for e in self.log_events() if e['event'] == 'DUPLICATE_CANDIDATE'][0]
        self.assertEqual(dup['duplicate_of'], rows[0]['doc_id'])
        self.assertTrue(os.path.exists(self.dr.p('90_Logs', ai.LOCK_NAME)))  # ロックファイルは常設(削除しない)
        probe = ai.RunLock(self.dr.p('90_Logs', ai.LOCK_NAME))
        self.assertTrue(probe.acquire({}))  # ロック自体は解放済み
        probe.release()



class TestMoveRetry(Base):
    """Inbox 整理の移動: SMB の一時的エラーは有限回リトライ。上限到達でも保全は取り消さない"""

    def setUp(self):
        super().setUp()
        self.sleeps = []
        self._orig_sleep, self._orig_rename = ai._sleep, os.rename
        ai._sleep = self.sleeps.append
        self.fail_left = 0

        def flaky_rename(src, dst):
            if '/00_Inbox/' in src and ('_imported' in dst or '_duplicate' in dst) and self.fail_left != 0:
                self.fail_left -= 1
                raise OSError(ai.errno.EDEVERR, 'Device error')
            return self._orig_rename(src, dst)
        os.rename = flaky_rename

    def tearDown(self):
        ai._sleep, os.rename = self._orig_sleep, self._orig_rename
        super().tearDown()

    def events(self, name):
        return [e for e in self.log_events() if e['event'] == name]

    def test_transient_error_recovers_with_backoff(self):
        self.put('a.pdf', b'a')
        self.fail_left = 2
        s = self.run_ingest()
        self.assertEqual(len(s['archived']), 1)
        self.assertEqual(s['move_pending'], [])
        self.assertEqual(self.sleeps, [1, 2])
        self.assertEqual(self.events('INBOX_MOVED')[0]['attempts'], 3)
        self.assertEqual(len(self.events('MOVE_RETRY')), 2)
        self.assertFalse(os.path.exists(self.dr.p('00_Inbox', 'a.pdf')))

    def test_retry_limit_keeps_registration_and_inbox_file(self):
        f = self.put('a.pdf', b'a')
        self.fail_left = -1  # 常に失敗
        s = self.run_ingest(move_delays=(1, 2, 4))
        self.assertEqual(self.sleeps, [1, 2, 4])  # 有限回(4 回試行)で止まる
        doc_id = self.ledger()[0]['doc_id']
        self.assertEqual(s['archived'], [(doc_id, 'a.pdf')])  # 正式登録は取り消さない
        self.assertEqual(s['move_pending'], [('a.pdf', 'archived', doc_id)])
        self.assertTrue(os.path.exists(f))  # 元ファイルは Inbox に残る(削除・上書きなし)
        pend = self.events('INBOX_MOVE_PENDING')[0]
        self.assertTrue(pend['move_pending'])
        self.assertEqual((pend['doc_id'], pend['attempts']), (doc_id, 4))
        self.assertIn('Device error', pend['move_error'])
        ok, problems, pending = ai.verify_archive(self.dr)
        self.assertEqual((ok, problems), (1, []))  # 保全状態は正常
        # 次回(エラー解消後): 同一ハッシュ・同名・_imported 未移動 → 重複候補にせず Inbox 整理
        self.fail_left = 0
        s2 = self.run_ingest()
        self.assertEqual((s2['archived'], s2['duplicates']), ([], []))
        self.assertEqual(s2['cleaned'], [(doc_id, 'a.pdf')])
        self.assertEqual(len(self.ledger()), 1)  # 新しい DocID は発行しない
        self.assertEqual(self.events('INBOX_CLEANUP')[0]['registered_intake_event_id'],
                         self.ledger()[0]['intake_event_id'])
        self.assertFalse(os.path.exists(f))

    def test_duplicate_move_retry_limit(self):
        self.put('a.pdf', b'a')
        self.run_ingest()
        f = self.put('a_copy.pdf', b'a')
        self.fail_left = -1
        s = self.run_ingest(move_delays=(1,))
        doc_id = self.ledger()[0]['doc_id']
        self.assertEqual(s['duplicates'], [('a_copy.pdf', doc_id)])
        self.assertEqual(s['move_pending'], [('a_copy.pdf', 'duplicate', doc_id)])
        self.assertTrue(os.path.exists(f))
        dup = self.events('DUPLICATE_CANDIDATE')[-1]
        self.assertEqual((dup['move_pending'], dup['moved_to'], dup['duplicate_of']), (True, '', doc_id))
        self.fail_left = 0
        s2 = self.run_ingest()  # 次回は別名なので重複候補として退避される
        self.assertEqual(s2['duplicates'], [('a_copy.pdf', doc_id)])
        self.assertEqual(s2['move_pending'], [])
        self.assertEqual(len(self.ledger()), 1)

    def test_non_transient_error_is_not_retried(self):
        self.put('a.pdf', b'a')
        orig = os.rename

        def denied(src, dst):
            if '/00_Inbox/' in src and '_imported' in dst:
                raise PermissionError(ai.errno.EPERM, 'Operation not permitted')
            return orig(src, dst)
        os.rename = denied
        try:
            s = self.run_ingest()
        finally:
            os.rename = orig
        self.assertEqual(self.sleeps, [])
        self.assertEqual(len(s['move_pending']), 1)
        self.assertEqual(self.events('INBOX_MOVE_PENDING')[0]['attempts'], 1)

    def test_resubmitted_after_successful_import_is_duplicate(self):
        # 同名・同内容でも _imported へ移動済みなら「残存」ではなく再投入 → 重複候補
        self.put('a.pdf', b'a')
        self.run_ingest()
        self.put('a.pdf', b'a')
        s = self.run_ingest()
        self.assertEqual(s['cleaned'], [])
        self.assertEqual(len(s['duplicates']), 1)


if __name__ == '__main__':
    unittest.main()
