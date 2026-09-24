# -*- coding: utf-8 -*-
"""nas_sandbox_e2e.py — 実 NAS の 35_Documents/99_Sandbox だけで取り込みの全項目を確認する

WHY: 単体テストは一時フォルダを NAS に見立てるため、SMB 固有の挙動(NFC/NFD・rename・close の遅延反映・
     ismount・flock)は確認できない。本番 Archive・本番台帳を汚さないよう、実 NAS では Sandbox だけを使う。

・target は sandbox に固定(本番を指定する手段を持たない)
・更新日時を人工的に書き換えない(SMB では close の遅延反映で上書きされるため)。
  投入後は「更新日時とサイズが変わらず、一定時間経過した」ことを実際に確認してから取り込む
・ダミーファイルの中身に実行タグを含め、過去の実行と内容が重ならないようにする
・L   前回までの実行で Sandbox の Inbox に残ったファイルの処理
・A〜F 正常取込・重複・日本語名・Fail Closed・原本不変・台帳/ログ
・G/H 同時実行の拒否・SIGKILL 後のロック解放   I partial の検出   R 移動リトライと上限
・S   SMB の rename エラー(Errno 83)の切り分け(99_Sandbox/_smb_diag/ のみ。取り込み処理は通さない)
・終了コード 0 = 全 PASS / 1 = FAIL あり(FAIL が出たらコミットせず報告する)
"""
import copy
import errno
import hashlib
import json
import os
import signal
import subprocess
import sys
import tempfile
import time
import unicodedata
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'Python'))
import aikos_intake as ai  # noqa: E402

MIN_AGE = 5          # この検証での取り込み見送り秒数(本番設定は 30。判定ロジックは同じ)
STABLE_TIMEOUT = 120
RESULTS = []


def check(label, cond, detail=''):
    RESULTS.append((label, bool(cond), detail))
    print(f"  [{'PASS' if cond else 'FAIL'}] {label}" + (f'  ({detail})' if detail and not cond else ''))
    return cond


def put(dr, name, data):
    dr.verify()  # 書き込み前に必ずマウント・目印を確認
    path = dr.p(ai.INBOX, name)
    dr._assert_inside(path)
    fd = os.open(path, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o644)
    with os.fdopen(fd, 'wb') as f:
        f.write(data)
    return path


def wait_stable(paths, min_age=MIN_AGE + 1):
    """更新日時・サイズが 2 回続けて同じで、かつ min_age 秒以上経過するまで待つ(SMB の close 反映待ち)"""
    deadline = time.time() + STABLE_TIMEOUT
    last = {}
    while time.time() < deadline:
        cur = {p: (os.stat(p).st_mtime, os.stat(p).st_size) for p in paths}
        if cur == last and all(time.time() - m >= min_age for m, _ in cur.values()):
            return True
        last = cur
        time.sleep(1)
    return False


def ingest(dr, local_log, **kw):
    return ai.ingest(dr, local_log, MIN_AGE, **kw)


def archive_files(dr):
    return sorted(os.path.join(dp, f) for dp, _, fs in os.walk(dr.p(ai.ARCHIVE)) for f in fs)


def events_of(dr, run_ids):
    out = []
    logdir = dr.p(ai.LOG_DIR)
    for name in sorted(os.listdir(logdir)):
        if name.endswith('.jsonl'):
            with open(os.path.join(logdir, name), encoding='utf-8') as f:
                out += [e for e in map(json.loads, f) if e.get('run_id') in run_ids]
    return out


def fingerprint(path):
    st = os.stat(path)
    return ai.sha256_of(path), st.st_mtime_ns, st.st_size


def check_duplicate(dr, label, s, name, expect_doc, ledger_n, arch_n, inbox_path):
    """重複候補: 新 DocID なし・新原本なし・既存 DocID 参照・_duplicate_candidates へ移動・ログあり"""
    check(f'{label}: 重複候補として判定', (name, expect_doc) in s['duplicates'], str(s))
    check(f'{label}: 新しい DocID・台帳行を作らない', len(ai.read_ledger(dr)) == ledger_n)
    check(f'{label}: 新しい Archive 原本を作らない', len(archive_files(dr)) == arch_n)
    ev = [e for e in events_of(dr, [s['run_id']]) if e['event'] == 'DUPLICATE_CANDIDATE'
          and e['original_filename'] == name]
    check(f'{label}: ログに既存 DocID を記録', len(ev) == 1 and ev[0]['duplicate_of'] == expect_doc, str(ev))
    moved = ev[0].get('moved_to') if ev else ''
    check(f'{label}: _duplicate_candidates へ移動済み(保留なし)',
          bool(moved) and moved.startswith('00_Inbox/_duplicate_candidates/') and ev[0]['move_pending'] is False
          and os.path.isfile(dr.p(moved)) and not os.path.exists(inbox_path), str(ev))


def main():
    cfg = ai.load_config()
    dr = ai.build_docroot(cfg, 'sandbox')
    if not dr.path.rstrip('/').endswith('/35_Documents/99_Sandbox'):
        print('Sandbox 以外を指しているため中止します:', dr.path)
        return 1
    dr.verify()
    tag = time.strftime('%Y%m%d%H%M%S')
    local_log = tempfile.mkdtemp(prefix='aikos_e2e_locallog_')
    print(f'Sandbox: {dr.path}\n実行タグ: {tag}\n')
    runs = []
    ledger_start = len(ai.read_ledger(dr))

    # ---------------- L. 前回までの残存ファイル
    print('L. 前回までに Inbox へ残ったファイルの処理')
    left = sorted(n for n in os.listdir(dr.p(ai.INBOX)) if not n.startswith(('.', '_')))
    print('   残存:', left)
    wait_stable([dr.p(ai.INBOX, n) for n in left])
    ledger_n, arch_n = len(ai.read_ledger(dr)), len(archive_files(dr))
    s = ingest(dr, local_log)
    runs.append(s['run_id'])
    by_name = {x['original_filename']: x for x in ai.read_ledger(dr)}
    for n in left:
        n_nfc = ai.nfc(n)
        if n_nfc.endswith('_empty.pdf'):
            check(f'L {n_nfc}: 0 バイトは失敗扱いで Inbox に残る',
                  any(f == n_nfc for f, _ in s['failed']) and os.path.exists(dr.p(ai.INBOX, n)))
        elif n_nfc.endswith('_partial.pdf'):
            check(f'L {n_nfc}: 前回 partial で失敗した元ファイルを正式登録', any(f == n_nfc for _, f in s['archived']), str(s))
        elif n_nfc in by_name:
            check_duplicate(dr, f'L {n_nfc}', s, n_nfc, by_name[n_nfc]['doc_id'], ledger_n + len(s['archived']),
                            arch_n + len(s['archived']), dr.p(ai.INBOX, n))
        else:
            check(f'L {n_nfc}: 処理された', not os.path.exists(dr.p(ai.INBOX, n)), str(s))
    check('L 移動保留なし', not s['move_pending'], str(s['move_pending']))

    # ---------------- A. 正常取込
    print('A. 正常取込')
    a1, a2 = f'E2E_{tag}_invoice.pdf', f'E2E_{tag}_receipt.png'
    d1, d2 = f'%PDF-1.4 dummy invoice {tag}'.encode(), f'PNG dummy receipt {tag}'.encode()
    check('A 投入ファイルが安定', wait_stable([put(dr, a1, d1), put(dr, a2, d2)]))
    s = ingest(dr, local_log)
    runs.append(s['run_id'])
    got = {n: d for d, n in s['archived']}
    check('A 2 件とも正式登録', set(got) == {a1, a2}, str(s))
    rows = {r['original_filename']: r for r in ai.read_ledger(dr)}
    for name, data in ((a1, d1), (a2, d2)):
        r = rows.get(name, {})
        path = dr.p(r.get('archive_path', 'missing'))
        check(f'A {name}: DocID 形式', ai.DOC_ID_RE.match(r.get('doc_id', '')))
        check(f'A {name}: 原本の中身とハッシュが一致', os.path.isfile(path) and open(path, 'rb').read() == data
              and r['sha256'] == hashlib.sha256(data).hexdigest())
        check(f'A {name}: _imported へ移動', not os.path.exists(dr.p(ai.INBOX, name)) and os.path.isfile(
            dr.p(ai.INBOX, ai.IMPORTED_DIR, r['ingested_at'][:7], f"{r['doc_id']}__{name}")))
    check('A DocID が重複しない', len(set(got.values())) == 2)
    check('A 移動保留なし', not s['move_pending'])
    arch_fp = {n: fingerprint(dr.p(rows[n]['archive_path'])) for n in (a1, a2)}

    # ---------------- B. 同一ファイルの再投入
    print('B. 同一ファイルの再投入')
    ledger_n, arch_n = len(ai.read_ledger(dr)), len(archive_files(dr))
    b_same = put(dr, a1, d1)                              # 同名・同内容
    b_copy_name = f'E2E_{tag}_invoice_copy.pdf'
    b_copy = put(dr, b_copy_name, d1)                     # 別名・同内容
    check('B 投入ファイルが安定', wait_stable([b_same, b_copy]))
    s = ingest(dr, local_log)
    runs.append(s['run_id'])
    check('B 新規登録 0 件', not s['archived'], str(s))
    check_duplicate(dr, 'B 同名', s, a1, got[a1], ledger_n, arch_n, b_same)
    check_duplicate(dr, 'B 別名', s, b_copy_name, got[a1], ledger_n, arch_n, b_copy)

    # ---------------- C. 日本語ファイル名(NFD 投入 → NFC 再投入。前回 Errno 83 が出た手順)
    print('C. 日本語ファイル名')
    jp_nfc = f'E2E_{tag}_請求書_株式会社ガイド_2026年10月分.pdf'
    jp_nfd = unicodedata.normalize('NFD', jp_nfc)
    dj = f'日本語ダミー {tag}'.encode()
    check('C NFD 投入ファイルが安定', wait_stable([put(dr, jp_nfd, dj)]))
    s = ingest(dr, local_log)
    runs.append(s['run_id'])
    r = {x['original_filename']: x for x in ai.read_ledger(dr)}.get(jp_nfc)
    check('C 台帳の元ファイル名が NFC', r is not None, str(s))
    if r:
        on_nas = os.listdir(os.path.dirname(dr.p(r['archive_path'])))
        check('C Archive 上のファイル名が NFC で読める', len(on_nas) == 1 and ai.nfc(on_nas[0]) == jp_nfc, repr(on_nas))
        check('C Archive 原本のハッシュ一致', ai.sha256_of(dr.p(r['archive_path'])) == hashlib.sha256(dj).hexdigest())
        check('C _imported へ移動(保留なし)', not s['move_pending'], str(s))
        ledger_n, arch_n = len(ai.read_ledger(dr)), len(archive_files(dr))
        c_nfc = put(dr, jp_nfc, dj)
        check('C NFC 再投入ファイルが安定', wait_stable([c_nfc]))
        s = ingest(dr, local_log)
        runs.append(s['run_id'])
        check_duplicate(dr, 'C NFC 再投入', s, jp_nfc, r['doc_id'], ledger_n, arch_n, c_nfc)
        retries = [e for e in events_of(dr, [s['run_id']]) if e['event'] == 'MOVE_RETRY']
        print(f'       (この移動での MOVE_RETRY: {len(retries)} 回)')

    # ---------------- D. Fail Closed
    print('D. Fail Closed')
    probe = put(dr, f'E2E_{tag}_failclosed_probe.pdf', f'probe {tag}'.encode())
    wait_stable([probe])
    n_ledger, n_events = len(ai.read_ledger(dr)), len(events_of(dr, runs))
    bad = copy.deepcopy(cfg)
    bad['targets']['sandbox']['marker_id'] = 'WRONG-MARKER'
    try:
        ingest(ai.build_docroot(bad, 'sandbox'), local_log)
        check('D-1 目印不一致で停止', False, '停止しなかった')
    except ai.MountError:
        check('D-1 目印不一致で停止', True)
    # SMB が外れて /Volumes/MomijiStore がローカルの通常フォルダになった状況(目印まで揃えても拒否する)
    fake = tempfile.mkdtemp(prefix='aikos_e2e_fake_volume_')
    fake_cfg = copy.deepcopy(cfg)
    fake_cfg['nas']['mount_point'] = fake
    for tgt in ('production', 'sandbox'):
        root = os.path.join(fake, fake_cfg['targets'][tgt]['doc_root_rel'])
        for sub in ai.SUBDIRS:
            os.makedirs(os.path.join(root, sub), exist_ok=True)
        with open(os.path.join(root, ai.MARKER_NAME), 'w', encoding='utf-8') as f:
            f.write(fake_cfg['targets'][tgt]['marker_id'] + '\n')
    with open(os.path.join(fake, fake_cfg['targets']['sandbox']['doc_root_rel'], ai.INBOX, 'local.pdf'), 'wb') as f:
        f.write(b'local')
    time.sleep(MIN_AGE + 1)
    before = sorted(os.path.join(dp, f) for dp, _, fs in os.walk(fake) for f in fs)
    try:
        ingest(ai.build_docroot(fake_cfg, 'sandbox'), local_log)
        check('D-2 マウントでないフォルダでは停止', False, '停止しなかった')
    except ai.MountError as e:
        check('D-2 マウントでないフォルダでは停止', True)
        print('       →', e)
    after = sorted(os.path.join(dp, f) for dp, _, fs in os.walk(fake) for f in fs)
    check('D-2 ローカルへ何も保存していない', before == after)
    bad_src = copy.deepcopy(cfg)
    bad_src['nas']['allowed_sources'] = ['//someone@other-nas/Other']
    try:
        ingest(ai.build_docroot(bad_src, 'sandbox'), local_log)
        check('D-3 接続元不一致で停止', False, '停止しなかった')
    except ai.MountError:
        check('D-3 接続元不一致で停止', True)
    check('D 停止中は NAS の台帳・ログに書いていない',
          len(ai.read_ledger(dr)) == n_ledger and len(events_of(dr, runs)) == n_events)
    check('D プローブは Inbox に残っている', os.path.exists(probe))
    with open(os.path.join(local_log, 'aikos_local_errors.jsonl'), encoding='utf-8') as f:
        local_ev = [json.loads(x)['event'] for x in f]
    check('D MOUNT_ERROR がローカルに 3 件記録', local_ev.count('MOUNT_ERROR') == 3, str(local_ev))
    s = ingest(dr, local_log)
    runs.append(s['run_id'])
    check('D 復旧後の実行でプローブを正式登録', any(n.endswith('failclosed_probe.pdf') for _, n in s['archived']), str(s))

    # ---------------- E. Archive 原本不変
    print('E. Archive 原本不変')
    for n, fp in arch_fp.items():
        check(f'E {n}: ハッシュ・更新日時・サイズが A 直後と同じ', fingerprint(dr.p(rows[n]['archive_path'])) == fp)
    ok, problems, pending = ai.verify_archive(dr)
    check('E verify 問題 0 件', not problems, str(problems))

    # ---------------- F. 台帳・ログ
    print('F. 台帳・ログ')
    empty = put(dr, f'E2E_{tag}_empty.pdf', b'')
    wait_stable([empty])
    s = ingest(dr, local_log)
    runs.append(s['run_id'])
    check('F 0 バイトは失敗扱いで Inbox に残る',
          any(n == os.path.basename(empty) for n, _ in s['failed']) and os.path.exists(empty), str(s))
    ledger = ai.read_ledger(dr)
    with open(ai.ledger_path(dr), encoding='utf-8-sig', newline='') as f:
        header = f.readline().strip().split(',')
    check('F 台帳ヘッダが仕様どおり', tuple(header) == ai.LEDGER_COLUMNS, str(header))
    mine = [x for x in ledger if x['original_filename'].startswith(f'E2E_{tag}_')]
    check('F この実行の正式登録が 4 件(A2・C1・D1)', len(mine) == 4, str(len(mine)))
    check('F 台帳は ARCHIVED のみ・DocID 一意',
          all(x['status'] == 'ARCHIVED' for x in ledger) and len({x['doc_id'] for x in ledger}) == len(ledger))
    evs = events_of(dr, runs)
    kinds = {e['event'] for e in evs}
    for need in ('RUN_START', 'INGEST_OK', 'INBOX_MOVED', 'DUPLICATE_CANDIDATE', 'INGEST_FAIL', 'RUN_END'):
        check(f'F ログに {need}', need in kinds)
    ok_ev = {e['doc_id']: e['intake_event_id'] for e in evs if e['event'] == 'INGEST_OK'}
    check('F 台帳とログを intake_event_id で突き合わせられる',
          all(ok_ev.get(x['doc_id']) == x['intake_event_id'] for x in mine))

    # ---------------- G. 同時実行 / H. SIGKILL 後のロック解放
    lock_path = dr.p(ai.LOG_DIR, ai.LOCK_NAME)
    check('G ロックファイルは常設(削除されていない)', os.path.isfile(lock_path))
    print('G. 同時実行')
    g_file = put(dr, f'E2E_{tag}_concurrent.pdf', f'concurrent {tag}'.encode())
    wait_stable([g_file])
    holder = subprocess.Popen(
        [sys.executable, '-c',
         'import sys,time; sys.path.insert(0, sys.argv[1]); import aikos_intake as ai;'
         'l = ai.RunLock(sys.argv[2]); print(l.acquire({"holder": "e2e-child"}), flush=True); time.sleep(120)',
         str(Path(ai.__file__).parent), lock_path], stdout=subprocess.PIPE, text=True)
    try:
        check('G 別プロセスがロックを取得', holder.stdout.readline().strip() == 'True')
        n_ledger = len(ai.read_ledger(dr))
        try:
            ingest(dr, local_log)
            check('G 保持中の取り込みは拒否', False, '拒否されなかった')
        except ai.MountError:
            check('G 保持中の取り込みは拒否', False, 'MountError')
        except ai.AikosError:
            check('G 保持中の取り込みは拒否', True)
        check('G 拒否時は台帳に書かず投入ファイルは Inbox に残る',
              len(ai.read_ledger(dr)) == n_ledger and os.path.exists(g_file))
        print('H. SIGKILL 後のロック解放')
        holder.send_signal(signal.SIGKILL)
        holder.wait(15)
    finally:
        if holder.poll() is None:
            holder.kill()
        holder.stdout.close()
    s = ingest(dr, local_log)
    runs.append(s['run_id'])
    check('H SIGKILL 後(後片付けなし)に次の取り込みが実行できる',
          any(n.endswith('_concurrent.pdf') for _, n in s['archived']), str(s))
    check('H ロックファイルは残っている(削除不要)', os.path.isfile(lock_path))
    normal = subprocess.run(
        [sys.executable, '-c',
         'import sys; sys.path.insert(0, sys.argv[1]); import aikos_intake as ai;'
         'l = ai.RunLock(sys.argv[2]); print(l.acquire({"holder": "e2e-normal"}))',
         str(Path(ai.__file__).parent), lock_path], capture_output=True, text=True)
    probe_lock = ai.RunLock(lock_path)
    got_lock = probe_lock.acquire({'holder': 'e2e-probe'})
    probe_lock.release()
    check('H 正常終了した別プロセスのロックも解放される', normal.stdout.strip() == 'True' and got_lock)

    # ---------------- I. partial の検出
    print('I. partial 発生時の検出')
    i_name = f'E2E_{tag}_partial.pdf'
    i_file = put(dr, i_name, f'partial {tag}'.encode())
    wait_stable([i_file])
    real_sha = ai.sha256_of

    def fail_on_staged(path):  # 85_Staging に書いた後の再計算だけを不一致にする(書き込み異常の模擬)
        return '0' * 64 if '/85_Staging/' in path else real_sha(path)
    ai.sha256_of = fail_on_staged
    try:
        s = ingest(dr, local_log)
    finally:
        ai.sha256_of = real_sha
    runs.append(s['run_id'])
    fail = [e for e in events_of(dr, [s['run_id']]) if e['event'] == 'INGEST_FAIL' and e['original_filename'] == i_name]
    check('I 失敗として記録', len(fail) == 1, str(s))
    staged = fail[0].get('staging_path', '') if fail else ''
    check('I 一時ファイルは 85_Staging に残る(Archive と区別)',
          staged.startswith('85_Staging/') and staged.endswith(ai.PARTIAL_SUFFIX) and os.path.isfile(dr.p(staged)))
    check('I ログに復旧情報(staging_path・inbox_path・intake_event_id)',
          bool(fail) and fail[0].get('inbox_path') and fail[0].get('intake_event_id'))
    check('I 台帳に載らない・DocID 未発行', all(x['original_filename'] != i_name for x in ai.read_ledger(dr)))
    check('I Archive に一時ファイルが無い', not [f for f in archive_files(dr) if f.endswith(ai.PARTIAL_SUFFIX)])
    check('I 元ファイルは Inbox に残る', os.path.exists(i_file))
    ok, problems, pending = ai.verify_archive(dr)
    check('I verify が未完了/要確認として検出', staged in [p[2] for p in pending]
          and fail[0]['intake_event_id'] in [p[0] for p in pending], str(pending))
    check('I verify の問題(原本の不整合)は 0 件', not problems, str(problems))
    s = ingest(dr, local_log)
    runs.append(s['run_id'])
    check('I Inbox の元ファイルから再取り込みで正式登録できる', any(n == i_name for _, n in s['archived']), str(s))
    check('I 一時ファイルは削除されず残る', os.path.isfile(dr.p(staged)))

    # ---------------- R. 移動リトライ(SMB の一時的エラーを模擬)
    print('R. 移動リトライ')
    real_rename = os.rename
    state = {'fail_left': 0}

    def flaky_rename(src, dst):
        if '/00_Inbox/' in src and ('/_imported/' in dst or '/_duplicate_candidates/' in dst) and state['fail_left']:
            state['fail_left'] -= 1
            raise OSError(errno.EDEVERR, 'Device error (simulated)')
        return real_rename(src, dst)

    def run_flaky(fail_left, delays=(1, 2, 4)):
        state['fail_left'] = fail_left
        os.rename = flaky_rename
        t0 = time.time()
        try:
            return ingest(dr, local_log, move_delays=delays), time.time() - t0
        finally:
            os.rename = real_rename

    r1 = put(dr, f'E2E_{tag}_retry_ok.pdf', f'retry ok {tag}'.encode())
    wait_stable([r1])
    s, took = run_flaky(2)
    runs.append(s['run_id'])
    ev = [e for e in events_of(dr, [s['run_id']]) if e['event'] in ('MOVE_RETRY', 'INBOX_MOVED')]
    check('R-1 2 回失敗 → 3 回目で移動成功(1 秒・2 秒の待機)',
          [e['event'] for e in ev] == ['MOVE_RETRY', 'MOVE_RETRY', 'INBOX_MOVED'] and ev[-1]['attempts'] == 3
          and [e['next_delay_seconds'] for e in ev[:2]] == [1, 2] and took >= 3 and not os.path.exists(r1), str(ev))
    r2_name = f'E2E_{tag}_retry_limit.pdf'
    r2 = put(dr, r2_name, f'retry limit {tag}'.encode())
    wait_stable([r2])
    s, took = run_flaky(-1)
    runs.append(s['run_id'])
    doc = next((d for d, n in s['archived'] if n == r2_name), None)
    pend = [e for e in events_of(dr, [s['run_id']]) if e['event'] == 'INBOX_MOVE_PENDING']
    check('R-2 上限到達: 4 回試行(1・2・4 秒)で停止', len(pend) == 1 and pend[0]['attempts'] == 4 and 7 <= took < 60,
          f'{pend} took={took:.1f}')
    check('R-2 正式登録は取り消さない(台帳・Archive あり)', doc is not None
          and any(x['doc_id'] == doc for x in ai.read_ledger(dr)))
    check('R-2 元ファイルは Inbox に残り move_pending/move_error を記録',
          os.path.exists(r2) and bool(pend) and pend[0]['move_pending'] is True and bool(pend[0]['move_error']))
    ok, problems, _ = ai.verify_archive(dr)
    check('R-2 保全状態は正常(verify 問題 0 件)', not problems, str(problems))
    ledger_n, arch_n = len(ai.read_ledger(dr)), len(archive_files(dr))
    s = ingest(dr, local_log)
    runs.append(s['run_id'])
    check('R-3 次回: 重複候補ではなく Inbox 整理として _imported へ', s['cleaned'] == [(doc, r2_name)]
          and not s['duplicates'] and not os.path.exists(r2), str(s))
    check('R-3 新しい DocID・原本を作らない', len(ai.read_ledger(dr)) == ledger_n and len(archive_files(dr)) == arch_n)
    r4_name = f'E2E_{tag}_retry_ok_dup.pdf'
    r4 = put(dr, r4_name, f'retry ok {tag}'.encode())  # R-1 と同内容の別名 → 重複候補
    wait_stable([r4])
    s, _ = run_flaky(-1, delays=(1,))
    runs.append(s['run_id'])
    check('R-4 重複候補の移動が上限到達 → Inbox に残し move_pending', os.path.exists(r4) and bool(s['move_pending'])
          and s['move_pending'][0][:2] == (r4_name, 'duplicate'), str(s))
    s = ingest(dr, local_log)
    runs.append(s['run_id'])
    check('R-4 次回の実行で _duplicate_candidates へ移動', not os.path.exists(r4) and not s['move_pending'], str(s))

    # ---------------- S. SMB Device error の切り分け(取り込み処理は通さない・削除しない)
    print('S. SMB rename エラーの切り分け(99_Sandbox/_smb_diag のみ)')
    diag = dr.p('_smb_diag', tag)
    dr.verify()
    dr._assert_inside(diag)
    os.makedirs(os.path.join(diag, 'moved'))

    def write(name, data, set_mtime=False):
        p = os.path.join(diag, name)
        dr._assert_inside(p)
        try:
            fd = os.open(p, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o644)
        except FileExistsError:
            return None  # 先行ファイルの移動が失敗して同名(正規化後)が残っている
        with os.fdopen(fd, 'wb') as f:
            f.write(data)
        if set_mtime:  # Finder のコピーは元の更新日時を設定し直す(close 直後の setattr)
            t = time.time() - 3600
            os.utime(p, (t, t))
        with open(p, 'rb') as f:  # 取り込み処理と同じく、移動前に中身を読む(ハッシュ計算相当)
            f.read()
        return p

    def move(p, dest_name):
        if p is None:
            return 'exists'
        try:
            real_rename(p, os.path.join(diag, 'moved', dest_name))
            return 'OK'
        except OSError as e:
            return f'errno{e.errno}'

    diag_rows = []
    for i in range(3):
        jp = f'D{i}_請求書_株式会社ガイド.pdf'
        # A: NFD → NFC(同名・直後)
        a_1 = move(write(unicodedata.normalize('NFD', 'A' + jp), b'a1'), f'A{i}_1.pdf')
        a_2 = move(write(unicodedata.normalize('NFC', 'A' + jp), b'a2'), f'A{i}_2.pdf')
        # B: NFC → NFD(同名・直後)
        b_1 = move(write(unicodedata.normalize('NFC', 'B' + jp), b'b1'), f'B{i}_1.pdf')
        b_2 = move(write(unicodedata.normalize('NFD', 'B' + jp), b'b2'), f'B{i}_2.pdf')
        # C: 書き込み直後の移動(同名の先行ファイルなし)
        c_ascii = move(write(f'C{i}_ascii.pdf', b'c'), f'C{i}_ascii.pdf')
        c_nfc = move(write(unicodedata.normalize('NFC', 'C' + jp), b'c'), f'C{i}_nfc.pdf')
        c_nfd = move(write(unicodedata.normalize('NFD', 'CD' + jp), b'c'), f'C{i}_nfd.pdf')
        c_utime = move(write(f'C{i}_utime.pdf', b'c', set_mtime=True), f'C{i}_utime.pdf')
        # D: 十分に安定した後の移動(NFD → NFC 同名を含む)
        p_ascii = write(f'D{i}_ascii.pdf', b'd')
        p_nfd = write(unicodedata.normalize('NFD', 'DD' + jp), b'd1')
        p_utime = write(f'D{i}_utime.pdf', b'd', set_mtime=True)
        wait_stable([x for x in (p_ascii, p_nfd, p_utime) if x], min_age=6)
        time.sleep(6)
        d_ascii = move(p_ascii, f'D{i}_ascii.pdf')
        d_utime = move(p_utime, f'D{i}_utime.pdf')
        d_1 = move(p_nfd, f'DD{i}_1.pdf')
        p_nfc = write(unicodedata.normalize('NFC', 'DD' + jp), b'd2')
        if p_nfc:
            wait_stable([p_nfc], min_age=6)
        d_2 = move(p_nfc, f'DD{i}_2.pdf')
        diag_rows.append((i, a_1, a_2, b_1, b_2, c_ascii, c_nfc, c_nfd, c_utime, d_ascii, d_utime, d_1, d_2))
    print('   回 | A:NFD→ | A:→NFC | B:NFC→ | B:→NFD | C:ascii | C:NFC | C:NFD | C:utime | D:ascii | D:utime | D:NFD→ | D:→NFC')
    for row in diag_rows:
        print('   ' + ' | '.join(str(x) for x in row))
    with open(os.path.join(local_log, 'smb_diag.json'), 'w', encoding='utf-8') as f:
        json.dump(diag_rows, f, ensure_ascii=False)

    # ---------------- 最終確認
    ok, problems, pending = ai.verify_archive(dr)
    print(f'\n最終 verify(Sandbox): 原本 OK {ok} / 問題 {len(problems)} / 未完了・要確認 {len(pending)}')
    for p_ in pending:
        print('   要確認', *p_)
    check('最終 verify 問題 0 件', not problems, str(problems))
    check('最終 台帳行数の増加 = この実行の正式登録数(L 含む)',
          len(ai.read_ledger(dr)) - ledger_start == sum(
              1 for e in events_of(dr, runs) if e['event'] == 'INGEST_OK'))

    n_fail = sum(1 for _, ok_, _ in RESULTS if not ok_)
    print(f'\n結果: {len(RESULTS) - n_fail} PASS / {n_fail} FAIL')
    return 1 if n_fail else 0


if __name__ == '__main__':
    sys.exit(main())
