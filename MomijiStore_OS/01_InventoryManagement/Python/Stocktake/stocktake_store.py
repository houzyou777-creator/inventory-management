# -*- coding: utf-8 -*-
"""stocktake_store.py — 棚卸しスキャン記録の保存(追記のみ・1件ごとに永続化)

設計(2026-09-25 承認):
    ・生データは消さない・書き換えない。訂正は「取消イベント」を追記して表す
    ・1件登録するたびに flush + fsync してから応答する(画面に「登録済み」と出た行は電源断でも残る)
    ・書き込むプロセスは1つだけ(ファイルロック)。複数の端末(Mac・iPad)は同じサーバーへ送る
    ・照合は jan_lookup の規則で行い、記録時の照合結果は「その時点の控え」として残す
      (集計は記録ファイル + 商品マスターから毎回照合し直す。未登録JANは後から自動で解決される)
    ・処理区分(棚卸/将来の入荷・移動)とロケーションの列を最初から持ち、在庫管理全体へ広げられるようにする

保存先: SourceData/Stocktake/(Git管理外・Mac ローカル)。NAS への書き込みは現行ルールで対象外。
"""
import csv
import fcntl
import hashlib
import os
import re
import threading
import unicodedata
from datetime import datetime

import jan_lookup as J
from jan_lookup import MP

DATA_DIR = os.path.join(MP.INV_SD, 'Stocktake')
SESSIONS_FILE = 'sessions.csv'
TEMP_ITEMS_FILE = 'temp_items.csv'
TEMP_LINKS_FILE = 'temp_links.csv'
EVENTS_DIR = 'events'
SUMMARY_DIR = 'summaries'
LOCK_FILE = '.writer.lock'

SESSION_COLUMNS = ['セッションID', '種別', '記録日時', '処理区分', '担当者', '端末ID', '備考']
EVENT_COLUMNS = ['イベントID', 'セッションID', '処理区分', 'イベント種別', '記録日時', '端末ID', '担当者',
                 'ロケーション', 'スキャン生値', '入力方法', 'コード種別', '正規化コード', '照合結果',
                 '内部管理ID', '商品名', '数量', '数量単位', '取消対象イベントID', '備考', 'マスター版']
TEMP_ITEM_COLUMNS = ['仮ID', '採番日時', 'セッションID', '端末ID', '担当者', 'メモ', 'ロケーション']
# 仮ID → 内部管理ID の紐付けは人が判断して追記する(最後の行が有効。内部管理IDが空なら紐付け解除)
TEMP_LINK_COLUMNS = ['仮ID', '内部管理ID', '紐付け日時', '確認者', '根拠']

OP_STOCKTAKE = '棚卸'
EV_REGISTER = '登録'
EV_CANCEL = '取消'
SES_START = '開始'
SES_END = '終了'
METHOD_SCAN = 'スキャン'
METHOD_CHOSEN = '候補選択'
METHOD_TEMP = '仮ID'
UNIT_SINGLE = '単品'
UNIT_SET = 'セット'
UNIT_PIECE = '個(未確定)'   # 未登録・仮ID・候補未選択(バラ/セット未選択を含む)。何の単位かは後で照合して決まる

TIME_FMT = '%Y-%m-%d %H:%M:%S'

# 保存先が本番(DATA_DIR)か練習用かは保存場所で決める(フラグの付け忘れで本番扱いにならないよう、既定以外はすべて練習)
ENV_PRODUCTION = 'production'
ENV_PRACTICE = 'practice'


def environment_of(data_dir):
    return ENV_PRODUCTION if os.path.realpath(data_dir) == os.path.realpath(DATA_DIR) else ENV_PRACTICE


def store_id_of(data_dir):
    """保存先ごとの識別子。練習用同士(別フォルダ)でも画面の状態が混ざらないように使う。"""
    return hashlib.sha256(os.path.realpath(data_dir).encode()).hexdigest()[:10]


class StoreError(Exception):
    """画面へ返す入力エラー。code で画面側の動き(再確認など)を切り替える。"""

    def __init__(self, message, code='invalid', detail=None):
        super().__init__(message)
        self.code = code
        self.detail = detail or {}


# ─────────────────────────────── CSV(追記のみ) ───────────────────────────────

def append_row(path, columns, row):
    new = not os.path.exists(path) or os.path.getsize(path) == 0
    # 新規だけ BOM 付き(Mac の Excel で文字化けしないように)。追記時に BOM を挟まない
    with open(path, 'a', encoding='utf-8-sig' if new else 'utf-8', newline='') as f:
        w = csv.writer(f)
        if new:
            w.writerow(columns)
        w.writerow([row.get(c, '') for c in columns])
        f.flush()
        os.fsync(f.fileno())


def read_rows(path, columns):
    if not os.path.exists(path) or os.path.getsize(path) == 0:
        return []
    with open(path, encoding='utf-8-sig', newline='') as f:
        reader = csv.DictReader(f)
        if reader.fieldnames != columns:
            # 列が変わったファイルを黙って読むと別の列を数量として扱いかねない
            raise StoreError(f'{os.path.basename(path)} の列構成が想定と違います: {reader.fieldnames}', 'schema')
        return [dict(r) for r in reader]


def clean_text(v, limit=80):
    s = unicodedata.normalize('NFKC', str(v or '')).strip()
    s = re.sub(r'[\x00-\x1f\x7f]', '', s)[:limit]
    # Excel で開いたときに数式として実行されないように(CSV インジェクション対策)
    return "'" + s if s[:1] in ('=', '+', '-', '@') else s


def validate_qty(qty, config, confirmed=False):
    """数量の検証。空欄と0を区別し、コードの誤入力と大きすぎる値を止める。"""
    qc = config['quantity']
    s = unicodedata.normalize('NFKC', str(qty if qty is not None else '')).strip()
    if s == '':
        raise StoreError('数量が空欄です。0個なら 0 を入力してください', 'qty_blank')
    if not s.isdigit():
        raise StoreError('数量は0以上の整数で入力してください', 'qty_invalid')
    if len(s) >= qc['code_like_min_digits']:
        raise StoreError('数量欄にJANなどのコードが入りました。数量を入力し直してください', 'qty_looks_like_code')
    v = int(s)
    if v > qc['hard_max']:
        raise StoreError(f'数量が上限({qc["hard_max"]})を超えています', 'qty_too_large')
    if v >= qc['confirm_threshold'] and not confirmed:
        raise StoreError(f'数量 {v} は大きい値です。確認のためもう一度入力してください', 'confirm_quantity', {'qty': v})
    return v


# ─────────────────────────────── ストア ───────────────────────────────

class Store:
    def __init__(self, master, config, data_dir=DATA_DIR, writer=True, clock=None):
        self.master = master
        self.config = config
        self.dir = data_dir
        self.environment = environment_of(data_dir)
        self.store_id = store_id_of(data_dir)
        self.clock = clock or datetime.now
        self._lock = threading.RLock()
        self._seq = {}
        self._lock_fh = None
        os.makedirs(os.path.join(self.dir, EVENTS_DIR), exist_ok=True)
        os.makedirs(os.path.join(self.dir, SUMMARY_DIR), exist_ok=True)
        if writer:
            self._acquire_writer_lock()
            links = self._p(TEMP_LINKS_FILE)
            if not os.path.exists(links):
                # 人が紐付けを書き込めるよう見出しだけ用意する
                with open(links, 'w', encoding='utf-8-sig', newline='') as f:
                    csv.writer(f).writerow(TEMP_LINK_COLUMNS)

    def _acquire_writer_lock(self):
        fh = open(self._p(LOCK_FILE), 'a')
        try:
            fcntl.flock(fh, fcntl.LOCK_EX | fcntl.LOCK_NB)
        except OSError:
            fh.close()
            raise StoreError('別の棚卸し画面(サーバー)が同じ保存先で起動中です。二重起動はできません', 'locked')
        self._lock_fh = fh

    def close(self):
        if self._lock_fh:
            fcntl.flock(self._lock_fh, fcntl.LOCK_UN)
            self._lock_fh.close()
            self._lock_fh = None

    def _p(self, *parts):
        return os.path.join(self.dir, *parts)

    def events_path(self, sid):
        return self._p(EVENTS_DIR, f'{sid}.csv')

    def _now(self):
        return self.clock().strftime(TIME_FMT)

    # ---------- セッション ----------

    def sessions(self):
        out = {}
        for r in read_rows(self._p(SESSIONS_FILE), SESSION_COLUMNS):
            sid = r['セッションID']
            if r['種別'] == SES_START:
                out[sid] = {'id': sid, 'started': r['記録日時'], 'op': r['処理区分'], 'staff': r['担当者'],
                            'device': r['端末ID'], 'closed': False, 'ended': ''}
            elif r['種別'] == SES_END and sid in out:
                out[sid]['closed'] = True
                out[sid]['ended'] = r['記録日時']
        return out

    def open_sessions(self):
        return [s for s in self.sessions().values() if not s['closed']]

    def start_session(self, staff, device, op=OP_STOCKTAKE, note=''):
        staff, device = self._who(staff, device)
        with self._lock:
            day = self.clock().strftime('%Y%m%d')
            n = sum(1 for sid in self.sessions() if sid.startswith(f'ST-{day}-')) + 1
            sid = f'ST-{day}-{n:02d}'
            append_row(self._p(SESSIONS_FILE), SESSION_COLUMNS,
                       {'セッションID': sid, '種別': SES_START, '記録日時': self._now(), '処理区分': op,
                        '担当者': staff, '端末ID': device, '備考': clean_text(note)})
            return sid

    def close_session(self, sid, staff, device):
        staff, device = self._who(staff, device)
        with self._lock:
            self._require_open(sid)
            append_row(self._p(SESSIONS_FILE), SESSION_COLUMNS,
                       {'セッションID': sid, '種別': SES_END, '記録日時': self._now(),
                        '処理区分': self.sessions()[sid]['op'], '担当者': staff, '端末ID': device})

    def _require_open(self, sid):
        s = self.sessions().get(sid)
        if not s:
            raise StoreError(f'セッション {sid} がありません', 'no_session')
        if s['closed']:
            raise StoreError(f'セッション {sid} は終了しています', 'session_closed')
        return s

    @staticmethod
    def _who(staff, device):
        staff, device = clean_text(staff, 40), clean_text(device, 40)
        if not staff or not device:
            raise StoreError('担当者と端末IDを入力してください', 'no_staff')
        return staff, device

    def _location(self, loc):
        s = J._clean(loc)
        if s and not re.fullmatch(self.config['location_pattern'], s):
            raise StoreError(f'ロケーション {s!r} の形式が不正です', 'bad_location')
        return s

    # ---------- 記録 ----------

    def events(self, sid):
        return read_rows(self.events_path(sid), EVENT_COLUMNS)

    def _next_event_id(self, sid):
        if sid not in self._seq:
            self._seq[sid] = len(self.events(sid))
        self._seq[sid] += 1
        return f'{sid}-{self._seq[sid]:05d}'

    def _append_event(self, sid, row):
        row['イベントID'] = self._next_event_id(sid)
        row['セッションID'] = sid
        row['記録日時'] = self._now()
        row['マスター版'] = self.master.version
        append_row(self.events_path(sid), EVENT_COLUMNS, row)
        return row

    def active_registrations(self, sid):
        evs = self.events(sid)
        cancelled = {e['取消対象イベントID'] for e in evs if e['イベント種別'] == EV_CANCEL}
        return [e for e in evs if e['イベント種別'] == EV_REGISTER and e['イベントID'] not in cancelled], cancelled

    def _check_duplicate(self, sid, device, code, qty, loc):
        """同じ端末で直前と同じ内容を短時間に登録したら確認を求める(二重登録の防止)。"""
        regs, _ = self.active_registrations(sid)
        mine = [e for e in regs if e['端末ID'] == device]
        if not mine:
            return
        last = mine[-1]
        if (last['正規化コード'], last['数量'], last['ロケーション']) != (code, str(qty), loc):
            return
        age = (self.clock() - datetime.strptime(last['記録日時'], TIME_FMT)).total_seconds()
        if age <= self.config['duplicate_guard_seconds']:
            raise StoreError('直前と同じ商品・数量・ロケーションです。追加で登録するなら確認してください',
                             'confirm_duplicate', {'event': last['イベントID']})

    def register(self, sid, staff, device, location, raw, qty, chosen_pid='',
                 qty_confirmed=False, dup_confirmed=False):
        staff, device = self._who(staff, device)
        with self._lock:
            ses = self._require_open(sid)
            loc = self._location(location)
            n = J.normalize_code(raw, self.config)
            note, method, pid = '', METHOD_SCAN, ''
            if n['kind'] == J.CODE_JAN:
                res = self.master.resolve_jan(n['code'])
                status, pid = res['status'], res['pid']
                if chosen_pid:
                    # 共通JANでセットPを選んだ場合も、ここでは換算しない(セットP×セット数のまま記録)
                    if chosen_pid not in res['candidates'] or status not in (J.R_DUPLICATE, J.R_SET_ONLY, J.R_SHARED):
                        raise StoreError(f'{chosen_pid} はこのJANの候補ではありません', 'bad_choice')
                    status, pid, method = J.R_CHOSEN, chosen_pid, METHOD_CHOSEN
                elif not pid and res['candidates']:
                    note = '候補: ' + ', '.join(res['candidates'])
            elif n['kind'] == J.CODE_TEMP:
                if n['code'] not in self.temp_items():
                    raise StoreError(f'仮ID {n["code"]} は発行されていません', 'unknown_temp')
                status, method = J.R_TEMP, METHOD_TEMP
                pid = self.temp_links().get(n['code'], '')
            else:
                raise StoreError(n['message'] or 'このコードは登録できません', 'bad_code', {'kind': n['kind']})
            q = validate_qty(qty, self.config, qty_confirmed)
            if not dup_confirmed:
                self._check_duplicate(sid, device, n['code'], q, loc)
            return self._append_event(sid, self._event_row(ses, staff, device, loc, raw, method, n, status, pid, q, note))

    def register_temp(self, sid, staff, device, location, memo, qty, qty_confirmed=False):
        """JANなし・読めない商品。仮IDを発行して同時に登録する(数量の検証が先。孤立した仮IDを作らない)。"""
        staff, device = self._who(staff, device)
        memo = clean_text(memo, 120)
        if not memo:
            raise StoreError('後で商品を特定できるよう、メモを入力してください(例: 青い箱 犬用ガム)', 'no_memo')
        with self._lock:
            ses = self._require_open(sid)
            loc = self._location(location)
            q = validate_qty(qty, self.config, qty_confirmed)
            tid = self._next_temp_id()
            append_row(self._p(TEMP_ITEMS_FILE), TEMP_ITEM_COLUMNS,
                       {'仮ID': tid, '採番日時': self._now(), 'セッションID': sid, '端末ID': device,
                        '担当者': staff, 'メモ': memo, 'ロケーション': loc})
            n = {'kind': J.CODE_TEMP, 'code': tid}
            return self._append_event(sid, self._event_row(ses, staff, device, loc, tid, METHOD_TEMP, n,
                                                           J.R_TEMP, '', q, memo))

    def _event_row(self, ses, staff, device, loc, raw, method, n, status, pid, q, note):
        prod = self.master.describe(pid) if pid else None
        if prod and prod['kind'] == J.KIND_SINGLE:
            unit = UNIT_SINGLE
        elif prod:
            unit = UNIT_SET
            note = note or 'セット換算待ち(構成表の完成後に単品へ換算)'
        else:
            unit = UNIT_PIECE
        return {'処理区分': ses['op'], 'イベント種別': EV_REGISTER, '端末ID': device, '担当者': staff,
                'ロケーション': loc, 'スキャン生値': clean_text(raw, 40), '入力方法': method,
                'コード種別': n['kind'], '正規化コード': n['code'], '照合結果': status, '内部管理ID': pid,
                '商品名': prod['name'] if prod else '', '数量': str(q), '数量単位': unit, '備考': note}

    def cancel(self, sid, staff, device, target_id=''):
        """取消イベントを追記する。元の行は残す。対象省略時はこの端末の直前の有効な登録。"""
        staff, device = self._who(staff, device)
        with self._lock:
            ses = self._require_open(sid)
            regs, cancelled = self.active_registrations(sid)
            if target_id:
                if target_id in cancelled:
                    raise StoreError(f'{target_id} は既に取消済みです', 'already_cancelled')
                hit = [e for e in regs if e['イベントID'] == target_id]
            else:
                hit = [e for e in regs if e['端末ID'] == device][-1:]
            if not hit:
                raise StoreError('取り消せる登録がありません', 'nothing_to_cancel')
            t = hit[0]
            row = {k: t[k] for k in ('ロケーション', 'スキャン生値', '入力方法', 'コード種別', '正規化コード',
                                     '照合結果', '内部管理ID', '商品名', '数量', '数量単位')}
            row.update({'処理区分': ses['op'], 'イベント種別': EV_CANCEL, '端末ID': device, '担当者': staff,
                        '取消対象イベントID': t['イベントID'], '備考': f'{t["イベントID"]} を取消'})
            return self._append_event(sid, row)

    def recent(self, sid, device='', limit=15):
        evs = self.events(sid)
        cancelled = {e['取消対象イベントID'] for e in evs if e['イベント種別'] == EV_CANCEL}
        out = [dict(e, 取消済=e['イベントID'] in cancelled) for e in evs if not device or e['端末ID'] == device]
        return out[::-1][:limit]

    # ---------- 仮ID ----------

    def temp_items(self):
        return {r['仮ID']: r for r in read_rows(self._p(TEMP_ITEMS_FILE), TEMP_ITEM_COLUMNS)}

    def temp_links(self):
        links = {}
        for r in read_rows(self._p(TEMP_LINKS_FILE), TEMP_LINK_COLUMNS):
            links[r['仮ID'].strip().upper()] = r['内部管理ID'].strip()
        return {k: v for k, v in links.items() if v}

    def _next_temp_id(self):
        prefix, digits = self.config['temp_id_prefix'], self.config['temp_id_digits']
        nums = [int(t[len(prefix):]) for t in self.temp_items() if t[len(prefix):].isdigit()]
        return f'{prefix}{(max(nums) if nums else 0) + 1:0{digits}d}'
