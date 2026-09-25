# -*- coding: utf-8 -*-
"""stocktake_app.py — JANスキャン棚卸し画面(ローカルWeb・標準ライブラリのみ)

使い方:
    python3 stocktake_app.py                 この Mac のブラウザだけで使う(http://127.0.0.1:8765/)
    python3 stocktake_app.py --lan           LAN 内の MacBook・iPad からも使う(アクセス用トークン付きURLを表示)
    python3 stocktake_app.py --data-dir DIR  保存先を変える(テスト・練習用)

この画面がすること / しないこと:
    ・する   : JAN → 商品マスター照合(読み取り専用)・数量の記録(追記のみ・1件ごとに保存)・取消イベント・集計xlsx
    ・しない : 商品マスター・在庫管理テーブルへの書き込み、Amazon・楽天・プライスターへの反映

なぜ Web 画面か(2026-09-25 承認 C案): 端末を Mac mini に固定せず、LAN 内の iPad 等からも同じ記録へ
書き込めるようにするため。書き込みはこのサーバー1か所に集め、端末ごとの保存(同期漏れ)を作らない。
"""
import argparse
import json
import os
import secrets
import socket
import sys
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, urlparse

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import jan_lookup as J  # noqa: E402
import stocktake_aggregate as A  # noqa: E402
import stocktake_store as S  # noqa: E402

STATIC_DIR = os.path.join(J.HERE, 'static')
LOCAL_HOSTS = ('127.0.0.1', 'localhost')
MAX_BODY = 64 * 1024


class App:
    """HTTP から切り離した処理本体(テストから直接呼べるように)。"""

    def __init__(self, store, token=''):
        self.store = store
        self.master = store.master
        self.config = store.config
        self.token = token

    # ---------- 照合(書き込みなし) ----------

    def lookup(self, code):
        n = J.normalize_code(code, self.config)
        out = {'kind': n['kind'], 'code': n['code'], 'message': n['message'],
               'status': '', 'product': None, 'candidates': []}
        if n['kind'] == J.CODE_JAN:
            res = self.master.resolve_jan(n['code'])
            out['status'] = res['status']
            out['product'] = self._product(res['pid'])
            out['candidates'] = [self._product(p) for p in res['candidates']]
        elif n['kind'] == J.CODE_TEMP:
            item = self.store.temp_items().get(n['code'])
            if not item:
                out.update(kind=J.CODE_INVALID, message=f'仮ID {n["code"]} は発行されていません')
            else:
                out['status'] = J.R_TEMP
                out['temp'] = {'id': n['code'], 'memo': item['メモ']}
                out['product'] = self._product(self.store.temp_links().get(n['code'], ''))
        return out

    def _product(self, pid):
        p = self.master.describe(pid) if pid else None
        if p and not self.config.get('show_cost', True):
            p['cost'] = None
        return p

    # ---------- ルーティング ----------

    def handle(self, method, path, query, body):
        s = self.store
        if method == 'GET' and path == '/api/info':
            return {'environment': s.environment, 'store_id': s.store_id,
                    'master_version': self.master.version, 'master_file': os.path.basename(self.master.path),
                    'products': len(self.master.products), 'commands': self.config['commands'],
                    'location_pattern': self.config['location_pattern'],
                    'quantity': self.config['quantity'], 'show_cost': self.config.get('show_cost', True)}
        if method == 'GET' and path == '/api/sessions':
            return {'sessions': s.open_sessions()}
        if method == 'GET' and path == '/api/recent':
            sid = query.get('session', [''])[0]
            dev = query.get('device', [''])[0]
            return {'events': s.recent(sid, dev) if sid else []}
        if method != 'POST':
            raise S.StoreError('not found', 'not_found')
        b = body
        who = (b.get('staff', ''), b.get('device', ''))
        if path == '/api/sessions':
            return {'session': s.start_session(*who)}
        if path == '/api/lookup':
            return self.lookup(b.get('code', ''))
        if path == '/api/register':
            ev = s.register(b.get('session', ''), *who, b.get('location', ''), b.get('raw', ''), b.get('qty', ''),
                            chosen_pid=b.get('chosen_pid', ''), qty_confirmed=bool(b.get('qty_confirmed')),
                            dup_confirmed=bool(b.get('dup_confirmed')))
            return {'event': ev}
        if path == '/api/register_temp':
            ev = s.register_temp(b.get('session', ''), *who, b.get('location', ''), b.get('memo', ''),
                                 b.get('qty', ''), qty_confirmed=bool(b.get('qty_confirmed')))
            return {'event': ev}
        if path == '/api/cancel':
            return {'event': s.cancel(b.get('session', ''), *who, b.get('target', ''))}
        if path in ('/api/aggregate', '/api/close'):
            sid = b.get('session', '')
            if path == '/api/close':
                s.close_session(sid, *who)
            result = A.run([sid], s)
            out = A.write_xlsx(result, [sid], s, os.path.join(s.dir, S.SUMMARY_DIR))
            return {'file': out, 'stats': result['stats'], 'checks': len(result['checks']),
                    'fingerprint': result['fingerprint']}
        raise S.StoreError('not found', 'not_found')


def make_handler(app, bind_host, port):
    class Handler(BaseHTTPRequestHandler):
        server_version = 'MomijiStocktake/1'

        def log_message(self, fmt, *args):   # 画面操作ごとのアクセスログは出さない(記録はCSVが正)
            pass

        def _send(self, code, obj=None, raw=None, ctype='application/json; charset=utf-8'):
            data = raw if raw is not None else json.dumps(obj, ensure_ascii=False, default=list).encode()
            self.send_response(code)
            self.send_header('Content-Type', ctype)
            self.send_header('Content-Length', str(len(data)))
            self.send_header('Cache-Control', 'no-store')
            self.send_header('X-Content-Type-Options', 'nosniff')
            self.end_headers()
            self.wfile.write(data)

        def _host_ok(self):
            # localhost 運用では Host を固定し、他サイトから localhost を突く攻撃(DNS rebinding)を防ぐ
            if bind_host not in LOCAL_HOSTS:
                return True
            return (self.headers.get('Host') or '') in (f'127.0.0.1:{port}', f'localhost:{port}')

        def _dispatch(self, method):
            if not self._host_ok():
                return self._send(403, {'error': 'host not allowed', 'code': 'forbidden'})
            u = urlparse(self.path)
            if method == 'GET' and u.path in ('/', '/index.html'):
                with open(os.path.join(STATIC_DIR, 'index.html'), 'rb') as f:
                    return self._send(200, raw=f.read(), ctype='text/html; charset=utf-8')
            if not u.path.startswith('/api/'):
                return self._send(404, {'error': 'not found', 'code': 'not_found'})
            if app.token and not secrets.compare_digest(self.headers.get('X-Stocktake-Token', ''), app.token):
                return self._send(401, {'error': 'アクセス用トークンが違います。表示されたURLから開き直してください',
                                        'code': 'token'})
            body = {}
            if method == 'POST':
                if not (self.headers.get('Content-Type') or '').startswith('application/json'):
                    return self._send(415, {'error': 'JSON で送ってください', 'code': 'content_type'})
                n = int(self.headers.get('Content-Length') or 0)
                if n > MAX_BODY:
                    return self._send(413, {'error': 'too large', 'code': 'too_large'})
                try:
                    body = json.loads(self.rfile.read(n) or b'{}')
                except ValueError:
                    return self._send(400, {'error': 'JSON が不正です', 'code': 'bad_json'})
            try:
                return self._send(200, app.handle(method, u.path, parse_qs(u.query), body))
            except S.StoreError as e:
                return self._send(404 if e.code == 'not_found' else 400,
                                  {'error': str(e), 'code': e.code, 'detail': e.detail})
            except Exception as e:  # noqa: BLE001  画面には「保存されていない」と必ず伝える
                return self._send(500, {'error': f'サーバーエラー(登録されていません): {e}', 'code': 'server'})

        def do_GET(self):
            self._dispatch('GET')

        def do_POST(self):
            self._dispatch('POST')

    return Handler


def lan_address():
    """LAN 側の IP を調べる(実際には送信しない UDP の connect で経路だけ引く)。"""
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as sk:
            sk.connect(('10.255.255.255', 1))
            return sk.getsockname()[0]
    except OSError:
        return '(IP不明)'


def main():
    cfg = J.load_config()
    ap = argparse.ArgumentParser(description='JANスキャン棚卸し画面')
    ap.add_argument('--lan', action='store_true', help='LAN 内の他端末(MacBook・iPad)からも使えるようにする')
    ap.add_argument('--port', type=int, default=cfg['server']['port'])
    ap.add_argument('--data-dir', default=S.DATA_DIR)
    a = ap.parse_args()

    master = J.load_master()
    store = S.Store(master, cfg, data_dir=a.data_dir)
    host = '0.0.0.0' if a.lan else cfg['server']['default_host']
    token = secrets.token_urlsafe(9) if a.lan else ''
    app = App(store, token)
    httpd = ThreadingHTTPServer((host, a.port), make_handler(app, host, a.port))
    print('JANスキャン棚卸し画面を起動しました(終了は Ctrl+C)')
    print(f'  商品マスター: {master.path}  版 {master.version}  {len(master.products)}件')
    print(f'  保存先      : {store.dir}')
    print(f'  環境        : {"本番" if store.environment == S.ENV_PRODUCTION else "練習用"}({store.environment})')
    print(f'  この Mac    : http://127.0.0.1:{a.port}/' + (f'?t={token}' if token else ''))
    if a.lan:
        print(f'  LAN 内端末  : http://{lan_address()}:{a.port}/?t={token}')
        print('  ※ LAN 内の誰でも開けるため、トークン付きURLは作業者にだけ伝えてください')
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        httpd.server_close()
        store.close()
        print('停止しました(登録済みのデータは保存されています)')


if __name__ == '__main__':
    main()
