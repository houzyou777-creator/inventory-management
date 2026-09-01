"""momiji-mcp — MomijiStore OS MCP Server

公式SDK(mcp)の標準サーバーAPIのみを使用する。独自実装はしない。

Phase3 STEP4: health_check
Phase3 STEP5-1: search_products(商品マスター検索・読み取り専用)

⚠️ 読み取り専用。Excelの書き込み・更新・削除は一切行わない。
   NAS上のExcelは「MCP検索用の読み取り専用スナップショット」であり正本ではない。
   正本は Mac側 SourceData/商品マスター_単品_v1.0.xlsx。
"""

from __future__ import annotations

import logging
import os
import secrets
import threading
import time
from pathlib import Path
from typing import Any

import openpyxl
from mcp.server import MCPServer

# starlette / uvicorn は mcp SDK の依存として既に入っている。
# 認証ミドルウェアと /health を足すために直接importする(新規の追加依存ではない)。
from starlette.middleware.base import BaseHTTPMiddleware
from starlette.requests import Request
from starlette.responses import JSONResponse

SERVICE_NAME = "momiji-mcp"

# --- 認証 -----------------------------------------------------------------
#  /mcp 配下は全てAPIキー必須。MCPのツールはすべて /mcp を通るため、
#  今後ツールを追加しても自動的に認証が適用される(個別対応は不要)。
#  /health のみ認証不要(監視用)。
API_KEY = os.environ.get("MOMIJI_API_KEY", "").strip()
API_KEY_HEADER = "X-API-Key"
PUBLIC_PATHS = frozenset({"/health"})

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
)
logger = logging.getLogger(SERVICE_NAME)

mcp = MCPServer(SERVICE_NAME)

# 読み取り専用マウント (:ro) されたスナップショット
PRODUCTS_XLSX = os.environ.get(
    "MOMIJI_PRODUCTS_XLSX", "/app/data/products/商品マスター_単品_v1.0.xlsx"
)
# 起動時点で「実際にどのパスを見るか」をログへ残す。
# 環境変数で上書きされていた場合、ここで判明する。
logger.info(
    "起動 参照パス=%s (環境変数指定=%s)",
    PRODUCTS_XLSX,
    "あり" if os.environ.get("MOMIJI_PRODUCTS_XLSX") else "なし(既定値)",
)

SHEET_MASTER = "商品マスター"
SHEET_LISTING = "出品テーブル"

# 商品マスター シートの列位置と、想定する列見出し(変更検知に使う)
PM_INTERNAL_ID, PM_JAN, PM_NAME = 0, 1, 2
PM_COST, PM_CHANNEL, PM_STATUS = 4, 5, 6
PM_HEADERS = {
    PM_INTERNAL_ID: "内部管理ID",
    PM_JAN: "JAN",
    PM_NAME: "商品名",
    PM_COST: "標準原価",
    PM_CHANNEL: "販売チャネル",
    PM_STATUS: "販売ステータス",
}

# 出品テーブル シートの列位置と、想定する列見出し
LT_INTERNAL_ID, LT_RAKUTEN_SKU, LT_ASIN = 1, 4, 5
LT_HEADERS = {
    LT_INTERNAL_ID: "内部管理ID",
    LT_RAKUTEN_SKU: "楽天SKU",
    LT_ASIN: "ASIN",
}

# 1回の検索で返す最大件数(応答の肥大を防ぐ)
# .env / docker-compose 側で変更できる。既定は 50 / 500。
DEFAULT_LIMIT = int(os.getenv("MOMIJI_DEFAULT_LIMIT", "50"))
MAX_LIMIT = int(os.getenv("MOMIJI_MAX_LIMIT", "500"))

# メモリキャッシュ(初回アクセス時にロード。ファイル更新で自動再ロード)
_cache: list[dict[str, Any]] = []
_cache_mtime: float | None = None
_cache_lock = threading.Lock()


class ProductDataError(Exception):
    """商品マスターを読めなかった。メッセージは利用者へそのまま返せる内容にする。"""


def _path_diagnostics(exc: BaseException) -> str:
    """アクセス失敗時に、原因を特定できるだけの情報を組み立てる。

    例外を握り潰すと調査ができなくなるため、実際の例外型・メッセージと、
    参照パス・存在確認・親ディレクトリの内容を必ず添える。
    ここで扱うのはコンテナ内部のパスであり、秘密情報は含まない。
    """
    target = Path(PRODUCTS_XLSX)
    parent = target.parent
    try:
        entries: Any = sorted(os.listdir(parent))
    except OSError as list_exc:
        entries = f"<一覧取得不可: {type(list_exc).__name__}: {list_exc}>"
    return (
        f"商品マスターへアクセスできません({type(exc).__name__}: {exc})。"
        f" 参照パス={PRODUCTS_XLSX}"
        f" / exists={target.exists()} is_file={target.is_file()}"
        f" / 親={parent} exists={parent.exists()} is_dir={parent.is_dir()}"
        f" / 親の内容={entries}"
    )


def _verify_headers(sheet, expected: dict[int, str], sheet_name: str) -> None:
    """列見出しが想定どおりかを確認する。

    列構成が変わったまま読み進めると誤った値を返してしまう。
    黙って間違えるより、明確に失敗させる(BL-7 記録なき変更を認めない)。
    """
    header = next(sheet.iter_rows(min_row=1, max_row=1, values_only=True), None)
    if header is None:
        raise ProductDataError(f"シート「{sheet_name}」が空です。")
    for index, name in expected.items():
        actual = header[index] if index < len(header) else None
        if actual != name:
            raise ProductDataError(
                f"シート「{sheet_name}」の列構成が想定と異なります"
                f"(列{index + 1}: 期待「{name}」/ 実際「{actual}」)。"
                "商品マスターの形式変更が疑われます。"
            )


def _load_products() -> list[dict[str, Any]]:
    """Excelを読み取り専用で開き、商品マスターと出品テーブルを突き合わせて返す。

    ASIN・楽天SKUは出品テーブル側にあり、1商品が複数の出品行を持つ場合がある
    (楽天とAmazonの両方など)。値が存在しない場合は推測せず None を返す。
    """
    logger.info("商品マスターを開きます path=%s", PRODUCTS_XLSX)
    try:
        wb = openpyxl.load_workbook(PRODUCTS_XLSX, read_only=True, data_only=True)
    except (FileNotFoundError, PermissionError, OSError) as exc:
        logger.exception("商品マスターを開けません path=%s", PRODUCTS_XLSX)
        raise ProductDataError(_path_diagnostics(exc)) from exc
    except Exception as exc:
        # 破損ファイル・非対応形式など
        logger.exception("商品マスターの読み込みに失敗 path=%s", PRODUCTS_XLSX)
        raise ProductDataError(
            f"商品マスターのファイルを開けませんでした({type(exc).__name__}: {exc})。"
            "ファイルが破損していないか確認してください。"
        ) from exc

    try:
        for sheet_name in (SHEET_MASTER, SHEET_LISTING):
            if sheet_name not in wb.sheetnames:
                raise ProductDataError(
                    f"シート「{sheet_name}」が見つかりません。"
                    f"存在するシート: {', '.join(wb.sheetnames)}"
                )
        _verify_headers(wb[SHEET_LISTING], LT_HEADERS, SHEET_LISTING)
        _verify_headers(wb[SHEET_MASTER], PM_HEADERS, SHEET_MASTER)

        listings: dict[str, dict[str, Any]] = {}
        for row in wb[SHEET_LISTING].iter_rows(min_row=2, values_only=True):
            pid = row[LT_INTERNAL_ID]
            if not pid:
                continue
            entry = listings.setdefault(pid, {"asin": None, "rakuten_sku": None})
            # 最初に見つかった非空の値を採用する
            if entry["asin"] is None and row[LT_ASIN]:
                entry["asin"] = str(row[LT_ASIN])
            if entry["rakuten_sku"] is None and row[LT_RAKUTEN_SKU]:
                entry["rakuten_sku"] = str(row[LT_RAKUTEN_SKU])

        products = []
        seen: set[str] = set()
        duplicates = 0
        for row in wb[SHEET_MASTER].iter_rows(min_row=2, values_only=True):
            pid = row[PM_INTERNAL_ID]
            if not pid:
                continue
            if pid in seen:
                duplicates += 1
            seen.add(pid)
            listing = listings.get(pid, {})
            products.append(
                {
                    "内部管理ID": pid,
                    "商品名": row[PM_NAME],
                    "JAN": str(row[PM_JAN]) if row[PM_JAN] is not None else None,
                    "ASIN": listing.get("asin"),
                    "楽天SKU": listing.get("rakuten_sku"),
                    "標準原価": row[PM_COST],
                    "販売チャネル": row[PM_CHANNEL],
                    "販売ステータス": row[PM_STATUS],
                }
            )
        if duplicates:
            # 正本側の問題なので落とさず警告する。該当行は全件返す。
            logger.warning("内部管理IDの重複を検出しました 件数=%d", duplicates)
        return products
    finally:
        wb.close()


def _ensure_loaded() -> list[dict[str, Any]]:
    """キャッシュを返す。未ロード、またはファイルが更新されていれば再ロードする。

    同時アクセスでの二重ロードを避けるためロックを取る。
    """
    global _cache, _cache_mtime
    try:
        mtime = os.path.getmtime(PRODUCTS_XLSX)
    except OSError as exc:
        # 例外を握り潰さない。実際の型・メッセージとパス診断を必ず残す。
        logger.exception("商品マスターへアクセスできません path=%s", PRODUCTS_XLSX)
        raise ProductDataError(_path_diagnostics(exc)) from exc

    with _cache_lock:
        if _cache_mtime != mtime:
            _cache = _load_products()
            _cache_mtime = mtime
            logger.info("商品マスターをロードしました 件数=%d", len(_cache))
        return _cache


@mcp.tool()
def health_check() -> dict:
    """MCPサーバーが応答することを確認する。"""
    return {"status": "ok", "service": SERVICE_NAME}


@mcp.tool()
def search_products(
    name: str | None = None,
    jan: str | None = None,
    internal_id: str | None = None,
    asin: str | None = None,
    rakuten_sku: str | None = None,
    limit: int = DEFAULT_LIMIT,
) -> dict:
    """商品マスターを検索する(読み取り専用)。

    商品名は部分一致、それ以外は完全一致。複数指定した場合はAND条件。
    条件を1つも指定しない場合は全件返さず、エラーを返す。

    Args:
        name: 商品名(部分一致)
        jan: JANコード(完全一致)
        internal_id: 内部管理ID(完全一致)
        asin: ASIN(完全一致)
        rakuten_sku: 楽天SKU(完全一致)
        limit: 返却する最大件数(既定50・上限500)。超過分は truncated=true で示す
    """
    started = time.perf_counter()
    conditions = {
        "name": name,
        "jan": jan,
        "internal_id": internal_id,
        "asin": asin,
        "rakuten_sku": rakuten_sku,
    }
    given = {k: v for k, v in conditions.items() if v}

    if not given:
        logger.info("search_products 条件なし — 検索を拒否しました")
        return {
            "error": "検索条件を1つ以上指定してください "
            "(name / jan / internal_id / asin / rakuten_sku)。"
            "全件返却は行いません。",
            "results": [],
            "count": 0,
        }

    try:
        products = _ensure_loaded()
    except ProductDataError as exc:
        logger.error("search_products 読み込み失敗: %s", exc)
        return {"error": str(exc), "results": [], "count": 0, "matched": 0}

    effective_limit = max(1, min(int(limit), MAX_LIMIT))
    results = []
    matched = 0
    for p in products:
        if name and (not p["商品名"] or name not in str(p["商品名"])):
            continue
        if jan and p["JAN"] != str(jan):
            continue
        if internal_id and p["内部管理ID"] != internal_id:
            continue
        if asin and p["ASIN"] != str(asin):
            continue
        if rakuten_sku and p["楽天SKU"] != str(rakuten_sku):
            continue
        matched += 1
        if len(results) < effective_limit:
            results.append(p)

    elapsed_ms = (time.perf_counter() - started) * 1000
    logger.info(
        "search_products 条件=%s 該当=%d件 返却=%d件 処理時間=%.1fms",
        given,
        matched,
        len(results),
        elapsed_ms,
    )
    return {
        "results": results,
        "count": len(results),
        "matched": matched,
        "truncated": matched > len(results),
        "elapsed_ms": round(elapsed_ms, 1),
    }


class ApiKeyMiddleware(BaseHTTPMiddleware):
    """/health 以外のすべてのパスでAPIキーを必須にする。

    MCPのツール呼び出しはすべて /mcp を通るため、この1箇所で
    既存・将来を問わず全ツールが保護される。
    """

    async def dispatch(self, request: Request, call_next):
        if request.url.path in PUBLIC_PATHS:
            return await call_next(request)

        provided = request.headers.get(API_KEY_HEADER, "")
        # タイミング攻撃を避けるため定数時間で比較する
        if not (provided and secrets.compare_digest(provided, API_KEY)):
            logger.warning(
                "認証失敗 path=%s header=%s",
                request.url.path,
                "あり" if provided else "なし",
            )
            return JSONResponse({"error": "Forbidden"}, status_code=403)

        return await call_next(request)


async def health_endpoint(request: Request) -> JSONResponse:
    """認証不要の監視用エンドポイント。データには一切触れない。"""
    return JSONResponse({"status": "ok", "service": SERVICE_NAME})


if __name__ == "__main__":
    # APIキー未設定なら起動しない(認証なしで公開される事故を防ぐ)
    if not API_KEY:
        raise SystemExit(
            "起動できません: 環境変数 MOMIJI_API_KEY が未設定です。"
            " .env に設定してください。"
        )

    import uvicorn

    app = mcp.streamable_http_app()
    app.add_middleware(ApiKeyMiddleware)
    app.add_route("/health", health_endpoint, methods=["GET"])

    logger.info(
        "起動します 認証=有効(%s ヘッダー必須) 認証不要パス=%s",
        API_KEY_HEADER,
        sorted(PUBLIC_PATHS),
    )
    # ポートはコンテナ内部で待ち受け、公開範囲は compose の ports で制御する。
    uvicorn.run(app, host="0.0.0.0", port=8000)
