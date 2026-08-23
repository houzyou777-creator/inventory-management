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
import time
from typing import Any

import openpyxl
from mcp.server import MCPServer

SERVICE_NAME = "momiji-mcp"

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

SHEET_MASTER = "商品マスター"
SHEET_LISTING = "出品テーブル"

# 商品マスター シートの列位置
PM_INTERNAL_ID, PM_JAN, PM_NAME = 0, 1, 2
PM_COST, PM_CHANNEL, PM_STATUS = 4, 5, 6

# 出品テーブル シートの列位置
LT_INTERNAL_ID, LT_RAKUTEN_SKU, LT_ASIN = 1, 4, 5

# メモリキャッシュ(初回アクセス時にロード。ファイル更新で自動再ロード)
_cache: list[dict[str, Any]] = []
_cache_mtime: float | None = None


def _load_products() -> list[dict[str, Any]]:
    """Excelを読み取り専用で開き、商品マスターと出品テーブルを突き合わせて返す。

    ASIN・楽天SKUは出品テーブル側にあり、1商品が複数の出品行を持つ場合がある
    (楽天とAmazonの両方など)。値が存在しない場合は推測せず None を返す。
    """
    wb = openpyxl.load_workbook(PRODUCTS_XLSX, read_only=True, data_only=True)
    try:
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
        for row in wb[SHEET_MASTER].iter_rows(min_row=2, values_only=True):
            pid = row[PM_INTERNAL_ID]
            if not pid:
                continue
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
        return products
    finally:
        wb.close()


def _ensure_loaded() -> list[dict[str, Any]]:
    """キャッシュを返す。未ロード、またはファイルが更新されていれば再ロードする。"""
    global _cache, _cache_mtime
    mtime = os.path.getmtime(PRODUCTS_XLSX)
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

    products = _ensure_loaded()
    results = []
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
        results.append(p)

    elapsed_ms = (time.perf_counter() - started) * 1000
    logger.info(
        "search_products 条件=%s ヒット=%d件 処理時間=%.1fms",
        given,
        len(results),
        elapsed_ms,
    )
    return {
        "results": results,
        "count": len(results),
        "elapsed_ms": round(elapsed_ms, 1),
    }


if __name__ == "__main__":
    # ポートはコンテナ内部(momiji_net)でのみ使用し、外部公開しない。
    mcp.run(transport="streamable-http", host="0.0.0.0", port=8000)
