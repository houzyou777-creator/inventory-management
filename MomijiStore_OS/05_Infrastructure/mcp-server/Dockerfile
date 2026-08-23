# =============================================================
#  momiji-mcp — MomijiStore OS MCP Server
#  Phase3 STEP4 / STEP5-1
# =============================================================
FROM python:3.12-slim

ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1

WORKDIR /app

# 依存だけ先に入れる(ソース変更時のリビルドを速くするため)
COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

# MCPServer が現在のSDKの正式APIであることを build 時点で確認する。
# SDKのAPIが変わっていれば、ここで build が失敗する(実行時まで発覚を遅らせない)。
RUN python -c "import mcp; from mcp.server import MCPServer; \
print('mcp', getattr(mcp, '__version__', 'unknown'), '/ MCPServer OK')"

COPY src/ ./src/

# 構文チェック。エラーがあれば build を失敗させる。
# ※ PYTHONDONTWRITEBYTECODE=1 のため .pyc は生成されないが、
#    構文エラーの検出(exit 1)は有効に働く。
RUN python -m compileall src

# 非rootで実行する(最小権限 — PROJECT_CHARTER §4)
RUN useradd --system --uid 10001 momiji && chown -R momiji:momiji /app
USER momiji

EXPOSE 8000

CMD ["python", "-u", "src/server.py"]
