# =============================================================
#  momiji-mcp — MomijiStore OS MCP Server
#  Phase3 STEP4 / STEP5-1
# =============================================================
FROM python:3.12-slim

ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1

WORKDIR /app

# 依存だけ先に入れる(ソース変更時のリビルドを速くするため)
# requirements.lock が正本。動作確認済みのバージョンで固定してビルドする。
# requirements.txt は「何を直接使うか」の宣言で、記録として同梱する。
COPY requirements.lock requirements.txt ./
RUN pip install --no-cache-dir -r requirements.lock

# MCPServer が現在のSDKの正式APIであることを build 時点で確認する。
# SDKのAPIが変わっていれば、ここで build が失敗する(実行時まで発覚を遅らせない)。
# psycopg は Intelligence Layer が使う。無いまま起動して、判断を記録しようと
# した瞬間に初めて失敗する、という事態を避けるため build 時に確認する。
RUN python -c "import mcp; from mcp.server import MCPServer; import psycopg; \
print('mcp', getattr(mcp, '__version__', 'unknown'), '/ MCPServer OK'); \
print('psycopg', psycopg.__version__, 'OK')"

COPY src/ ./src/

# 構文チェック。エラーがあれば build を失敗させる。
# ※ PYTHONDONTWRITEBYTECODE=1 のため .pyc は生成されないが、
#    構文エラーの検出(exit 1)は有効に働く。
RUN python -m compileall src

# 検証を build に組み込む。壊れたロジックはイメージにならない。
# Deploy.sh を通すたびに自動で走るため、テストの実行忘れが起きない。
# ※ DBは使わない範囲(入力検証・SQL組み立て)のみ。DB本体の制約は
#    deploy/Migrate.sh の「構造の確認」がNAS上で検証する。
COPY tests/ ./tests/
RUN python -m unittest discover -s tests -v

# 非rootで実行する(最小権限 — PROJECT_CHARTER §4)
RUN useradd --system --uid 10001 momiji && chown -R momiji:momiji /app
USER momiji

EXPOSE 8000

CMD ["python", "-u", "src/server.py"]
