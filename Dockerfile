# outlook-mcp — stdio MCP サーバ
#
# 認証情報はイメージに焼かない。.env もトークンキャッシュも実行時に渡す。
#
#   docker build -t outlook-mcp .
#
#   # 初回ログイン(ブラウザでコードを入力する。-it が要る)
#   docker run -it --rm \
#     -e OUTLOOK_CLIENT_ID=<あなたのクライアントID> \
#     -v outlook-mcp-token:/app/data \
#     -e OUTLOOK_TOKEN_CACHE=/app/data/token_cache.json \
#     outlook-mcp python login.py
#
#   # MCP クライアントから起動するとき(stdio なので -i が要る。-t は付けない)
#   docker run -i --rm \
#     -e OUTLOOK_CLIENT_ID=<あなたのクライアントID> \
#     -v outlook-mcp-token:/app/data \
#     -e OUTLOOK_TOKEN_CACHE=/app/data/token_cache.json \
#     outlook-mcp
#
# トークンキャッシュはメールボックスの鍵に相当する。名前付きボリュームに置き、
# イメージには絶対に含めないこと。

FROM python:3.13-slim

ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1

WORKDIR /app

# 依存を先に入れてレイヤーキャッシュを効かせる
COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

COPY outlook_server.py outlook_auth.py login.py ./

# root で動かさない。トークンキャッシュの置き場だけ書き込み可能にする。
RUN useradd --create-home --uid 10001 app \
    && mkdir -p /app/data \
    && chown -R app:app /app
USER app

# stdio サーバ。標準入出力がそのまま MCP のトランスポートになるため、
# 起動時に何も標準出力へ書かない。認証情報が無くても起動は成功し、
# ツール一覧には応答する(各ツールが設定手順を案内する文字列を返す)。
CMD ["python", "outlook_server.py"]
