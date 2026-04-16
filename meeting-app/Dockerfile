FROM python:3.13-slim

WORKDIR /app

# uv のインストール
COPY --from=ghcr.io/astral-sh/uv:latest /uv /usr/local/bin/uv

# 依存関係のインストール（キャッシュ活用のため先にコピー）
COPY pyproject.toml uv.lock ./
RUN uv sync --frozen --no-dev

# アプリコードのコピー
COPY . .

EXPOSE 8000

CMD ["uv", "run", "uvicorn", "app:APP", "--host", "0.0.0.0", "--port", "8000"]
