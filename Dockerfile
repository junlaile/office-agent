# syntax=docker/dockerfile:1
# ============================================================
# office-agent 镜像
#
# 构建:
#   docker build -t office-agent .
#   # GitHub API 限流时（下载 OfficeCLI 需要访问 GitHub Releases）:
#   docker build --build-arg GITHUB_TOKEN=ghp_xxx -t office-agent .
#   # 若本地已执行过 scripts/fetch_officecli.py，bin/officecli 会被直接
#   # 复用，构建时无需再访问 GitHub。
#
# 运行:
#   # Web API（默认入口，监听 8000）:
#   docker run --rm -p 8000:8000 -e LLM_API_KEY=sk-xxx \
#       -v "$PWD/output:/app/output" office-agent
#   # 交互式 CLI:
#   docker run --rm -it -e LLM_API_KEY=sk-xxx \
#       -v "$PWD/output:/app/output" office-agent office-agent
# ============================================================

# ---------- 构建阶段：装依赖 + 准备 OfficeCLI ----------
FROM python:3.13-slim-bookworm AS builder

# uv 直接从官方镜像拷贝二进制
COPY --from=ghcr.io/astral-sh/uv:0.9 /uv /uvx /bin/

ENV UV_COMPILE_BYTECODE=1 \
    UV_LINK_MODE=copy \
    UV_PYTHON_DOWNLOADS=never

WORKDIR /app

# 先只拷依赖清单，充分利用 Docker 层缓存
COPY pyproject.toml uv.lock ./
RUN --mount=type=cache,target=/root/.cache/uv \
    uv sync --frozen --no-dev --no-install-project

# 再拷完整项目并安装项目本身（src 布局，editable 安装指向 /app/src）
COPY . .
RUN --mount=type=cache,target=/root/.cache/uv \
    uv sync --frozen --no-dev

# 下载 OfficeCLI 二进制到 bin/（构建上下文里已有则直接复用）
# 注意：不要在本阶段运行 officecli——builder 没装 ICU，会直接崩溃
ARG GITHUB_TOKEN
RUN if [ ! -x bin/officecli ]; then \
        GITHUB_TOKEN="$GITHUB_TOKEN" .venv/bin/python scripts/fetch_officecli.py; \
    fi

# ---------- 运行阶段 ----------
FROM python:3.13-slim-bookworm AS runtime

# officecli 是自包含 .NET 二进制，运行需要 ICU / libstdc++
RUN apt-get update \
    && apt-get install -y --no-install-recommends \
        ca-certificates \
        libicu72 \
        libstdc++6 \
    && rm -rf /var/lib/apt/lists/*

# 非 root 运行（--create-home：.NET 单文件自解压需要可写的 HOME）
RUN useradd --create-home --uid 1000 app

WORKDIR /app

# 整体拷贝 /app：.venv、src/、template/、bin/officecli、pyproject.toml
COPY --from=builder --chown=app:app /app /app

# 生成文档的落地目录（建议挂载卷持久化）
RUN mkdir -p /app/output && chown app:app /app/output

ENV PATH="/app/.venv/bin:$PATH" \
    OUTPUT_DIR=/app/output \
    API_HOST=0.0.0.0 \
    API_PORT=8000

USER app

# 构建期冒烟测试：officecli 可执行（同时预热 .NET 单文件自解压缓存）
RUN bin/officecli --version

EXPOSE 8000

HEALTHCHECK --interval=30s --timeout=5s --start-period=10s --retries=3 \
    CMD python -c "import urllib.request,os; urllib.request.urlopen(f'http://127.0.0.1:{os.environ.get(\"API_PORT\",\"8000\")}/health', timeout=4)"

# 默认启动 Web API；交互式 CLI 见文件头部注释
CMD ["office-agent-api"]
