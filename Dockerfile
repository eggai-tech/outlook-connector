# syntax=docker/dockerfile:1

# --- builder: resolve and install dependencies with uv -----------------------
FROM python:3.13-slim AS builder

# uv, pinned via the official distroless image
COPY --from=ghcr.io/astral-sh/uv:0.11.26 /uv /uvx /bin/

# git + ssh are needed to fetch the private `outlook-helper` dependency
# (ssh://git@github.com/eggai-tech/outlook-helper.git) during `uv sync`.
RUN apt-get update \
    && apt-get install -y --no-install-recommends git openssh-client \
    && rm -rf /var/lib/apt/lists/* \
    && mkdir -p -m 0700 /root/.ssh \
    && ssh-keyscan github.com >> /root/.ssh/known_hosts

ENV UV_COMPILE_BYTECODE=1 \
    UV_LINK_MODE=copy \
    UV_PYTHON_DOWNLOADS=0

WORKDIR /app

# Install dependencies only (the project itself is a uv "virtual" project, so
# nothing to build). `--mount=type=ssh` forwards the SSH agent for the private
# git dependency; layer is cached until uv.lock/pyproject.toml change.
RUN --mount=type=cache,target=/root/.cache/uv \
    --mount=type=bind,source=uv.lock,target=uv.lock \
    --mount=type=bind,source=pyproject.toml,target=pyproject.toml \
    --mount=type=ssh \
    uv sync --frozen --no-dev --no-install-project

# --- runtime -----------------------------------------------------------------
FROM python:3.13-slim

WORKDIR /app

# Bring in the resolved virtualenv, then the application source.
COPY --from=builder /app/.venv /app/.venv
COPY . /app

ENV PATH="/app/.venv/bin:$PATH" \
    PYTHONUNBUFFERED=1

# Runs the long-running connector; exits cleanly on SIGTERM/SIGINT.
ENTRYPOINT ["python", "main.py"]
