# syntax=docker/dockerfile:1

# --- builder: resolve and install dependencies with uv -----------------------
FROM python:3.13-slim AS builder

# uv, pinned via the official distroless image
COPY --from=ghcr.io/astral-sh/uv:0.11.26 /uv /uvx /bin/

ENV UV_COMPILE_BYTECODE=1 \
    UV_LINK_MODE=copy \
    UV_PYTHON_DOWNLOADS=0

WORKDIR /app

# Third-party dependencies only. Every workspace member's manifest has to be
# readable for `--frozen` to validate the lock, but no first-party source is
# installed yet, so this layer stays cached until a manifest or the lock changes.
# No git/ssh: `outlook-helper` is in-tree at libs/, not a private git dependency.
COPY pyproject.toml uv.lock README.md ./
COPY libs/outlook-helper/pyproject.toml libs/outlook-helper/README.md ./libs/outlook-helper/
RUN --mount=type=cache,target=/root/.cache/uv \
    uv sync --frozen --no-dev --no-install-workspace

# First-party source, then install outlook-connector and outlook-helper.
# `--no-editable` copies them into the venv instead of linking back to /app, so
# the runtime stage needs nothing but the venv itself.
COPY src ./src
COPY libs ./libs
RUN --mount=type=cache,target=/root/.cache/uv \
    uv sync --frozen --no-dev --no-editable

# --- runtime -----------------------------------------------------------------
FROM python:3.13-slim

WORKDIR /app

# The venv is self-contained: outlook_connector and outlook_helper are installed
# into it, so no application source is copied into this stage.
COPY --from=builder /app/.venv /app/.venv

ENV PATH="/app/.venv/bin:$PATH" \
    PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1

# Drop privileges. Nothing is written to /app at runtime; config.yaml is
# bind-mounted read-only by docker-compose.
RUN useradd --create-home --uid 10001 connector
USER connector

# Runs the long-running connector; exits cleanly on SIGTERM/SIGINT.
# No CMD: the compose service must not append arguments to this entrypoint.
ENTRYPOINT ["outlook-connector"]
