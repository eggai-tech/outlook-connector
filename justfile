# outlook-connector — local build & run.
#
#   just up      bring the whole stack up from scratch
#   just         list every recipe

set shell := ["bash", "-euo", "pipefail", "-c"]

# List the available recipes.
default:
    @just --list --unsorted

# --- lifecycle ---------------------------------------------------------------

# Bring the stack up from scratch: seed config, build the image, start everything.
up: init build
    docker compose up -d
    @echo
    @just ps
    @echo
    @echo "Follow the connector with:  just logs"

# Same, but stay attached to the logs (Ctrl-C stops the stack).
up-fg: init build
    docker compose up --abort-on-container-exit

# Stop and remove the containers, keeping the postgres/redis volumes.
down:
    docker compose down --remove-orphans

# Stop, wipe the postgres/redis volumes and the built image, then start clean.
reset: nuke up

# Destroy everything this stack owns: containers, volumes and the local image.
nuke:
    docker compose down --remove-orphans --volumes --rmi local

# Restart just the connector (picks up an edited config.yaml).
restart:
    docker compose restart connector

# --- build -------------------------------------------------------------------

# Build the connector image.
build:
    docker compose build connector

# Rebuild from zero, ignoring every cached layer.
rebuild:
    docker compose build --no-cache --pull connector

# --- inspection --------------------------------------------------------------

# Show the state of every service.
ps:
    docker compose ps

# Tail the connector logs (pass a service name to follow a different one).
logs service="connector":
    docker compose logs -f --tail=100 {{service}}

# Open a shell inside the running connector container.
sh:
    docker compose exec connector bash

# Print the fully-resolved compose configuration.
config:
    docker compose config

# --- local development -------------------------------------------------------

# Install the whole workspace (connector + libs/outlook-helper) with dev deps.
sync:
    uv sync --all-packages --all-groups

# Run the connector test suite on the host (no containers involved).
test *args:
    uv run pytest {{args}}

# The two suites cannot share a pytest run: each defines a tests/test_config.py
# and neither has an __init__.py, so collection collides on module name. Running
# from the member's directory gives it its own rootdir and config.

# Run the vendored outlook-helper test suite.
test-helper *args:
    uv run --directory libs/outlook-helper pytest {{args}}

# Both suites.
test-all: test test-helper

# `dev` honours config.yaml as-is (kafka by default) and starts no containers,
# so whatever broker it names has to be reachable already. Use `just run` to
# point a host-side connector at the compose redis instead.

# Install everything and start the connector on the host: seed, sync, run.
dev: init sync
    #!/usr/bin/env bash
    set -euo pipefail
    # The connector itself does not read .env (only compose does), so export it
    # here — that is where the AZURE_* credentials live.
    set -a; source .env; set +a
    exec uv run outlook-connector

# Run the connector on the host against the compose redis (needs `just up`).
run:
    BUS__TRANSPORT=redis BUS__BROKER_URL=redis://localhost:${REDIS_HOST_PORT:-36379}/0 uv run outlook-connector

# --- prerequisites -----------------------------------------------------------

# Seed config.yaml and .env from the checked-in examples if they are missing.
init:
    @bash scripts/seed.sh
