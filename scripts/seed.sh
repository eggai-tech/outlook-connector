#!/usr/bin/env bash
# Seed config.yaml and .env from the checked-in examples if they are missing.
set -euo pipefail
cd "$(dirname "$0")/.."

if [[ ! -f config.yaml ]]; then
    cp config.yaml.example config.yaml
    echo "seeded config.yaml from config.yaml.example — edit it (mailbox, bus)"
fi

if [[ ! -f .env ]]; then
    cp env.example .env
    echo "seeded .env from env.example — fill in the AZURE_* credentials"
fi
