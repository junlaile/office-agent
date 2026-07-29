#!/usr/bin/env bash
set -euo pipefail

# Install uv
curl -LsSf https://astral.sh/uv/install.sh | sh
export PATH="$HOME/.local/bin:$PATH"

# Install Python 3.13 and sync all dependencies (including dev)
cd /workspace
uv python install 3.13
uv sync --dev

echo "✓ Environment ready. Run: uv run pytest -q"
