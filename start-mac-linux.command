#!/bin/bash
# Starts the CSV cleaner in your browser. First start takes 1-2 minutes.
cd "$(dirname "$0")"
if ! command -v uv >/dev/null 2>&1; then
  echo "Installing uv (Python package manager, one-time)..."
  curl -LsSf https://astral.sh/uv/install.sh | sh
  export PATH="$HOME/.local/bin:$PATH"
fi
uv run --python 3.12 --with-requirements requirements.txt streamlit run app.py
