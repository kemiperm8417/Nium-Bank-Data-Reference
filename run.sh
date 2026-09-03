#!/usr/bin/env bash
# Bank Reference Data — macOS / Linux launcher.
#
#   ./run.sh                              start the web app (opens in your browser)
#   ./run.sh --countries US,GB,IN         export straight to Excel, no browser
#   ./run.sh --sepa --out sepa.xlsx       any refdata.py arguments work here
#
# First run creates a private .venv in this folder and installs requirements.
# You must be on the Nium VPN — the API is not reachable from elsewhere.
set -euo pipefail
cd "$(dirname "$0")"

if ! command -v python3 >/dev/null 2>&1; then
  echo "python3 not found. Install it from https://www.python.org/downloads/ and re-run." >&2
  exit 1
fi

if [ ! -x .venv/bin/python ]; then
  echo "Creating virtual environment (.venv)…"
  python3 -m venv .venv
fi
# shellcheck disable=SC1091
source .venv/bin/activate

if [ ! -f .venv/.deps-installed ] || [ requirements.txt -nt .venv/.deps-installed ]; then
  echo "Installing requirements…"
  pip install --quiet --upgrade pip
  pip install --quiet -r requirements.txt
  touch .venv/.deps-installed
fi

export STREAMLIT_BROWSER_GATHER_USAGE_STATS=false

if [ $# -gt 0 ]; then
  exec python refdata.py "$@"
fi

echo "Starting the web app — press Ctrl+C to stop."
exec streamlit run app.py
