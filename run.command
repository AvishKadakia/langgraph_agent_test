#!/bin/bash
# ░░ Agent QA Eval Suite — macOS launcher ░░
# Double-click this file to run the eval suite locally. No make or Docker needed.
# First run sets things up (a minute); after that it starts in seconds.
cd "$(dirname "$0")"

echo "🧪  Agent QA Eval Suite"
echo "------------------------------------------"

# 1) Python 3 ---------------------------------------------------------------
if ! command -v python3 >/dev/null 2>&1; then
  echo "❌  Python 3 is required but wasn't found."
  echo "    Install it from https://www.python.org/downloads/  (or: brew install python)"
  echo "    then double-click this file again."
  read -r -p "Press Return to close…"; exit 1
fi

# 2) Azure CLI --------------------------------------------------------------
if ! command -v az >/dev/null 2>&1; then
  echo "❌  Azure CLI is required but wasn't found."
  echo "    Install it: https://learn.microsoft.com/cli/azure/install-azure-cli-macos"
  echo "    (or: brew install azure-cli) then double-click this file again."
  read -r -p "Press Return to close…"; exit 1
fi

# 3) One-time setup: virtual env + dependencies -----------------------------
if [ ! -d .venv ]; then
  echo "📦  First-time setup (installing — about a minute)…"
  python3 -m venv .venv || { echo "❌  Could not create the environment."; read -r -p "Press Return…"; exit 1; }
  ./.venv/bin/python -m pip install --quiet --upgrade pip
  ./.venv/bin/python -m pip install --quiet -r requirements.txt || { echo "❌  Dependency install failed."; read -r -p "Press Return…"; exit 1; }
fi

# 4) Config -----------------------------------------------------------------
[ -f .env ] || { cp .env.example .env; echo "📝  Created .env from the template."; }

# 5) Azure sign-in (only if needed) -----------------------------------------
if ! az account show >/dev/null 2>&1; then
  echo "🔐  Opening Azure sign-in in your browser…"
  az login >/dev/null || { echo "❌  Azure sign-in failed."; read -r -p "Press Return…"; exit 1; }
fi
echo "✅  Signed in as: $(az account show --query user.name -o tsv 2>/dev/null)"

# 6) Launch -----------------------------------------------------------------
PORT=$(grep -E '^APP_PORT=' .env | cut -d= -f2); PORT=${PORT:-8080}
echo ""
echo "🚀  Opening  http://localhost:$PORT"
echo "    Keep this window open while you use the app. Close it to stop."
echo "------------------------------------------"
( sleep 3; open "http://localhost:$PORT" ) &
exec ./.venv/bin/python -m uvicorn app.main:app --host 127.0.0.1 --port "$PORT"
