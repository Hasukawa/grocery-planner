#!/bin/bash
set -u
cd "$(dirname "$0")"
if [ ! -d .venv ]; then
  echo "Virtualenv not found at .venv — run the setup steps in README.md first."
  echo
  read -n1 -s -r -p "Press any key to close."
  exit 1
fi
source .venv/bin/activate
python ocado_tuesday.py
status=$?
echo
read -n1 -s -r -p "Done (exit $status). Press any key to close."
