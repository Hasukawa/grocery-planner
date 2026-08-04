#!/bin/bash
set -u
cd "$(dirname "$0")"
source .venv/bin/activate
python probe_trolley.py
status=$?
echo
read -n1 -s -r -p "Done (exit $status). Press any key to close."
