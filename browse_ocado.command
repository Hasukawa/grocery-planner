#!/bin/bash
set -u
cd "$(dirname "$0")"
source .venv/bin/activate
python browse_ocado.py
