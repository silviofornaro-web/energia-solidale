#!/usr/bin/env bash
set -o errexit

PYTHON_BIN="${PYTHON_BIN:-python3}"

"$PYTHON_BIN" -m pip install -r requirements-django.txt
"$PYTHON_BIN" manage.py collectstatic --no-input
