#!/usr/bin/env bash
set -o errexit

PYTHON_BIN="${PYTHON_BIN:-python3}"

"$PYTHON_BIN" -m pip install --upgrade pip
"$PYTHON_BIN" -m pip install -r requirements.txt

if [ -f requirements-django.txt ]; then
  "$PYTHON_BIN" -m pip install -r requirements-django.txt
fi

if [ -f manage.py ]; then
  "$PYTHON_BIN" manage.py collectstatic --no-input
fi
