#!/usr/bin/env bash
set -o errexit

PYTHON_BIN="${PYTHON_BIN:-python3}"

"$PYTHON_BIN" -m pip install --upgrade pip
"$PYTHON_BIN" -m pip install -r requirements-django.txt

if [ -f manage.py ]; then
  "$PYTHON_BIN" manage.py migrate --no-input
  "$PYTHON_BIN" manage.py collectstatic --no-input
fi
