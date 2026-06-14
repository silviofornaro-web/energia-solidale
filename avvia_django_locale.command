#!/bin/bash

set -e

cd "$(dirname "$0")"

if [ -x "./.venv-django/bin/python" ] && ./.venv-django/bin/python -c "import django" >/dev/null 2>&1; then
    PYTHON_BIN="./.venv-django/bin/python"
elif [ -d "./.venv-django/lib/python3.12/site-packages" ] && [ -x "/opt/homebrew/opt/python@3.12/bin/python3.12" ]; then
    PYTHON_BIN="/opt/homebrew/opt/python@3.12/bin/python3.12"
    export PYTHONPATH="$(pwd)/.venv-django/lib/python3.12/site-packages${PYTHONPATH:+:$PYTHONPATH}"
elif [ -x "./venv/bin/python" ] && ./venv/bin/python -c "import django" >/dev/null 2>&1; then
    PYTHON_BIN="./venv/bin/python"
else
    echo "Django non e disponibile in un ambiente Python valido. Ripristina .venv-django oppure reinstalla requirements-django.txt."
    exit 1
fi

"$PYTHON_BIN" manage.py migrate --noinput
"$PYTHON_BIN" manage.py runserver 127.0.0.1:8000
