#!/bin/sh
set -eu

PROJECT_DIRECTORY=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
cd "$PROJECT_DIRECTORY"

if ! command -v python3 >/dev/null 2>&1; then
    printf '%s\n' "Python 3 is required. Install it from https://www.python.org/downloads/macos/ and run this file again."
    exit 1
fi

python3 -c 'import sys; raise SystemExit("Python 3.10 or newer is required.") if sys.version_info < (3, 10) else None'

if [ ! -x ".venv/bin/python" ]; then
    python3 -m venv .venv
fi

".venv/bin/python" -m pip install --disable-pip-version-check --upgrade pip
".venv/bin/python" -m pip install --disable-pip-version-check -r requirements.txt
exec ".venv/bin/python" app.py "$@"