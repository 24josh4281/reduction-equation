#!/bin/sh
set -e
cd "$(dirname "$0")"
python3 -m venv .venv
.venv/bin/python -m pip install --upgrade pip --quiet
.venv/bin/python -m pip install -r requirements.txt --quiet
PYTHONUTF8=1 .venv/bin/python scripts/monitoring_plan.py check examples/plan_example.txt
echo ""
echo "설치 완료. 편집기 실행: .venv/bin/python scripts/equation_editor/server.py"
