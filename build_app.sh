#!/bin/bash
set -e
cd /Users/v8zoqrk/excel
source venv/bin/activate
pyinstaller "Excel AI Solver.spec" --clean
echo "Build complete: dist/Excel AI Solver.app"
