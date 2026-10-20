# Excel AI Solver

A macOS menu bar application that inspects Excel statistics worksheets (Estonian
"tööleht" exercises), plans the solution with Claude, computes exact results with
NumPy/SciPy, and writes a solved copy of the workbook.

The AI never does arithmetic — it only identifies tasks and plans the method. All
numbers are computed deterministically in Python, so results are reproducible and
verifiable.

## Architecture

```
src/
  app.py          rumps menu bar app + state machine (IDLE→RUNNING→PAUSED/COMPLETED/FAILED)
  config.py       settings (JSON) + API key in macOS Keychain + tkinter settings window
  schemas.py      pydantic models + tool_use JSON schemas for structured AI output
  math_engine.py  pure NumPy/SciPy statistics (correlation, regression, CIs, tests)
  workbook.py     .xls inspection (xlrd), .xlsx output + charts (openpyxl), xlwings (optional)
  ai_client.py    Claude client: analysis + revision questions, token/cost tracking, budget, retries
  solver.py       pipeline inspect→analyze(AI)→solve(Python)→verify→write; Immediate + Gradual modes
  job_manager.py  SQLite job-state persistence (pause/resume/cancel/recovery)
```

The pipeline:

1. **Inspect** — read the workbook; extract numeric columns, task prose, detected task types.
2. **Analyze (1 AI call)** — Claude returns a structured `WorkbookAnalysis` (tasks, data ranges, Excel formulas, Estonian explanations) via `tool_use`.
3. **Solve (Python)** — each `TaskPlan` is mapped to a `math_engine` function.
4. **Verify** — sanity checks (e.g. −1 ≤ r ≤ 1, 0 ≤ R² ≤ 1).
5. **Write** — a `<name>_lahendatud.xlsx` copy with computed cells, formulas, explanations, and charts.

### Modes

- **Lahenda kohe** (Solve now) — `ImmediateSolver` runs every task at once.
- **Lahenda järk-järgult** (Solve gradually) — `GradualSolver` runs one batch per interval (default 300 s), supporting pause/resume/cancel.

## Setup

```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

Set the API key either in the Settings window (stored in Keychain) or via the
`ANTHROPIC_API_KEY` environment variable.

## Run

```bash
source venv/bin/activate
python run_app.py
```

A 📊 icon appears in the menu bar. With Microsoft Excel open, the app targets the
active workbook; otherwise it shows a file picker.

## Configuration

`~/Library/Application Support/ExcelAISolver/config.json`

| Field | Default | Meaning |
|---|---|---|
| `model` | `claude-sonnet-4-5` | Claude model ID |
| `budget_per_workbook` | `1.0` | USD cost cap per workbook (enforced before each AI call) |
| `pricing` | per-model `{input, output}` $/1M tokens | cost estimation table |
| `gradual_interval_seconds` | `300` | seconds between gradual-mode batches |

API key: macOS Keychain (`service=ExcelAISolver`, `username=anthropic_api_key`),
falling back to `ANTHROPIC_API_KEY`.

Jobs: `~/Library/Application Support/ExcelAISolver/jobs.db` (SQLite).

## Tests

```bash
source venv/bin/activate
python -m pytest tests/ -v
```

The math engine is validated against the bundled `samples/tooleht5.xls`:

| Quantity | Reference |
|---|---|
| Pearson r | 0.88498417 |
| Intercept a | 4.04725078 |
| Slope b | 0.85251932 |
| R² | 0.78319698 |
| ŷ(x=12.5) | 14.70374226 |

Tests never call the real Claude API — the AI client is mocked.

## Build a standalone .app

```bash
bash build_app.sh
# -> dist/Excel AI Solver.app
```

The bundle sets `LSUIElement=true` (menu-bar-only, no Dock icon),
`CFBundleIdentifier=com.excelaisolver.app`, and an Apple Events usage description
for controlling Excel.

## Notes

- `xlrd>=2.0` reads only `.xls`; output is always `.xlsx` via `openpyxl`, so Excel
  is not required to produce results.
- `xlwings` is optional — every use is guarded, so the app works without a running
  Excel (falling back to a file picker and skipping the live open/refresh).
