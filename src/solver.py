"""Orchestration pipeline: inspect -> analyze (AI) -> solve (Python) -> verify -> write.

ImmediateSolver runs everything at once. GradualSolver runs one batch per
interval (default 300s). Both share the same pipeline and support progress
callbacks so the menu bar can report status without blocking.
"""

from __future__ import annotations

import time
from typing import Any, Callable

from . import math_engine as me
from . import workbook as wb_mod
from . import excel_writer as ew
from .schemas import TaskPlan, TaskResult, TaskType, WorkbookAnalysis

ProgressCallback = Callable[[str, int, int], None]


def _noop(message: str, done: int, total: int) -> None:  # pragma: no cover
    pass


def _resolve_xy(task: TaskPlan, inspection: wb_mod.WorkbookInspection) -> tuple[list[float], list[float]]:
    """Resolve the x and y numeric columns for a task.

    Preference order: explicit inputs on the task (x/y lists), then the first
    two numeric columns of the task's source sheet.
    """
    inputs = task.inputs or {}
    x = inputs.get("x")
    y = inputs.get("y")
    if isinstance(x, list) and isinstance(y, list) and x and y:
        return [float(v) for v in x], [float(v) for v in y]

    sheet = next(
        (s for s in inspection.sheets if s.name == task.source_sheet),
        inspection.sheets[0] if inspection.sheets else None,
    )
    if sheet is None or len(sheet.numeric_columns) < 2:
        raise ValueError(f"Cannot resolve x/y data for task {task.task_id}")
    cols = list(sheet.numeric_columns.values())
    # Trim to equal length in case columns differ by a trailing cell.
    n = min(len(cols[0]), len(cols[1]))
    return cols[0][:n], cols[1][:n]


def solve_task(task: TaskPlan, inspection: wb_mod.WorkbookInspection) -> TaskResult:
    """Run the Python computation for one planned task."""
    result = TaskResult(
        task_id=task.task_id,
        task_type=task.task_type,
        source_sheet=task.source_sheet,
        description=task.description,
        output_location=task.output_location or {"sheet": "Arvutused", "start_row": 1},
        explanation=task.explanation,
    )

    try:
        if task.task_type == TaskType.CORRELATION:
            x, y = _resolve_xy(task, inspection)
            r, p = me.pearson_correlation(x, y)
            result.values = {"r": r, "p_value": p}

        elif task.task_type == TaskType.REGRESSION:
            x, y = _resolve_xy(task, inspection)
            result.values = me.linear_regression(x, y)

        elif task.task_type == TaskType.CONFIDENCE_INTERVAL:
            x, y = _resolve_xy(task, inspection)
            result.values = me.regression_confidence_intervals(x, y)

        elif task.task_type in (TaskType.PREDICTION, TaskType.PREDICTION_INTERVAL):
            x, y = _resolve_xy(task, inspection)
            reg = me.linear_regression(x, y)
            x_new = float((task.inputs or {}).get("x_new", 0.0))
            y_pred = me.predict_y(reg["intercept"], reg["slope"], x_new)
            vals: dict[str, Any] = {
                "intercept": reg["intercept"],
                "slope": reg["slope"],
                "x_new": x_new,
                "y_pred": y_pred,
            }
            if task.task_type == TaskType.PREDICTION_INTERVAL:
                lo, hi, extrap = me.prediction_interval(x, y, x_new)
                vals.update({"pi_lower": lo, "pi_upper": hi, "is_extrapolation": extrap})
            result.values = vals

        elif task.task_type == TaskType.HYPOTHESIS_TEST:
            inputs = task.inputs or {}
            if "x2" in inputs:
                result.values = me.t_test_two_sample(inputs["x1"], inputs["x2"])
            else:
                data = inputs.get("data") or _resolve_xy(task, inspection)[0]
                mu0 = float(inputs.get("mu0", 0.0))
                result.values = me.t_test_one_sample(data, mu0)

        elif task.task_type == TaskType.DESCRIPTIVE_STATS:
            x, _ = _resolve_xy(task, inspection)
            result.values = me.descriptive_stats(x)

        elif task.task_type == TaskType.CHART:
            result.values = {"chart": True}

        else:
            result.values = {"note": "No computation for task type 'other'."}

        # Materialize the planned Excel formulas with their computed values.
        result.written_cells = _materialize_cells(task, result.values)
        result.verified = _verify(task, result)
    except Exception as exc:
        result.error = str(exc)
        result.verified = False
    return result


def _materialize_cells(task: TaskPlan, values: dict[str, Any]) -> list[dict[str, Any]]:
    """Pair the AI's planned formulas with computed values for writing."""
    cells: list[dict[str, Any]] = []
    for item in task.excel_formulas:
        cells.append(
            {
                "cell": item.get("cell"),
                "formula": item.get("formula"),
            }
        )
    # If the plan had no formulas, drop the scalar values into the output sheet.
    if not cells:
        start_row = int((task.output_location or {}).get("start_row", 1))
        row = start_row
        for key, val in values.items():
            if isinstance(val, (int, float, str, bool)):
                cells.append({"cell": f"A{row}", "value": f"{key}"})
                cells.append({"cell": f"B{row}", "value": val})
                row += 1
    return cells


def _verify(task: TaskPlan, result: TaskResult) -> bool:
    """Lightweight sanity checks on computed values."""
    v = result.values
    if result.error:
        return False
    if task.task_type == TaskType.CORRELATION:
        r = v.get("r")
        return r is not None and -1.0 <= r <= 1.0
    if task.task_type == TaskType.REGRESSION:
        r2 = v.get("r_squared")
        return r2 is not None and 0.0 <= r2 <= 1.0
    return bool(v)


class _BasePipeline:
    def __init__(self, ai_client, config, progress: ProgressCallback | None = None):
        self.ai = ai_client
        self.config = config
        self.progress = progress or _noop

    def inspect(self, path: str) -> wb_mod.WorkbookInspection:
        self.progress("Inspecting workbook", 0, 1)
        return wb_mod.inspect_workbook(path)

    def analyze(self, inspection: wb_mod.WorkbookInspection) -> WorkbookAnalysis:
        self.progress("Analyzing with AI", 0, 1)
        return self.ai.analyze_workbook(inspection.to_ai_payload())

    def write_results(
        self,
        inspection: wb_mod.WorkbookInspection,
        results: list[TaskResult],
        source_path: str,
    ) -> str:
        out_path = wb_mod.create_output_copy(source_path)
        wb = wb_mod.build_output_workbook(inspection, out_path)
        for res in results:
            wb_mod.write_task_result(wb, res)
        return wb_mod.save_workbook(wb, out_path)


class ImmediateSolver(_BasePipeline):
    """Run the full pipeline for all tasks in one go."""

    def run(self, path: str) -> dict[str, Any]:
        inspection = self.inspect(path)
        analysis = self.analyze(inspection)
        total = len(analysis.tasks)
        results: list[TaskResult] = []
        for i, task in enumerate(analysis.tasks):
            self.progress(f"Solving {task.task_id}", i, total)
            results.append(solve_task(task, inspection))

        self.progress("Kirjutan tulemused", total, total)

        # Backup original before modifying
        backup_path = ew.backup_workbook(path)

        # Write solution sheets in-place into the active workbook
        xw_book = wb_mod.get_xw_book(path)
        sheet_report = ew.write_all_sheets(xw_book, analysis, results, inspection)

        # QA verification
        original_sheet = inspection.sheets[0].name if inspection.sheets else ""
        qa_issues = ew.verify_workbook(xw_book, original_sheet, sheet_report["written"])

        ew.save_workbook(xw_book)
        self.progress("Completed", total, total)

        return {
            "output_path": path,
            "backup_path": backup_path,
            "results": results,
            "analysis": analysis,
            "cost": self.ai.total_cost,
            "sheet_report": sheet_report,
            "qa_issues": qa_issues,
        }


class GradualSolver(_BasePipeline):
    """Run tasks in batches, one batch per interval (default 300s)."""

    def __init__(self, ai_client, config, progress=None, batch_size: int = 1, sleep_fn=time.sleep):
        super().__init__(ai_client, config, progress)
        self.batch_size = batch_size
        self.sleep_fn = sleep_fn
        self.interval = config.gradual_interval_seconds

    def run(self, path: str, should_continue: Callable[[], bool] | None = None) -> dict[str, Any]:
        should_continue = should_continue or (lambda: True)
        inspection = self.inspect(path)
        analysis = self.analyze(inspection)
        tasks = analysis.tasks
        total = len(tasks)

        # Backup before any modification
        backup_path = ew.backup_workbook(path)
        xw_book = wb_mod.get_xw_book(path)

        results: list[TaskResult] = []
        idx = 0
        while idx < total:
            if not should_continue():
                self.progress("Paused", idx, total)
                break
            batch = tasks[idx : idx + self.batch_size]
            for task in batch:
                self.progress(f"Solving {task.task_id}", idx, total)
                results.append(solve_task(task, inspection))
                idx += 1

            # Write all completed results so far (sheets regenerated each batch).
            ew.write_all_sheets(xw_book, analysis, results, inspection)
            ew.save_workbook(xw_book)

            if idx < total:
                self.sleep_fn(self.interval)

        original_sheet = inspection.sheets[0].name if inspection.sheets else ""
        sheet_report = ew.write_all_sheets(xw_book, analysis, results, inspection)
        qa_issues = ew.verify_workbook(xw_book, original_sheet, sheet_report["written"])
        ew.save_workbook(xw_book)

        self.progress("Completed" if idx >= total else "Paused", idx, total)
        return {
            "output_path": path,
            "backup_path": backup_path,
            "results": results,
            "analysis": analysis,
            "cost": self.ai.total_cost,
            "completed": idx >= total,
            "sheet_report": sheet_report,
            "qa_issues": qa_issues,
        }
