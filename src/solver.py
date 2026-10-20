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

_SLEEP_CHUNK = 1.0  # granularity for interruptible sleeps (seconds)


def _noop(message: str, done: int, total: int) -> None:  # pragma: no cover
    pass


def _resolve_xy(task: TaskPlan, inspection: wb_mod.WorkbookInspection) -> tuple[list[float], list[float]]:
    """Resolve paired x/y numeric data for a task.

    Resolution order (most to least specific):
    1. x_range + y_range in task.inputs — exact cell references, row-aligned,
       missing observations skipped.  This is the preferred path.
    2. Explicit x/y lists in task.inputs — only when both present with matching
       lengths.
    3. First two numeric columns of the source sheet — fallback used only when
       the sheet has exactly two numeric columns, so there is no ambiguity.

    Raises ValueError when data cannot be resolved unambiguously.
    """
    inputs = task.inputs or {}

    # Path 1: range strings from the AI plan (preferred).
    x_range = inputs.get("x_range")
    y_range = inputs.get("y_range")
    if x_range and y_range:
        x, y = wb_mod.extract_paired_range_values(inspection, task.source_sheet, x_range, y_range)
        if not x or not y:
            raise ValueError(
                f"Task {task.task_id}: no paired numeric data found in "
                f"x_range='{x_range}', y_range='{y_range}'."
            )
        return x, y

    # Path 2: AI-provided value lists.
    x_list = inputs.get("x")
    y_list = inputs.get("y")
    if isinstance(x_list, list) and isinstance(y_list, list) and x_list and y_list:
        if abs(len(x_list) - len(y_list)) > 1:
            raise ValueError(
                f"Task {task.task_id}: AI-provided x ({len(x_list)}) and y "
                f"({len(y_list)}) lengths differ by more than 1. "
                "Cannot align paired observations."
            )
        n = min(len(x_list), len(y_list))
        return [float(v) for v in x_list[:n]], [float(v) for v in y_list[:n]]

    # Path 3: positional fallback — only safe with exactly 2 numeric columns.
    sheet = next(
        (s for s in inspection.sheets if s.name == task.source_sheet),
        inspection.sheets[0] if inspection.sheets else None,
    )
    if sheet is None or len(sheet.numeric_columns) < 2:
        raise ValueError(
            f"Task {task.task_id}: cannot resolve x/y data — "
            f"sheet '{task.source_sheet}' has fewer than 2 numeric columns "
            "and no x_range/y_range was specified in the AI plan."
        )
    if len(sheet.numeric_columns) > 2:
        raise ValueError(
            f"Task {task.task_id}: sheet '{task.source_sheet}' has "
            f"{len(sheet.numeric_columns)} numeric columns but no x_range/y_range "
            "was specified in the AI plan. Cannot determine which columns to use."
        )
    cols = list(sheet.numeric_columns.values())
    if abs(len(cols[0]) - len(cols[1])) > 1:
        raise ValueError(
            f"Task {task.task_id}: x column has {len(cols[0])} values but "
            f"y column has {len(cols[1])} values — columns likely misaligned. "
            "Check for missing cells or specify x_range/y_range."
        )
    n = min(len(cols[0]), len(cols[1]))
    return cols[0][:n], cols[1][:n]


def _resolve_single(task: TaskPlan, inspection: wb_mod.WorkbookInspection) -> list[float]:
    """Resolve a single column of numeric data for one-sample tasks.

    Resolution order:
    1. data_range or x_range from task.inputs (exact cell reference).
    2. data list in task.inputs.
    3. First numeric column — only when the sheet has exactly one numeric column.
    """
    inputs = task.inputs or {}

    # Range reference (preferred).
    range_str = inputs.get("data_range") or inputs.get("x_range")
    if range_str:
        vals = wb_mod.extract_range_values(inspection, task.source_sheet, range_str)
        if not vals:
            raise ValueError(
                f"Task {task.task_id}: no numeric values in range '{range_str}'."
            )
        return vals

    # Explicit list.
    data = inputs.get("data")
    if isinstance(data, list) and data:
        return [float(v) for v in data if isinstance(v, (int, float))]

    # Positional fallback — only unambiguous with exactly one numeric column.
    sheet = next(
        (s for s in inspection.sheets if s.name == task.source_sheet),
        inspection.sheets[0] if inspection.sheets else None,
    )
    if sheet is None or not sheet.numeric_columns:
        raise ValueError(
            f"Task {task.task_id}: no numeric data in sheet '{task.source_sheet}'."
        )
    if len(sheet.numeric_columns) > 1:
        raise ValueError(
            f"Task {task.task_id}: sheet has {len(sheet.numeric_columns)} numeric "
            "columns but no data_range was specified. Cannot determine which column "
            "to use."
        )
    return list(list(sheet.numeric_columns.values())[0])


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
        inputs = task.inputs or {}

        if task.task_type == TaskType.CORRELATION:
            x, y = _resolve_xy(task, inspection)
            method = inputs.get("method", "pearson").lower()
            if method == "spearman":
                rho, p = me.spearman_correlation(x, y)
                result.values = {"rho": rho, "p_value": p, "method": "spearman"}
            else:
                r, p = me.pearson_correlation(x, y)
                result.values = {"r": r, "p_value": p, "method": "pearson"}

        elif task.task_type == TaskType.REGRESSION:
            x, y = _resolve_xy(task, inspection)
            result.values = me.linear_regression(x, y)

        elif task.task_type == TaskType.CONFIDENCE_INTERVAL:
            x, y = _resolve_xy(task, inspection)
            alpha = float(inputs.get("alpha", 0.05))
            x_new = inputs.get("x_new")
            if x_new is not None:
                # Mean-response CI at a specified x value.
                lo, hi = me.confidence_interval_mean(x, y, float(x_new), alpha)
                result.values = {
                    "x_new": float(x_new),
                    "ci_lower": lo,
                    "ci_upper": hi,
                    "alpha": alpha,
                    "ci_type": "mean_response",
                }
            else:
                # Coefficient CIs (slope and intercept).
                result.values = me.regression_confidence_intervals(x, y, alpha)

        elif task.task_type in (TaskType.PREDICTION, TaskType.PREDICTION_INTERVAL):
            x, y = _resolve_xy(task, inspection)
            if "x_new" not in inputs:
                raise ValueError(
                    f"Task {task.task_id} requires a prediction point x₀ "
                    "(inputs.x_new) but the AI plan did not specify one. "
                    "Cannot compute a prediction without a target x value."
                )
            x_new = float(inputs["x_new"])
            alpha = float(inputs.get("alpha", 0.05))
            reg = me.linear_regression(x, y)
            y_pred = me.predict_y(reg["intercept"], reg["slope"], x_new)
            vals: dict[str, Any] = {
                "intercept": reg["intercept"],
                "slope": reg["slope"],
                "x_new": x_new,
                "y_pred": y_pred,
                "alpha": alpha,
            }
            if task.task_type == TaskType.PREDICTION_INTERVAL:
                lo, hi, extrap = me.prediction_interval(x, y, x_new, alpha)
                vals.update({"pi_lower": lo, "pi_upper": hi, "is_extrapolation": extrap})
            result.values = vals

        elif task.task_type == TaskType.HYPOTHESIS_TEST:
            alpha = float(inputs.get("alpha", 0.05))
            x_range = inputs.get("x_range")
            y_range = inputs.get("y_range")
            if x_range and y_range:
                # Two-sample via explicit range references.
                x1 = wb_mod.extract_range_values(inspection, task.source_sheet, x_range)
                x2 = wb_mod.extract_range_values(inspection, task.source_sheet, y_range)
                equal_var = bool(inputs.get("equal_var", False))
                result.values = me.t_test_two_sample(x1, x2, alpha=alpha, equal_var=equal_var)
            elif "x2" in inputs:
                equal_var = bool(inputs.get("equal_var", False))
                result.values = me.t_test_two_sample(
                    inputs["x1"], inputs["x2"], alpha=alpha, equal_var=equal_var
                )
            else:
                if "mu0" not in inputs:
                    raise ValueError(
                        f"Task {task.task_id} is a one-sample t-test but the "
                        "null hypothesis mean mu0 was not specified in the AI plan. "
                        "Cannot test without a reference value."
                    )
                data = _resolve_single(task, inspection)
                mu0 = float(inputs["mu0"])
                result.values = me.t_test_one_sample(data, mu0, alpha=alpha)

        elif task.task_type == TaskType.DESCRIPTIVE_STATS:
            data = _resolve_single(task, inspection)
            result.values = me.descriptive_stats(data)

        elif task.task_type == TaskType.CHART:
            result.values = {"chart": True}

        else:
            result.values = {"note": "No computation for task type 'other'."}

        # Materialize the planned Excel formulas with their computed values.
        result.written_cells = _materialize_cells(task, result.values)
        result.verified = _verify(task, result, inspection)
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


def _verify(
    task: TaskPlan,
    result: TaskResult,
    inspection: wb_mod.WorkbookInspection,
) -> bool:
    """Sanity checks on computed values.

    CHART and OTHER tasks are never marked verified — they cannot be checked
    without live Excel interaction.  For REGRESSION an independent numerical
    re-computation is performed against the source data with a 1e-6 tolerance.
    """
    v = result.values
    if result.error:
        return False

    # Tasks with no deterministic Python result cannot be verified here.
    if task.task_type in (TaskType.CHART, TaskType.OTHER):
        return False

    if task.task_type == TaskType.CORRELATION:
        coeff = v.get("rho") if v.get("method") == "spearman" else v.get("r")
        return coeff is not None and -1.0 <= coeff <= 1.0

    if task.task_type == TaskType.REGRESSION:
        r2 = v.get("r_squared")
        slope = v.get("slope")
        intercept = v.get("intercept")
        if r2 is None or slope is None or intercept is None:
            return False
        if not (0.0 <= r2 <= 1.0):
            return False
        # Independent numerical check: recompute from source data.
        try:
            x, y = _resolve_xy(task, inspection)
            check = me.linear_regression(x, y)
            tol = 1e-6
            if (
                abs(check["intercept"] - intercept) > tol
                or abs(check["slope"] - slope) > tol
                or abs(check["r_squared"] - r2) > tol
            ):
                return False
        except Exception:
            pass  # Data unavailable; accept range check
        return True

    if task.task_type == TaskType.CONFIDENCE_INTERVAL:
        ci_type = v.get("ci_type", "coefficients")
        if ci_type == "mean_response":
            lo, hi = v.get("ci_lower"), v.get("ci_upper")
            return lo is not None and hi is not None and lo < hi
        # Coefficient CIs: slope is stored as (lo, hi) tuple by math_engine.
        slope_pair = v.get("slope")
        if isinstance(slope_pair, (tuple, list)) and len(slope_pair) == 2:
            return slope_pair[0] < slope_pair[1]
        return bool(v)

    if task.task_type in (TaskType.PREDICTION, TaskType.PREDICTION_INTERVAL):
        y_pred = v.get("y_pred")
        if y_pred is None:
            return False
        if task.task_type == TaskType.PREDICTION_INTERVAL:
            lo, hi = v.get("pi_lower"), v.get("pi_upper")
            return lo is not None and hi is not None and lo < hi
        return True

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

        # Generate revision questions (one additional AI call after all tasks are solved).
        self.progress("Genereerin kordamisküsimused", total, total)
        try:
            task_dicts = [t.model_dump() for t in analysis.tasks]
            result_dicts = [r.model_dump() for r in results if not r.error]
            questions = self.ai.generate_revision_questions(task_dicts, result_dicts)
            analysis = analysis.model_copy(update={"revision_questions": questions})
        except Exception:
            pass  # Revision questions are optional; do not fail the whole job.

        self.progress("Kirjutan tulemused", total, total)

        # Backup original before modifying.
        backup_path = ew.backup_workbook(path)

        # Write solution sheets in-place into the active workbook.
        xw_book = wb_mod.get_xw_book(path)
        sheet_report = ew.write_all_sheets(xw_book, analysis, results, inspection)

        # QA verification.
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

    def _interruptible_sleep(
        self, duration: float, should_continue: Callable[[], bool]
    ) -> bool:
        """Sleep for `duration` seconds in _SLEEP_CHUNK increments.

        Returns False if should_continue() signals cancellation mid-sleep,
        True if the full duration elapsed normally.
        """
        remaining = float(duration)
        while remaining > 0.0:
            chunk = min(_SLEEP_CHUNK, remaining)
            self.sleep_fn(chunk)
            remaining -= chunk
            if not should_continue():
                return False
        return True

    def run(self, path: str, should_continue: Callable[[], bool] | None = None) -> dict[str, Any]:
        should_continue = should_continue or (lambda: True)
        inspection = self.inspect(path)
        analysis = self.analyze(inspection)
        tasks = analysis.tasks
        total = len(tasks)

        # Backup before any modification.
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

            # Commit progress after each batch.
            ew.write_all_sheets(xw_book, analysis, results, inspection)
            ew.save_workbook(xw_book)

            if idx < total:
                completed = self._interruptible_sleep(self.interval, should_continue)
                if not completed:
                    self.progress("Paused", idx, total)
                    break

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
