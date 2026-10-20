"""Tests for the orchestration pipeline with a mocked AI client."""

import math
from pathlib import Path

import pytest
from unittest.mock import MagicMock

from unittest.mock import Mock

from src import workbook as wb_mod
from src import excel_writer as ew
from src.config import Config
from src.schemas import TaskPlan, TaskType, WorkbookAnalysis
from src.solver import GradualSolver, ImmediateSolver, solve_task
from tests.fixtures import (
    FIREARMS,
    HOMICIDES,
    REF_INTERCEPT,
    REF_PEARSON_R,
    REF_R_SQUARED,
    REF_SLOPE,
)

TOL = 1e-6


def _inspection():
    insp = wb_mod.WorkbookInspection(path="x.xls")
    sheet = wb_mod.SheetData(name="Tooleht")
    sheet.numeric_columns = {"A": list(FIREARMS), "B": list(HOMICIDES)}
    insp.sheets.append(sheet)
    return insp


class FakeAI:
    """Mock AI client: returns a canned analysis, never hits the network."""

    def __init__(self, tasks):
        self._tasks = tasks
        self.total_cost = 0.0
        self.analyze_calls = 0

    def analyze_workbook(self, payload):
        self.analyze_calls += 1
        return WorkbookAnalysis(sheets=["Tooleht"], tasks=self._tasks)

    def generate_revision_questions(self, tasks, results):
        return [{"question": "q", "answer": "a"}]


def _config(tmp_path):
    cfg = Config(path=tmp_path / "config.json")
    cfg._data["gradual_interval_seconds"] = 300
    return cfg


def _patch_excel_writer(monkeypatch):
    """Mock the in-place excel_writer path so the solver never touches Excel."""
    monkeypatch.setattr(ew, "backup_workbook", lambda path: "/tmp/backup.xls")
    monkeypatch.setattr(wb_mod, "get_xw_book", lambda path: Mock())
    monkeypatch.setattr(
        ew, "write_all_sheets",
        lambda xw_book, analysis, results, inspection: {
            "written": ["Lahendus"], "skipped": [], "errors": [],
        },
    )
    monkeypatch.setattr(ew, "verify_workbook", lambda xw_book, orig, expected: [])
    monkeypatch.setattr(ew, "save_workbook", lambda xw_book: None)


def test_solve_correlation_task():
    task = TaskPlan(
        task_id="t1",
        source_sheet="Tooleht",
        task_type=TaskType.CORRELATION,
        description="Pearson",
    )
    res = solve_task(task, _inspection())
    assert res.verified
    assert math.isclose(res.values["r"], REF_PEARSON_R, abs_tol=TOL)


def test_solve_regression_task():
    task = TaskPlan(
        task_id="t2",
        source_sheet="Tooleht",
        task_type=TaskType.REGRESSION,
        description="OLS",
    )
    res = solve_task(task, _inspection())
    assert math.isclose(res.values["intercept"], REF_INTERCEPT, abs_tol=TOL)
    assert math.isclose(res.values["slope"], REF_SLOPE, abs_tol=TOL)


def test_solve_prediction_task():
    task = TaskPlan(
        task_id="t3",
        source_sheet="Tooleht",
        task_type=TaskType.PREDICTION,
        description="predict",
        inputs={"x_new": 12.5},
    )
    res = solve_task(task, _inspection())
    assert math.isclose(res.values["y_pred"], 14.70374226, abs_tol=1e-5)


def test_solve_task_records_error_on_bad_data():
    # Sheet without two numeric columns -> resolution fails, recorded as error.
    insp = wb_mod.WorkbookInspection(path="x.xls")
    insp.sheets.append(wb_mod.SheetData(name="Empty"))
    task = TaskPlan(
        task_id="e1", source_sheet="Empty", task_type=TaskType.CORRELATION, description="x"
    )
    res = solve_task(task, insp)
    assert res.error is not None
    assert res.verified is False


def test_immediate_solver_end_to_end(tmp_path, monkeypatch):
    tasks = [
        TaskPlan(task_id="t1", source_sheet="Tooleht", task_type=TaskType.CORRELATION, description="r"),
        TaskPlan(task_id="t2", source_sheet="Tooleht", task_type=TaskType.REGRESSION, description="reg"),
    ]
    ai = FakeAI(tasks)
    cfg = _config(tmp_path)

    src = tmp_path / "book.xls"
    src.write_text("x")
    # Patch inspection to avoid reading a real .xls.
    monkeypatch.setattr(wb_mod, "inspect_workbook", lambda path: _inspection())
    _patch_excel_writer(monkeypatch)

    progress_log = []
    solver = ImmediateSolver(ai, cfg, progress=lambda m, d, t: progress_log.append((m, d, t)))
    result = solver.run(str(src))

    assert ai.analyze_calls == 1
    assert len(result["results"]) == 2
    # In-place: output path is the original file, no new file created.
    assert result["output_path"] == str(src)
    assert ("Completed", 2, 2) in progress_log


def test_gradual_solver_schedules_with_mocked_time(tmp_path, monkeypatch):
    tasks = [
        TaskPlan(task_id=f"t{i}", source_sheet="Tooleht", task_type=TaskType.CORRELATION, description="r")
        for i in range(3)
    ]
    ai = FakeAI(tasks)
    cfg = _config(tmp_path)

    src = tmp_path / "book.xls"
    src.write_text("x")
    monkeypatch.setattr(wb_mod, "inspect_workbook", lambda path: _inspection())
    _patch_excel_writer(monkeypatch)

    sleeps = []
    solver = GradualSolver(ai, cfg, batch_size=1, sleep_fn=lambda s: sleeps.append(s))
    result = solver.run(str(src))

    assert result["completed"] is True
    assert len(result["results"]) == 3
    # 3 tasks, one per batch => 2 inter-batch sleeps totalling 2 × 300 = 600s.
    # Sleeps are chunked into 1s increments for cancellation responsiveness.
    assert sum(sleeps) == pytest.approx(600.0)
    assert all(s <= 1.0 for s in sleeps)


def test_gradual_solver_commits_after_each_batch(tmp_path, monkeypatch):
    """GradualSolver must write the workbook once per batch plus a final write."""
    tasks = [
        TaskPlan(task_id=f"t{i}", source_sheet="Tooleht", task_type=TaskType.CORRELATION, description="r")
        for i in range(3)
    ]
    ai = FakeAI(tasks)
    cfg = _config(tmp_path)
    src = tmp_path / "book.xls"
    src.write_text("x")
    monkeypatch.setattr(wb_mod, "inspect_workbook", lambda path: _inspection())

    monkeypatch.setattr(ew, "backup_workbook", lambda path: "/tmp/backup.xls")
    monkeypatch.setattr(wb_mod, "get_xw_book", lambda path: Mock())
    monkeypatch.setattr(ew, "verify_workbook", lambda xw_book, orig, expected: [])
    monkeypatch.setattr(ew, "save_workbook", lambda xw_book: None)

    write_calls: list[int] = []
    monkeypatch.setattr(
        ew, "write_all_sheets",
        lambda xw_book, analysis, results, inspection: (
            write_calls.append(len(results)),
            {"written": ["Lahendus"], "skipped": [], "errors": []},
        )[1],
    )

    solver = GradualSolver(ai, cfg, batch_size=1, sleep_fn=lambda s: None)
    result = solver.run(str(src))

    assert result["completed"] is True
    # 3 batches (one write each) + one final write after the loop = 4 calls.
    # Each carries the cumulative result count.
    assert write_calls == [1, 2, 3, 3]


def test_gradual_solver_pause_via_should_continue(tmp_path, monkeypatch):
    tasks = [
        TaskPlan(task_id=f"t{i}", source_sheet="Tooleht", task_type=TaskType.CORRELATION, description="r")
        for i in range(3)
    ]
    ai = FakeAI(tasks)
    cfg = _config(tmp_path)
    src = tmp_path / "book.xls"
    src.write_text("x")
    monkeypatch.setattr(wb_mod, "inspect_workbook", lambda path: _inspection())
    _patch_excel_writer(monkeypatch)

    calls = {"n": 0}

    def should_continue():
        calls["n"] += 1
        return calls["n"] <= 1  # allow first batch, pause before second

    solver = GradualSolver(ai, cfg, batch_size=1, sleep_fn=lambda s: None)
    result = solver.run(str(src), should_continue=should_continue)

    assert result["completed"] is False
    assert len(result["results"]) == 1


# ================================================================ new regression tests


def test_missing_x_new_raises():
    """Prediction task with no x_new must raise, not silently predict at x=0."""
    task = TaskPlan(
        task_id="t_pred",
        source_sheet="Tooleht",
        task_type=TaskType.PREDICTION,
        description="predict",
        inputs={},  # no x_new
    )
    res = solve_task(task, _inspection())
    assert res.error is not None
    assert "x_new" in res.error or "x₀" in res.error
    assert res.verified is False


def test_missing_mu0_raises():
    """One-sample t-test without mu0 must raise, not default to 0."""
    insp = wb_mod.WorkbookInspection(path="x.xls")
    sheet = wb_mod.SheetData(name="Tooleht")
    sheet.numeric_columns = {"A": [1.0, 2.0, 3.0, 4.0, 5.0]}
    insp.sheets.append(sheet)
    task = TaskPlan(
        task_id="t_hyp",
        source_sheet="Tooleht",
        task_type=TaskType.HYPOTHESIS_TEST,
        description="t-test",
        inputs={"data": [1.0, 2.0, 3.0, 4.0, 5.0]},  # no mu0
    )
    res = solve_task(task, insp)
    assert res.error is not None
    assert "mu0" in res.error
    assert res.verified is False


def test_chart_task_not_verified():
    """CHART tasks must never be marked verified — can't verify without live Excel."""
    task = TaskPlan(
        task_id="tc",
        source_sheet="Tooleht",
        task_type=TaskType.CHART,
        description="chart",
    )
    res = solve_task(task, _inspection())
    assert res.verified is False


def test_other_task_not_verified():
    """OTHER tasks must never be marked verified."""
    task = TaskPlan(
        task_id="to",
        source_sheet="Tooleht",
        task_type=TaskType.OTHER,
        description="other",
    )
    res = solve_task(task, _inspection())
    assert res.verified is False


def test_regression_verify_independent_check():
    """_verify for REGRESSION recomputes from source data and rejects tampered values."""
    from src.solver import _verify
    task = TaskPlan(
        task_id="t_reg",
        source_sheet="Tooleht",
        task_type=TaskType.REGRESSION,
        description="reg",
    )
    insp = _inspection()

    # Correct values → verified.
    result_ok = solve_task(task, insp)
    assert result_ok.verified is True

    # Tampered intercept → should fail independent check.
    from src.schemas import TaskResult as TR
    tampered = TR(
        task_id="t_reg",
        task_type=TaskType.REGRESSION,
        source_sheet="Tooleht",
        values={
            "intercept": 99.9,    # wrong
            "slope": REF_SLOPE,
            "r_squared": REF_R_SQUARED,
        },
    )
    assert _verify(task, tampered, insp) is False


def test_spearman_dispatch():
    """CORRELATION task with method=spearman dispatches to spearman_correlation."""
    task = TaskPlan(
        task_id="t_sp",
        source_sheet="Tooleht",
        task_type=TaskType.CORRELATION,
        description="Spearman",
        inputs={"method": "spearman"},
    )
    res = solve_task(task, _inspection())
    assert res.error is None
    assert "rho" in res.values
    assert res.values.get("method") == "spearman"
    assert -1.0 <= res.values["rho"] <= 1.0
    assert res.verified is True


def test_mean_response_ci_dispatch():
    """CONFIDENCE_INTERVAL with x_new dispatches to confidence_interval_mean."""
    task = TaskPlan(
        task_id="t_ci",
        source_sheet="Tooleht",
        task_type=TaskType.CONFIDENCE_INTERVAL,
        description="mean CI",
        inputs={"x_new": 12.5},
    )
    res = solve_task(task, _inspection())
    assert res.error is None
    assert res.values.get("ci_type") == "mean_response"
    assert "ci_lower" in res.values and "ci_upper" in res.values
    assert res.values["ci_lower"] < res.values["ci_upper"]
    # Mean-response CI must be narrower than prediction interval at the same point.
    from src import math_engine as me
    pi_lo, pi_hi, _ = me.prediction_interval(list(FIREARMS), list(HOMICIDES), 12.5)
    assert res.values["ci_lower"] > pi_lo
    assert res.values["ci_upper"] < pi_hi


def test_prediction_alpha_from_inputs():
    """PREDICTION_INTERVAL uses the alpha specified in task inputs."""
    base_task = TaskPlan(
        task_id="t_pi",
        source_sheet="Tooleht",
        task_type=TaskType.PREDICTION_INTERVAL,
        description="PI",
        inputs={"x_new": 5.0},  # default alpha=0.05
    )
    wide_task = TaskPlan(
        task_id="t_pi90",
        source_sheet="Tooleht",
        task_type=TaskType.PREDICTION_INTERVAL,
        description="PI 90%",
        inputs={"x_new": 5.0, "alpha": 0.10},
    )
    insp = _inspection()
    res_95 = solve_task(base_task, insp)
    res_90 = solve_task(wide_task, insp)
    assert res_95.error is None
    assert res_90.error is None
    # 95% interval must be wider than 90% interval.
    width_95 = res_95.values["pi_upper"] - res_95.values["pi_lower"]
    width_90 = res_90.values["pi_upper"] - res_90.values["pi_lower"]
    assert width_95 > width_90


def test_confidence_interval_alpha_from_inputs():
    """CONFIDENCE_INTERVAL uses the alpha specified in task inputs."""
    task_90 = TaskPlan(
        task_id="t_ci90",
        source_sheet="Tooleht",
        task_type=TaskType.CONFIDENCE_INTERVAL,
        description="CI 90%",
        inputs={"alpha": 0.10},
    )
    res = solve_task(task_90, _inspection())
    assert res.error is None
    assert math.isclose(res.values.get("alpha", 0.05), 0.10, abs_tol=1e-9)


def test_resolve_xy_rejects_large_length_mismatch():
    """_resolve_xy raises when column lengths differ by more than 1."""
    from src.solver import _resolve_xy
    insp = wb_mod.WorkbookInspection(path="x.xls")
    sheet = wb_mod.SheetData(name="S")
    sheet.numeric_columns = {"A": [1.0, 2.0, 3.0, 4.0, 5.0], "B": [1.0, 2.0]}
    insp.sheets.append(sheet)
    task = TaskPlan(task_id="t1", source_sheet="S", task_type=TaskType.REGRESSION, description="x")
    with pytest.raises(ValueError, match="misaligned"):
        _resolve_xy(task, insp)


def test_negative_slope_regression():
    """Regression with negative slope is computed and verified correctly."""
    x = [1.0, 2.0, 3.0, 4.0, 5.0]
    y = [5.0, 4.0, 3.0, 2.0, 1.0]
    insp = wb_mod.WorkbookInspection(path="x.xls")
    sheet = wb_mod.SheetData(name="S")
    sheet.numeric_columns = {"A": x, "B": y}
    insp.sheets.append(sheet)
    task = TaskPlan(task_id="t1", source_sheet="S", task_type=TaskType.REGRESSION, description="neg")
    res = solve_task(task, insp)
    assert res.error is None
    assert res.values["slope"] < 0
    assert res.verified is True


def test_generate_revision_questions_called(tmp_path, monkeypatch):
    """ImmediateSolver must call generate_revision_questions() and attach questions."""
    tasks = [
        TaskPlan(task_id="t1", source_sheet="Tooleht", task_type=TaskType.CORRELATION, description="r"),
    ]
    ai = FakeAI(tasks)
    call_log = []
    original_gen = ai.generate_revision_questions

    def tracking_gen(task_dicts, result_dicts):
        call_log.append((len(task_dicts), len(result_dicts)))
        return original_gen(task_dicts, result_dicts)

    ai.generate_revision_questions = tracking_gen

    src = tmp_path / "book.xls"
    src.write_text("x")
    monkeypatch.setattr(wb_mod, "inspect_workbook", lambda path: _inspection())
    _patch_excel_writer(monkeypatch)

    cfg = _config(tmp_path)
    solver = ImmediateSolver(ai, cfg)
    result = solver.run(str(src))

    assert len(call_log) == 1, "generate_revision_questions must be called exactly once"
    questions = result["analysis"].revision_questions
    assert len(questions) >= 1


def test_gradual_solver_cancels_during_sleep(tmp_path, monkeypatch):
    """GradualSolver stops mid-interval when should_continue returns False."""
    tasks = [
        TaskPlan(task_id=f"t{i}", source_sheet="Tooleht", task_type=TaskType.CORRELATION, description="r")
        for i in range(3)
    ]
    ai = FakeAI(tasks)
    cfg = _config(tmp_path)
    src = tmp_path / "book.xls"
    src.write_text("x")
    monkeypatch.setattr(wb_mod, "inspect_workbook", lambda path: _inspection())
    _patch_excel_writer(monkeypatch)

    # Allow first batch; cancel on 2nd should_continue call (during sleep).
    calls = {"n": 0}

    def should_continue():
        calls["n"] += 1
        # First call: start of loop → allow; subsequent calls during sleep → deny.
        return calls["n"] == 1

    sleep_calls = []
    solver = GradualSolver(ai, cfg, batch_size=1, sleep_fn=lambda s: sleep_calls.append(s))
    result = solver.run(str(src), should_continue=should_continue)

    assert result["completed"] is False
    assert len(result["results"]) == 1
    # Solver must have entered the sleep loop but exited early (not slept the full 300s).
    assert 0 < sum(sleep_calls) < 300


def test_get_xw_book_no_fallback_to_active():
    """get_xw_book must raise, not return the active book, when path is not found."""
    import types

    class FakeBooks:
        def __iter__(self):
            return iter([])

    class FakeXw:
        books = FakeBooks()

    import src.workbook as wbm
    original = wbm._import_xlwings
    wbm._import_xlwings = lambda: FakeXw()
    try:
        with pytest.raises(RuntimeError, match="ei leitud"):
            wbm.get_xw_book("/nonexistent/book.xlsx")
    finally:
        wbm._import_xlwings = original
