"""Tests for the orchestration pipeline with a mocked AI client."""

import math
from pathlib import Path

import pytest

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
    # 3 tasks, one per batch => 2 inter-batch sleeps of 300s each.
    assert sleeps == [300, 300]


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
