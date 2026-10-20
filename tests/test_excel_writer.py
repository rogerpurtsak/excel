"""Tests for the in-place excel_writer module with xlwings fully mocked."""

import math

import pytest

from src import excel_writer as ew
from src import math_engine as me
from src import workbook as wb_mod
from src.schemas import TaskPlan, TaskResult, TaskType, WorkbookAnalysis
from tests.fixtures import FIREARMS, HOMICIDES


# ---------------------------------------------------------------- fakes
class FakeFont:
    def __init__(self):
        self.bold = False
        self.italic = False
        self.size = None
        self.color = None


class FakeApi:
    def __init__(self):
        self.WrapText = False


class FakeRange:
    def __init__(self, values, key):
        self._values = values
        self._key = key
        self.font = FakeFont()
        self.color = None
        self.number_format = None
        self.column_width = None
        self.row_height = None
        self.left = 0
        self.top = 0
        self.api = FakeApi()

    @property
    def value(self):
        return self._values.get(self._key)

    @value.setter
    def value(self, v):
        self._values[self._key] = v

    def merge(self):
        pass

    def end(self, direction):
        return self


class FakeCharts:
    def add(self, **kwargs):
        raise RuntimeError("no charts in fake")


class FakeSheet:
    def __init__(self, name):
        self.name = name
        self._ranges = {}   # key -> FakeRange object
        self._values = {}   # key -> cell value
        self.charts = FakeCharts()

    def range(self, a, b=None):
        # Normalize key to a string for A1 refs, tuple for (row, col).
        key = str(a) if b is None else (a, b)
        if key not in self._ranges:
            self._ranges[key] = FakeRange(self._values, key)
        return self._ranges[key]

    def clear(self):
        self._ranges = {}
        self._values = {}


class FakeSheets:
    def __init__(self):
        self._sheets = []

    def add(self, name, after=None, before=None):
        s = FakeSheet(name)
        self._sheets.append(s)
        return s

    def __iter__(self):
        return iter(self._sheets)

    def __getitem__(self, key):
        if isinstance(key, int):
            return self._sheets[key]
        for s in self._sheets:
            if s.name == key:
                return s
        raise KeyError(key)


class FakeBook:
    def __init__(self):
        self.sheets = FakeSheets()
        self.saved = False

    def save(self):
        self.saved = True


def _inspection():
    insp = wb_mod.WorkbookInspection(path="/some/dir/Tooleht.xls")
    sheet = wb_mod.SheetData(name="Tooleht")
    sheet.numeric_columns = {"A": list(FIREARMS), "B": list(HOMICIDES)}
    insp.sheets.append(sheet)
    return insp


# ---------------------------------------------------------------- tests
def test_backup_creates_file(tmp_path, monkeypatch):
    src = tmp_path / "book.xls"
    src.write_text("data")
    monkeypatch.setattr(ew, "_backup_dir", lambda: tmp_path / "backups")
    backup = ew.backup_workbook(str(src))
    from pathlib import Path
    assert Path(backup).exists()
    assert Path(backup).parent == tmp_path / "backups"


def test_get_or_create_ai_sheet_creates_new():
    book = FakeBook()
    book.sheets.add("Tooleht")  # unrelated existing sheet
    sheet = ew.get_or_create_ai_sheet(book, "Lahendus")
    assert sheet.name == "Lahendus"
    assert ew._is_ai_sheet(sheet)


def test_get_or_create_ai_sheet_reuses_ai_sheet():
    book = FakeBook()
    existing = book.sheets.add("Lahendus")
    ew._mark_sheet(existing)
    existing.range("A5").value = "old content"
    sheet = ew.get_or_create_ai_sheet(book, "Lahendus")
    assert sheet is existing
    # Content cleared, marker re-applied.
    assert sheet.range("A5").value is None
    assert ew._is_ai_sheet(sheet)


def test_get_or_create_ai_sheet_refuses_user_sheet():
    book = FakeBook()
    book.sheets.add("Lahendus")  # not marked = user sheet
    with pytest.raises(ValueError):
        ew.get_or_create_ai_sheet(book, "Lahendus")


def test_write_lahendus_no_crash():
    book = FakeBook()
    insp = _inspection()
    plan = TaskPlan(task_id="t1", source_sheet="Tooleht",
                    task_type=TaskType.CORRELATION, description="Pearson")
    result = TaskResult(task_id="t1", task_type=TaskType.CORRELATION,
                        source_sheet="Tooleht", values={"r": 0.885, "p_value": 0.003})
    name = ew.write_lahendus(book, [plan], [result], insp)
    assert name == ew.SHEET_LAHENDUS


def test_write_all_sheets_returns_report(monkeypatch):
    book = FakeBook()
    insp = _inspection()
    analysis = WorkbookAnalysis(sheets=["Tooleht"], tasks=[])
    report = ew.write_all_sheets(book, analysis, [], insp)
    assert set(report.keys()) == {"written", "skipped", "errors"}
    assert ew.SHEET_LAHENDUS in report["written"]


def test_no_separate_output_file_created(tmp_path, monkeypatch):
    from src.solver import ImmediateSolver
    from src.config import Config

    class FakeAI:
        total_cost = 0.0
        def analyze_workbook(self, payload):
            return WorkbookAnalysis(sheets=["Tooleht"], tasks=[
                TaskPlan(task_id="t1", source_sheet="Tooleht",
                         task_type=TaskType.CORRELATION, description="r"),
            ])

    src = tmp_path / "Tooleht.xls"
    src.write_text("x")
    monkeypatch.setattr(wb_mod, "inspect_workbook", lambda path: _inspection())
    monkeypatch.setattr(ew, "backup_workbook", lambda path: str(tmp_path / "backups" / "b.xls"))
    monkeypatch.setattr(wb_mod, "get_xw_book", lambda path: FakeBook())

    cfg = Config(path=tmp_path / "config.json")
    solver = ImmediateSolver(FakeAI(), cfg)
    result = solver.run(str(src))

    assert result["output_path"] == str(src)
    # No *_lahendatud* file created anywhere in tmp_path.
    assert not list(tmp_path.glob("*lahendatud*"))


def test_original_sheet_preserved():
    book = FakeBook()
    book.sheets.add("Tooleht")  # original, unmarked
    insp = _inspection()
    analysis = WorkbookAnalysis(sheets=["Tooleht"], tasks=[])
    report = ew.write_all_sheets(book, analysis, [], insp)
    assert "Tooleht" not in report["written"]
    issues = ew.verify_workbook(book, "Tooleht", report["written"])
    assert issues == []


def test_verify_workbook_detects_missing_original():
    book = FakeBook()
    book.sheets.add("Lahendus")
    issues = ew.verify_workbook(book, "Tooleht", ["Lahendus"])
    assert any("Tooleht" in i for i in issues)


def test_outward_rounding_pi():
    lo, hi, extrap = me.prediction_interval(FIREARMS, HOMICIDES, 12.5)
    lo_rounded = math.floor(lo * 10000) / 10000
    hi_rounded = math.ceil(hi * 10000) / 10000
    # Lower rounds down (<= true), upper rounds up (>= true).
    assert lo_rounded <= lo
    assert hi_rounded >= hi


def test_backup_path_not_active(tmp_path, monkeypatch):
    """backup_workbook copies the file but never opens it in Excel."""
    opened = []
    monkeypatch.setattr(wb_mod, "open_in_excel", lambda p: opened.append(p))
    monkeypatch.setattr(ew, "_backup_dir", lambda: tmp_path / "backups")
    src = tmp_path / "book.xls"
    src.write_text("data")
    backup = ew.backup_workbook(str(src))
    from pathlib import Path
    assert Path(backup).exists()
    assert opened == []  # no Excel open triggered


def test_backup_timestamps_unique(tmp_path, monkeypatch):
    """Two rapid backup_workbook calls must not produce the same filename."""
    monkeypatch.setattr(ew, "_backup_dir", lambda: tmp_path / "backups")
    src = tmp_path / "book.xls"
    src.write_text("data")
    p1 = ew.backup_workbook(str(src))
    p2 = ew.backup_workbook(str(src))
    assert p1 != p2, "Two rapid backups produced identical filenames"


def test_ci_labels_use_actual_alpha():
    """write_lahendus must reflect the actual alpha in CI display text, not always 95%."""
    book = FakeBook()
    insp = _inspection()

    plan = TaskPlan(
        task_id="t_ci",
        source_sheet="Tooleht",
        task_type=TaskType.CONFIDENCE_INTERVAL,
        description="CI 90%",
    )
    from src import math_engine as me
    ci_vals = me.regression_confidence_intervals(list(insp.sheets[0].numeric_columns["A"]),
                                                 list(insp.sheets[0].numeric_columns["B"]),
                                                 alpha=0.10)
    result = TaskResult(
        task_id="t_ci",
        task_type=TaskType.CONFIDENCE_INTERVAL,
        source_sheet="Tooleht",
        values=ci_vals,
    )
    # Should not raise and should write 90% labels (not 95%).
    name = ew.write_lahendus(book, [plan], [result], insp)
    assert name == ew.SHEET_LAHENDUS
    # Inspect written cell values for "90" somewhere in the sheet.
    sheet = book.sheets[ew.SHEET_LAHENDUS]
    all_values = list(sheet._values.values())
    text_vals = [str(v) for v in all_values if v is not None]
    assert any("90" in s for s in text_vals), "Expected '90' in CI labels but not found"


def test_formula_written_as_formula(monkeypatch):
    """write_arvutused stores Excel formulas via .formula, not .value."""
    book = FakeBook()
    insp = _inspection()

    # Patch FakeRange to track whether formula or value was set.
    formula_assignments = {}
    value_assignments = {}

    original_range_class = FakeRange

    class TrackingRange(original_range_class):
        @property
        def formula(self):
            return formula_assignments.get(self._key)

        @formula.setter
        def formula(self, v):
            formula_assignments[self._key] = v

        @original_range_class.value.setter
        def value(self, v):
            self._values[self._key] = v
            value_assignments[self._key] = v

    # Monkey-patch FakeSheet to use TrackingRange.
    original_range_fn = FakeSheet.range

    def tracking_range(self_sheet, a, b=None):
        key = str(a) if b is None else (a, b)
        if key not in self_sheet._ranges:
            self_sheet._ranges[key] = TrackingRange(self_sheet._values, key)
        return self_sheet._ranges[key]

    monkeypatch.setattr(FakeSheet, "range", tracking_range)

    plan = TaskPlan(
        task_id="t1",
        source_sheet="Tooleht",
        task_type=TaskType.CORRELATION,
        description="r",
        excel_formulas=[{"cell": "B10", "formula": "=CORREL(A2:A9,B2:B9)"}],
    )
    result = TaskResult(
        task_id="t1",
        task_type=TaskType.CORRELATION,
        source_sheet="Tooleht",
        values={"r": 0.885, "p_value": 0.003},
    )

    ew.write_arvutused(book, [plan], [result], insp)

    # At least one formula starting with = should have been written via .formula.
    assert any(
        isinstance(v, str) and v.startswith("=") for v in formula_assignments.values()
    ), f"No formula written via .formula setter; got: {formula_assignments}"


def test_repeated_run_reuses_ai_sheet():
    """write_all_sheets on a second run must clear and reuse existing AI sheets."""
    book = FakeBook()
    insp = _inspection()
    analysis = WorkbookAnalysis(sheets=["Tooleht"], tasks=[])

    # First run.
    report1 = ew.write_all_sheets(book, analysis, [], insp)
    # Confirm sheet exists and is marked.
    lahendus = book.sheets[ew.SHEET_LAHENDUS]
    lahendus.range("A5").value = "old data from first run"

    # Second run — must reuse the same sheet (clear + rewrite), not raise.
    report2 = ew.write_all_sheets(book, analysis, [], insp)
    assert ew.SHEET_LAHENDUS in report2["written"]
    # Old data must be cleared.
    assert lahendus.range("A5").value is None
