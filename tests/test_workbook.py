"""Tests for workbook inspection and output-copy naming."""

import os
from pathlib import Path

import pytest

from src import workbook as wb_mod

SAMPLE = str(Path(__file__).resolve().parents[1] / "samples" / "tooleht5.xls")


def test_column_letter():
    assert wb_mod._column_letter(0) == "A"
    assert wb_mod._column_letter(1) == "B"
    assert wb_mod._column_letter(25) == "Z"
    assert wb_mod._column_letter(26) == "AA"


@pytest.mark.skipif(wb_mod.xlrd is None, reason="xlrd not installed")
def test_inspect_sample_workbook():
    inspection = wb_mod.inspect_workbook(SAMPLE)
    assert inspection.sheets, "expected at least one sheet"
    sheet = inspection.sheets[0]
    # Two numeric columns of firearms vs homicides.
    assert len(sheet.numeric_columns) >= 2
    cols = list(sheet.numeric_columns.values())
    assert 11.6 in cols[0]
    assert 13.1 in cols[1]
    # Estonian headers should appear in the text blocks.
    blob = " ".join(sheet.text_blocks)
    assert "Automaatrelvad" in blob
    assert "Mõrvad" in blob


@pytest.mark.skipif(wb_mod.xlrd is None, reason="xlrd not installed")
def test_detect_tasks_mentions_regression():
    inspection = wb_mod.inspect_workbook(SAMPLE)
    # The exercise asks for a regression line, so "regression" should be detected.
    assert "regression" in inspection.detected_tasks


def test_create_output_copy_naming(tmp_path):
    src = tmp_path / "book.xls"
    src.write_text("x")
    out = wb_mod.create_output_copy(str(src))
    assert out.endswith("book_lahendatud.xlsx")


def test_create_output_copy_collision(tmp_path):
    src = tmp_path / "book.xls"
    src.write_text("x")
    first = Path(wb_mod.create_output_copy(str(src)))
    first.write_text("y")  # occupy the first candidate
    out2 = wb_mod.create_output_copy(str(src))
    assert out2.endswith("book_lahendatud_2.xlsx")


def test_detect_active_workbook_without_xlwings(monkeypatch):
    # Force the xlwings import to fail -> graceful None.
    monkeypatch.setattr(wb_mod, "_import_xlwings", lambda: None)
    assert wb_mod.detect_active_workbook() is None


def test_open_in_excel_without_xlwings(monkeypatch):
    monkeypatch.setattr(wb_mod, "_import_xlwings", lambda: None)
    assert wb_mod.open_in_excel("/tmp/whatever.xlsx") is None


@pytest.mark.skipif(wb_mod.openpyxl is None, reason="openpyxl not installed")
def test_flush_raises_xlwings_error_never_falls_back(tmp_path):
    """Regression: when xw_book is not None and xlwings write fails, flush_results_to_excel
    must raise XlwingsWriteError. It must never write via openpyxl as a fallback while
    Excel holds the file open.
    """
    from src.schemas import TaskResult, TaskType
    from unittest.mock import MagicMock, patch

    inspection = wb_mod.WorkbookInspection(path="x.xls")
    inspection.sheets.append(wb_mod.SheetData(name="S"))
    out_path = str(tmp_path / "out.xlsx")
    openpyxl_wb = wb_mod.build_output_workbook(inspection, out_path)
    wb_mod.save_workbook(openpyxl_wb, out_path)

    task_result = TaskResult(
        task_id="t1",
        task_type=TaskType.CORRELATION,
        source_sheet="S",
        description="corr",
        written_cells=[{"cell": "A1", "value": 0.885}],
        output_location={"sheet": "Arvutused"},
        explanation="",
    )

    # Simulate a live xw_book whose write raises.
    broken_xw_book = MagicMock()
    broken_xw_book.sheets.__iter__ = MagicMock(return_value=iter([]))
    broken_xw_book.sheets.__contains__ = MagicMock(return_value=False)
    broken_xw_book.sheets.add = MagicMock(side_effect=RuntimeError("Excel write error"))

    openpyxl_save_calls = []
    with patch.object(wb_mod, "save_workbook", side_effect=lambda wb, p: openpyxl_save_calls.append(p)):
        with pytest.raises(wb_mod.XlwingsWriteError):
            wb_mod.flush_results_to_excel(openpyxl_wb, broken_xw_book, [task_result], out_path)

    # openpyxl must not have been called — the file was not written externally.
    assert openpyxl_save_calls == [], (
        "flush_results_to_excel must not fall back to openpyxl when xw_book is not None"
    )


@pytest.mark.skipif(wb_mod.openpyxl is None, reason="openpyxl not installed")
def test_flush_uses_openpyxl_when_no_xw_book(tmp_path):
    """When xw_book is None (Excel not running), openpyxl+save is the correct path."""
    from src.schemas import TaskResult, TaskType

    inspection = wb_mod.WorkbookInspection(path="x.xls")
    inspection.sheets.append(wb_mod.SheetData(name="S"))
    out_path = str(tmp_path / "out.xlsx")
    openpyxl_wb = wb_mod.build_output_workbook(inspection, out_path)
    wb_mod.save_workbook(openpyxl_wb, out_path)

    task_result = TaskResult(
        task_id="t1",
        task_type=TaskType.CORRELATION,
        source_sheet="S",
        description="corr",
        written_cells=[{"cell": "D1", "value": 0.885}],
        output_location={"sheet": "Arvutused"},
        explanation="",
    )
    # xw_book=None -> openpyxl path, no exception
    wb_mod.flush_results_to_excel(openpyxl_wb, None, [task_result], out_path)
    reopened = wb_mod.openpyxl.load_workbook(out_path)
    assert reopened["Arvutused"]["D1"].value == 0.885


@pytest.mark.skipif(wb_mod.openpyxl is None, reason="openpyxl not installed")
def test_build_and_save_output(tmp_path):
    from src.schemas import TaskResult, TaskType

    inspection = wb_mod.WorkbookInspection(path="x.xls")
    sheet = wb_mod.SheetData(name="Tooleht", rows=[["Automaatrelvad", "Mõrvad"], [1.0, 2.0]])
    inspection.sheets.append(sheet)

    out_path = str(tmp_path / "out.xlsx")
    wb = wb_mod.build_output_workbook(inspection, out_path)
    result = TaskResult(
        task_id="t1",
        task_type=TaskType.CORRELATION,
        source_sheet="Tooleht",
        description="corr",
        written_cells=[{"cell": "D1", "value": 0.885}],
        output_location={"sheet": "Arvutused"},
        explanation="selgitus",
    )
    wb_mod.write_task_result(wb, result)
    wb_mod.save_workbook(wb, out_path)
    assert os.path.exists(out_path)

    reopened = wb_mod.openpyxl.load_workbook(out_path)
    assert "Arvutused" in reopened.sheetnames
    assert reopened["Arvutused"]["D1"].value == 0.885
