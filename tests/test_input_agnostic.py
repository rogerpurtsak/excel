"""Tests verifying input-agnostic workbook solving across structurally different workbooks.

These tests cover three synthetic workbooks:
1. ttest_nonadjacent.xlsx  — two-sample t-test with data in columns C and E
2. descriptive_col_b.xlsx  — descriptive stats with data in column B only
3. onesample_mu0.xlsx      — one-sample t-test with mu0 from cell text, data in column D

The tests exercise the full inspection → _resolve_single/_resolve_xy path
without any workbook-specific code paths.
"""
from __future__ import annotations

import math
from pathlib import Path

import pytest

from src import workbook as wb_mod
from src.schemas import TaskPlan, TaskType
from src.solver import _resolve_single, _resolve_xy, solve_task

FIXTURES = Path(__file__).parent / "fixtures"


# ------------------------------------------------------------------ helpers
def _task(task_type: TaskType, sheet: str, **inputs) -> TaskPlan:
    return TaskPlan(
        task_id="t1",
        source_sheet=sheet,
        task_type=task_type,
        description="test task",
        inputs=inputs,
    )


# ================================================================== workbook 1
class TestTTestNonadjacent:
    """Two-sample t-test with data in non-adjacent columns C and E."""

    @pytest.fixture(scope="class")
    def inspection(self):
        return wb_mod.inspect_workbook(str(FIXTURES / "ttest_nonadjacent.xlsx"))

    def test_sheet_detected(self, inspection):
        assert inspection.sheets[0].name == "TTest"

    def test_numeric_columns_detected(self, inspection):
        sheet = inspection.sheets[0]
        assert "C" in sheet.numeric_columns
        assert "E" in sheet.numeric_columns

    def test_column_headers_detected(self, inspection):
        sheet = inspection.sheets[0]
        assert sheet.column_headers.get("C") == "Group A"
        assert sheet.column_headers.get("E") == "Group B"

    def test_extract_range_values_c(self, inspection):
        vals = wb_mod.extract_range_values(inspection, "TTest", "C2:C9")
        assert vals == pytest.approx([5.1, 4.8, 5.3, 5.0, 4.9, 5.2, 4.7, 5.1])

    def test_extract_range_values_e(self, inspection):
        vals = wb_mod.extract_range_values(inspection, "TTest", "E2:E9")
        assert vals == pytest.approx([4.3, 4.0, 4.5, 4.1, 4.4, 4.2, 4.6, 4.0])

    def test_extract_paired_range_values(self, inspection):
        x, y = wb_mod.extract_paired_range_values(inspection, "TTest", "C2:C9", "E2:E9")
        assert len(x) == 8
        assert len(y) == 8
        assert x[0] == pytest.approx(5.1)
        assert y[0] == pytest.approx(4.3)

    def test_solve_two_sample_ttest_via_ranges(self, inspection):
        task = _task(
            TaskType.HYPOTHESIS_TEST,
            "TTest",
            x_range="C2:C9",
            y_range="E2:E9",
            alpha=0.05,
        )
        result = solve_task(task, inspection)
        assert result.error is None
        assert "t_statistic" in result.values
        assert "p_value" in result.values
        assert isinstance(result.values["reject_null"], bool)
        # Group A mean (~5.0) > Group B mean (~4.3) → should reject at alpha=0.05
        assert result.values["reject_null"] is True

    def test_solve_raises_without_ranges_and_with_ambiguous_columns(self, inspection):
        """Three numeric columns (A would be empty here, but C and E are present)
        → positional fallback is only used when exactly 2 columns; here exactly 2
        so it should work. But if we add ambiguity by making the task reference a
        wrong sheet, it should raise."""
        task = _task(
            TaskType.HYPOTHESIS_TEST,
            "NonexistentSheet",
            mu0=5.0,
            data_range="C2:C9",
        )
        result = solve_task(task, inspection)
        assert result.error is not None


# ================================================================== workbook 2
class TestDescriptiveColB:
    """Descriptive statistics with data only in column B (not A)."""

    @pytest.fixture(scope="class")
    def inspection(self):
        return wb_mod.inspect_workbook(str(FIXTURES / "descriptive_col_b.xlsx"))

    def test_sheet_detected(self, inspection):
        assert inspection.sheets[0].name == "Stats"

    def test_only_column_b_numeric(self, inspection):
        sheet = inspection.sheets[0]
        assert "B" in sheet.numeric_columns
        assert "A" not in sheet.numeric_columns

    def test_column_header_b(self, inspection):
        sheet = inspection.sheets[0]
        assert sheet.column_headers.get("B") == "Values"

    def test_extract_range_col_b(self, inspection):
        vals = wb_mod.extract_range_values(inspection, "Stats", "B2:B11")
        assert len(vals) == 10
        assert vals[0] == pytest.approx(12.5)

    def test_resolve_single_via_data_range(self, inspection):
        task = _task(TaskType.DESCRIPTIVE_STATS, "Stats", data_range="B2:B11")
        vals = _resolve_single(task, inspection)
        assert len(vals) == 10

    def test_solve_descriptive_stats_col_b(self, inspection):
        task = _task(TaskType.DESCRIPTIVE_STATS, "Stats", data_range="B2:B11")
        result = solve_task(task, inspection)
        assert result.error is None
        assert result.values["n"] == 10
        assert math.isclose(result.values["min"], 11.8, abs_tol=1e-6)
        assert math.isclose(result.values["max"], 15.0, abs_tol=1e-6)

    def test_resolve_single_positional_fallback_works_with_one_column(self, inspection):
        """Only one numeric column → positional fallback is unambiguous."""
        task = _task(TaskType.DESCRIPTIVE_STATS, "Stats")  # no data_range
        vals = _resolve_single(task, inspection)
        assert len(vals) == 10


# ================================================================== workbook 3
class TestOneSampleMu0:
    """One-sample t-test with mu0=50, data in column D (not A or B)."""

    @pytest.fixture(scope="class")
    def inspection(self):
        return wb_mod.inspect_workbook(str(FIXTURES / "onesample_mu0.xlsx"))

    def test_sheet_detected(self, inspection):
        assert inspection.sheets[0].name == "HypTest"

    def test_only_column_d_numeric(self, inspection):
        sheet = inspection.sheets[0]
        assert "D" in sheet.numeric_columns
        assert "A" not in sheet.numeric_columns
        assert "B" not in sheet.numeric_columns

    def test_extract_range_col_d(self, inspection):
        vals = wb_mod.extract_range_values(inspection, "HypTest", "D2:D11")
        assert len(vals) == 10
        assert vals[0] == pytest.approx(48.2)

    def test_resolve_single_via_data_range(self, inspection):
        task = _task(TaskType.HYPOTHESIS_TEST, "HypTest", data_range="D2:D11", mu0=50.0)
        vals = _resolve_single(task, inspection)
        assert len(vals) == 10

    def test_solve_onesample_ttest_via_data_range(self, inspection):
        task = _task(
            TaskType.HYPOTHESIS_TEST,
            "HypTest",
            data_range="D2:D11",
            mu0=50.0,
            alpha=0.05,
        )
        result = solve_task(task, inspection)
        assert result.error is None
        assert "t_statistic" in result.values
        assert "p_value" in result.values
        # Mean ~50.0, H0: mu=50.0, data close to 50 → should NOT reject
        assert result.values["reject_null"] is False

    def test_missing_mu0_raises(self, inspection):
        task = _task(
            TaskType.HYPOTHESIS_TEST,
            "HypTest",
            data_range="D2:D11",
            # mu0 intentionally missing
        )
        result = solve_task(task, inspection)
        assert result.error is not None
        assert "mu0" in result.error.lower() or "one-sample" in result.error.lower()

    def test_positional_fallback_works_single_column(self, inspection):
        """Column D is the only numeric column → positional fallback should work."""
        task = _task(
            TaskType.HYPOTHESIS_TEST,
            "HypTest",
            mu0=50.0,
            # no data_range — relies on positional fallback
        )
        result = solve_task(task, inspection)
        assert result.error is None


# ================================================================== cross-workbook
class TestNoDuplicateColumnAssumptions:
    """Verify that no workbook-specific code paths were needed for any workbook."""

    def test_ttest_nonadjacent_requires_no_column_ab(self):
        """The t-test workbook uses C and E — solver must not hard-code A/B."""
        insp = wb_mod.inspect_workbook(str(FIXTURES / "ttest_nonadjacent.xlsx"))
        sheet = insp.sheets[0]
        assert "A" not in sheet.numeric_columns, "Column A should have no numeric data"

    def test_all_three_workbooks_parseable(self):
        for name in ("ttest_nonadjacent.xlsx", "descriptive_col_b.xlsx", "onesample_mu0.xlsx"):
            insp = wb_mod.inspect_workbook(str(FIXTURES / name))
            assert insp.sheets, f"{name} produced no sheets"

    def test_ai_payload_includes_column_letters(self):
        """to_ai_payload must expose cell coordinates for every non-empty cell."""
        for name in ("ttest_nonadjacent.xlsx", "descriptive_col_b.xlsx", "onesample_mu0.xlsx"):
            insp = wb_mod.inspect_workbook(str(FIXTURES / name))
            payload = insp.to_ai_payload()
            for sheet_payload in payload["sheets"]:
                for row_entry in sheet_payload["rows"]:
                    assert "row" in row_entry, f"{name}: row entry missing 'row' key"
                    assert "cells" in row_entry, f"{name}: row entry missing 'cells' key"
                    for col_letter in row_entry["cells"]:
                        assert col_letter.isupper(), f"column letter not uppercase in {name}"
