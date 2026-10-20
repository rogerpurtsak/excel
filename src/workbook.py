"""Excel workbook inspection and editing.

Inspection of legacy .xls uses xlrd (read-only). Output is always written as
.xlsx via openpyxl so results are portable and don't require Excel. If xlwings
and a running Excel are available, we can additionally open/refresh the file in
Excel, but every xlwings use is guarded so the package works without Excel.
"""

from __future__ import annotations

import os
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

try:
    import xlrd
except Exception:  # pragma: no cover
    xlrd = None  # type: ignore[assignment]

try:
    import openpyxl
    from openpyxl.chart import ScatterChart, Reference, Series
except Exception:  # pragma: no cover
    openpyxl = None  # type: ignore[assignment]

OUTPUT_SUFFIX = "_lahendatud"  # Estonian for "solved"


class XlwingsWriteError(RuntimeError):
    """Raised when xlwings fails to write to a workbook that is already open in Excel.

    The caller must pause the job rather than falling back to openpyxl, which
    would overwrite the file externally while Excel holds it open.
    """


@dataclass
class SheetData:
    name: str
    rows: list[list[Any]] = field(default_factory=list)
    text_blocks: list[str] = field(default_factory=list)
    numeric_columns: dict[str, list[float]] = field(default_factory=dict)


@dataclass
class WorkbookInspection:
    path: str
    sheets: list[SheetData] = field(default_factory=list)
    detected_tasks: list[str] = field(default_factory=list)

    @property
    def sheet_names(self) -> list[str]:
        return [s.name for s in self.sheets]

    def to_ai_payload(self) -> dict[str, Any]:
        """Compact representation handed to the AI for analysis."""
        return {
            "path": os.path.basename(self.path),
            "sheets": [
                {
                    "name": s.name,
                    "rows": s.rows,
                    "text_blocks": s.text_blocks,
                    "numeric_columns": {k: v for k, v in s.numeric_columns.items()},
                }
                for s in self.sheets
            ],
            "detected_tasks": self.detected_tasks,
        }


def _column_letter(idx: int) -> str:
    """0-based column index -> Excel letter (0 -> A)."""
    letters = ""
    idx += 1
    while idx:
        idx, rem = divmod(idx - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def inspect_workbook(path: str) -> WorkbookInspection:
    """Read a workbook (.xls via xlrd, .xlsx via openpyxl) into a WorkbookInspection."""
    ext = Path(path).suffix.lower()
    if ext == ".xls":
        return _inspect_xls(path)
    if ext == ".xlsx":
        return _inspect_xlsx(path)
    raise ValueError(f"Unsupported workbook extension: {ext}")


def _collect_sheet(name: str, grid: list[list[Any]]) -> SheetData:
    sheet = SheetData(name=name, rows=grid)
    # Gather text blocks (task prose) and numeric columns.
    numeric_by_col: dict[int, list[float]] = {}
    for row in grid:
        for col_idx, value in enumerate(row):
            if isinstance(value, str) and value.strip():
                if len(value.strip()) > 3:
                    sheet.text_blocks.append(value.strip())
            elif isinstance(value, (int, float)) and not isinstance(value, bool):
                numeric_by_col.setdefault(col_idx, []).append(float(value))
    for col_idx, vals in sorted(numeric_by_col.items()):
        if len(vals) >= 2:
            sheet.numeric_columns[_column_letter(col_idx)] = vals
    return sheet


def _inspect_xls(path: str) -> WorkbookInspection:
    if xlrd is None:
        raise RuntimeError("xlrd is required to read .xls files")
    book = xlrd.open_workbook(path)
    inspection = WorkbookInspection(path=path)
    for sheet in book.sheets():
        grid: list[list[Any]] = []
        for r in range(sheet.nrows):
            row: list[Any] = []
            for c in range(sheet.ncols):
                cell = sheet.cell(r, c)
                row.append(cell.value if cell.value != "" else None)
            grid.append(row)
        inspection.sheets.append(_collect_sheet(sheet.name, grid))
    inspection.detected_tasks = _detect_tasks(inspection)
    return inspection


def _inspect_xlsx(path: str) -> WorkbookInspection:
    if openpyxl is None:
        raise RuntimeError("openpyxl is required to read .xlsx files")
    book = openpyxl.load_workbook(path, data_only=True)
    inspection = WorkbookInspection(path=path)
    for ws in book.worksheets:
        grid = [[cell.value for cell in row] for row in ws.iter_rows()]
        inspection.sheets.append(_collect_sheet(ws.title, grid))
    inspection.detected_tasks = _detect_tasks(inspection)
    return inspection


# Estonian keywords that signal the kind of statistics task being requested.
_TASK_KEYWORDS = {
    "regressioon": "regression",
    "korrelatsioon": "correlation",
    "usaldus": "confidence_interval",
    "ennusta": "prediction",
    "prognoos": "prediction",
    "hüpotees": "hypothesis_test",
}


def _detect_tasks(inspection: WorkbookInspection) -> list[str]:
    found: list[str] = []
    for sheet in inspection.sheets:
        blob = " ".join(sheet.text_blocks).lower()
        for kw, task in _TASK_KEYWORDS.items():
            if kw in blob and task not in found:
                found.append(task)
    return found


def create_output_copy(source_path: str) -> str:
    """Return a non-colliding output path with the _lahendatud suffix and .xlsx extension.

    Does not create the file; callers write to the returned path.
    """
    src = Path(source_path)
    base = src.stem + OUTPUT_SUFFIX
    candidate = src.with_name(base + ".xlsx")
    counter = 2
    while candidate.exists():
        candidate = src.with_name(f"{base}_{counter}.xlsx")
        counter += 1
    return str(candidate)


def build_output_workbook(inspection: WorkbookInspection, output_path: str):
    """Create an openpyxl workbook mirroring the source data, return the Workbook object."""
    if openpyxl is None:
        raise RuntimeError("openpyxl is required to build output workbooks")
    wb = openpyxl.Workbook()
    # Remove the default sheet; we recreate from source.
    default = wb.active
    wb.remove(default)
    for sheet in inspection.sheets:
        ws = wb.create_sheet(title=sheet.name[:31] or "Sheet")
        for row in sheet.rows:
            ws.append(list(row))
    if not wb.worksheets:
        wb.create_sheet(title="Arvutused")
    return wb


def write_task_result(wb, task_result) -> None:
    """Write a TaskResult's cells into the workbook (openpyxl Workbook)."""
    loc = task_result.output_location or {}
    target_name = loc.get("sheet") or "Arvutused"
    if target_name in wb.sheetnames:
        ws = wb[target_name]
    else:
        ws = wb.create_sheet(title=target_name[:31])

    for cell in task_result.written_cells:
        ref = cell.get("cell")
        if not ref:
            continue
        # Prefer the formula (Excel re-evaluates); fall back to the value.
        ws[ref] = cell.get("formula") or cell.get("value")

    # Append a human-readable explanation block below the data.
    if task_result.explanation:
        start = (ws.max_row or 0) + 2
        ws.cell(row=start, column=1, value=f"[{task_result.task_id}] {task_result.description}")
        ws.cell(row=start + 1, column=1, value=task_result.explanation)


def create_chart(wb, chart_spec: dict[str, Any]) -> None:
    """Create a scatter chart with a regression line from a chart spec.

    chart_spec keys: sheet, x_range, y_range, title, anchor.
    """
    if openpyxl is None:
        return
    sheet_name = chart_spec.get("sheet")
    if sheet_name not in wb.sheetnames:
        return
    ws = wb[sheet_name]
    chart = ScatterChart()
    chart.title = chart_spec.get("title", "Scatter")
    chart.x_axis.title = chart_spec.get("x_title", "X")
    chart.y_axis.title = chart_spec.get("y_title", "Y")
    chart.style = 13

    xr = chart_spec.get("x_range")
    yr = chart_spec.get("y_range")
    if not (xr and yr):
        return
    xref = Reference(ws, min_col=xr["col"], min_row=xr["min_row"], max_row=xr["max_row"])
    yref = Reference(ws, min_col=yr["col"], min_row=yr["min_row"], max_row=yr["max_row"])
    series = Series(yref, xref, title=chart_spec.get("series_title", "data"))
    series.marker.symbol = "circle"
    series.graphicalProperties.line.noFill = True
    chart.series.append(series)
    ws.add_chart(chart, chart_spec.get("anchor", "H2"))


def save_workbook(wb, output_path: str) -> str:
    """Save an openpyxl workbook to output_path and return the path."""
    wb.save(output_path)
    return output_path


def _write_result_xlwings(xw_book, task_result) -> None:
    """Write a single TaskResult into an already-open xlwings Book."""
    loc = task_result.output_location or {}
    sheet_name = (loc.get("sheet") or "Arvutused")[:31]
    sheet_names = [s.name for s in xw_book.sheets]
    if sheet_name not in sheet_names:
        xw_book.sheets.add(sheet_name)
    ws = xw_book.sheets[sheet_name]
    for cell in task_result.written_cells:
        ref = cell.get("cell")
        if not ref:
            continue
        formula = cell.get("formula")
        if formula:
            ws.range(ref).formula = formula
        else:
            ws.range(ref).value = cell.get("value")
    if task_result.explanation:
        try:
            last_row = ws.range("A1").end("down").row + 2
        except Exception:
            last_row = 3
        ws.range(f"A{last_row}").value = f"[{task_result.task_id}] {task_result.description}"
        ws.range(f"A{last_row + 1}").value = task_result.explanation


def flush_results_to_excel(openpyxl_wb, xw_book, new_results: list, out_path: str) -> None:
    """Commit a batch of new TaskResults to Excel.

    When the workbook is already open in Excel (xw_book is not None), writes
    exclusively through xlwings so changes appear live without a file-lock
    conflict. If xlwings fails, raises XlwingsWriteError — the caller must
    pause the job. Falling back to openpyxl while Excel holds the file open
    is never allowed.

    When xw_book is None (Excel is not running or file was not opened there),
    writes via openpyxl and saves to disk.
    """
    if xw_book is not None:
        try:
            for result in new_results:
                _write_result_xlwings(xw_book, result)
            xw_book.save()
            return
        except XlwingsWriteError:
            raise
        except Exception as exc:
            raise XlwingsWriteError(
                f"xlwings kirjutamine ebaõnnestus, töö peatatud: {exc}"
            ) from exc

    for result in new_results:
        write_task_result(openpyxl_wb, result)
    save_workbook(openpyxl_wb, out_path)


# ------------------------------------------------------------------ xlwings (optional)
def _import_xlwings():
    try:
        import xlwings  # noqa: F401

        return xlwings
    except Exception:
        return None


def open_in_excel(path: str):
    """Open a workbook in Excel via xlwings. No-op (returns None) if unavailable."""
    xw = _import_xlwings()
    if xw is None:
        return None
    try:
        return xw.Book(path)
    except Exception:
        return None


def get_xw_book(path: str):
    """Return the xlwings Book for the given path, which must already be open in Excel.
    Searches all open books by resolved path; falls back to the active book.
    Raises RuntimeError with Estonian message if not found.
    """
    xw = _import_xlwings()
    if xw is None:
        raise RuntimeError("xlwings pole saadaval. Veendu, et Microsoft Excel on paigaldatud.")
    from pathlib import Path as _Path
    norm = str(_Path(path).resolve())
    try:
        for book in xw.books:
            try:
                if str(_Path(book.fullname).resolve()) == norm:
                    return book
            except Exception:
                continue
        active = xw.books.active
        if active:
            return active
    except Exception as exc:
        pass
    raise RuntimeError(
        f"Töövihikut '{_Path(path).name}' ei leitud Excelis. "
        "Veendu, et fail on Microsoft Excelis avatud."
    )


def detect_active_workbook() -> str | None:
    """Return the path of the active Excel workbook via xlwings, or None."""
    xw = _import_xlwings()
    if xw is None:
        return None
    try:
        book = xw.books.active
        return book.fullname
    except Exception:
        return None
