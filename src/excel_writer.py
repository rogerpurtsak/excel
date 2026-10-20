"""In-place worksheet writer for Excel AI Solver.

Academic formatting matched to the Estonian university reference workbook style:
  - Arial 11pt throughout
  - Light yellow (#FFFFCC) for raw input data
  - Light green (#CCFFCC) for intermediate calculated values
  - Yellow (#FFFF00) for sum/total rows
  - Light cyan (#CCFFFF) for mean rows
  - No dark banners, no dashboard layout, no AI-slop decoration

All modifications go directly into the already-open workbook via xlwings.
"""
from __future__ import annotations

import math
import re
import shutil
from datetime import datetime
from pathlib import Path
from typing import Any

from .schemas import TaskResult, TaskType, WorkbookAnalysis
from . import workbook as wb_mod

AI_MARKER = "ExcelAISolver:generated"
_MARKER_CELL = "ZZ1"

SHEET_LAHENDUS   = "Lahendus"
SHEET_ARVUTUSED  = "Arvutused ja kontroll"
SHEET_REGRESSION = "Regression"
SHEET_KORDAMIS   = "Kordamisküsimused"
SHEET_VASTUSED   = "Vastused"

# Academic color palette — RGB tuples for xlwings
_COL_INPUT   = (255, 255, 204)  # light yellow  — raw input data
_COL_CALC    = (204, 255, 204)  # light green   — intermediate calculated values
_COL_TOTAL   = (255, 255,   0)  # yellow        — sums / totals
_COL_MEAN    = (204, 255, 255)  # light cyan    — means
_COL_NSIZE   = (204, 235, 197)  # light sage    — sample size row
_COL_RESULT  = (255, 255, 153)  # bright yellow — key answer
_COL_ERROR   = (255, 200, 200)  # light red     — error messages

_FONT    = "Arial"
_FONT_SZ = 11


# ---------------------------------------------------------------- CI helpers
def _ci_value(values: dict[str, Any], flat_key: str) -> Any:
    """Fetch a CI bound from a flat key or the nested tuple form produced by math_engine."""
    if flat_key in values:
        return values[flat_key]
    mapping = {
        "intercept_ci_lower": ("intercept", 0),
        "intercept_ci_upper": ("intercept", 1),
        "slope_ci_lower":     ("slope", 0),
        "slope_ci_upper":     ("slope", 1),
    }
    base, idx = mapping.get(flat_key, (None, None))
    if base is not None:
        pair = values.get(base)
        if isinstance(pair, (list, tuple)) and len(pair) > idx:
            return pair[idx]
    return ""


# ---------------------------------------------------------------- backup
def _backup_dir() -> Path:
    return Path.home() / "Library" / "Application Support" / "ExcelAISolver" / "backups"


def backup_workbook(path: str) -> str:
    """Copy the file to the backups directory with a timestamped name."""
    src = Path(path)
    dest_dir = _backup_dir()
    dest_dir.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    dest = dest_dir / f"{src.stem}_{ts}{src.suffix}"
    shutil.copy2(src, dest)
    return str(dest)


# ---------------------------------------------------------------- sheet helpers
def _is_ai_sheet(sheet) -> bool:
    try:
        return sheet.range(_MARKER_CELL).value == AI_MARKER
    except Exception:
        return False


def _mark_sheet(sheet) -> None:
    try:
        cell = sheet.range(_MARKER_CELL)
        cell.value = AI_MARKER
        try:
            cell.font.color = (255, 255, 255)
        except Exception:
            pass
        try:
            cell.color = (255, 255, 255)
        except Exception:
            pass
    except Exception:
        pass


def get_or_create_ai_sheet(xw_book, name: str):
    """Return an AI-owned sheet. Reuses (and clears) if previously marked. Refuses user sheets."""
    existing = {s.name: s for s in xw_book.sheets}
    if name in existing:
        sheet = existing[name]
        if _is_ai_sheet(sheet):
            try:
                sheet.clear()
            except Exception:
                pass
            _mark_sheet(sheet)
            return sheet
        raise ValueError(
            f"Leht '{name}' on juba olemas ja seda on käsitsi muudetud. "
            "Nimeta see ümber või kustuta, et lahendus saaks selle uuesti luua."
        )
    try:
        last = xw_book.sheets[-1]
        sheet = xw_book.sheets.add(name, after=last)
    except Exception:
        sheet = xw_book.sheets.add(name)
    _mark_sheet(sheet)
    return sheet


# ---------------------------------------------------------------- cell primitives
def _c(sheet, row: int, col: int, value, bold=False, italic=False,
       color=None, fmt=None) -> None:
    """Write one cell. col is 1-based (1=A, 2=B, ...)."""
    try:
        cell = sheet.range((row, col))
        cell.value = value
        try:
            cell.font.name = _FONT
            cell.font.size = _FONT_SZ
        except Exception:
            pass
        if bold:
            try:
                cell.font.bold = True
            except Exception:
                pass
        if italic:
            try:
                cell.font.italic = True
            except Exception:
                pass
        if color is not None:
            try:
                cell.color = color
            except Exception:
                pass
        if fmt:
            try:
                cell.number_format = fmt
            except Exception:
                pass
    except Exception:
        pass


def _set_col_widths(sheet, widths: dict) -> None:
    for col, w in widths.items():
        try:
            sheet.range(f"{col}:{col}").column_width = w
        except Exception:
            pass


# ---------------------------------------------------------------- layout primitives
def _section_title(sheet, row: int, text: str) -> None:
    """Bold section title — plain text, no background."""
    _c(sheet, row, 1, text, bold=True)


def _step(sheet, row: int, text: str) -> None:
    """Plain-text calculation step label."""
    _c(sheet, row, 1, text)


def _conclusion(sheet, row: int, text: str) -> None:
    """Italic conclusion / interpretation — no background color."""
    _c(sheet, row, 1, text, italic=True)


def _kv(sheet, row: int, label: str, value,
        fmt: str | None = None, highlight: bool = False,
        col_label: int = 1, col_val: int = 2) -> None:
    """Label–value pair. Optional yellow highlight on the value cell."""
    try:
        lc = sheet.range((row, col_label))
        lc.value = label
        lc.font.name = _FONT
        lc.font.size = _FONT_SZ
    except Exception:
        pass
    try:
        vc = sheet.range((row, col_val))
        vc.value = value
        vc.font.name = _FONT
        vc.font.size = _FONT_SZ
        if fmt:
            try:
                vc.number_format = fmt
            except Exception:
                pass
        if highlight:
            try:
                vc.color = _COL_RESULT
            except Exception:
                pass
    except Exception:
        pass


def _write_formula_cell(sheet, row: int, col: int, formula: str,
                         color=None) -> None:
    """Write an Excel formula via .formula; fall back to .value."""
    try:
        cell = sheet.range((row, col))
        if isinstance(formula, str) and formula.startswith("="):
            try:
                cell.formula = formula
            except Exception:
                cell.value = formula
        else:
            cell.value = formula
        cell.font.name = _FONT
        cell.font.size = _FONT_SZ
        if color is not None:
            try:
                cell.color = color
            except Exception:
                pass
    except Exception:
        pass


# ---------------------------------------------------------------- data table
def _col_letter_from_range(range_str: str) -> str | None:
    """Extract the column letter from a range string like 'C2:C9' → 'C'."""
    m = re.match(r"^([A-Za-z]+)\d+", range_str.strip())
    if m:
        letter = m.group(1).upper()
        # Normalise to 0-based index then back to ensure consistent form
        try:
            idx = wb_mod._col_letter_to_idx(letter)
            return wb_mod._column_letter(idx)
        except Exception:
            return letter
    return None


def _mini_data_table(sheet, start_row: int, src_sheet,
                      col_keys: list | None = None,
                      max_rows: int = 25) -> int:
    """Write a compact data table with academic colors.

    Column A: row labels (Summa, Keskmine, Valimimaht).
    Columns B+: data values.
    Returns the next available row after the last table row.
    """
    cols = src_sheet.numeric_columns
    if not cols:
        return start_row

    if col_keys:
        col_items = [(k, cols[k]) for k in col_keys if k in cols]
    else:
        col_items = list(cols.items())[:2]

    if not col_items:
        return start_row

    n = min(max_rows, min(len(v) for _, v in col_items))
    if n == 0:
        return start_row

    row = start_row

    # Header row — labels in B, C, …
    for ci, (letter, _) in enumerate(col_items):
        header = src_sheet.column_headers.get(letter, letter)
        _c(sheet, row, 2 + ci, header, bold=True, color=_COL_INPUT)
    row += 1

    # Data rows
    for ri in range(n):
        for ci, (_, vals) in enumerate(col_items):
            if ri < len(vals):
                _c(sheet, row, 2 + ci, vals[ri], color=_COL_INPUT)
        row += 1

    # Sum row
    _c(sheet, row, 1, "Summa", color=_COL_TOTAL)
    for ci, (_, vals) in enumerate(col_items):
        _c(sheet, row, 2 + ci, sum(vals[:n]), bold=True, color=_COL_TOTAL)
    row += 1

    # Mean row
    _c(sheet, row, 1, "Keskmine", color=_COL_MEAN)
    for ci, (_, vals) in enumerate(col_items):
        mean_val = sum(vals[:n]) / n
        _c(sheet, row, 2 + ci, mean_val, color=_COL_MEAN, fmt="0.0000")
    row += 1

    # Sample size
    _c(sheet, row, 1, "Valimimaht", color=_COL_NSIZE)
    _c(sheet, row, 2, n, color=_COL_NSIZE)
    row += 1

    return row


def _resolve_col_keys(plan, inspection) -> list | None:
    """Try to derive column letters for the data table from task inputs."""
    inputs = plan.inputs or {}
    x_range = inputs.get("x_range")
    y_range = inputs.get("y_range")
    data_range = inputs.get("data_range")

    keys = []
    for rng in [x_range, y_range, data_range]:
        if rng:
            k = _col_letter_from_range(rng)
            if k:
                keys.append(k)
    return keys if keys else None


# ---------------------------------------------------------------- task section writers
def _task_source_sheet(plan, inspection):
    return next(
        (s for s in inspection.sheets if s.name == plan.source_sheet),
        inspection.sheets[0] if inspection.sheets else None,
    )


def _write_corr_section(sheet, row: int, plan, result, inspection) -> int:
    v = result.values
    method = v.get("method", "pearson")
    coeff_key = "rho" if method == "spearman" else "r"
    coeff = v.get(coeff_key, "")
    p_val = v.get("p_value", "")

    src = _task_source_sheet(plan, inspection)
    if src and src.numeric_columns:
        col_keys = _resolve_col_keys(plan, inspection) or list(src.numeric_columns.keys())[:2]
        row = _mini_data_table(sheet, row, src, col_keys=col_keys)
        row += 1

    method_est = ("Spearmani astakkorrelatsioonikordajat"
                  if method == "spearman"
                  else "lineaarset korrelatsioonikordajat")
    _step(sheet, row, f"Leiame {method_est}.")
    row += 1

    _kv(sheet, row, f"Korrelatsioonikordaja {coeff_key} =", coeff,
        fmt="0.0000", highlight=True)
    row += 1
    if p_val != "":
        _kv(sheet, row, "p-väärtus =", p_val, fmt="0.0000")
        row += 1

    formulas = [f for f in (plan.excel_formulas or [])
                if f.get("formula", "").startswith("=")]
    if formulas:
        row += 1
        _step(sheet, row, "Kontroll Exceli funktsiooni abil:")
        row += 1
        for f in formulas[:2]:
            _write_formula_cell(sheet, row, 2, f.get("formula", ""), color=_COL_CALC)
            row += 1

    if isinstance(coeff, (int, float)):
        row += 1
        abs_r = abs(coeff)
        direction = "positiivne" if coeff > 0 else "negatiivne"
        strength = ("väga tugev" if abs_r > 0.8 else
                    "tugev" if abs_r > 0.6 else
                    "mõõdukas" if abs_r > 0.4 else "nõrk")
        _conclusion(sheet, row,
                    f"Järeldus. Korrelatsioonikordaja {coeff_key} = {coeff:.4f} "
                    f"näitab {strength} {direction} lineaarset seost muutujate vahel.")
        row += 1

    return row


def _write_reg_section(sheet, row: int, plan, result, inspection) -> int:
    v = result.values
    a = v.get("intercept", "")
    b = v.get("slope", "")
    r2 = v.get("r_squared", "")

    src = _task_source_sheet(plan, inspection)
    if src and src.numeric_columns:
        col_keys = _resolve_col_keys(plan, inspection) or list(src.numeric_columns.keys())[:2]
        row = _mini_data_table(sheet, row, src, col_keys=col_keys)
        row += 1

    _step(sheet, row, "Leiame regressioonisirge y = bx + a parameetrid.")
    row += 1
    _kv(sheet, row, "Vabaliige a =", a, fmt="0.00000000", highlight=True)
    row += 1
    _kv(sheet, row, "Regressioonikoefitsient b =", b, fmt="0.00000000", highlight=True)
    row += 1
    _kv(sheet, row, "Determinatsioonikordaja R² =", r2, fmt="0.0000", highlight=True)
    row += 1
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        eq = f"ŷ = {a:.4f} + {b:.4f} · x"
        _kv(sheet, row, "Regressioonivõrrand:", eq)
        row += 1

    formulas = [f for f in (plan.excel_formulas or [])
                if f.get("formula", "").startswith("=")]
    if formulas:
        row += 1
        _step(sheet, row, "Kontroll Exceli funktsioonide abil:")
        row += 1
        for f in formulas[:4]:
            label = f.get("cell", "")
            _kv(sheet, row, label, None)
            _write_formula_cell(sheet, row, 2, f.get("formula", ""), color=_COL_CALC)
            row += 1

    if isinstance(r2, (int, float)):
        row += 1
        _conclusion(sheet, row,
                    f"Järeldus. R² = {r2:.4f} — {r2*100:.1f}% sõltuva muutuja "
                    "dispersioonist on seletatav sõltumatu muutujaga.")
        row += 1

    return row


def _write_ci_section(sheet, row: int, plan, result) -> int:
    v = result.values
    alpha = v.get("alpha", 0.05)
    conf_pct = round((1 - alpha) * 100)
    ci_type = v.get("ci_type", "coefficients")

    _step(sheet, row, f"Leiame {conf_pct}%-lised usalduspiirid.")
    row += 1

    if ci_type == "mean_response":
        x_new = v.get("x_new", "")
        _kv(sheet, row, "x₀ =", x_new, fmt="0.00")
        row += 1
        _kv(sheet, row, f"Keskväärtuse {conf_pct}% UI alumine piir:",
            v.get("ci_lower", ""), fmt="0.0000", highlight=True)
        row += 1
        _kv(sheet, row, f"Keskväärtuse {conf_pct}% UI ülemine piir:",
            v.get("ci_upper", ""), fmt="0.0000", highlight=True)
        row += 1
    else:
        ci_data = [
            (f"Vabaliige a — alumine piir ({conf_pct}%):", "intercept_ci_lower"),
            (f"Vabaliige a — ülemine piir ({conf_pct}%):", "intercept_ci_upper"),
            (f"Koef. b — alumine piir ({conf_pct}%):",     "slope_ci_lower"),
            (f"Koef. b — ülemine piir ({conf_pct}%):",     "slope_ci_upper"),
        ]
        for label, key in ci_data:
            val = _ci_value(v, key)
            _kv(sheet, row, label, val, fmt="0.0000",
                highlight=(val not in ("", None)))
            row += 1

    return row


def _write_pred_section(sheet, row: int, plan, result) -> int:
    v = result.values
    x_new = v.get("x_new", "")
    y_pred = v.get("y_pred", "")
    alpha = v.get("alpha", 0.05)
    conf_pct = round((1 - alpha) * 100)

    _step(sheet, row, f"Leiame prognoosi x₀ = {x_new} korral.")
    row += 1
    _kv(sheet, row, "Punkthinnang ŷ =", y_pred, fmt="0.0000", highlight=True)
    row += 1

    if result.task_type == TaskType.PREDICTION_INTERVAL:
        lo = v.get("pi_lower", "")
        hi = v.get("pi_upper", "")
        extrap = v.get("is_extrapolation", False)
        if isinstance(lo, (int, float)):
            lo = math.floor(lo * 10000) / 10000
        if isinstance(hi, (int, float)):
            hi = math.ceil(hi * 10000) / 10000
        _kv(sheet, row, f"{conf_pct}% ennustusintervall — alumine piir:",
            lo, fmt="0.0000", highlight=True)
        row += 1
        _kv(sheet, row, f"{conf_pct}% ennustusintervall — ülemine piir:",
            hi, fmt="0.0000", highlight=True)
        row += 1
        if extrap:
            _conclusion(sheet, row,
                        f"Märkus. Ekstrapolatsioon: x₀ = {x_new} on väljaspool vaatlusandmete vahemikku.")
            row += 1

    return row


def _write_hyp_section(sheet, row: int, plan, result, inspection) -> int:
    v = result.values
    alpha = v.get("alpha", 0.05)
    inputs = plan.inputs or {}

    src = _task_source_sheet(plan, inspection)
    if src and src.numeric_columns:
        x_range = inputs.get("x_range") or inputs.get("data_range")
        y_range = inputs.get("y_range")
        if x_range:
            xk = _col_letter_from_range(x_range)
            yk = _col_letter_from_range(y_range) if y_range else None
            col_keys = [k for k in [xk, yk] if k]
            row = _mini_data_table(sheet, row, src, col_keys=col_keys or None)
        else:
            row = _mini_data_table(sheet, row, src)
        row += 1

    _step(sheet, row, "Hüpoteeside kontrollimine.")
    row += 1

    mu0 = inputs.get("mu0")
    t_stat = v.get("t_statistic", "")
    p_val = v.get("p_value", "")
    reject = v.get("reject_null", None)
    mean1 = v.get("mean1", "")
    mean2 = v.get("mean2", "")
    df = v.get("df", "")

    if mu0 is not None:
        _kv(sheet, row, "H₀: μ =", mu0, fmt="0.00")
        row += 1
    if mean1 != "":
        lbl = "Valimi 1 keskmine x̄₁ =" if mean2 != "" else "Valimi keskmine x̄ ="
        _kv(sheet, row, lbl, mean1, fmt="0.0000")
        row += 1
    if mean2 != "":
        _kv(sheet, row, "Valimi 2 keskmine x̄₂ =", mean2, fmt="0.0000")
        row += 1
    if df != "":
        _kv(sheet, row, "Vabadusastmed df =", df, fmt="0")
        row += 1

    _kv(sheet, row, "t-statistik =", t_stat, fmt="0.0000", highlight=True)
    row += 1
    _kv(sheet, row, "p-väärtus =", p_val, fmt="0.0000", highlight=True)
    row += 1
    _kv(sheet, row, "Olulisustase α =", alpha, fmt="0.00")
    row += 1

    if isinstance(reject, bool) and isinstance(p_val, (int, float)):
        row += 1
        if reject:
            _conclusion(sheet, row,
                        f"Järeldus. Kuna p = {p_val:.4f} < α = {alpha}, lükatakse nullhüpotees tagasi.")
        else:
            _conclusion(sheet, row,
                        f"Järeldus. Kuna p = {p_val:.4f} ≥ α = {alpha}, ei ole alust nullhüpoteesi tagasi lükata.")
        row += 1

    return row


def _write_desc_section(sheet, row: int, plan, result, inspection) -> int:
    v = result.values
    inputs = plan.inputs or {}

    src = _task_source_sheet(plan, inspection)
    if src and src.numeric_columns:
        data_range = inputs.get("data_range") or inputs.get("x_range")
        if data_range:
            k = _col_letter_from_range(data_range)
            col_keys = [k] if k else None
        else:
            col_keys = list(src.numeric_columns.keys())[:1]
        row = _mini_data_table(sheet, row, src, col_keys=col_keys)
        row += 1

    _step(sheet, row, "Kirjeldav statistika:")
    row += 1

    key_stats = {"min", "max", "mean", "median", "std"}
    for label, key, fmt in [
        ("n =",                      "n",      "0"),
        ("Miinimum =",               "min",    "0.0000"),
        ("Maksimum =",               "max",    "0.0000"),
        ("Aritmeetiline keskmine =", "mean",   "0.0000"),
        ("Mediaan =",                "median", "0.0000"),
        ("Standardhälve s =",        "std",    "0.0000"),
        ("Dispersioon s² =",         "var",    "0.0000"),
    ]:
        val = v.get(key)
        if val is not None:
            _kv(sheet, row, label, val, fmt=fmt, highlight=(key in key_stats))
            row += 1

    return row


# ---------------------------------------------------------------- sheet writers
def write_lahendus(xw_book, task_plans, results, inspection) -> str:
    sheet = get_or_create_ai_sheet(xw_book, SHEET_LAHENDUS)
    _set_col_widths(sheet, {"A": 38, "B": 16, "C": 16, "D": 16, "E": 16})

    result_by_id = {r.task_id: r for r in results}
    row = 1

    if not task_plans:
        try:
            sheet.range("A1").value = "Ülesandeid ei tuvastatud."
        except Exception:
            pass
        row += 1

    for task_idx, plan in enumerate(task_plans):
        if task_idx > 0:
            row += 1  # blank row between tasks

        result = result_by_id.get(plan.task_id)

        # Task description as plain bold heading
        _section_title(sheet, row, plan.description[:120])
        row += 1

        if result is None or result.error:
            err = result.error if result else "ootel"
            _c(sheet, row, 1, f"Lahendamata: {err}", color=_COL_ERROR)
            row += 2
            continue

        tt = result.task_type

        if tt == TaskType.CORRELATION:
            row = _write_corr_section(sheet, row, plan, result, inspection)
        elif tt == TaskType.REGRESSION:
            row = _write_reg_section(sheet, row, plan, result, inspection)
        elif tt == TaskType.CONFIDENCE_INTERVAL:
            row = _write_ci_section(sheet, row, plan, result)
        elif tt in (TaskType.PREDICTION, TaskType.PREDICTION_INTERVAL):
            row = _write_pred_section(sheet, row, plan, result)
        elif tt == TaskType.HYPOTHESIS_TEST:
            row = _write_hyp_section(sheet, row, plan, result, inspection)
        elif tt == TaskType.DESCRIPTIVE_STATS:
            row = _write_desc_section(sheet, row, plan, result, inspection)
        elif tt == TaskType.CHART:
            _step(sheet, row, "Graafik: vt Regression leht.")
            row += 1
        else:
            for k, val in result.values.items():
                if isinstance(val, (int, float, str, bool)):
                    _kv(sheet, row, k, val)
                    row += 1

        # AI explanation (only if non-trivial and not already shown as conclusion)
        expl = (result.explanation or "").strip()
        if expl and len(expl) > 20:
            row += 1
            _conclusion(sheet, row, expl[:400])
            row += 1

    return SHEET_LAHENDUS


def write_arvutused(xw_book, task_plans, results, inspection) -> str:
    """Detailed formula verification sheet — compact, no banners."""
    sheet = get_or_create_ai_sheet(xw_book, SHEET_ARVUTUSED)
    _set_col_widths(sheet, {"A": 14, "B": 42, "C": 18, "D": 18})

    plan_by_id = {p.task_id: p for p in task_plans}
    row = 1

    _section_title(sheet, row, "Arvutused ja kontroll")
    row += 2

    for result in results:
        if result.error:
            _c(sheet, row, 1, f"{result.task_id}: {result.error}", color=_COL_ERROR)
            row += 2
            continue

        plan = plan_by_id.get(result.task_id)
        if not plan or not plan.excel_formulas:
            continue

        _section_title(sheet, row, f"{result.task_id}: {result.description[:80]}")
        row += 1
        _step(sheet, row, "Kontroll Exceli funktsiooni abil:")
        row += 1

        for f in plan.excel_formulas[:10]:
            cell_ref = f.get("cell", "")
            formula = f.get("formula", "")
            if not formula:
                continue
            _c(sheet, row, 1, cell_ref)
            _write_formula_cell(sheet, row, 2, formula, color=_COL_CALC)
            row += 1

        row += 1

    return SHEET_ARVUTUSED


def write_regression(xw_book, task_plans, results, inspection) -> str:
    """Regression-specific sheet with academic table layout."""
    sheet = get_or_create_ai_sheet(xw_book, SHEET_REGRESSION)
    _set_col_widths(sheet, {"A": 32, "B": 16, "C": 16, "D": 16, "E": 16})

    def _first(*types):
        for r in results:
            if not r.error and r.task_type in types and r.values:
                return r
        return None

    reg_result = _first(TaskType.REGRESSION)
    ci_result  = _first(TaskType.CONFIDENCE_INTERVAL)
    pred_result = _first(TaskType.PREDICTION, TaskType.PREDICTION_INTERVAL)

    # Identify data columns
    first_sheet = inspection.sheets[0] if inspection.sheets else None
    num_cols = list(first_sheet.numeric_columns.items()) if first_sheet else []
    x_vals = list(num_cols[0][1]) if len(num_cols) >= 1 else []
    y_vals = list(num_cols[1][1]) if len(num_cols) >= 2 else []
    n_pts = min(len(x_vals), len(y_vals))
    x_vals, y_vals = x_vals[:n_pts], y_vals[:n_pts]

    # Regression statistics
    row = 1
    if reg_result:
        rv = reg_result.values
        a   = rv.get("intercept")
        b   = rv.get("slope")
        r2  = rv.get("r_squared")
        n   = rv.get("n", n_pts)
        se_a = rv.get("intercept_std_error") or rv.get("se_intercept")
        se_b = rv.get("slope_std_error") or rv.get("se_slope")

        _section_title(sheet, row, "Regressioonistatistika")
        row += 1
        for label, val, fmt in [
            ("Valimi maht n =",              n,   "0"),
            ("Vabaliige a =",                a,   "0.00000000"),
            ("Regressioonikoefitsient b =",  b,   "0.00000000"),
            ("Determinatsioonikordaja R² =", r2,  "0.00000000"),
        ]:
            _kv(sheet, row, label, val, fmt=fmt, highlight=(val is not None))
            row += 1
        if se_a is not None:
            _kv(sheet, row, "Vabaliige a standardviga =", se_a, fmt="0.00000000")
            row += 1
        if se_b is not None:
            _kv(sheet, row, "Koefitsient b standardviga =", se_b, fmt="0.00000000")
            row += 1
        if a is not None and b is not None:
            _kv(sheet, row, "Regressioonivõrrand:", f"ŷ = {a:.4f} + {b:.4f} · x")
            row += 1
        row += 1

    # Confidence intervals
    if ci_result:
        cv = ci_result.values
        ci_alpha = cv.get("alpha", 0.05)
        ci_pct = round((1 - ci_alpha) * 100)
        _section_title(sheet, row, f"Usaldusintervallid ({ci_pct}%)")
        row += 1
        ci_type = cv.get("ci_type", "coefficients")
        if ci_type == "mean_response":
            _kv(sheet, row, f"x₀ =", cv.get("x_new", ""), fmt="0.00")
            row += 1
            _kv(sheet, row, f"Keskväärtuse {ci_pct}% UI alumine:",
                cv.get("ci_lower"), fmt="0.0000", highlight=True)
            row += 1
            _kv(sheet, row, f"Keskväärtuse {ci_pct}% UI ülemine:",
                cv.get("ci_upper"), fmt="0.0000", highlight=True)
            row += 1
        else:
            for label, key in [
                (f"Vabaliige a alumine ({ci_pct}%):", "intercept_ci_lower"),
                (f"Vabaliige a ülemine ({ci_pct}%):", "intercept_ci_upper"),
                (f"Koef. b alumine ({ci_pct}%):",     "slope_ci_lower"),
                (f"Koef. b ülemine ({ci_pct}%):",     "slope_ci_upper"),
            ]:
                val = _ci_value(cv, key)
                _kv(sheet, row, label, val, fmt="0.0000",
                    highlight=(val not in ("", None)))
                row += 1
        row += 1

    # Prediction
    if pred_result:
        pv = pred_result.values
        _section_title(sheet, row, "Prognoos")
        row += 1
        x_new = pv.get("x_new")
        y_pred = pv.get("y_pred")
        alpha = pv.get("alpha", 0.05)
        conf_pct = round((1 - alpha) * 100)
        _kv(sheet, row, "x₀ =", x_new, fmt="0.00")
        row += 1
        _kv(sheet, row, "Punkthinnang ŷ =", y_pred, fmt="0.0000", highlight=True)
        row += 1
        if pred_result.task_type == TaskType.PREDICTION_INTERVAL:
            lo = pv.get("pi_lower")
            hi = pv.get("pi_upper")
            extrap = pv.get("is_extrapolation", False)
            if isinstance(lo, (int, float)):
                lo = math.floor(lo * 10000) / 10000
            if isinstance(hi, (int, float)):
                hi = math.ceil(hi * 10000) / 10000
            _kv(sheet, row, f"{conf_pct}% PI alumine:", lo, fmt="0.0000", highlight=True)
            row += 1
            _kv(sheet, row, f"{conf_pct}% PI ülemine:", hi, fmt="0.0000", highlight=True)
            row += 1
            if extrap:
                _conclusion(sheet, row,
                            "Ekstrapolatsioon: x₀ on väljaspool vaatlusandmete vahemikku.")
                row += 1
        row += 1

    # Chart
    if reg_result and x_vals and y_vals:
        row += 1
        _create_scatter_with_trendline(
            sheet, row, x_vals, y_vals,
            x_label=num_cols[0][0] if num_cols else "x",
            y_label=num_cols[1][0] if len(num_cols) > 1 else "y",
            title="Hajuvusdiagramm koos regressioonisirgega",
        )

    return SHEET_REGRESSION


def _create_scatter_with_trendline(sheet, row_offset, x_vals, y_vals,
                                    x_label, y_label, title) -> None:
    n = len(x_vals)
    try:
        sheet.range((row_offset, 10)).value = x_label
        sheet.range((row_offset, 11)).value = y_label
        for i in range(n):
            sheet.range((row_offset + 1 + i, 10)).value = x_vals[i]
            sheet.range((row_offset + 1 + i, 11)).value = y_vals[i]
    except Exception:
        pass
    try:
        chart = sheet.charts.add(
            left=sheet.range("A1").left,
            top=sheet.range(f"A{row_offset}").top,
            width=380, height=260,
        )
        chart.chart_type = "xy_scatter"
        data_range = sheet.range((row_offset, 10), (row_offset + n, 11))
        chart.set_source_data(data_range)
        try:
            chart.api.HasTitle = True
            chart.api.ChartTitle.Text = title
        except Exception:
            pass
        try:
            sc = chart.api.SeriesCollection(1)
            tl = sc.Trendlines().Add()
            tl.Type = -4132  # xlLinear (not 5 = xlExponential)
            tl.DisplayEquation = True
            tl.DisplayRSquared = True
        except Exception:
            pass
    except Exception:
        pass


def write_kordamisküsimused(xw_book, questions: list[dict]) -> str:
    sheet = get_or_create_ai_sheet(xw_book, SHEET_KORDAMIS)
    _set_col_widths(sheet, {"A": 6, "B": 62, "C": 18, "D": 18})

    _section_title(sheet, 1, "Kordamisküsimused")
    _step(sheet, 2, "Vasta küsimustele. Vastused on eraldi lehel 'Vastused'.")
    row = 3

    for i, q in enumerate(questions, 1):
        question_text = q.get("question", "")
        options = q.get("options", [])
        _c(sheet, row, 1, f"K{i}", bold=True)
        _c(sheet, row, 2, question_text, bold=True)
        row += 1
        for j, opt in enumerate(options[:4]):
            _c(sheet, row, 1, f"  {chr(65+j)})")
            _c(sheet, row, 2, opt)
            row += 1
        row += 1

    return SHEET_KORDAMIS


def write_vastused(xw_book, questions: list[dict]) -> str:
    sheet = get_or_create_ai_sheet(xw_book, SHEET_VASTUSED)
    _set_col_widths(sheet, {"A": 6, "B": 18, "C": 50})

    _section_title(sheet, 1, "Vastused")
    for ci, label in enumerate(["K#", "Vastus", "Selgitus"], 1):
        _c(sheet, 2, ci, label, bold=True)
    row = 3

    for i, q in enumerate(questions, 1):
        correct = q.get("correct_option", "")
        answer = q.get("answer", "")
        display = f"{correct}) {answer}" if correct and answer else (correct or answer)
        explanation = q.get("explanation", "")
        _c(sheet, row, 1, f"K{i}")
        _c(sheet, row, 2, display, color=_COL_RESULT)
        try:
            cc = sheet.range((row, 3))
            cc.value = explanation
            cc.font.name = _FONT
            cc.font.size = _FONT_SZ
            try:
                cc.api.WrapText = True
            except Exception:
                pass
        except Exception:
            pass
        row += 1

    return SHEET_VASTUSED


def write_all_sheets(xw_book, analysis: WorkbookAnalysis,
                     results: list[TaskResult], inspection) -> dict:
    """Write all relevant sheets. Returns {"written": list, "skipped": list, "errors": list}."""
    written, skipped, errors = [], [], []

    task_plans = analysis.tasks

    try:
        write_lahendus(xw_book, task_plans, results, inspection)
        written.append(SHEET_LAHENDUS)
    except ValueError as e:
        skipped.append((SHEET_LAHENDUS, str(e)))
    except Exception as e:
        errors.append((SHEET_LAHENDUS, str(e)))

    has_regression = any(r.task_type == TaskType.REGRESSION for r in results if not r.error)
    has_ci = any(r.task_type == TaskType.CONFIDENCE_INTERVAL for r in results if not r.error)
    has_pred = any(r.task_type in (TaskType.PREDICTION, TaskType.PREDICTION_INTERVAL)
                   for r in results if not r.error)
    has_formulas = any(
        p.excel_formulas for p in task_plans
        if any(r.task_id == p.task_id and not r.error for r in results)
    )

    if has_formulas:
        try:
            write_arvutused(xw_book, task_plans, results, inspection)
            written.append(SHEET_ARVUTUSED)
        except ValueError as e:
            skipped.append((SHEET_ARVUTUSED, str(e)))
        except Exception as e:
            errors.append((SHEET_ARVUTUSED, str(e)))

    if has_regression or has_ci or has_pred:
        try:
            write_regression(xw_book, task_plans, results, inspection)
            written.append(SHEET_REGRESSION)
        except ValueError as e:
            skipped.append((SHEET_REGRESSION, str(e)))
        except Exception as e:
            errors.append((SHEET_REGRESSION, str(e)))

    questions = getattr(analysis, "revision_questions", []) or []
    if questions:
        for fn, name in [(write_kordamisküsimused, SHEET_KORDAMIS),
                          (write_vastused, SHEET_VASTUSED)]:
            try:
                fn(xw_book, questions)
                written.append(name)
            except ValueError as e:
                skipped.append((name, str(e)))
            except Exception as e:
                errors.append((name, str(e)))

    return {"written": written, "skipped": skipped, "errors": errors}


def verify_workbook(xw_book, original_sheet: str, expected_sheets: list[str]) -> list[str]:
    """Run QA checks. Returns list of issue strings (empty means all OK)."""
    issues = []
    existing = {s.name for s in xw_book.sheets}

    if original_sheet and original_sheet not in existing:
        issues.append(f"Algne leht '{original_sheet}' puudub!")
    elif original_sheet:
        orig = xw_book.sheets[original_sheet]
        if _is_ai_sheet(orig):
            issues.append(
                f"Algne leht '{original_sheet}' on märgistatud AI-genereerituna "
                "— võimalik, et kirjutati üle."
            )

    for name in expected_sheets:
        if name not in existing:
            issues.append(f"Oodatav leht '{name}' puudub.")

    if SHEET_LAHENDUS in existing:
        try:
            val = xw_book.sheets[SHEET_LAHENDUS].range("A1").value
            if not val:
                issues.append(f"Leht '{SHEET_LAHENDUS}' on tühi.")
        except Exception:
            pass

    return issues


def save_workbook(xw_book) -> None:
    xw_book.save()
