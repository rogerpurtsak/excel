"""In-place worksheet writer for Excel AI Solver.

All modifications go directly into the already-open workbook via xlwings.
Never writes to the file externally with openpyxl while Excel has it open.
"""
from __future__ import annotations

import math
import shutil
from datetime import datetime
from pathlib import Path
from typing import Any

from .schemas import TaskResult, TaskType, WorkbookAnalysis
from . import workbook as wb_mod

# Marker written to cell ZZ1 of every AI-generated sheet.
# This lets re-runs detect their own sheets without touching user content.
AI_MARKER = "ExcelAISolver:generated"
_MARKER_CELL = "ZZ1"

# Standard sheet names
SHEET_LAHENDUS      = "Lahendus"
SHEET_ARVUTUSED     = "Arvutused ja kontroll"
SHEET_REGRESSION    = "Regression"
SHEET_KORDAMIS      = "Kordamisküsimused"
SHEET_VASTUSED      = "Vastused"

# Color palette — RGB tuples for xlwings
_BG_H1      = (0,   70,  127)   # dark navy
_FG_H1      = (255, 255, 255)   # white
_BG_H2      = (91,  155, 213)   # medium blue
_FG_H2      = (255, 255, 255)
_BG_H3      = (217, 225, 242)   # light blue
_BG_RESULT  = (255, 255, 153)   # yellow highlight
_BG_FORMULA = (235, 241, 222)   # light green
_BG_GRAY    = (242, 242, 242)   # light gray


# ---------------------------------------------------------------- CI helpers
def _ci_value(values: dict[str, Any], flat_key: str) -> Any:
    """Fetch a CI bound either from a flat key (e.g. 'slope_ci_lower') or from
    the nested tuple form produced by math_engine ('slope': (lo, hi)).
    Returns '' when unavailable.
    """
    if flat_key in values:
        return values[flat_key]
    # Map flat keys onto (base, index).
    mapping = {
        "intercept_ci_lower": ("intercept", 0),
        "intercept_ci_upper": ("intercept", 1),
        "slope_ci_lower": ("slope", 0),
        "slope_ci_upper": ("slope", 1),
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
    """Copy the file to the backups directory with a timestamped name.
    Returns the backup path string. Does NOT open the backup in Excel.
    """
    src = Path(path)
    dest_dir = _backup_dir()
    dest_dir.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
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
    """Return an AI-owned sheet named `name`.

    Reuses (and clears) a sheet we previously marked. Refuses to overwrite a
    user-edited sheet of the same name. Creates a new sheet otherwise.
    """
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
    # Create a new sheet after the last one.
    try:
        last = xw_book.sheets[-1]
        sheet = xw_book.sheets.add(name, after=last)
    except Exception:
        sheet = xw_book.sheets.add(name)
    _mark_sheet(sheet)
    return sheet


# ---------------------------------------------------------------- formatting
def _col_letter(n: int) -> str:
    return chr(64 + n)


def _set_col_widths(sheet, widths: dict) -> None:
    for col, w in widths.items():
        try:
            sheet.range(f"{col}:{col}").column_width = w
        except Exception:
            pass


def _h1(sheet, row, text, n_cols=5) -> None:
    try:
        rng = sheet.range(f"A{row}:{_col_letter(n_cols)}{row}")
        try:
            rng.merge()
        except Exception:
            pass
        cell = sheet.range(f"A{row}")
        cell.value = text
        cell.font.bold = True
        cell.font.size = 14
        cell.font.color = _FG_H1
        cell.color = _BG_H1
        try:
            sheet.range(f"A{row}").row_height = 24
        except Exception:
            pass
    except Exception:
        pass


def _h2(sheet, row, text, n_cols=5) -> None:
    try:
        rng = sheet.range(f"A{row}:{_col_letter(n_cols)}{row}")
        try:
            rng.merge()
        except Exception:
            pass
        cell = sheet.range(f"A{row}")
        cell.value = text
        cell.font.bold = True
        cell.font.size = 12
        cell.font.color = _FG_H2
        cell.color = _BG_H2
    except Exception:
        pass


def _kv(sheet, row, key, value, num_fmt=None, highlight=False, col_key=1, col_val=2) -> None:
    try:
        kcell = sheet.range((row, col_key))
        kcell.value = key
        kcell.font.bold = True
    except Exception:
        pass
    try:
        vcell = sheet.range((row, col_val))
        vcell.value = value
        if num_fmt:
            try:
                vcell.number_format = num_fmt
            except Exception:
                pass
        if highlight:
            try:
                vcell.color = _BG_RESULT
                vcell.font.bold = True
            except Exception:
                pass
    except Exception:
        pass


def _formula_row(sheet, row, label, formula_str, computed_value, num_fmt=None) -> None:
    try:
        a = sheet.range(f"A{row}")
        a.value = label
        a.font.bold = True
    except Exception:
        pass
    try:
        b = sheet.range(f"B{row}")
        b.value = formula_str
        b.color = _BG_FORMULA
    except Exception:
        pass
    try:
        c = sheet.range(f"C{row}")
        c.value = computed_value
        c.color = _BG_RESULT
        if num_fmt:
            try:
                c.number_format = num_fmt
            except Exception:
                pass
    except Exception:
        pass


def _explanation_row(sheet, row, text, n_cols=5) -> None:
    try:
        rng = sheet.range(f"A{row}:{_col_letter(n_cols)}{row}")
        try:
            rng.merge()
        except Exception:
            pass
        cell = sheet.range(f"A{row}")
        cell.value = text
        cell.font.italic = True
        cell.font.size = 10
        cell.color = _BG_GRAY
        try:
            cell.api.WrapText = True
        except Exception:
            pass
        try:
            sheet.range(f"A{row}").row_height = 40
        except Exception:
            pass
    except Exception:
        pass


# ---------------------------------------------------------------- sheet writers
def write_lahendus(xw_book, task_plans, results, inspection) -> str:
    sheet = get_or_create_ai_sheet(xw_book, SHEET_LAHENDUS)
    _set_col_widths(sheet, {"A": 35, "B": 18, "C": 18, "D": 18, "E": 40})

    _h1(sheet, 1, "Lahendus")
    _kv(sheet, 2, "Töövihik:", Path(inspection.path).name)
    row = 4

    result_by_id = {r.task_id: r for r in results}

    for plan in task_plans:
        result = result_by_id.get(plan.task_id)
        _h2(sheet, row, f"{plan.task_id}: {plan.description[:80]}", n_cols=5)
        row += 1

        if result is None or result.error:
            try:
                c = sheet.range(f"A{row}")
                c.value = f"⚠ Lahendamata: {result.error if result else 'ootel'}"
                c.color = (255, 200, 200)
            except Exception:
                pass
            row += 2
            continue

        v = result.values

        if result.task_type == TaskType.CORRELATION:
            r_val = v.get("r", "")
            p_val = v.get("p_value", "")
            _kv(sheet, row, "Pearsoni korrelatsioonikordaja r:", r_val, num_fmt="0.0000", highlight=True)
            row += 1
            _kv(sheet, row, "p-väärtus:", p_val, num_fmt="0.0000")
            row += 1
            if isinstance(r_val, (int, float)):
                abs_r = abs(r_val)
                direction = "positiivne" if r_val > 0 else "negatiivne"
                if abs_r > 0.8:
                    strength = "väga tugev"
                elif abs_r > 0.6:
                    strength = "tugev"
                elif abs_r > 0.4:
                    strength = "mõõdukas"
                else:
                    strength = "nõrk"
                interp = f"Korrelatsioonikordaja r = {r_val:.4f} näitab {strength} {direction} lineaarset seost muutujate vahel."
                _explanation_row(sheet, row, interp)
                row += 1

        elif result.task_type == TaskType.REGRESSION:
            a = v.get("intercept", "")
            b = v.get("slope", "")
            r2 = v.get("r_squared", "")
            _kv(sheet, row, "Vabaliige a:", a, num_fmt="0.00000000", highlight=True)
            row += 1
            _kv(sheet, row, "Regressioonikoefitsient b:", b, num_fmt="0.00000000", highlight=True)
            row += 1
            _kv(sheet, row, "Determinatsioonikordaja R²:", r2, num_fmt="0.0000", highlight=True)
            row += 1
            if isinstance(a, (int, float)) and isinstance(b, (int, float)):
                eq = f"Regressioonisirgue võrrand:  ŷ = {a:.4f} + {b:.4f} · x"
                _kv(sheet, row, "Võrrand:", eq)
                row += 1
                if isinstance(r2, (int, float)):
                    interp = f"R² = {r2:.4f} tähendab, et {r2*100:.1f}% sõltuva muutuja dispersioonist on seletatav sõltumatu muutujaga."
                    _explanation_row(sheet, row, interp)
                    row += 1

        elif result.task_type == TaskType.CONFIDENCE_INTERVAL:
            for key_label, val_key, fmt in [
                ("Vabaliikme a alumine piir (95%):", "intercept_ci_lower", "0.0000"),
                ("Vabaliikme a ülemine piir (95%):", "intercept_ci_upper", "0.0000"),
                ("Koefitsiendi b alumine piir (95%):", "slope_ci_lower", "0.0000"),
                ("Koefitsiendi b ülemine piir (95%):", "slope_ci_upper", "0.0000"),
            ]:
                val = _ci_value(v, val_key)
                _kv(sheet, row, key_label, val, num_fmt=fmt, highlight=bool(val != ""))
                row += 1
            _explanation_row(sheet, row, "Usaldusintervallid on arvutatud 95% usaldusnivool (α=0.05, kahepoolne).")
            row += 1

        elif result.task_type in (TaskType.PREDICTION, TaskType.PREDICTION_INTERVAL):
            x_new = v.get("x_new", "")
            y_pred = v.get("y_pred", "")
            _kv(sheet, row, "Prognoospunkt x₀:", x_new, num_fmt="0.0")
            row += 1
            _kv(sheet, row, "Punkthinnang ŷ:", y_pred, num_fmt="0.0000", highlight=True)
            row += 1
            if result.task_type == TaskType.PREDICTION_INTERVAL:
                lo = v.get("pi_lower", "")
                hi = v.get("pi_upper", "")
                extrap = v.get("is_extrapolation", False)
                # Outward rounding to the 4 displayed decimals.
                if isinstance(lo, (int, float)):
                    lo = math.floor(lo * 10000) / 10000
                if isinstance(hi, (int, float)):
                    hi = math.ceil(hi * 10000) / 10000
                _kv(sheet, row, "95% ennustusintervall — alumine piir:", lo, num_fmt="0.0000", highlight=True)
                row += 1
                _kv(sheet, row, "95% ennustusintervall — ülemine piir:", hi, num_fmt="0.0000", highlight=True)
                row += 1
                if extrap:
                    _explanation_row(sheet, row, f"⚠ Ekstrapolatsioon: x₀={x_new} on väljaspool vaatlusandmete vahemikku. Tulemus on prognoos, mitte interpolatsioon.")
                    row += 1

        elif result.task_type == TaskType.CHART:
            try:
                c = sheet.range(f"A{row}")
                c.value = "Graafik: vt Regression leht"
                c.font.italic = True
            except Exception:
                pass
            row += 1

        else:
            for k, val in v.items():
                if isinstance(val, (int, float, str)):
                    _kv(sheet, row, k, val)
                    row += 1

        if result.explanation and len(result.explanation) > 10:
            _explanation_row(sheet, row, result.explanation[:500])
            row += 1

        row += 1  # blank row between tasks

    return SHEET_LAHENDUS


def write_arvutused(xw_book, task_plans, results, inspection) -> str:
    sheet = get_or_create_ai_sheet(xw_book, SHEET_ARVUTUSED)
    _set_col_widths(sheet, {"A": 40, "B": 22, "C": 22, "D": 22})

    _h1(sheet, 1, "Arvutused ja kontroll")
    row = 2

    plan_by_id = {p.task_id: p for p in task_plans}

    # Source data
    _h2(sheet, row, "Sisendandmed")
    row += 1
    source_sheet = next(
        (s for s in inspection.sheets if len(s.numeric_columns) >= 2),
        inspection.sheets[0] if inspection.sheets else None,
    )
    if source_sheet and source_sheet.numeric_columns:
        cols = list(source_sheet.numeric_columns.items())
        for ci, (letter, _vals) in enumerate(cols, 1):
            try:
                hc = sheet.range((row, ci))
                hc.value = letter
                hc.font.bold = True
                hc.color = _BG_GRAY
            except Exception:
                pass
        row += 1
        n = max((len(v) for _k, v in cols), default=0)
        for ri in range(n):
            for ci, (_letter, vals) in enumerate(cols, 1):
                if ri < len(vals):
                    try:
                        sheet.range((row, ci)).value = vals[ri]
                    except Exception:
                        pass
            row += 1
        row += 1  # blank after data

    row += 1
    for result in results:
        if result.error:
            try:
                c = sheet.range(f"A{row}")
                c.value = f"⚠ {result.task_id}: {result.error}"
                c.color = (255, 200, 200)
            except Exception:
                pass
            row += 2
            continue

        _h2(sheet, row, f"{result.task_id}: {result.description[:60]}")
        row += 1

        v = result.values
        for k, val in v.items():
            if isinstance(val, bool):
                _kv(sheet, row, k, str(val))
            elif isinstance(val, (int, float)):
                fmt = "0.00000000" if abs(val) > 0.0001 else "0.00000E+00"
                _kv(sheet, row, k, val, num_fmt=fmt,
                    highlight=(k in ("r", "intercept", "slope", "r_squared", "y_pred", "pi_lower", "pi_upper")))
            elif isinstance(val, str):
                _kv(sheet, row, k, val)
            else:
                continue
            row += 1

        plan = plan_by_id.get(result.task_id)
        if plan and plan.excel_formulas:
            _h2(sheet, row, "Exceli valemid (kontroll)", n_cols=4)
            row += 1
            for f in plan.excel_formulas[:8]:
                cell_ref = f.get("cell", "")
                formula = f.get("formula", "")
                if cell_ref and formula:
                    try:
                        sheet.range(f"A{row}").value = cell_ref
                        fb = sheet.range(f"B{row}")
                        fb.value = formula
                        fb.color = _BG_FORMULA
                    except Exception:
                        pass
                    row += 1

        row += 1  # blank

    return SHEET_ARVUTUSED


def write_regression(xw_book, task_plans, results, inspection) -> str:
    sheet = get_or_create_ai_sheet(xw_book, SHEET_REGRESSION)
    _set_col_widths(sheet, {"A": 35, "B": 20, "C": 20, "D": 20, "E": 20})

    _h1(sheet, 1, "Regressioonianalüüs")
    row = 2

    first_sheet = inspection.sheets[0] if inspection.sheets else None
    num_cols = list(first_sheet.numeric_columns.items()) if first_sheet else []
    x_vals = num_cols[0][1] if len(num_cols) >= 1 else []
    y_vals = num_cols[1][1] if len(num_cols) >= 2 else []
    n_pts = min(len(x_vals), len(y_vals))
    x_vals = list(x_vals[:n_pts])
    y_vals = list(y_vals[:n_pts])

    def _first(*types):
        for r in results:
            if not r.error and r.task_type in types and r.values:
                return r
        return None

    reg_result = _first(TaskType.REGRESSION)
    ci_result = _first(TaskType.CONFIDENCE_INTERVAL)
    pred_result = _first(TaskType.PREDICTION, TaskType.PREDICTION_INTERVAL)

    # Data table
    _h2(sheet, row, "Sisendandmed")
    row += 1
    for ci, (letter, _vals) in enumerate(num_cols[:2], 1):
        try:
            hc = sheet.range((row, ci))
            hc.value = letter
            hc.font.bold = True
            hc.color = _BG_GRAY
        except Exception:
            pass
    row += 1
    for ri in range(n_pts):
        try:
            if len(x_vals) > ri:
                sheet.range((row, 1)).value = x_vals[ri]
            if len(y_vals) > ri:
                sheet.range((row, 2)).value = y_vals[ri]
        except Exception:
            pass
        row += 1

    # Regression summary
    row += 1
    _h2(sheet, row, "Regressioonikokkuvõte")
    row += 1
    if reg_result:
        rv = reg_result.values
        a = rv.get("intercept")
        b = rv.get("slope")
        r2 = rv.get("r_squared")
        n = rv.get("n", len(x_vals))
        se_a = rv.get("intercept_std_error", rv.get("se_intercept"))
        se_b = rv.get("slope_std_error", rv.get("se_slope"))

        rows_to_write = [
            ("n (vaatluste arv)", n, "0"),
            ("Vabaliige a", a, "0.00000000"),
            ("Regressioonikoefitsient b", b, "0.00000000"),
            ("Determinatsioonikordaja R²", r2, "0.00000000"),
        ]
        if se_a:
            rows_to_write.append(("Vabaliige SE", se_a, "0.00000000"))
        if se_b:
            rows_to_write.append(("Koef. b SE", se_b, "0.00000000"))

        for label, val, fmt in rows_to_write:
            _kv(sheet, row, label, val, num_fmt=fmt, highlight=(val is not None))
            row += 1

        if a is not None and b is not None:
            eq = f"ŷ = {a:.4f} + {b:.4f} · x"
            _kv(sheet, row, "Regressioonivõrrand:", eq)
            row += 1

        note = "Märkus: Need tulemused on arvutatud Pythoni (NumPy/SciPy) abil, mis on matemaatiliselt samaväärne Exceli Regression tööriistaga."
        _explanation_row(sheet, row, note)
        row += 1

    # Confidence intervals
    if ci_result:
        row += 1
        _h2(sheet, row, "Regressioonikoefitsientide usaldusintervallid (95%)")
        row += 1
        cv = ci_result.values
        ci_rows = [
            ("Vabaliige a — alumine piir", _ci_value(cv, "intercept_ci_lower"), "0.0000"),
            ("Vabaliige a — ülemine piir", _ci_value(cv, "intercept_ci_upper"), "0.0000"),
            ("Koef. b — alumine piir", _ci_value(cv, "slope_ci_lower"), "0.0000"),
            ("Koef. b — ülemine piir", _ci_value(cv, "slope_ci_upper"), "0.0000"),
        ]
        for label, val, fmt in ci_rows:
            has = val != "" and val is not None
            _kv(sheet, row, label, val, num_fmt=fmt, highlight=has)
            row += 1

    # Prediction
    if pred_result:
        row += 1
        _h2(sheet, row, "Prognoos")
        row += 1
        pv = pred_result.values
        _kv(sheet, row, "Prognoosi x₀:", pv.get("x_new"), num_fmt="0.0")
        row += 1
        _kv(sheet, row, "Punkthinnang ŷ:", pv.get("y_pred"), num_fmt="0.0000", highlight=True)
        row += 1
        if pred_result.task_type == TaskType.PREDICTION_INTERVAL:
            lo = pv.get("pi_lower")
            hi = pv.get("pi_upper")
            extrap = pv.get("is_extrapolation", False)
            if isinstance(lo, (int, float)):
                lo = math.floor(lo * 10000) / 10000
            if isinstance(hi, (int, float)):
                hi = math.ceil(hi * 10000) / 10000
            _kv(sheet, row, "95% PI — alumine:", lo, num_fmt="0.0000", highlight=True)
            row += 1
            _kv(sheet, row, "95% PI — ülemine:", hi, num_fmt="0.0000", highlight=True)
            row += 1
            if extrap:
                _explanation_row(sheet, row, "⚠ Ekstrapolatsioon: x₀ on väljaspool vaatlusandmete vahemikku.")
                row += 1

    # Chart
    if reg_result and x_vals and y_vals:
        row += 2
        _create_scatter_with_trendline(
            sheet, row, x_vals, y_vals,
            x_label=num_cols[0][0] if num_cols else "x",
            y_label=num_cols[1][0] if len(num_cols) > 1 else "y",
            title="Hajuvusdiagramm koos regressioonisirgega",
        )

    return SHEET_REGRESSION


def _create_scatter_with_trendline(sheet, row_offset, x_vals, y_vals, x_label, y_label, title) -> None:
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
            width=420, height=300,
        )
        chart.chart_type = 'xy_scatter'
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
            tl.Type = 5  # xlLinear
            tl.DisplayEquation = True
            tl.DisplayRSquared = True
        except Exception:
            pass
    except Exception:
        pass


def write_kordamisküsimused(xw_book, questions: list[dict]) -> str:
    sheet = get_or_create_ai_sheet(xw_book, SHEET_KORDAMIS)
    _set_col_widths(sheet, {"A": 8, "B": 60, "C": 20, "D": 20, "E": 20})

    _h1(sheet, 1, "Kordamisküsimused")
    _h2(sheet, 2, "Vasta küsimustele. Vastused on eraldi lehel 'Vastused'.", n_cols=5)
    row = 3

    for i, q in enumerate(questions, 1):
        question_text = q.get("question", "")
        options = q.get("options", [])
        try:
            ac = sheet.range(f"A{row}")
            ac.value = f"K{i}"
            ac.font.bold = True
            bc = sheet.range(f"B{row}")
            bc.value = question_text
            bc.font.bold = True
            bc.color = _BG_H3
        except Exception:
            pass
        row += 1

        if options:
            for j, opt in enumerate(options[:4]):
                letter = chr(65 + j)
                try:
                    sheet.range(f"A{row}").value = f"  {letter})"
                    sheet.range(f"B{row}").value = opt
                except Exception:
                    pass
                row += 1

        row += 1  # blank between questions

    return SHEET_KORDAMIS


def write_vastused(xw_book, questions: list[dict]) -> str:
    sheet = get_or_create_ai_sheet(xw_book, SHEET_VASTUSED)
    _set_col_widths(sheet, {"A": 8, "B": 20, "C": 40, "D": 50})

    _h1(sheet, 1, "Vastused")
    for ci, label in enumerate(["K#", "Vastus", "Selgitus"], 1):
        try:
            hc = sheet.range((2, ci))
            hc.value = label
            hc.font.bold = True
            hc.color = _BG_GRAY
        except Exception:
            pass
    row = 3

    for i, q in enumerate(questions, 1):
        answer = q.get("answer", "")
        explanation = q.get("explanation", q.get("answer", ""))
        try:
            sheet.range(f"A{row}").value = f"K{i}"
            bc = sheet.range(f"B{row}")
            bc.value = answer
            bc.color = _BG_RESULT
            cc = sheet.range(f"C{row}")
            cc.value = explanation
            try:
                cc.api.WrapText = True
            except Exception:
                pass
        except Exception:
            pass
        row += 1

    return SHEET_VASTUSED


def write_all_sheets(xw_book, analysis: WorkbookAnalysis, results: list[TaskResult], inspection) -> dict:
    """Write all relevant sheets. Returns {"written": list, "skipped": list, "errors": list}."""
    written, skipped, errors = [], [], []

    has_regression = any(r.task_type == TaskType.REGRESSION for r in results if not r.error)
    has_ci = any(r.task_type == TaskType.CONFIDENCE_INTERVAL for r in results if not r.error)
    has_pred = any(r.task_type in (TaskType.PREDICTION, TaskType.PREDICTION_INTERVAL) for r in results if not r.error)
    needs_regression_sheet = has_regression or has_ci or has_pred

    task_plans = analysis.tasks

    try:
        write_lahendus(xw_book, task_plans, results, inspection)
        written.append(SHEET_LAHENDUS)
    except ValueError as e:
        skipped.append((SHEET_LAHENDUS, str(e)))
    except Exception as e:
        errors.append((SHEET_LAHENDUS, str(e)))

    try:
        write_arvutused(xw_book, task_plans, results, inspection)
        written.append(SHEET_ARVUTUSED)
    except ValueError as e:
        skipped.append((SHEET_ARVUTUSED, str(e)))
    except Exception as e:
        errors.append((SHEET_ARVUTUSED, str(e)))

    if needs_regression_sheet:
        try:
            write_regression(xw_book, task_plans, results, inspection)
            written.append(SHEET_REGRESSION)
        except ValueError as e:
            skipped.append((SHEET_REGRESSION, str(e)))
        except Exception as e:
            errors.append((SHEET_REGRESSION, str(e)))

    questions = getattr(analysis, "revision_questions", []) or []
    if questions:
        try:
            write_kordamisküsimused(xw_book, questions)
            written.append(SHEET_KORDAMIS)
        except ValueError as e:
            skipped.append((SHEET_KORDAMIS, str(e)))
        except Exception as e:
            errors.append((SHEET_KORDAMIS, str(e)))

        try:
            write_vastused(xw_book, questions)
            written.append(SHEET_VASTUSED)
        except ValueError as e:
            skipped.append((SHEET_VASTUSED, str(e)))
        except Exception as e:
            errors.append((SHEET_VASTUSED, str(e)))

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
            issues.append(f"Algne leht '{original_sheet}' on märgistatud AI-genereerituna — võimalik, et kirjutati üle.")

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
