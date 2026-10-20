"""Statistical computations backed by NumPy/SciPy.

Every public function is pure: it takes plain sequences / numbers and returns
plain Python values (floats, dicts). No Excel or AI dependencies here, so this
module is trivially unit-testable against known reference values.

Reference values (verified against tooleht5.xls, DO NOT hardcode):
    Pearson r  ~= 0.88498417
    Intercept  ~= 4.04725078
    Slope      ~= 0.85251932
    R^2        ~= 0.78319698
    y(x=12.5)  ~= 14.70374226
"""

from __future__ import annotations

from typing import Sequence

import numpy as np
from scipy import stats


ArrayLike = Sequence[float]


def _as_arrays(x: ArrayLike, y: ArrayLike) -> tuple[np.ndarray, np.ndarray]:
    xa = np.asarray(x, dtype=float)
    ya = np.asarray(y, dtype=float)
    if xa.shape != ya.shape:
        raise ValueError(f"x and y must have equal length, got {xa.shape} and {ya.shape}")
    if xa.size < 2:
        raise ValueError("need at least 2 data points")
    return xa, ya


def pearson_correlation(x: ArrayLike, y: ArrayLike) -> tuple[float, float]:
    """Return (r, p_value) for Pearson's correlation coefficient."""
    xa, ya = _as_arrays(x, y)
    r, p = stats.pearsonr(xa, ya)
    return float(r), float(p)


def spearman_correlation(x: ArrayLike, y: ArrayLike) -> tuple[float, float]:
    """Return (rho, p_value) for Spearman's rank correlation."""
    xa, ya = _as_arrays(x, y)
    rho, p = stats.spearmanr(xa, ya)
    return float(rho), float(p)


def linear_regression(x: ArrayLike, y: ArrayLike) -> dict[str, float]:
    """Ordinary least squares y = intercept + slope * x.

    Returns a dict with intercept, slope, r_squared, and the standard errors of
    both coefficients.
    """
    xa, ya = _as_arrays(x, y)
    res = stats.linregress(xa, ya)
    n = xa.size
    # scipy gives stderr of slope and intercept_stderr directly.
    return {
        "intercept": float(res.intercept),
        "slope": float(res.slope),
        "r_value": float(res.rvalue),
        "r_squared": float(res.rvalue**2),
        "p_value": float(res.pvalue),
        "slope_std_error": float(res.stderr),
        "intercept_std_error": float(res.intercept_stderr),
        "n": int(n),
    }


def regression_confidence_intervals(
    x: ArrayLike, y: ArrayLike, alpha: float = 0.05
) -> dict[str, tuple[float, float]]:
    """Confidence intervals for the regression intercept and slope.

    Uses the t distribution with n-2 degrees of freedom.
    """
    xa, ya = _as_arrays(x, y)
    reg = linear_regression(xa, ya)
    n = reg["n"]
    df = n - 2
    t_crit = float(stats.t.ppf(1 - alpha / 2, df))

    slope_lo = reg["slope"] - t_crit * reg["slope_std_error"]
    slope_hi = reg["slope"] + t_crit * reg["slope_std_error"]
    int_lo = reg["intercept"] - t_crit * reg["intercept_std_error"]
    int_hi = reg["intercept"] + t_crit * reg["intercept_std_error"]

    return {
        "alpha": alpha,
        "t_critical": t_crit,
        "df": df,
        "slope": (slope_lo, slope_hi),
        "intercept": (int_lo, int_hi),
    }


def predict_y(intercept: float, slope: float, x_new: float) -> float:
    """Point prediction y = intercept + slope * x_new."""
    return float(intercept + slope * x_new)


def _regression_prediction_pieces(
    xa: np.ndarray, ya: np.ndarray
) -> tuple[float, float, float, float, int]:
    """Return (intercept, slope, s_err, mean_x, n) shared by interval helpers."""
    res = stats.linregress(xa, ya)
    n = xa.size
    y_hat = res.intercept + res.slope * xa
    residuals = ya - y_hat
    # Standard error of the estimate (residual standard deviation).
    s_err = float(np.sqrt(np.sum(residuals**2) / (n - 2)))
    return float(res.intercept), float(res.slope), s_err, float(np.mean(xa)), n


def prediction_interval(
    x: ArrayLike, y: ArrayLike, x_new: float, alpha: float = 0.05
) -> tuple[float, float, bool]:
    """Prediction interval for a single new observation at x_new.

    Returns (lower, upper, is_extrapolation). is_extrapolation is True when
    x_new falls outside the observed range of x.
    """
    xa, ya = _as_arrays(x, y)
    intercept, slope, s_err, mean_x, n = _regression_prediction_pieces(xa, ya)
    df = n - 2
    t_crit = float(stats.t.ppf(1 - alpha / 2, df))
    ssx = float(np.sum((xa - mean_x) ** 2))

    y_pred = intercept + slope * x_new
    # Prediction interval includes the +1 term for the new observation's own variance.
    margin = t_crit * s_err * np.sqrt(1 + 1 / n + (x_new - mean_x) ** 2 / ssx)
    is_extrap = bool(x_new < np.min(xa) or x_new > np.max(xa))
    return float(y_pred - margin), float(y_pred + margin), is_extrap


def confidence_interval_mean(
    x: ArrayLike, y: ArrayLike, x_new: float, alpha: float = 0.05
) -> tuple[float, float]:
    """Confidence interval for the mean response E[y|x_new] at x_new."""
    xa, ya = _as_arrays(x, y)
    intercept, slope, s_err, mean_x, n = _regression_prediction_pieces(xa, ya)
    df = n - 2
    t_crit = float(stats.t.ppf(1 - alpha / 2, df))
    ssx = float(np.sum((xa - mean_x) ** 2))

    y_pred = intercept + slope * x_new
    # Mean-response interval omits the +1 term.
    margin = t_crit * s_err * np.sqrt(1 / n + (x_new - mean_x) ** 2 / ssx)
    return float(y_pred - margin), float(y_pred + margin)


def descriptive_stats(data: ArrayLike) -> dict[str, float]:
    """Common descriptive statistics for a 1-D sample."""
    arr = np.asarray(data, dtype=float)
    if arr.size == 0:
        raise ValueError("data is empty")
    return {
        "n": int(arr.size),
        "mean": float(np.mean(arr)),
        "median": float(np.median(arr)),
        "std": float(np.std(arr, ddof=1)) if arr.size > 1 else 0.0,
        "variance": float(np.var(arr, ddof=1)) if arr.size > 1 else 0.0,
        "min": float(np.min(arr)),
        "max": float(np.max(arr)),
        "range": float(np.max(arr) - np.min(arr)),
        "sum": float(np.sum(arr)),
        "q1": float(np.percentile(arr, 25)),
        "q3": float(np.percentile(arr, 75)),
    }


def t_test_one_sample(
    data: ArrayLike, mu0: float, alpha: float = 0.05, two_tailed: bool = True
) -> dict[str, float | bool]:
    """One-sample t-test of H0: mean == mu0."""
    arr = np.asarray(data, dtype=float)
    alternative = "two-sided" if two_tailed else "greater"
    t_stat, p_value = stats.ttest_1samp(arr, mu0, alternative=alternative)
    df = arr.size - 1
    return {
        "t_statistic": float(t_stat),
        "p_value": float(p_value),
        "df": int(df),
        "mean": float(np.mean(arr)),
        "mu0": float(mu0),
        "alpha": alpha,
        "reject_null": bool(p_value < alpha),
    }


def t_test_two_sample(
    x1: ArrayLike, x2: ArrayLike, equal_var: bool = False, alpha: float = 0.05
) -> dict[str, float | bool]:
    """Two-sample t-test (Welch's by default) of H0: mean1 == mean2."""
    a = np.asarray(x1, dtype=float)
    b = np.asarray(x2, dtype=float)
    t_stat, p_value = stats.ttest_ind(a, b, equal_var=equal_var)
    return {
        "t_statistic": float(t_stat),
        "p_value": float(p_value),
        "mean1": float(np.mean(a)),
        "mean2": float(np.mean(b)),
        "equal_var": bool(equal_var),
        "alpha": alpha,
        "reject_null": bool(p_value < alpha),
    }


def chi_square_test(
    observed: ArrayLike, expected: ArrayLike | None = None, alpha: float = 0.05
) -> dict[str, float | bool]:
    """Chi-square goodness-of-fit test.

    If expected is None, a uniform distribution is assumed.
    """
    obs = np.asarray(observed, dtype=float)
    if expected is None:
        exp = np.full_like(obs, obs.sum() / obs.size)
    else:
        exp = np.asarray(expected, dtype=float)
    chi2, p_value = stats.chisquare(obs, exp)
    df = obs.size - 1
    return {
        "chi_square": float(chi2),
        "p_value": float(p_value),
        "df": int(df),
        "alpha": alpha,
        "reject_null": bool(p_value < alpha),
    }
