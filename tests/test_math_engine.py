"""Tests for the statistics engine against known values from tooleht5.xls."""

import math

import pytest

from src import math_engine as me
from tests.fixtures import (
    FIREARMS,
    HOMICIDES,
    REF_INTERCEPT,
    REF_PEARSON_R,
    REF_PRED_AT_12_5,
    REF_R_SQUARED,
    REF_SLOPE,
)

TOL = 1e-6


def test_pearson_correlation():
    r, p = me.pearson_correlation(FIREARMS, HOMICIDES)
    assert math.isclose(r, REF_PEARSON_R, abs_tol=TOL)
    assert 0.0 <= p <= 1.0


def test_linear_regression_coefficients():
    reg = me.linear_regression(FIREARMS, HOMICIDES)
    assert math.isclose(reg["intercept"], REF_INTERCEPT, abs_tol=TOL)
    assert math.isclose(reg["slope"], REF_SLOPE, abs_tol=TOL)
    assert math.isclose(reg["r_squared"], REF_R_SQUARED, abs_tol=TOL)


def test_r_squared_equals_r_squared():
    r, _ = me.pearson_correlation(FIREARMS, HOMICIDES)
    reg = me.linear_regression(FIREARMS, HOMICIDES)
    assert math.isclose(r * r, reg["r_squared"], abs_tol=TOL)


def test_predict_y_at_12_5():
    y = me.predict_y(REF_INTERCEPT, REF_SLOPE, 12.5)
    assert math.isclose(y, REF_PRED_AT_12_5, abs_tol=1e-5)


def test_predict_y_from_regression():
    reg = me.linear_regression(FIREARMS, HOMICIDES)
    y = me.predict_y(reg["intercept"], reg["slope"], 12.5)
    assert math.isclose(y, REF_PRED_AT_12_5, abs_tol=TOL)


def test_spearman_correlation_range():
    rho, p = me.spearman_correlation(FIREARMS, HOMICIDES)
    assert -1.0 <= rho <= 1.0
    assert 0.0 <= p <= 1.0


def test_regression_confidence_intervals_bracket_estimate():
    ci = me.regression_confidence_intervals(FIREARMS, HOMICIDES)
    reg = me.linear_regression(FIREARMS, HOMICIDES)
    lo, hi = ci["slope"]
    assert lo < reg["slope"] < hi
    lo_i, hi_i = ci["intercept"]
    assert lo_i < reg["intercept"] < hi_i
    assert ci["df"] == len(FIREARMS) - 2


def test_prediction_interval_contains_point_and_flags_extrapolation():
    lo, hi, extrap = me.prediction_interval(FIREARMS, HOMICIDES, 12.5)
    assert lo < REF_PRED_AT_12_5 < hi
    # 12.5 is beyond max firearms (11.6) -> extrapolation.
    assert extrap is True


def test_prediction_interval_within_range_not_extrapolation():
    _, _, extrap = me.prediction_interval(FIREARMS, HOMICIDES, 5.0)
    assert extrap is False


def test_confidence_interval_mean_narrower_than_prediction():
    ci_lo, ci_hi = me.confidence_interval_mean(FIREARMS, HOMICIDES, 5.0)
    pi_lo, pi_hi, _ = me.prediction_interval(FIREARMS, HOMICIDES, 5.0)
    assert (ci_hi - ci_lo) < (pi_hi - pi_lo)


def test_descriptive_stats():
    stats = me.descriptive_stats(FIREARMS)
    assert stats["n"] == 8
    assert math.isclose(stats["min"], 0.6, abs_tol=TOL)
    assert math.isclose(stats["max"], 11.6, abs_tol=TOL)
    assert math.isclose(stats["mean"], sum(FIREARMS) / len(FIREARMS), abs_tol=TOL)


def test_t_test_one_sample():
    res = me.t_test_one_sample(HOMICIDES, mu0=5.0)
    assert "t_statistic" in res and "p_value" in res
    assert res["df"] == len(HOMICIDES) - 1
    assert isinstance(res["reject_null"], bool)


def test_t_test_two_sample():
    res = me.t_test_two_sample(FIREARMS, HOMICIDES)
    assert "t_statistic" in res
    assert math.isclose(res["mean1"], sum(FIREARMS) / len(FIREARMS), abs_tol=TOL)


def test_chi_square_uniform_default():
    res = me.chi_square_test([10, 12, 8, 10])
    assert res["chi_square"] >= 0.0
    assert res["df"] == 3


def test_chi_square_uses_alpha():
    """chi_square_test must use the provided alpha, not a hardcoded 0.05.

    The reject_null flag must equal (p_value < alpha) for every alpha value.
    """
    for obs, alpha in [
        ([10, 10, 10, 10], 0.05),    # uniform: large p -> should not reject
        ([10, 10, 10, 10], 0.99),    # same data, huge alpha -> should reject
        ([1, 50, 1, 1], 0.05),       # skewed: very small p -> should reject
        ([1, 50, 1, 1], 1e-8),       # same skewed data, tiny alpha
    ]:
        res = me.chi_square_test(obs, alpha=alpha)
        assert math.isclose(res["alpha"], alpha), f"alpha not stored for obs={obs}"
        expected_reject = res["p_value"] < alpha
        assert res["reject_null"] is expected_reject, (
            f"reject_null={res['reject_null']} but p={res['p_value']:.6f}, alpha={alpha}"
        )


def test_spearman_correlation_values():
    """spearman_correlation returns rho in [-1, 1] and a valid p-value."""
    rho, p = me.spearman_correlation(FIREARMS, HOMICIDES)
    assert -1.0 <= rho <= 1.0
    assert 0.0 <= p <= 1.0
    # Firearms vs homicides should be positively correlated (Spearman too).
    assert rho > 0.5


def test_confidence_interval_mean_narrower_than_pi():
    """Mean-response CI must always be contained within the prediction interval."""
    for x_new in [3.0, 5.0, 11.6]:
        ci_lo, ci_hi = me.confidence_interval_mean(FIREARMS, HOMICIDES, x_new)
        pi_lo, pi_hi, _ = me.prediction_interval(FIREARMS, HOMICIDES, x_new)
        assert ci_lo > pi_lo, f"CI lower not > PI lower at x={x_new}"
        assert ci_hi < pi_hi, f"CI upper not < PI upper at x={x_new}"


def test_prediction_interval_alpha_respected():
    """Wider alpha (looser) produces wider PI than default 95%."""
    lo_95, hi_95, _ = me.prediction_interval(FIREARMS, HOMICIDES, 5.0, alpha=0.05)
    lo_90, hi_90, _ = me.prediction_interval(FIREARMS, HOMICIDES, 5.0, alpha=0.10)
    # 95% interval must be wider than 90%.
    assert hi_95 - lo_95 > hi_90 - lo_90


def test_mismatched_lengths_raise():
    with pytest.raises(ValueError):
        me.pearson_correlation([1, 2, 3], [1, 2])


def test_too_few_points_raise():
    with pytest.raises(ValueError):
        me.linear_regression([1.0], [2.0])
