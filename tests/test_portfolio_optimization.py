from __future__ import annotations

import numpy as np
import pandas as pd

from institutional_research.src.portfolio_optimization import (
    _confidence_scalar,
    efficient_frontier,
    expected_return_inputs,
    maximum_history_returns,
    optimize_suite,
    robust_covariance,
)


def _synthetic(seed=7, rows=1500):
    rng = np.random.default_rng(seed)
    dates = pd.bdate_range("2019-01-01", periods=rows)
    market = rng.normal(0.00035, 0.009, rows)
    data = {}
    for i, ticker in enumerate(["AAA", "BBB", "CCC", "DDD", "EEE"]):
        data[ticker] = (0.55 + i * 0.08) * market + rng.normal(0.00015 + i * 0.00003, 0.006 + i * 0.0006, rows)
    return pd.DataFrame(data, index=dates), pd.Series(market, index=dates, name="SPY")


def _holdings():
    return pd.DataFrame({
        "Ticker": ["AAA", "BBB", "CCC", "DDD", "EEE"],
        "Weight": [0.20] * 5,
        "Sector": ["Tech", "Tech", "Financials", "Health", "Industrials"],
    })


def test_robust_covariance_is_psd_and_annualized():
    returns, _ = _synthetic()
    cov, meta = robust_covariance(returns)
    assert cov.shape == (5, 5)
    assert meta["method"].startswith("20Y pairwise")
    eig = np.linalg.eigvalsh(cov.to_numpy())
    assert eig.min() >= -1e-8
    assert np.diag(cov).min() > 0


def test_maximum_history_fallback_preserves_assets_and_coverage():
    returns, _ = _synthetic(rows=800)
    panel, coverage = maximum_history_returns(
        list(returns.columns), returns, history_db=None, max_years=20
    )
    assert list(panel.columns) == list(returns.columns)
    assert len(coverage) == 5
    assert set(coverage["Source"]) == {"Public price fallback"}
    assert coverage["Observations"].min() >= 700


def test_expected_return_inputs_work_without_ml_database():
    returns, benchmark = _synthetic()
    out, meta = expected_return_inputs(
        returns, benchmark, pd.DataFrame(), list(returns.columns), history_db=None
    )
    assert len(out) == 5
    assert out["ClassicalExpectedReturn"].notna().all()
    assert out["BlendedMLExpectedReturn"].notna().all()
    assert np.allclose(out["ClassicalExpectedReturn"], out["BlendedMLExpectedReturn"])
    assert meta["ml_covered_assets"] == 0
    assert _confidence_scalar("High") > _confidence_scalar("Moderate") > _confidence_scalar("Low")


def test_optimizer_builds_requested_classical_and_ml_portfolios():
    returns, benchmark = _synthetic()
    cov, _ = robust_covariance(returns)
    inputs, _ = expected_return_inputs(
        returns, benchmark, pd.DataFrame(), list(returns.columns), history_db=None
    )
    summary, weights, meta = optimize_suite(
        holdings=_holdings(),
        expected_bounds=pd.DataFrame(),
        expected_inputs=inputs,
        classical_cov=cov,
        regime_cov=cov,
        risk_free_rate=0.03,
        max_position=0.40,
        max_sector=0.60,
        turnover_limit=0.80,
        risk_aversion=4.0,
    )
    names = set(summary["Portfolio"])
    required = {
        "Minimum Variance",
        "Maximum Sharpe — Classical",
        "Mean-Variance — Classical",
        "ML Maximum Sharpe",
        "ML Mean-Variance",
        "Regime-Aware ML Maximum Sharpe",
    }
    assert required.issubset(names)
    assert meta["status"] == "PASS"

    optimized = weights[weights["Portfolio"].isin(required)]
    sums = optimized.groupby("Portfolio")["TargetWeight"].sum()
    assert np.allclose(sums.to_numpy(), 1.0, atol=1e-5)
    assert optimized["TargetWeight"].min() >= -1e-8
    assert optimized["TargetWeight"].max() <= 0.40001

    current_vol = float(summary.loc[summary["Portfolio"] == "Current Portfolio", "ExpectedVolatility"].iloc[0])
    min_var_vol = float(summary.loc[summary["Portfolio"] == "Minimum Variance", "ExpectedVolatility"].iloc[0])
    assert min_var_vol <= current_vol + 1e-8


def test_efficient_frontier_returns_finite_ordered_points():
    returns, benchmark = _synthetic()
    cov, _ = robust_covariance(returns)
    inputs, _ = expected_return_inputs(
        returns, benchmark, pd.DataFrame(), list(returns.columns), history_db=None
    )
    mu = inputs.set_index("Ticker")["ClassicalExpectedReturn"]
    frontier = efficient_frontier(
        mu=mu,
        cov=cov,
        holdings=_holdings(),
        expected_bounds=pd.DataFrame(),
        max_position=0.40,
        max_sector=0.60,
        turnover_limit=None,
        risk_free_rate=0.03,
        points=25,
    )
    assert len(frontier) >= 10
    assert frontier["ExpectedReturn"].notna().all()
    assert frontier["ExpectedVolatility"].gt(0).all()
    assert frontier["TargetReturn"].is_monotonic_increasing
