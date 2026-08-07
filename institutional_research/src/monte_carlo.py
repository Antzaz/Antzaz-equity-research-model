from __future__ import annotations

import numpy as np
import pandas as pd


def bootstrap_portfolio(
    asset_returns: pd.DataFrame,
    weights: pd.Series,
    simulations: int = 5000,
    horizon_days: int = 252,
    seed: int = 42,
):
    r = asset_returns.dropna(how="any")
    if len(r) < 30:
        raise ValueError("Not enough complete historical return observations for Monte Carlo.")

    w = weights.reindex(r.columns).astype(float)
    w = w / w.sum()
    portfolio_daily = r.mul(w, axis=1).sum(axis=1).to_numpy()

    rng = np.random.default_rng(seed)
    sampled = rng.choice(portfolio_daily, size=(simulations, horizon_days), replace=True)
    terminal = np.prod(1 + sampled, axis=1) - 1

    summary = {
        "simulations": int(simulations),
        "horizon_days": int(horizon_days),
        "mean_return": float(np.mean(terminal)),
        "median_return": float(np.median(terminal)),
        "p05_return": float(np.quantile(terminal, 0.05)),
        "p10_return": float(np.quantile(terminal, 0.10)),
        "p25_return": float(np.quantile(terminal, 0.25)),
        "p75_return": float(np.quantile(terminal, 0.75)),
        "p90_return": float(np.quantile(terminal, 0.90)),
        "p95_return": float(np.quantile(terminal, 0.95)),
        "probability_loss": float(np.mean(terminal < 0)),
        "probability_loss_20pct": float(np.mean(terminal < -0.20)),
        "probability_gain_20pct": float(np.mean(terminal > 0.20)),
    }
    distribution = pd.DataFrame({
        "Simulation": np.arange(1, simulations + 1),
        "TerminalReturn": terminal,
    })
    return summary, distribution
