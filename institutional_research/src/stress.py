from __future__ import annotations

import pandas as pd


def beta_stress_test(
    holdings: pd.DataFrame,
    asset_risk: pd.DataFrame,
    scenarios: list[dict],
) -> pd.DataFrame:
    beta_map = asset_risk.set_index("Ticker")["Beta"].to_dict()
    rows = []
    for scenario in scenarios:
        name = scenario["name"]
        market_shock = float(scenario.get("benchmark_shock", 0.0))
        idio = float(scenario.get("idiosyncratic_shock", 0.0))
        portfolio_impact = 0.0

        for _, row in holdings.iterrows():
            ticker = row["Ticker"]
            beta = beta_map.get(ticker)
            if pd.isna(beta):
                beta = 1.0
            shocked_return = beta * market_shock + idio
            contribution = row["Weight"] * shocked_return
            portfolio_impact += contribution
            rows.append({
                "Scenario": name,
                "Ticker": ticker,
                "Weight": row["Weight"],
                "Beta": beta,
                "BenchmarkShock": market_shock,
                "IdiosyncraticShock": idio,
                "EstimatedHoldingReturn": shocked_return,
                "PortfolioContribution": contribution,
            })

        rows.append({
            "Scenario": name,
            "Ticker": "PORTFOLIO",
            "Weight": 1.0,
            "Beta": None,
            "BenchmarkShock": market_shock,
            "IdiosyncraticShock": idio,
            "EstimatedHoldingReturn": portfolio_impact,
            "PortfolioContribution": portfolio_impact,
        })

    return pd.DataFrame(rows)
