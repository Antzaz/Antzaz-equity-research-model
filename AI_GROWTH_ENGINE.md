# AI Growth Forecast Engine

The AI Growth Forecast Engine adds an institutional-style bridge between AI evidence, machine-learning forecasts, and valuation expectations.

## Architecture

The engine deliberately separates three questions:

1. **What is the company actually saying / reporting about AI?**
   - The evidence layer reads the latest `research_data/<TICKER>/kpi_history.json` snapshot and optional files under `research_data/<TICKER>/ai_sources/`.
   - Without an API key it uses deterministic, auditable keyword/signal extraction.
   - With `--llm` it uses OpenAI Structured Outputs to classify the same evidence into normalized demand, monetization, adoption, efficiency, capex-burden, and risk scores.
   - The LLM is not allowed to infer undisclosed AI revenue.

2. **What does the fundamental history predict?**
   - LightGBM forecasts next-fiscal-year revenue growth and next-fiscal-year FCF growth.
   - Training observations come from the existing point-in-time SQLite ML history store.
   - Targets require a contiguous next fiscal year.
   - The final ~20% of observations are reserved as a chronological holdout and training targets are purged when they would not have been knowable before the holdout.
   - Elastic Net is the required simple benchmark. LightGBM confidence is downgraded when it fails to match the Elastic Net holdout MAE.
   - SHAP explains the current LightGBM forecast; native feature importance is the fallback if SHAP is unavailable.

3. **Is AI growth better or worse than what valuation requires?**
   - The AI evidence scores create a deliberately bounded adjustment to the independent LightGBM fundamental forecast.
   - The AI-adjusted FCF growth forecast is compared with the current reverse-DCF implied annual FCF growth hurdle.
   - The difference is the **AI Expectations Gap**.

## Why the AI overlay is bounded

The project currently has much more historical financial data than dated, standardized AI KPI history. A fully supervised “AI model” would therefore overstate the amount of evidence available.

Until enough dated AI observations have accumulated, the architecture is:

`LightGBM fundamental forecast + bounded AI evidence adjustment`

rather than:

`Sparse AI mentions -> unconstrained neural-network forecast`

Once AI KPI history becomes deep enough across companies and quarters, the normalized AI scores can be promoted into the supervised feature set and tested out of sample.

## Outputs

Each run writes:

- `ml_runs/<TICKER>/<timestamp>/ai_growth_results.json`
- a separate `AI Growth Forecast` sheet in the research workbook unless `--no-workbook-write` is used.

The sheet shows AI signal scores, LightGBM next-FY revenue/FCF growth, AI-adjusted growth, reverse-DCF implied FCF growth, the AI Expectations Gap, LightGBM vs Elastic Net holdout MAE, SHAP/feature drivers, evidence, and governance notes.

It does **not** overwrite the authoritative DCF assumptions.

## Commands

Deterministic AI extraction + LightGBM + reverse DCF:

```powershell
python .\research.py GOOGL --ml
```

Full agent research plus LLM AI extraction plus LightGBM:

```powershell
$env:OPENAI_API_KEY="YOUR_KEY"
python .\research.py GOOGL --ml --ai
```

Run only the AI Growth engine:

```powershell
python .\ai_growth_research.py GOOGL
```

Run only the AI Growth engine with structured LLM extraction:

```powershell
$env:OPENAI_API_KEY="YOUR_KEY"
python .\ai_growth_research.py GOOGL --llm
```

The default extraction model is `gpt-5.6-luna` because extraction is a structured, high-volume task. Use `--model` to override it.

## Data maturity and interpretation

A forecast is a research diagnostic, not a price target. The most important validation fields are the chronological holdout MAE and the comparison with Elastic Net. If history is insufficient, the engine reports `INSUFFICIENT_DATA` instead of manufacturing a forecast.

The AI Expectations Gap is:

`AI-adjusted FCF growth forecast - reverse-DCF implied FCF growth`

A positive gap does not automatically mean “buy”; it means the model forecast is above the simplified market-implied FCF hurdle. It still needs normal thesis, valuation, quality, balance-sheet, risk, and portfolio-construction review.
