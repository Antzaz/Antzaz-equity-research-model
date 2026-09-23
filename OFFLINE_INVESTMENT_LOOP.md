# Offline Investment Loop

This is the private/offline layer only. It is designed to connect company research, portfolio sizing, transactions and later outcomes without exposing sensitive position economics to the public showcase.

## 1. Company research accountability

The guarded company workbook already creates:

- Forecast Accountability — point-in-time base forecasts and matured forecast error;
- Earnings & Revisions — estimate revisions and public earnings-surprise history when available;
- Capital Allocation — FCF, capex, buybacks less SBC, dividends, M&A, ROIC, incremental ROIC and ROIC-WACC spread where supported;
- Valuation History — historical valuation context using available price/fundamental history;
- SOTP Framework — evidence-gated segment valuation framework that stays REVIEW until explicit segment economics/multiples are supportable;
- Thesis Timeline — durable catalyst/thesis-change monitoring.

Durable local histories live under research_data/<TICKER>/ and are ignored from public export.

## 2. Realized portfolio truth layer

Create the private transaction ledger:

```powershell
cd institutional_research
Copy-Item transaction_ledger_template.csv transaction_ledger.csv
```

Use DEPOSIT/WITHDRAW rows for external cash flows and BUY/SELL rows for trades. The accounting engine reconstructs point-in-time shares, cash, NAV and weights using raw closes plus corporate actions. It calculates TWR, annualized TWR, MWR/XIRR, turnover and implementation shortfall when reference prices are supplied.

## 3. Realized attribution

With a complete transaction ledger, the project calculates security-level realized contribution using point-in-time weights.

For sector allocation/selection/interaction attribution, also create:

```powershell
Copy-Item benchmark_sector_history_template.csv benchmark_sector_history.csv
```

Current benchmark weights are never backfilled as historical truth.

## 4. Research-to-portfolio bridge

The institutional portfolio layer discovers the latest private workbook for every current holding and derives a transparent expected-return bridge from:

- base fair-value convergence;
- dividend yield;
- bounded net share-reduction / buyback yield;
- workbook data-quality confidence;
- overall investment score;
- realized forecast-accuracy history.

Expected return is shrunk toward the configured benchmark expected return when confidence is weak. Manual expected_returns.csv remains authoritative where explicitly supplied.

## 5. Position sizing and thesis budget

The Research Decision Loop exports:

- CurrentWeight;
- RiskContribution;
- ExpectedAlpha;
- ResearchConfidence;
- DownsideReference;
- SuggestedMin / SuggestedMidpoint / SuggestedMax;
- RangeStatus.

It also compares capital weight, covariance-based risk weight and expected-alpha budget.

## 6. Fundamental portfolio scenarios

Company-model bear/base/bull or P10/P90 values are aggregated into portfolio-level bottom-up scenarios.

For explicit analyst scenarios:

```powershell
Copy-Item fundamental_scenarios_template.csv fundamental_scenarios.csv
```

Missing security shocks are never inferred.

## 7. Decision-journal learning

Before material trades, create:

```powershell
Copy-Item portfolio_decision_journal_template.csv portfolio_decision_journal.csv
```

The project matures entries at 3M, 6M and 12M and measures decision alpha, correctness, expected-return error, conviction buckets, sectors and thesis categories.

## 8. Transaction-cost-aware rebalancing

Optimizer targets are passed through an offline rebalance gate using:

- configured rebalance threshold;
- estimated spread;
- square-root ADV impact diagnostic;
- expected-alpha benefit;
- minimum benefit/cost ratio.

The output is a research diagnostic, not an execution instruction.

## 9. Running the full private loop

```powershell
cd institutional_research
python run_research.py
python -m streamlit run app.py
```

Use the Realized Portfolio and Research Decision Loop pages for the new outputs.

## Privacy boundary

The public showcase validator blocks transaction, cost-basis, share-count, decision-journal and other private accounting fields. The public showcase repository is not modified by this offline investment loop.
