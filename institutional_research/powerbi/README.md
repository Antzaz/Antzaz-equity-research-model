# Power BI setup

The research engine writes Power BI-friendly flat files into:

`institutional_research/outputs/latest/`

## Recommended method: Folder / CSV connector

In Power BI Desktop:

1. Run `python run_research.py`.
2. Choose **Get data**.
3. Use **Text/CSV** for an individual table or **Folder** for the whole `outputs/latest` directory.
4. Load the tables you want.
5. Refresh after rerunning the Python research engine.

Recommended tables:

- `holdings_analysis.csv`
- `portfolio_timeseries.csv`
- `risk_contribution.csv`
- `correlation_matrix.csv`
- `factor_scores.csv`
- `factor_exposure.csv`
- `sector_exposure.csv`
- `stress_tests.csv`
- `monte_carlo_distribution.csv`
- `reverse_dcf.csv`
- `forecast_accuracy.csv`

The CSV approach deliberately keeps Python computation separate from Power BI presentation.

## Direct Python scripts in Power BI

Power BI Desktop can also run Python scripts, but this project does not require that integration.
The exported-CSV architecture is easier to debug and keeps the same analytical outputs usable by
Power BI, Streamlit, Excel, Tableau, or another visualization layer.

## Suggested report pages

1. Executive dashboard
2. Holdings and P&L
3. Risk and drawdown
4. Correlation and concentration
5. Factor exposures
6. Stress scenarios
7. Monte Carlo distribution
8. Reverse DCF / market-implied expectations
9. Forecast accuracy
