# Persistent ML Historical Database

The machine-learning layer can now maintain a local point-in-time SQLite database at:

`ml_data/ml_history.sqlite`

The database is local and gitignored. API keys are read only from environment variables and are never written into SQLite, Excel, GitHub, logs, or source files.

## Alpha Vantage key on Windows

The existing project convention is `ALPHAVANTAGE_API_KEY`.

Persist the key for the current Windows user:

```powershell
[Environment]::SetEnvironmentVariable("ALPHAVANTAGE_API_KEY","PASTE_YOUR_KEY_HERE","User")
```

Load it into the current PowerShell session without reopening the terminal:

```powershell
$env:ALPHAVANTAGE_API_KEY=[Environment]::GetEnvironmentVariable("ALPHAVANTAGE_API_KEY","User")
```

Verify without printing the secret:

```powershell
if ($env:ALPHAVANTAGE_API_KEY) { "Alpha Vantage key loaded" } else { "Alpha Vantage key missing" }
```

Alpha Vantage currently documents a free usage ceiling of up to 25 requests per day for most datasets. The default enrichment budget is therefore 20 calls/run, leaving a small buffer for the normal equity-research consensus layer.

## Optional SEC contact header

For deep US historical fundamentals, configure a compliant SEC User-Agent with a real contact email:

```powershell
[Environment]::SetEnvironmentVariable("SEC_USER_AGENT","EquityResearch your-real-email@example.com","User")
$env:SEC_USER_AGENT=[Environment]::GetEnvironmentVariable("SEC_USER_AGENT","User")
```

Do not use the placeholder email literally.

## Step 1 — free/public historical database

After pulling the newest project version:

```powershell
cd "C:\Users\Antza\Documents\Antzaz-equity-research-model"
python -m pip install -r .\requirements.txt
python .\ml_history.py bootstrap --universe sp500 --limit 500 --years 20 --call-budget 20
```

The bootstrap is resumable. Re-running it does not intentionally create duplicate rows.

Source roles:

- Yahoo Finance: bulk historical daily prices and fallback annual statements.
- SEC Company Facts: deep annual US fundamentals when `SEC_USER_AGENT` is configured.
- Alpha Vantage: normalized annual statements plus quarterly earnings actual/estimate/surprise history, quota-aware.
- Federal Reserve FRED: official macro history.

The database stores:

- prices;
- annual fundamentals;
- earnings actuals and prevailing estimates where available;
- macro series;
- point-in-time ML features;
- valuation factors;
- consensus snapshots;
- ML predictions for later outcome grading.

Point-in-time annual fundamentals use a conservative availability lag when an exact filed/reported timestamp is unavailable. The feature builder never uses a forward target dated before the feature observation date.

### Useful incremental commands

Prices only:

```powershell
python .\ml_history.py backfill-prices --universe sp500 --limit 500 --years 20
```

Fundamentals only:

```powershell
python .\ml_history.py backfill-fundamentals --universe sp500 --limit 500
```

Spend today's Alpha Vantage budget on symbols not yet enriched:

```powershell
python .\ml_history.py enrich-alpha --universe sp500 --limit 500 --call-budget 20
```

Refresh official macro history:

```powershell
python .\ml_history.py backfill-macro
```

Rebuild ML features after adding data:

```powershell
python .\ml_history.py build-features --benchmark SPY
```

Inspect database coverage:

```powershell
python .\ml_history.py status
```

## How ML uses the database

Normal ML research remains:

```powershell
python .\research.py GOOGL --ml
```

When `ml_data/ml_history.sqlite` contains enough point-in-time observations, `ml_research.py` automatically prefers it for:

- Expected 12-month excess-return training;
- historical valuation factors;
- earnings-surprise training;
- portfolio return history when coverage is sufficient.

If the database is not ready, the old live-data construction remains the fallback.

The expected-return database adds valuation features to the existing quality/growth/momentum feature set:

- Price / Sales;
- earnings yield;
- FCF yield;
- book-to-market;
- EV / Sales.

Every ML run with a persistent database also writes its predictions into the local `predictions` table so later versions can grade matured forecasts.

## Step 2 — FMP analyst-data layer

The adapter is already implemented but remains inactive until `FMP_API_KEY` exists.

When you obtain a key, store it the same way:

```powershell
[Environment]::SetEnvironmentVariable("FMP_API_KEY","PASTE_YOUR_KEY_HERE","User")
$env:FMP_API_KEY=[Environment]::GetEnvironmentVariable("FMP_API_KEY","User")
```

Then run:

```powershell
python .\ml_history.py enrich-fmp --universe sp500 --limit 500 --call-budget 200
```

The FMP adapter stores:

- revenue consensus mean/low/high;
- EPS consensus mean/low/high;
- analyst counts where supplied;
- historical actual-vs-estimated earnings observations;
- timestamped provider snapshots.

A crucial limitation is retained deliberately: a provider's old fiscal-year consensus row is not automatically treated as a true historical point-in-time revision series. Genuine revision history requires dated snapshots. Repeated FMP/Alpha snapshots will build that history prospectively; only provider fields with valid as-of dates should be used for historical backtests.

## Survivorship-bias warning

`--universe sp500` currently starts from a current S&P 500 constituent snapshot. That is useful for building a large training set but is not a survivorship-free historical universe. Do not interpret a backtest on only today's survivors as institutional-grade evidence. Historical constituents/delistings should be added later from a provider that supports them reliably.
