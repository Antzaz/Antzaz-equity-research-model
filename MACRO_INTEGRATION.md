# Three-system macro research bridge

The repositories are connected without sharing private holdings:

1. **Private Global Macro Lab** — builds point-in-time FRED/ALFRED regimes and
   exports a strict, latest-completed-month JSON context. Its own scheduled
   publisher can write only `data/macro_context.json` into the showcase.
2. **Public investment showcase** — displays a dated, observational macro card.
   This JSON is independent of, and must never modify, the curated 15-holding
   AMD portfolio or its analytics.
3. **Main equity / institutional portfolio research** — explicitly downloads
   that exact sanitized public JSON into the *gitignored* local
   `institutional_research/.offline_cache/macro_context.json` cache.

Online synchronization (not part of offline startup):

```powershell
python automation/sync_macro_context.py
python institutional_research/run_research.py
python research.py GOOGL --no-ml
```

The portfolio's `outputs/latest/summary.json` includes an additive,
read-only `macro_context` object. Each research agent manifest includes it
as contextual evidence; the company DCF, WACC, price forecasts, historical
backtests, securities, weights, alpha and position economics are unchanged.
If the bridge is missing/stale, the context is UNAVAILABLE and the existing
portfolio / research runs continue. No current macro regime is invented.

The data contract is **exactly** 11 allowlisted fields. It rejects holdings,
ticker symbols, portfolio costs, future data, outdated months and unknown
economic regimes. This is a descriptive signal, not a point-in-time ML training
feature, because retroactive regime labels must not leak into historical models.

Live three-repository updates require the private Macro Lab's
`FRED_API_KEY` and limited `SHOWCASE_REPO_TOKEN` secrets, a passing macro
publisher run, and successful sync on the equity research machine or hosted
workflow. No tokens or private data are stored in the JSON.
