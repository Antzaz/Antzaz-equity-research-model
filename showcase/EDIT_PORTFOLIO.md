# Edit the recruiter portfolio

The recruiter-facing app reads **only** `showcase/data/recruiter_portfolio.json` for the company names, theses and public values you choose to show.

## Fastest workflow

1. Open `showcase/data/recruiter_portfolio.json`.
2. Fill in the portfolio description and metrics you want public.
3. Add one object under `companies` for each company.
4. Set `"published": true` only when that company is ready for recruiters.
5. Commit the file. A deployed Streamlit app will update from the repository automatically.

Use `recruiter_portfolio.example.json` as a copy/paste template.

## Fields you can edit per company

- `company`, `ticker`, `sector`, `weight`
- `rating` — e.g. BUY / HOLD / SELL or your own wording
- `current_price`, `fair_value`, `upside`, `currency`
- `conviction`, `investment_horizon`, `last_reviewed`
- `one_line_thesis`
- `thesis` — main investment-thesis bullets
- `catalysts`
- `risks`
- `falsification_conditions`
- `scorecard`
- `valuation_scenarios`
- `key_kpis`
- `institutional_comparison`
- `notes`
- `sources`
- `full_report_url`

If `upside` is `null`, the app calculates it from current price and fair value.

## Privacy switch

`"published": false` keeps a company out of the recruiter view. This lets you prepare research cards in the same file and publish them only when ready.

Do not place private brokerage details, account values, credentials, API keys, transaction history or cost basis in this public data file.
