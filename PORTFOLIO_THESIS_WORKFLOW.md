# Portfolio Thesis Workflow

The recruiter-facing Streamlit app can show real company names, portfolio weights and a documented investment thesis for each published holding.

## Private source files

Keep these only on your machine / private workflow:

- `institutional_research/portfolio.csv`
- `institutional_research/portfolio_thesis.xlsx`

Both are gitignored.

## Workbook structure

### Portfolio Philosophy

Use this sheet for portfolio-level reasoning such as:

- investment philosophy
- portfolio objective
- benchmark
- time horizon
- research process
- selection criteria
- position sizing
- diversification
- risk management
- sell discipline
- monitoring and review
- portfolio edge
- what you avoid

Set `Publish = Yes` to include the completed section in the recruiter dashboard.

### Company Theses

One row per current portfolio company. The workbook supports:

- Publish control
- Ticker (private join key)
- Company
- Status
- Time horizon
- Conviction
- Business Quality score
- Moat score
- Management / Capital Allocation score
- Balance Sheet score
- Growth score
- Valuation score
- Risk / Resilience score
- Composite Score
- Expected Annual Return
- Investment Thesis
- Why I Own It
- Competitive Advantage / Moat
- Growth Drivers
- Valuation Rationale
- Catalysts
- Key Risks
- Falsification / Sell Condition
- Monitoring KPI
- Review Date
- Public Notes
- Private Notes (never published)

Rows marked `Publish = Yes` are eligible for publication, but completely empty thesis rows are ignored.

## Save the workbook

Save the completed template as:

```text
C:\Users\Antza\Documents\Antzaz-equity-research-model\institutional_research\portfolio_thesis.xlsx
```

## Sync portfolio + thesis + refresh online analytics

From the repository root:

```powershell
cd "C:\Users\Antza\Documents\Antzaz-equity-research-model"
git pull
powershell -ExecutionPolicy Bypass -File .\automation\sync_portfolio_secret.ps1
```

The command:

1. validates local `portfolio.csv`;
2. updates the private portfolio GitHub Actions secret;
3. detects `portfolio_thesis.xlsx`;
4. extracts only explicitly recruiter-safe thesis fields;
5. excludes `Private Notes (never published)`;
6. stores the compact public thesis payload in a private GitHub Actions secret;
7. triggers the portfolio refresh workflow.

## Public snapshot

The GitHub workflow combines the latest production portfolio analytics with the recruiter-safe thesis payload.

Public snapshot fields can include:

- company name
- sector
- portfolio weight
- risk contribution
- portfolio metrics
- portfolio philosophy
- company investment thesis
- research scorecard
- conviction
- expected return
- risks
- sell/falsification condition
- monitoring KPI

The public snapshot excludes:

- ticker symbols
- shares
- average cost / cost basis
- exact portfolio value
- unrealized P&L
- transactions
- private notes
- credentials
- private Excel workbooks

## Streamlit tabs

The recruiter app includes:

- `Portfolio Analytics`
- `Investment Thesis`
- `Equity Research`
- `Methodology`

The Investment Thesis tab provides a portfolio-philosophy overview plus a company selector for detailed security-level reasoning and scorecards.
