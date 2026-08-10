# Live private hosting — portfolio-driven daily refresh

The live architecture deliberately keeps personal portfolio data out of the public repository.

## What refreshes automatically

`.github/workflows/daily-portfolio-refresh.yml` runs every day at **08:17 Europe/Helsinki** and can also be started manually from GitHub Actions.

Each run:

1. reconstructs `institutional_research/portfolio.csv` from an encrypted GitHub Actions secret;
2. derives the active ticker universe directly from the portfolio;
3. refreshes the institutional portfolio analysis;
4. runs `update_model.py` only for those active portfolio companies;
5. packages portfolio outputs and the latest company workbooks into one AES-256 encrypted ZIP with generic filenames;
6. uploads only that encrypted ZIP as the `private-live-data` Actions artifact.

Ticker values are masked in Actions before model refreshes begin. Portfolio data is never committed to Git.

## Required GitHub Actions secrets

Repository **Settings → Secrets and variables → Actions → New repository secret**.

### `PORTFOLIO_CSV_B64`

Create it from your local private portfolio file in PowerShell:

```powershell
$bytes = [System.IO.File]::ReadAllBytes("C:\Users\Antza\Documents\Antzaz-equity-research-model\institutional_research\portfolio.csv")
[Convert]::ToBase64String($bytes) | Set-Clipboard
```

Paste the clipboard contents as the secret value.

Whenever you change the portfolio, update this secret. The next daily run will automatically research the new active holdings and stop refreshing removed holdings.

### `LIVE_BUNDLE_PASSWORD`

Use a long random password. It encrypts the private daily bundle with AES-256. Do not put this password in the repository.

Example local password generation:

```powershell
-join ((48..57)+(65..90)+(97..122) | Get-Random -Count 40 | ForEach-Object {[char]$_})
```

### `SEC_USER_AGENT` (recommended)

Use a descriptive SEC User-Agent containing a real contact email, for example:

```text
Antzaz Equity Research your-email@example.com
```

## First test

After the secrets exist:

1. Open the repository on GitHub.
2. Open **Actions**.
3. Select **Daily private portfolio refresh**.
4. Choose **Run workflow**.
5. Confirm the run succeeds and produces one artifact named `private-live-data`.

The artifact contains only an encrypted ZIP. Ticker-to-file mapping is stored inside the encrypted manifest.

## Hosted private research portal

Deploy **one private Streamlit app** using this entrypoint:

```text
institutional_research/live_dashboard.py
```

The portal explicitly exposes three views from the same encrypted daily bundle:

- **Portfolio Dashboard** — risk, construction, factors, attribution, liquidity, Monte Carlo, reverse DCF, forecasts, etc.
- **Alpha Analysis** — CAPM/Jensen alpha, FF3, Carhart 4, FF5, style-proxy alpha, rolling alpha, t-statistics and factor decomposition.
- **Company Research** — only companies in the private portfolio universe, including historical financials, segment analysis, news and downloadable Excel models.

Do not deploy `dashboard.py` directly for the hosted version. The live entrypoint first downloads/decrypts the newest successful daily bundle and then routes to the selected research view.

Add these Streamlit secrets:

```toml
[live_data]
repository = "Antzaz/Antzaz-equity-research-model"
github_token = "YOUR_FINE_GRAINED_GITHUB_TOKEN"
bundle_password = "THE_SAME_VALUE_AS_LIVE_BUNDLE_PASSWORD"
workflow_file = "daily-portfolio-refresh.yml"
```

The GitHub token only needs permission to read this repository's Actions metadata/artifacts. Use the narrowest fine-grained token possible.

## Sharing with a friend

Keep the hosted portal private. In Streamlit Community Cloud, open the app's **Share / Settings → Sharing** controls and add your friend's email as a **viewer**.

A viewer should not need the GitHub token, bundle password, Actions secrets, or repository write access. They use the private Streamlit URL and authenticate through the access method offered by Streamlit.

Do **not** give repository write/admin access merely to let someone view the app. Developer access can also grant app-management capabilities. If your friend wants to run their own portfolio rather than view yours, have them fork/copy the public code and configure their own private portfolio secrets instead of sharing yours.

## Privacy model

Public GitHub contains only code and templates. Private holdings live in GitHub Actions secrets and inside AES-256-encrypted artifacts. The daily artifact uses generic workbook names (`company_01.xlsx`, etc.); the ticker mapping is inside the encrypted manifest.

The generated daily artifact is retained for 14 days. The hosted portal always selects the newest non-expired artifact from the latest successful workflow run.

## Important limitation

The portfolio file is currently supplied through a GitHub Actions secret. Therefore changing `portfolio.csv` locally does **not** automatically update GitHub's copy. After a portfolio change, update `PORTFOLIO_CSV_B64` once. A later enhancement can move holdings to a private database/editable portfolio UI so portfolio changes themselves sync automatically.
