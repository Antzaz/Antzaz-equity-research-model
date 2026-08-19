# Recruiter-Facing Streamlit Deployment

This guide turns the project into a clean recruiter-facing setup:

```text
Private research repository
        ↓
real portfolio analytics
        ↓
sanitization + validation
        ↓
public showcase repository
        ↓
public Streamlit Community Cloud app
        ↓
resume / LinkedIn / job applications
```

## Target architecture

### Private repository

`Antzaz/Antzaz-equity-research-model`

Keep here:

- equity-research engine and source code
- private portfolio inputs
- research workbooks
- full company models
- private Streamlit portal
- GitHub Actions secrets

### Public repository

`Antzaz/Antzaz-investment-research-showcase`

Expose only:

- `app.py`
- `requirements.txt`
- `README.md`
- `.gitignore`
- `data/portfolio_snapshot.json`

The JSON contains sanitized real aggregate portfolio analytics. Holdings are anonymized.

## Important current-state note

The main research repository should be private if the goal is to show recruiters the finished work without exposing the implementation.

On GitHub:

1. Open `Antzaz/Antzaz-equity-research-model`.
2. Open **Settings**.
3. Scroll to **Danger Zone**.
4. Choose **Change repository visibility**.
5. Change the repository to **Private** and confirm the visibility-change warnings.

Do this only after the separate public showcase repository has been created and verified.

## Step 1 — refresh the real portfolio locally

From PowerShell:

```powershell
cd "C:\Users\Antza\Documents\Antzaz-equity-research-model"
git pull
python -m pip install -r .\institutional_research\requirements.txt
python .\institutional_research\run_research.py
```

This must finish successfully before recruiter publication.

## Step 2 — build and validate the recruiter-safe snapshot

Run:

```powershell
python .\automation\build_showcase_snapshot.py
python .\automation\validate_showcase.py
```

Validation must say:

```text
Recruiter showcase validation passed.
```

The validator checks that:

- core portfolio return/risk metrics exist
- multiple anonymous holdings exist
- growth-history data exists
- portfolio weights approximately sum to 100%
- the snapshot is marked as real sanitized analytics
- forbidden ticker/company/cost-basis/market-value/transaction fields are absent

## Step 3 — create the clean public showcase folder

Run:

```powershell
powershell -ExecutionPolicy Bypass -File .\automation\export_showcase.ps1
```

The script now fails rather than silently publishing a demo portfolio when real validated analytics are unavailable.

Default destination:

```text
C:\Users\Antza\Documents\Antzaz-investment-research-showcase
```

## Step 4 — test the recruiter version locally

Run:

```powershell
cd "$HOME\Documents\Antzaz-investment-research-showcase"
python -m pip install -r .\requirements.txt
python -m streamlit run .\app.py
```

Before publishing, verify:

- Portfolio Analytics loads
- six headline metrics render
- portfolio-vs-benchmark chart renders
- anonymous portfolio-weight pie chart renders
- capital weight vs risk contribution renders
- alpha chart renders when alpha data is available
- factor exposure chart renders
- stress scenarios render
- no ticker symbols, company names, cost basis, shares or portfolio value appear

## Step 5 — create the separate public GitHub repository

With GitHub CLI installed and authenticated:

```powershell
cd "$HOME\Documents\Antzaz-investment-research-showcase"
gh repo create Antzaz-investment-research-showcase --public --source . --remote origin --push
```

If you prefer the GitHub website, create a new public repository named:

```text
Antzaz-investment-research-showcase
```

Then push the exported folder to its `main` branch.

Do not add the full research repository as the public app source.

## Step 6 — configure automatic sanitized updates

The existing workflow `.github/workflows/daily-portfolio-refresh.yml` already knows how to publish only the sanitized snapshot to the public showcase repository.

### Create a fine-grained GitHub token

Create a fine-grained personal access token with:

- resource owner: `Antzaz`
- repository access: only `Antzaz-investment-research-showcase`
- repository permission: **Contents — Read and write**

Use the shortest practical expiration and rotate the token when needed.

### Add the token to the private research repository

In `Antzaz/Antzaz-equity-research-model`:

1. **Settings**
2. **Secrets and variables**
3. **Actions**
4. **Secrets**
5. **New repository secret**

Name:

```text
SHOWCASE_REPO_TOKEN
```

Value: the fine-grained token.

### Add the public repository variable

In the same **Secrets and variables → Actions** area, open **Variables** and create:

```text
SHOWCASE_REPOSITORY
```

with value:

```text
Antzaz/Antzaz-investment-research-showcase
```

## Step 7 — test automatic publication

In the private repository:

1. Open **Actions**.
2. Select **Daily private portfolio refresh**.
3. Choose **Run workflow**.
4. Wait for the run to finish.
5. Open the public showcase repository.
6. Confirm that `data/portfolio_snapshot.json` received a new commit titled similar to `Refresh sanitized portfolio analytics`.

The workflow now validates the public snapshot before publishing it.

## Step 8 — deploy the public Streamlit app

Use Streamlit Community Cloud.

1. Sign in with GitHub and connect your public repositories.
2. Choose **Create app**.
3. Choose the existing-app option.
4. Repository: `Antzaz/Antzaz-investment-research-showcase`.
5. Branch: `main`.
6. Entrypoint file: `app.py`.
7. Choose a short professional subdomain if available, for example something based on `anton-investment-research`.
8. No Streamlit secrets are required for the public showcase.
9. Deploy.
10. In the app's sharing settings, verify that the app is public.

Streamlit Community Cloud follows the GitHub repository, so future commits to the public showcase should update the deployed app automatically.

## Step 9 — recruiter test

Open the Streamlit URL in an incognito/private browser window where you are not signed in.

The recruiter test passes only if:

- the app opens without authentication
- no GitHub access is required
- the portfolio visuals render
- the portfolio snapshot is real and sanitized
- no sensitive position-level information is visible
- the page works on both desktop and mobile widths

## Step 10 — add the link to your resume

Use the public Streamlit URL as the main project link, not the private GitHub repository.

Suggested resume label:

```text
Interactive Equity Research & Portfolio Analytics Platform
```

Optional secondary link:

```text
Public project repository
```

The Streamlit link should be the primary call-to-action because recruiters can evaluate the project without reading code.

## Ongoing maintenance

After the one-time setup:

- private portfolio stays in the private research repository / Actions secret
- daily workflow refreshes the real analytics
- sanitizer removes identifiers and sensitive economics
- validator blocks unsafe/incomplete snapshots
- sanitized JSON is pushed to the public showcase repository
- Streamlit automatically picks up the public repository update

That means the recruiter link can stay the same while the displayed portfolio analytics refresh over time.
