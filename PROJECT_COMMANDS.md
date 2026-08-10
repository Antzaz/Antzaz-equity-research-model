# Project launch commands

Assumes the repository is located at:

`C:\Users\Antza\Documents\Antzaz-equity-research-model`

## Equity research model

Replace `UBS` with any ticker:

```powershell
cd "C:\Users\Antza\Documents\Antzaz-equity-research-model"; git pull; python .\update_model.py UBS
```

## Portfolio research — refresh data and launch local dashboard

```powershell
cd "C:\Users\Antza\Documents\Antzaz-equity-research-model\institutional_research"; python .\run_research.py; python -m streamlit run .\dashboard.py
```

## Private unified live portal

Requires the `[live_data]` Streamlit secrets and a successful `Daily private portfolio refresh` GitHub Actions run:

```powershell
cd "C:\Users\Antza\Documents\Antzaz-equity-research-model"; git pull; python -m streamlit run .\institutional_research\live_dashboard.py
```

The private portal contains:

- Portfolio Dashboard
- Alpha Analysis
- Company Research

## Resume-safe public showcase

```powershell
cd "C:\Users\Antza\Documents\Antzaz-equity-research-model"; git pull; python -m pip install -r .\showcase\requirements.txt; python -m streamlit run .\showcase\app.py
```

## Export only the public showcase into a clean separate folder

```powershell
cd "C:\Users\Antza\Documents\Antzaz-equity-research-model"; git pull; powershell -ExecutionPolicy Bypass -File .\automation\export_showcase.ps1
```

The default output folder is:

`C:\Users\Antza\Documents\Antzaz-investment-research-showcase`

If GitHub CLI (`gh`) is installed and authenticated, publish the exported folder as a new public repository with:

```powershell
cd "$HOME\Documents\Antzaz-investment-research-showcase"; gh repo create Antzaz-investment-research-showcase --public --source . --remote origin --push
```

## Recommended hosting structure

- Main repository: PRIVATE
- Streamlit private app: `institutional_research/live_dashboard.py`
- Friend access: invite as Streamlit viewer, not GitHub collaborator
- Resume/public app: deploy `app.py` from the separate public showcase repository
