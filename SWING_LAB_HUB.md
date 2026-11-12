# Antzaz Research Hub — Swing Trading Lab

The Windows **Antzaz Research Hub** (`desktop_research_hub.py` in the main equity-research repository) now has a fourth independent tool: **Swing Trading Lab**.

## One-time setup

1. Extract the standalone `Antzaz-swing-trading-lab` ZIP into `Documents\Antzaz-swing-trading-lab`, next to `Documents\Antzaz-equity-research-model`.
2. In the swing lab folder, sign in to GitHub CLI (`gh auth login`) and run `powershell -ExecutionPolicy Bypass -File .\publish_to_github.ps1`. This creates the separate **private** `Antzaz/Antzaz-swing-trading-lab` repository. This step must be completed once; there is currently no GitHub repository at that path.
3. Open your existing Research Hub from the desktop after pulling the latest main equity-research code (`git pull --ff-only origin main`). The new **Swing Trading Lab** card has **Install / Update Swing Lab**, **Open Swing Dashboard**, and **Refresh Swing Data**.
4. **Install / Update Swing Lab** verifies the swing repository's origin and main branch, uses safe Git fast-forward updates, and installs a dedicated `.venv` if necessary. If you didn't extract locally first, it can clone the private project after you publish it, provided Git authentication is working.
5. **Refresh Swing Data** downloads genuine Yahoo price history and generates a cached scan. **Open Swing Dashboard** opens http://localhost:8503 from saved data, without requiring an internet connection.

The older Portfolio Dashboard continues on port **8501**, and Macro Lab on **8502**. No orders are placed; Swing Lab remains research-only.

## Going forward

- **Update from GitHub** in Research Hub updates main repo code and an existing published Swing Lab checkout if present; missing projects are skipped.
- **Sync All Data (Online)** tries to pull swing code, fetch real Yahoo Finance history and update its scan, while continuing to update Macro and Portfolio even when Yahoo is unavailable.
- On an installed Swing Lab checkout you can also run `update_project.bat` or `update_and_dashboard.bat` independently of the Hub.
- To use a custom project folder, set `ANTZAZ_SWING_LAB_PATH` before launching the Hub.

Safety checks reject unexpected swing repo origins, non-main branches and ZIP-only folders without Git metadata. Historical Yahoo caches and private research outputs are gitignored and **not** pushed to GitHub. Separate Windows PCs need their own data refresh. If GitHub cannot be reached, existing dashboards can still open offline, but the code update and Yahoo data refresh require a connection.
