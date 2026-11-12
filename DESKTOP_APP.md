# Windows Desktop Launcher

The private repository includes a small Windows desktop launcher so the latest equity-research and portfolio tools can be used like a normal app.

## First setup on each computer

1. Install Git (or GitHub Desktop) and Python 3.
2. Clone the private Antzaz-equity-research-model repository.
3. Open the cloned folder.
4. Double-click install_desktop_app.bat.

The installer updates the clone when possible, installs the equity-research and portfolio dependencies, and creates a desktop shortcut called Antzaz Investment Research.

After that, use the desktop shortcut instead of PowerShell.

## Launcher buttons

- Update from GitHub: fast-forwards from origin/main.
- Run Equity Research: runs the full guarded research.py ticker workflow.
- Build Model Only: runs the deterministic guarded workbook build.
- Open Portfolio Dashboard: refreshes the private portfolio analytics and starts Streamlit.
- Open Latest Workbook: opens the newest generated workbook for the ticker entered in the app.
- Open Project Folder.

Update from GitHub before launch is enabled by default. The desktop shortcut also tries a quiet fast-forward update before the window opens, so a computer that has not been used recently starts from the newest available repo version.

## Portfolio portability

institutional_research/portfolio.csv is tracked in this PRIVATE repository so an authenticated clone on another computer starts with the same holdings, shares and average costs.

Do not copy that file to the public showcase repository.

If you later edit portfolio.csv and want those edits to appear on another computer, commit and push the private file before changing computers. The launcher intentionally does not silently overwrite or auto-push private portfolio changes.

## Git safety

Updates use git pull --ff-only. If a computer has unresolved local changes, the app refuses to overwrite them. Commit or stash those changes and update again.

## Removing the desktop shortcut

Delete Antzaz Investment Research from the Windows desktop. This does not delete the repository.

## Current Antzaz Research Hub: four research tools

The desktop shortcut launches `desktop_research_hub.py`, an **offline-first**
Research Hub with **Portfolio Research**, **Bonds & Macro**, **Equity Research**
and **Swing Trading Lab** (separate, research-only, Finnish and US).

The Swing card has **Install / Update Swing Lab**, **Open Swing Dashboard** and
**Refresh Swing Data**. Its separate dashboard runs on localhost port **8503**,
alongside Portfolio on 8501 and Macro on 8502. It uses a dedicated Python
virtual environment and GitHub repository. **Update from GitHub** also pulls an
installed Swing Lab checkout; **Sync All Data (Online)** attempts a swing market
refresh without blocking portfolio research when Yahoo is unavailable.

The new Swing Lab repository **must be published once** to a separate **private**
GitHub repository before Git-based updates or automatic cloning work. Download
the latest ZIP and run its `publish_to_github.ps1` script after authenticating
GitHub CLI. See [SWING_LAB_HUB.md](SWING_LAB_HUB.md) for complete instructions.

**Correction to legacy notes above:** opening the desktop shortcut does not
automatically update from GitHub. The current Research Hub intentionally starts
offline and updates only through explicit update and synchronization actions.
The main GitHub repository visibility should be verified in GitHub Settings;
do not assume that it is private or store private holdings in tracked files.
