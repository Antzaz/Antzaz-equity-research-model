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
