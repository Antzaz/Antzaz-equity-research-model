"""Safe paths and commands for the independent Antzaz Swing Trading Lab.

This module is intentionally stdlib-only: Windows Research Hub imports it without
loading the swing project's optional scientific dependencies.
"""
from __future__ import annotations

import os
from pathlib import Path
import re

SWING_REPOSITORY = "https://github.com/Antzaz/Antzaz-swing-trading-lab.git"
SWING_FOLDER_NAME = "Antzaz-swing-trading-lab"
SWING_DASHBOARD_PORT = 8503
SWING_DASHBOARD_URL = f"http://localhost:{SWING_DASHBOARD_PORT}"
_ORIGIN_PATTERN = re.compile(
    r"^(?:https://github\.com/|git@github\.com:|ssh://git@github\.com/)"
    r"antzaz/antzaz-swing-trading-lab(?:\.git)?/?$",
    re.IGNORECASE,
)


def swing_root(main_repo_root: Path) -> Path:
    """Use a sibling checkout, or an explicit user's path override."""
    override = os.environ.get("ANTZAZ_SWING_LAB_PATH", "").strip()
    if override:
        return Path(override).expanduser().resolve()
    return (Path(main_repo_root).resolve().parent / SWING_FOLDER_NAME).resolve()


def valid_swing_origin(remote_url: str) -> bool:
    """Never pull an unrelated GitHub repository in a reused project folder."""
    return bool(_ORIGIN_PATTERN.fullmatch((remote_url or "").strip()))


def swing_project_ready(root: Path) -> bool:
    return all((root / p).is_file() for p in ("pyproject.toml", "dashboard.py", "swinglab/cli.py"))


def swing_git_checkout_present(root: Path) -> bool:
    return (root / ".git").exists() and swing_project_ready(root)


def swing_python(root: Path) -> Path:
    """Always use the swing project's own isolated Python environment."""
    return Path(root) / ".venv" / ("Scripts/python.exe" if os.name == "nt" else "bin/python")


def swing_clone_command(root: Path) -> list[str]:
    return ["git", "clone", "--branch", "main", SWING_REPOSITORY, str(root)]


def swing_pull_command() -> list[str]:
    return ["git", "pull", "--ff-only", "origin", "main"]


def swing_refresh_commands(root: Path) -> list[list[str]]:
    py = str(swing_python(root))
    return [
        [py, "-m", "swinglab.cli", "fetch"],
        [py, "-m", "swinglab.cli", "scan", "--cache-only"],
    ]


def swing_dashboard_command(root: Path) -> list[str]:
    return [
        str(swing_python(root)), "-m", "streamlit", "run", "dashboard.py",
        "--server.address", "127.0.0.1", "--server.port", str(SWING_DASHBOARD_PORT),
        "--server.headless", "true", "--browser.gatherUsageStats", "false",
    ]
