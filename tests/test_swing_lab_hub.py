from pathlib import Path

import swing_lab_hub as lab


def test_commands_and_sibling_default(tmp_path, monkeypatch):
    monkeypatch.delenv("ANTZAZ_SWING_LAB_PATH", raising=False)
    root = lab.swing_root(tmp_path / "equity")
    assert root == tmp_path / "Antzaz-swing-trading-lab"
    assert lab.swing_pull_command() == ["git", "pull", "--ff-only", "origin", "main"]
    assert lab.swing_clone_command(root)[0:3] == ["git", "clone", "--branch"]
    assert lab.swing_dashboard_command(root)[-6:] == [
        "--server.port", "8503", "--server.headless", "true",
        "--browser.gatherUsageStats", "false",
    ]
    assert lab.swing_refresh_commands(root)[0][-2:] == ["swinglab.cli", "fetch"]


def test_override_and_safety(tmp_path, monkeypatch):
    override = tmp_path / "custom"
    monkeypatch.setenv("ANTZAZ_SWING_LAB_PATH", str(override))
    assert lab.swing_root(tmp_path / "equity") == override
    for trusted in (
        "https://github.com/Antzaz/Antzaz-swing-trading-lab.git",
        "git@github.com:Antzaz/Antzaz-swing-trading-lab.git",
        "ssh://git@github.com/Antzaz/Antzaz-swing-trading-lab",
    ):
        assert lab.valid_swing_origin(trusted)
    for untrusted in (
        "", "file:///tmp/Antzaz-swing-trading-lab",
        "https://github.com/other/Antzaz-swing-trading-lab.git",
        "https://github.com/Antzaz/Antzaz-swing-trading-lab-extra",
    ):
        assert not lab.valid_swing_origin(untrusted)


def test_checkout_requires_git_and_code(tmp_path):
    (tmp_path / ".git").mkdir()
    assert not lab.swing_git_checkout_present(tmp_path)
    for relpath in ("pyproject.toml", "dashboard.py", "swinglab/cli.py"):
        path = tmp_path / relpath
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text("pass", encoding="utf-8")
    assert lab.swing_git_checkout_present(tmp_path)
    assert ".venv" in lab.swing_python(tmp_path).parts


def test_hub_has_all_three_actions_and_sync():
    source = (Path(__file__).resolve().parents[1] / "desktop_research_hub.py").read_text("utf-8")
    for phrase in (
        "Open Swing Dashboard", "Install / Update Swing Lab",
        "Refresh Swing Data", "def _sync_swing_code(", "def _swing_task(",
        "def _refresh_swing_data(", "self._sync_swing_code(allow_clone=False)",
    ):
        assert phrase in source
