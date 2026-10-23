from __future__ import annotations

import importlib.util
from pathlib import Path

import pytest

ROOT=Path(__file__).resolve().parents[1]
spec=importlib.util.spec_from_file_location("desktop_launcher",ROOT/"desktop_launcher.py")
assert spec and spec.loader
module=importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)


def test_desktop_launcher_commands_are_guarded():
    assert module.normalize_ticker(" googl ")=="GOOGL"
    assert module.git_update_command()==["git","pull","--ff-only","origin","main"]
    full=module.full_research_command("msft")
    assert Path(full[1]).name=="research.py"
    assert full[-1]=="MSFT"
    model=module.model_only_command("tsm")
    assert Path(model[1]).name=="commodity_safe_runner.py"
    assert model[-1]=="TSM"


def test_desktop_launcher_rejects_shell_injection():
    with pytest.raises(ValueError):
        module.normalize_ticker("GOOGL & calc")
