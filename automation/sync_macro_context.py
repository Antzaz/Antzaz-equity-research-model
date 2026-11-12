from __future__ import annotations

"""Explicit online-only bridge: public sanitized Macro Lab context -> local cache.

Never call from the offline desktop launcher. Never fetch from private GitHub repos.
"""

import argparse
import json
from pathlib import Path
import sys
from urllib.request import Request, urlopen

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))
from institutional_research.src.macro_bridge import validate_macro_context

SOURCE_URL = (
    "https://raw.githubusercontent.com/Antzaz/"
    "Antzaz-investment-research-showcase/main/data/macro_context.json"
)
DEFAULT_DEST = Path(__file__).resolve().parents[1] / "institutional_research/.offline_cache/macro_context.json"


def sync(destination: Path, url: str = SOURCE_URL) -> dict:
    request = Request(url, headers={"User-Agent": "Antzaz-macro-context-bridge/1"})
    with urlopen(request, timeout=15) as response:
        data = json.loads(response.read(16384).decode("utf-8"))
    checked = validate_macro_context(data)
    destination.parent.mkdir(parents=True, exist_ok=True)
    temporary = destination.with_name(destination.name + ".tmp")
    temporary.write_text(json.dumps(checked, indent=2) + "\n", encoding="utf-8")
    temporary.replace(destination)
    return checked


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="Cache sanitized macro regime context from public showcase")
    parser.add_argument("--output", type=Path, default=DEFAULT_DEST)
    args = parser.parse_args(argv)
    result = sync(args.output)
    print(f"Synced macro context: {result['regime']} as of {result['as_of']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
