from __future__ import annotations

"""Build one encrypted bundle containing the latest private portfolio outputs and
one equity-research workbook per active portfolio company.

The archive uses generic filenames so portfolio tickers are not exposed through the
ZIP directory listing or GitHub Actions artifact metadata. The ticker-to-file mapping
lives inside the encrypted manifest.
"""

import argparse
import json
from datetime import datetime, timezone
from pathlib import Path

import pyzipper

from portfolio_universe import portfolio_tickers


def latest_model(models_dir: Path, ticker: str) -> Path | None:
    files = list(models_dir.glob(f"{ticker}_Equity_Research_*.xlsx"))
    return max(files, key=lambda p: p.stat().st_mtime) if files else None


def build_bundle(portfolio_csv: Path, portfolio_outputs: Path, models_dir: Path, output: Path, password: str) -> dict:
    tickers = portfolio_tickers(portfolio_csv)
    output.parent.mkdir(parents=True, exist_ok=True)
    manifest = {
        "generated_utc": datetime.now(timezone.utc).isoformat(),
        "company_count": len(tickers),
        "companies": [],
        "portfolio_output_dir": "portfolio_outputs",
    }

    with pyzipper.AESZipFile(
        output,
        mode="w",
        compression=pyzipper.ZIP_DEFLATED,
        encryption=pyzipper.WZ_AES,
    ) as zf:
        zf.setpassword(password.encode("utf-8"))
        zf.setencryption(pyzipper.WZ_AES, nbits=256)

        if portfolio_outputs.exists():
            for path in sorted(portfolio_outputs.rglob("*")):
                if path.is_file():
                    rel = path.relative_to(portfolio_outputs)
                    zf.write(path, arcname=str(Path("portfolio_outputs") / rel))

        for idx, ticker in enumerate(tickers, 1):
            model = latest_model(models_dir, ticker)
            generic = f"company_{idx:02d}.xlsx"
            entry = {"ticker": ticker, "file": None, "status": "missing"}
            if model is not None:
                arcname = str(Path("equity_models") / generic)
                zf.write(model, arcname=arcname)
                entry.update({"file": arcname, "status": "ok"})
            manifest["companies"].append(entry)

        zf.writestr("manifest.json", json.dumps(manifest, indent=2).encode("utf-8"))

    return manifest


def main() -> None:
    p = argparse.ArgumentParser()
    p.add_argument("--portfolio", default="institutional_research/portfolio.csv")
    p.add_argument("--portfolio-outputs", default="institutional_research/outputs/latest")
    p.add_argument("--models-dir", default="updated_models")
    p.add_argument("--output", required=True)
    p.add_argument("--password", required=True)
    args = p.parse_args()

    manifest = build_bundle(
        Path(args.portfolio), Path(args.portfolio_outputs), Path(args.models_dir), Path(args.output), args.password
    )
    # Deliberately print counts only; never disclose portfolio tickers in CI logs.
    ok = sum(1 for x in manifest["companies"] if x["status"] == "ok")
    print(f"Encrypted live bundle created: {ok}/{manifest['company_count']} company models included")


if __name__ == "__main__":
    main()
