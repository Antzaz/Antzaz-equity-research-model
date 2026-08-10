from __future__ import annotations

"""Return the active equity-research universe from the private portfolio CSV.

The script prints tickers only; it never prints shares, costs, weights, or notes.
It is used by the private daily refresh workflow so the equity-research universe
always follows the portfolio rather than a separately maintained ticker list.
"""

import argparse
import re
from pathlib import Path

import pandas as pd

TICKER_RE = re.compile(r"^[A-Z0-9.\-]{1,15}$")


def portfolio_tickers(path: str | Path) -> list[str]:
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(path)

    df = pd.read_csv(path, comment="#")
    if "Ticker" not in df.columns:
        raise ValueError("portfolio.csv must contain a Ticker column")

    tickers = df["Ticker"].astype(str).str.upper().str.strip()
    active = pd.Series(True, index=df.index)

    # If position fields are present, ignore explicit zero/closed rows. Missing values
    # remain eligible because the portfolio project supports either Shares or ManualWeight.
    shares = pd.to_numeric(df.get("Shares"), errors="coerce") if "Shares" in df.columns else None
    weights = pd.to_numeric(df.get("ManualWeight"), errors="coerce") if "ManualWeight" in df.columns else None
    if shares is not None or weights is not None:
        positive_shares = (shares > 0) if shares is not None else pd.Series(False, index=df.index)
        positive_weights = (weights > 0) if weights is not None else pd.Series(False, index=df.index)
        unspecified = pd.Series(True, index=df.index)
        if shares is not None:
            unspecified &= shares.isna()
        if weights is not None:
            unspecified &= weights.isna()
        active = positive_shares | positive_weights | unspecified

    out: list[str] = []
    for ticker in tickers[active]:
        if ticker in {"", "NAN", "NONE"}:
            continue
        if not TICKER_RE.fullmatch(ticker):
            raise ValueError(f"Invalid ticker in portfolio.csv: {ticker!r}")
        if ticker not in out:
            out.append(ticker)

    if not out:
        raise ValueError("portfolio.csv contains no active tickers")
    return out


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("portfolio", nargs="?", default="institutional_research/portfolio.csv")
    args = parser.parse_args()
    for ticker in portfolio_tickers(args.portfolio):
        print(ticker)


if __name__ == "__main__":
    main()
