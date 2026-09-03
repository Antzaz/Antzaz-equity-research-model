from __future__ import annotations

"""Build a recruiter-safe public fundamentals payload for the showcase.

This deliberately publishes only company names and public market/research characteristics.
It never serializes tickers, share counts, cost basis, market values, P&L, transactions,
credentials, private notes, or private workbooks. Reverse DCF is business-model gated so a
bank/insurer/REIT is never presented with an industrial FCF hurdle as if it were meaningful.
"""

import csv
import json
import time
from datetime import datetime, timezone
from pathlib import Path
import sys

import numpy as np
import yfinance as yf

ROOT = Path(__file__).resolve().parents[1]
PORTFOLIO = ROOT / "institutional_research" / "portfolio.csv"
CONFIG = ROOT / "institutional_research" / "config.json"
DEST = ROOT / "showcase" / "data" / "public_fundamentals.json"

sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "institutional_research"))
from business_model_registry import get_business_model_policy, reverse_dcf_applicability_message  # noqa: E402
from src.reverse_dcf import solve_implied_growth  # noqa: E402


def _tickers() -> list[str]:
    if not PORTFOLIO.exists():
        raise SystemExit(f"Missing private portfolio input: {PORTFOLIO}")
    with PORTFOLIO.open("r", encoding="utf-8-sig", newline="") as f:
        rows = csv.DictReader(line for line in f if not line.lstrip().startswith("#"))
        values = [str(row.get("Ticker") or "").strip().upper() for row in rows]
    return list(dict.fromkeys(t for t in values if t))


def _num(value):
    try:
        value = float(value)
        return value if np.isfinite(value) else None
    except Exception:
        return None


def _fetch_info(ticker: str, attempts: int = 3) -> dict:
    last_error = None
    for attempt in range(attempts):
        try:
            info = yf.Ticker(ticker).info or {}
            if info:
                return info
        except Exception as exc:
            last_error = exc
        if attempt < attempts - 1:
            time.sleep(1.5 * (attempt + 1))
    if last_error:
        print(f"WARNING: public fundamentals unavailable for one holding after retries: {type(last_error).__name__}")
    return {}


def _forward_pe(info: dict):
    direct = _num(info.get("forwardPE"))
    if direct is not None:
        return direct
    price = _num(info.get("currentPrice") or info.get("regularMarketPrice"))
    forward_eps = _num(info.get("forwardEps"))
    if price is not None and forward_eps is not None and forward_eps > 0:
        return price / forward_eps
    return None


def main() -> None:
    tickers = _tickers()
    cfg = json.loads(CONFIG.read_text(encoding="utf-8"))
    dcf = cfg.get("reverse_dcf", {})

    companies: list[dict] = []
    for ticker in tickers:
        info = _fetch_info(ticker)
        if not info:
            continue

        company = str(info.get("longName") or info.get("shortName") or "").strip()
        if not company:
            print("WARNING: skipping one holding because no public company name was returned")
            continue

        policy = get_business_model_policy(
            ticker, info.get("sector"), info.get("industry"), company
        )
        market_cap = _num(info.get("marketCap"))
        fcf = _num(info.get("freeCashflow"))
        cash = _num(info.get("totalCash")) or 0.0
        debt = _num(info.get("totalDebt")) or 0.0
        implied = None
        reverse_status = reverse_dcf_applicability_message(policy)
        if policy.reverse_dcf_allowed and market_cap is not None and fcf is not None:
            implied_raw = solve_implied_growth(
                target_market_cap=market_cap,
                fcf0=fcf,
                years=int(dcf.get("years", 10)),
                wacc=float(dcf.get("wacc", 0.09)),
                terminal_growth=float(dcf.get("terminal_growth", 0.03)),
                cash=cash,
                debt=debt,
                min_growth=float(dcf.get("min_growth", -0.20)),
                max_growth=float(dcf.get("max_growth", 0.50)),
            )
            implied = _num(implied_raw)
            reverse_status = "Solved" if implied is not None else "Not meaningful / insufficient public FCF data"

        aliases = []
        if company == "JPMorgan Chase & Co.":
            aliases.append("JPM")
        if company.lower().startswith("sanofi"):
            aliases.append("SNY")

        row = {
            "company": company,
            "aliases": aliases,
            "forward_pe": _forward_pe(info),
            "revenue_growth": _num(info.get("revenueGrowth")),
            "operating_margin": _num(info.get("operatingMargins")),
            "roe": _num(info.get("returnOnEquity")),
            "reverse_dcf": {
                "implied_annual_fcf_growth": implied,
                "wacc": _num(dcf.get("wacc")) if policy.reverse_dcf_allowed else None,
                "terminal_growth": _num(dcf.get("terminal_growth")) if policy.reverse_dcf_allowed else None,
                "forecast_years": int(dcf.get("years", 10)),
                "status": reverse_status,
                "primary_valuation": policy.primary_valuation,
            },
        }
        companies.append(row)

    if len(companies) < max(1, len(tickers) // 2):
        raise SystemExit(
            f"Only {len(companies)} of {len(tickers)} holdings returned public fundamentals; refusing to publish a mostly empty refresh."
        )

    payload = {
        "generated_utc": datetime.now(timezone.utc).isoformat(),
        "source_note": (
            "Public market characteristics refreshed independently from the full portfolio run. "
            "Forward P/E uses the public forward P/E field or price/forward-EPS fallback. "
            "Reverse DCF uses the configured simplified FCF model only for business models where industrial FCF is economically appropriate; otherwise it is explicitly N/M."
        ),
        "companies": companies,
    }
    DEST.parent.mkdir(parents=True, exist_ok=True)
    DEST.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"Recruiter-safe public fundamentals written to: {DEST}")
    print(f"Companies exported: {len(companies)}")
    for field in ["forward_pe", "revenue_growth", "operating_margin", "roe"]:
        print(f"{field}: {sum(1 for x in companies if x.get(field) is not None)} / {len(companies)}")
    print(
        "reverse_dcf: "
        f"{sum(1 for x in companies if (x.get('reverse_dcf') or {}).get('implied_annual_fcf_growth') is not None)} / {len(companies)}"
    )


if __name__ == "__main__":
    main()
