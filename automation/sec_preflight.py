from __future__ import annotations

"""Validate the SEC automated-access identity used by unattended research jobs.

The SEC asks automated clients to declare a descriptive User-Agent with contact
information and to stay within its fair-access limits. This helper deliberately
never prints the configured identity because it may contain a personal email.
"""

import argparse
import os
import re

import requests

SEC_TEST_URL = "https://www.sec.gov/files/company_tickers.json"
EMAIL_RE = re.compile(r"[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}", re.IGNORECASE)


def validate_user_agent(value: str | None) -> tuple[bool, str]:
    ua = (value or "").strip()
    if not ua:
        return False, "SEC_USER_AGENT is not configured"
    if len(ua) < 12:
        return False, "SEC_USER_AGENT is too short to be a descriptive automated-access identity"
    if not EMAIL_RE.search(ua):
        return False, "SEC_USER_AGENT should include a real contact email"
    return True, "configured"


def probe_sec(user_agent: str, timeout: int = 20) -> tuple[bool, str]:
    headers = {
        "User-Agent": user_agent,
        "Accept-Encoding": "gzip, deflate",
        "Accept": "application/json",
    }
    try:
        response = requests.get(SEC_TEST_URL, headers=headers, timeout=timeout)
        response.raise_for_status()
        payload = response.json()
        if not isinstance(payload, dict) or not payload:
            return False, "SEC returned an unexpected response"
    except Exception as exc:
        return False, f"SEC probe failed: {type(exc).__name__}: {exc}"
    return True, "SEC automated access is reachable"


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="Check SEC automated-access configuration")
    parser.add_argument(
        "--strict",
        action="store_true",
        help="Return non-zero when SEC identity or connectivity is unavailable",
    )
    parser.add_argument(
        "--no-network",
        action="store_true",
        help="Validate only the User-Agent format; do not probe SEC.gov",
    )
    args = parser.parse_args(argv)

    ua = (os.getenv("SEC_USER_AGENT") or "").strip()
    valid, note = validate_user_agent(ua)
    if not valid:
        print(f"[sec] DISABLED: {note}. Yahoo fundamentals remain the transparent fallback.")
        return 2 if args.strict else 0

    if args.no_network:
        print("[sec] ENABLED: declared SEC automated-access identity is configured.")
        return 0

    reachable, note = probe_sec(ua)
    if not reachable:
        print(f"[sec] DEGRADED: {note}. Yahoo fundamentals remain the transparent fallback.")
        return 3 if args.strict else 0

    print("[sec] ENABLED: declared identity validated and SEC public data is reachable.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
