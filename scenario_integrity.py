from __future__ import annotations

"""Cross-company scenario-assumption integrity guards.

The base generator historically clipped operating margins to 1%..50%. That is unsafe for both
very high-margin businesses (payments/software) and loss-making businesses (biotech/turnarounds):
it silently replaces observed economics with an arbitrary template bound. This guard runs only
on freshly generated default scenarios and restores margin paths around the latest reported
operating margin when the legacy bound would have changed it.
"""

from typing import Any

from business_model_registry import get_business_model_policy


def _finite(v: Any) -> float | None:
    try:
        x = float(v)
        return x if x == x and abs(x) != float('inf') else None
    except Exception:
        return None


def _fade(start: float, end: float, n: int = 10) -> list[float]:
    if n <= 1:
        return [end]
    return [start + (end - start) * i / (n - 1) for i in range(n)]


def _latest_margin(hist: dict | None) -> tuple[int | None, float | None]:
    rows = hist or {}
    for year in sorted(rows, reverse=True):
        row = rows.get(year) or {}
        rev = _finite(row.get('revenue')); op = _finite(row.get('op'))
        if rev and op is not None and rev > 0:
            return int(year), float(op / rev)
    return None, None


def repair_default_scenario_margin_paths(wb, ticker: str, hist: dict | None, info: dict | None = None) -> dict:
    result = {'changed': False, 'ticker': str(ticker).upper().strip()}
    if 'Three-Case Scenarios' not in wb.sheetnames:
        return result
    year, margin = _latest_margin(hist)
    if margin is None:
        return result

    try:
        cd = wb['Company Data']
        policy = get_business_model_policy(ticker, cd['B6'].value, cd['B7'].value, cd['B5'].value)
    except Exception:
        policy = get_business_model_policy(ticker)
    result.update({'business_model': policy.key, 'latest_year': year, 'latest_operating_margin': margin})

    # Banks and insurance companies do not use industrial operating-margin scenario paths.
    if not policy.industrial_fcf_primary:
        return result

    # If the legacy 1%..50% clamp did not bind, preserve the existing generated scenario.
    if 0.01 <= margin <= 0.50:
        return result

    ws = wb['Three-Case Scenarios']
    low, high = -0.80, 0.90
    clip = lambda x: max(low, min(high, float(x)))
    blocks = [range(2, 12), range(14, 24), range(26, 36)]
    paths = [
        _fade(clip(margin - 0.03), clip(margin - 0.04)),
        _fade(clip(margin), clip(margin + 0.01)),
        _fade(clip(margin + 0.01), clip(margin + 0.04)),
    ]
    for cols, values in zip(blocks, paths):
        for col, value in zip(cols, values):
            ws.cell(14, col).value = value
            ws.cell(14, col).number_format = '0.0%;[Red](0.0%);-'

    result['changed'] = True
    result['reason'] = (
        f'Legacy 1%-50% operating-margin clamp would have replaced observed FY{year} margin '
        f'{margin:.1%}; scenario paths were rebuilt around the reported margin.'
    )
    try:
        wb.calculation.calcMode = 'auto'; wb.calculation.fullCalcOnLoad = True; wb.calculation.forceFullCalc = True
    except Exception:
        pass
    return result
