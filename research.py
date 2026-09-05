"""Run the agent-assisted equity research pipeline.

Examples:
    python research.py GOOGL
    python research.py GOOGL --ai
    python research.py GOOGL --no-ml
    python research.py GOOGL --skip-model

The guarded commodity_safe_runner.py entry point runs the existing deterministic model after
installing cross-border data-integrity checks and commodity-company valuation normalization.
Specialist agents inspect the generated workbook, maintain KPI history, monitor thesis evidence,
check registered data sources, and run research QA. LLM reasoning remains opt-in.

Machine learning is enabled by default for every standard research search. The default ML mode
uses the project's maximum supported shared-history target: up to 500 names, 20 years of price
history, the full six-model quantitative stack, persistent point-in-time history, walk-forward
validation and the LightGBM AI Growth Forecast layer. Use --no-ml only when a deterministic-only
workbook is intentionally desired.

Generated workbooks also receive a non-destructive public-evidence recovery and presentation
pass. It fills avoidable qualitative evidence gaps from issuer/regulatory sources, keeps genuinely
non-comparable fields as N/A/REVIEW instead of estimating them, and consolidates the workbook
presentation without altering valuation inputs.
"""

from __future__ import annotations

import argparse
import json
import re
import subprocess
import sys
from datetime import datetime
from pathlib import Path

from agent_research import (
    AgentContext,
    FilingsAgent,
    KPIEarningsAgent,
    OpenAIResearchClient,
    ResearchQAAgent,
    SourceHealthAgent,
    ThesisMonitorAgent,
)
from chart_excel_compat import apply_chart_compatibility_fix
from chart_readability_v2 import apply_chart_readability
from workbook_enhancements import apply_workbook_enhancements

BASE = Path(__file__).resolve().parent
UPDATED_MODELS = BASE / "updated_models"
RUNS_DIR = BASE / "research_runs"


def valid_ticker(raw: str) -> str:
    ticker = raw.upper().strip()
    if not re.fullmatch(r"[A-Z0-9.\-]{1,10}", ticker):
        raise argparse.ArgumentTypeError("Enter only a ticker, e.g. GOOGL, MSFT, NVDA, TSM, or SIE.DE.")
    return ticker


def latest_workbook(ticker: str) -> Path:
    candidates = sorted(
        UPDATED_MODELS.glob(f"{ticker}_Equity_Research_*.xlsx"),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    if not candidates:
        raise FileNotFoundError(f"No generated workbook found for {ticker} in {UPDATED_MODELS}")
    return candidates[0]


def run_model(ticker: str) -> None:
    print(f"[model] Building guarded deterministic research workbook for {ticker}...")
    subprocess.run([sys.executable, str(BASE / "commodity_safe_runner.py"), ticker], cwd=BASE, check=True)


def run_ml(ticker: str, workbook: Path, *, use_llm: bool = False) -> None:
    print(f"[ml] Running maximum machine-learning research layer for {ticker} (500 names / 20 years target)...")
    subprocess.run(
        [
            sys.executable,
            str(BASE / "ml_research.py"),
            ticker,
            "--workbook",
            str(workbook),
            "--max-universe",
            "500",
            "--auto-history-limit",
            "500",
            "--auto-history-years",
            "20",
            "--auto-history-batch-size",
            "500",
        ],
        cwd=BASE,
        check=True,
    )
    print(f"[ml] Running AI Growth Forecast layer for {ticker}...")
    cmd = [
        sys.executable,
        str(BASE / "ai_growth_research.py"),
        ticker,
        "--workbook",
        str(workbook),
    ]
    if use_llm:
        cmd.append("--llm")
    subprocess.run(cmd, cwd=BASE, check=True)


def enhance_workbook(ticker: str, workbook: Path, *, reason: str) -> dict | None:
    """Apply evidence recovery/presentation cleanup without making the research run fragile."""
    try:
        result = apply_workbook_enhancements(workbook, ticker)
        print(
            f"[workbook] {reason}: workforce={result.get('workforce_status')}, "
            f"annual_filing={'yes' if result.get('annual_filing') else 'no'}, "
            f"hidden_support={','.join(result.get('hidden_support_sheets') or []) or 'none'}"
        )
        return result
    except Exception as exc:
        print(f"[workbook] WARNING: {reason} enhancement failed without changing model inputs: {exc}")
        return None


def fix_chart_rendering(workbook: Path, *, reason: str) -> dict | None:
    """Repair ML/AI chart sources and OOXML axes for Excel compatibility."""
    try:
        result = apply_chart_compatibility_fix(workbook)
        touched = ",".join(result.get("chart_helper_sheets") or []) or "none"
        count = int(result.get("charts_repaired") or 0)
        print(f"[workbook] {reason}: Excel chart compatibility repaired {count} chart(s) on {touched}")
        return result
    except Exception as exc:
        print(f"[workbook] WARNING: {reason} chart compatibility fix failed: {exc}")
        return None


def fix_chart_layout(ticker: str, workbook: Path, *, reason: str) -> dict | None:
    """Run last so no later ML/AI writer can stack or scatter charts again."""
    try:
        result = apply_chart_readability(workbook, ticker)
        print(
            f"[workbook] {reason}: rebuilt {int(result.get('charts_rebuilt') or 0)} chart(s); "
            f"layout={result.get('sheet_counts',{})}"
        )
        return result
    except Exception as exc:
        print(f"[workbook] WARNING: {reason} chart layout cleanup failed: {exc}")
        return None


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Agent-assisted equity research orchestrator")
    parser.add_argument("ticker", type=valid_ticker)
    parser.add_argument("--skip-model", action="store_true", help="Reuse the newest existing workbook instead of running the guarded deterministic model")
    parser.add_argument("--ai", action="store_true", help="Enable evidence-bound OpenAI reasoning. Requires OPENAI_API_KEY. The default ML run then also uses structured AI-growth evidence extraction.")
    ml_group = parser.add_mutually_exclusive_group()
    ml_group.add_argument("--ml", dest="ml", action="store_true", help="Run the maximum quantitative ML layer (default; retained for backwards-compatible commands).")
    ml_group.add_argument("--no-ml", dest="ml", action="store_false", help="Skip ML/AI Growth and build a deterministic-only research workbook.")
    parser.set_defaults(ml=True)
    parser.add_argument("--model", help="OpenAI model override for the research agents. The AI-growth extractor uses its cost-efficient structured-extraction default unless run directly with ai_growth_research.py --model.")
    parser.add_argument("--strict", action="store_true", help="Return a non-zero exit code when Research QA fails.")
    return parser


def render_report(ticker: str, workbook: Path, results: list, ai_model: str | None, ml_enabled: bool = False) -> str:
    lines = [
        f"# {ticker} Agent Research Report",
        "",
        f"Generated: {datetime.now().astimezone().isoformat(timespec='seconds')}",
        f"Workbook: `{workbook}`",
        f"AI reasoning: {'enabled (' + ai_model + ')' if ai_model else 'disabled; deterministic agent checks only'}",
        f"Machine learning: {'maximum mode enabled; six models with a 500-name / 20-year persistent-history target' if ml_enabled else 'disabled by --no-ml'}",
        "",
        "## Agent results",
        "",
    ]
    for result in results:
        lines.extend([f"### {result.name} — {result.status}", "", result.summary, ""])
        if result.ai_narrative:
            lines.extend(["**AI evidence review**", "", result.ai_narrative, ""])
        if result.artifacts:
            lines.append("Artifacts: " + ", ".join(f"`{x}`" for x in result.artifacts))
            lines.append("")
    lines.extend([
        "## Governance",
        "",
        "- The agents and ML models do not execute trades.",
        "- AI narrative and ML output do not overwrite financial inputs or DCF assumptions.",
        "- Public-evidence recovery does not overwrite valuation inputs or invent missing quantitative facts.",
        "- The AI Growth Forecast uses the LLM only for evidence extraction; LightGBM produces the fundamental growth forecast.",
        "- The AI evidence adjustment is bounded until enough dated AI KPI history exists for supervised training.",
        "- LightGBM is checked against an Elastic Net baseline on a chronological holdout and confidence is downgraded when it fails to match the baseline.",
        "- Reported facts, calculations, model estimates, and inference remain separate.",
        "- ML walk-forward testing and data-readiness gates are preferred to filling missing outputs with fabricated data.",
        "- Cross-border annual data passes completed-fiscal-year and scale-integrity checks before valuation.",
        "- Commodity producers use a separate normalization overlay so peak commodity/acquisition years are not extrapolated as secular growth.",
        "- Market-share records preserve their market definition and source; non-comparable market share is N/A rather than estimated.",
        "- Valuation changes remain analyst-reviewed and are calculated by the deterministic model.",
        "",
    ])
    return "\n".join(lines)


def main() -> int:
    args = build_parser().parse_args()
    ticker = args.ticker
    if not args.skip_model:
        run_model(ticker)
    workbook = latest_workbook(ticker)

    if args.skip_model:
        enhance_workbook(ticker, workbook, reason="reused workbook")
        fix_chart_rendering(workbook, reason="reused workbook")
        fix_chart_layout(ticker, workbook, reason="reused workbook final layout")

    stamp = datetime.now().astimezone().strftime("%Y%m%d_%H%M%S")
    run_dir = RUNS_DIR / ticker / stamp
    run_dir.mkdir(parents=True, exist_ok=True)

    ai_client = OpenAIResearchClient(args.model) if args.ai else None
    ctx = AgentContext(
        ticker=ticker,
        repo_root=BASE,
        workbook_path=workbook,
        run_dir=run_dir,
        use_ai=bool(ai_client),
        ai_model=ai_client.model if ai_client else None,
    )
    agents = [
        SourceHealthAgent(),
        FilingsAgent(ai_client),
        KPIEarningsAgent(ai_client),
        ThesisMonitorAgent(ai_client),
        ResearchQAAgent(ai_client),
    ]

    results = []
    for agent in agents:
        print(f"[agent] {agent.name}...")
        try:
            result = agent.run(ctx)
        except Exception as exc:
            from agent_research import AgentResult
            result = AgentResult(agent.name, "FAIL", f"Agent failed: {exc}", {"error": repr(exc)})
        results.append(result)
        print(f"        {result.status}: {result.summary}")

    if args.ml:
        try:
            run_ml(ticker, workbook, use_llm=bool(ai_client))
            enhance_workbook(ticker, workbook, reason="post-ML final polish")
            fix_chart_rendering(workbook, reason="post-ML Excel chart compatibility")
            # This MUST be the final chart writer. ML/AI layers can recreate charts, so running
            # the canonical layout after them prevents duplicated anchors and scattered graphs.
            fix_chart_layout(ticker, workbook, reason="post-ML canonical chart layout")
        except Exception as exc:
            print(f"[ml] WARNING: ML / AI Growth layer failed without changing the deterministic valuation model: {exc}")
            if args.strict:
                return 3

    manifest = {
        "ticker": ticker,
        "workbook": str(workbook),
        "ai_model": ctx.ai_model,
        "ml_enabled": bool(args.ml),
        "ml_analysis_mode": "maximum_500_names_20_years" if args.ml else None,
        "ai_growth_enabled": bool(args.ml),
        "ai_growth_llm_extraction": bool(args.ml and ai_client),
        "data_integrity_runner": "commodity_safe_runner.py",
        "public_evidence_and_workbook_polish": True,
        "final_chart_layout": True,
        "results": [x.to_dict() for x in results],
    }
    (run_dir / "manifest.json").write_text(json.dumps(manifest, indent=2, ensure_ascii=False, default=str) + "\n", encoding="utf-8")
    report = render_report(ticker, workbook, results, ctx.ai_model, bool(args.ml))
    report_path = run_dir / "agent_report.md"
    report_path.write_text(report, encoding="utf-8")

    print(f"[done] Agent report: {report_path}")
    qa = next((x for x in results if x.name == "Research QA / Skeptic Agent"), None)
    return 2 if args.strict and qa and qa.status == "FAIL" else 0


if __name__ == "__main__":
    raise SystemExit(main())
