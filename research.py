"""Run the agent-assisted equity research pipeline.

Examples:
    python research.py GOOGL
    python research.py GOOGL --ai
    python research.py GOOGL --skip-model --ai

The existing update_model.py remains the deterministic source of workbook calculations.
Specialist agents inspect the generated workbook, maintain KPI history, monitor thesis
evidence, check registered data sources, and run research QA. LLM reasoning is opt-in.
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

BASE = Path(__file__).resolve().parent
UPDATED_MODELS = BASE / "updated_models"
RUNS_DIR = BASE / "research_runs"


def valid_ticker(raw: str) -> str:
    ticker = raw.upper().strip()
    if not re.fullmatch(r"[A-Z0-9.\-]{1,10}", ticker):
        raise argparse.ArgumentTypeError("Enter only a ticker, e.g. GOOGL, MSFT, NVDA, or TSM.")
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
    print(f"[model] Building deterministic research workbook for {ticker}...")
    subprocess.run([sys.executable, str(BASE / "update_model.py"), ticker], cwd=BASE, check=True)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Agent-assisted equity research orchestrator")
    parser.add_argument("ticker", type=valid_ticker)
    parser.add_argument("--skip-model", action="store_true", help="Reuse the newest existing workbook instead of running update_model.py")
    parser.add_argument("--ai", action="store_true", help="Enable evidence-bound OpenAI reasoning. Requires OPENAI_API_KEY.")
    parser.add_argument("--model", help="OpenAI model override. Otherwise uses OPENAI_RESEARCH_MODEL or the project default.")
    parser.add_argument("--strict", action="store_true", help="Return a non-zero exit code when Research QA fails.")
    return parser


def render_report(ticker: str, workbook: Path, results: list, ai_model: str | None) -> str:
    lines = [
        f"# {ticker} Agent Research Report",
        "",
        f"Generated: {datetime.now().astimezone().isoformat(timespec='seconds')}",
        f"Workbook: `{workbook}`",
        f"AI reasoning: {'enabled (' + ai_model + ')' if ai_model else 'disabled; deterministic agent checks only'}",
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
        "- The agents do not execute trades.",
        "- AI narrative does not overwrite financial inputs or DCF assumptions.",
        "- Reported facts, calculations, and inference remain separate.",
        "- Market-share records preserve their market definition and source.",
        "- Valuation changes remain analyst-reviewed and are calculated by the existing deterministic model.",
        "",
    ])
    return "\n".join(lines)


def main() -> int:
    args = build_parser().parse_args()
    ticker = args.ticker
    if not args.skip_model:
        run_model(ticker)
    workbook = latest_workbook(ticker)

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

    manifest = {
        "ticker": ticker,
        "workbook": str(workbook),
        "ai_model": ctx.ai_model,
        "results": [x.to_dict() for x in results],
    }
    (run_dir / "manifest.json").write_text(json.dumps(manifest, indent=2, ensure_ascii=False, default=str) + "\n", encoding="utf-8")
    report = render_report(ticker, workbook, results, ctx.ai_model)
    report_path = run_dir / "agent_report.md"
    report_path.write_text(report, encoding="utf-8")

    print(f"[done] Agent report: {report_path}")
    qa = next((x for x in results if x.name == "Research QA / Skeptic Agent"), None)
    return 2 if args.strict and qa and qa.status == "FAIL" else 0


if __name__ == "__main__":
    raise SystemExit(main())
