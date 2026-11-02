from __future__ import annotations

from datetime import date, datetime, timedelta, timezone
import json
import sqlite3
from pathlib import Path

import pytest
from openpyxl import Workbook

from automation.check_ml_health import inspect
from automation.verify_agent_runs import verify
from agent_research import (
    AgentContext,
    FilingsAgent,
    KPIEarningsAgent,
    ResearchQAAgent,
    ThesisMonitorAgent,
)


def _ml_db(path: Path, *, last_market_date: date, recent_symbols=12, total_symbols=12):
    with sqlite3.connect(path) as con:
        con.executescript(
            "CREATE TABLE symbols(symbol TEXT PRIMARY KEY);"
            "CREATE TABLE prices(symbol TEXT,date TEXT,adj_close REAL,close REAL,fetched_at TEXT);"
            "CREATE TABLE features(symbol TEXT,as_of TEXT);"
        )
        con.executemany("INSERT INTO symbols VALUES(?)", [(f"S{i}",) for i in range(total_symbols)])
        stamp=datetime.now(timezone.utc).isoformat()
        con.executemany(
            "INSERT INTO prices VALUES(?,?,?,?,?)",
            [(f"S{i}", last_market_date.isoformat(), 100.0, 100.0, stamp) for i in range(recent_symbols)],
        )
        con.execute("INSERT INTO features VALUES(?,?)", ("S0", last_market_date.isoformat()))


def test_ml_health_passes_recent_complete_data(tmp_path):
    db=tmp_path/"history.sqlite"
    _ml_db(db,last_market_date=date.today()-timedelta(days=1))
    status=inspect(db,max_age_days=7,min_coverage=0.60)
    assert status["status"]=="PASS"
    assert status["recent_symbol_count"]==12


def test_ml_health_rejects_stale_prices(tmp_path):
    db=tmp_path/"history.sqlite"
    _ml_db(db,last_market_date=date.today()-timedelta(days=20))
    with pytest.raises(RuntimeError,match="stale"):
        inspect(db)


def test_ml_health_rejects_incomplete_coverage(tmp_path):
    db=tmp_path/"history.sqlite"
    _ml_db(db,last_market_date=date.today(),recent_symbols=4,total_symbols=12)
    with pytest.raises(RuntimeError,match="coverage"):
        inspect(db)


def test_five_agents_execute_on_finnish_esef_workbook(tmp_path):
    wb=Workbook()
    ws=wb.active
    ws.title="Filings"
    ws.append(["Form","Period","Filed","URL","Source"])
    ws.append(["ESEF AFR","2025-12-31","2026-03-01","https://example.test/esef.json","Finnish OAM / ESEF"])
    evidence=wb.create_sheet("AI Impact Analysis")
    evidence.append(["KPI / Evidence","Current","Unit / Comparison","Signal","Investment Read-through","As Of","Source URL","Data Type"])
    evidence.append(["Reported operating profitability","5.5","EUR m","positive","Improvement",date.today().isoformat(),"https://example.test/report","issuer"])
    filename=tmp_path/"YIT.HE_Equity_Research.xlsx"
    wb.save(filename)

    ctx=AgentContext(
        ticker="YIT.HE",repo_root=tmp_path,workbook_path=filename,
        run_dir=tmp_path/"research_runs"/"YIT.HE"/"latest",
    )
    agents=[FilingsAgent(),KPIEarningsAgent(),ThesisMonitorAgent(),ResearchQAAgent()]
    results=[agent.run(ctx) for agent in agents]
    assert results[0].status=="PASS"
    assert results[0].details["filings"][0]["form"]=="ESEF AFR"
    assert all(x.status in {"PASS","WARN"} for x in results)

    # Include a source-health result as the fifth contract. Source probes are tested
    # independently to avoid live HTTP calls inside the smoke suite.
    names=["Source Health Agent"]+[x.name for x in results]
    manifest={"ticker":"YIT.HE","results":[{"name":n,"status":"PASS"} for n in names]}
    location=tmp_path/"runs"/"YIT.HE"/"latest"
    location.mkdir(parents=True)
    (location/"manifest.json").write_text(json.dumps(manifest),encoding="utf-8")
    checked=verify(tmp_path/"runs",["YIT.HE"])
    assert checked["expected_agent_runs"]==5


def test_agent_verifier_rejects_missing_agent(tmp_path):
    folder=tmp_path/"YIT.HE"/"20261102"
    folder.mkdir(parents=True)
    (folder/"manifest.json").write_text(
        json.dumps({"ticker":"YIT.HE","results":[{"name":"Source Health Agent","status":"PASS"}]}),
        encoding="utf-8",
    )
    with pytest.raises(RuntimeError,match="expected research agents"):
        verify(tmp_path,["YIT.HE"])
