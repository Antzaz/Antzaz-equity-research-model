from __future__ import annotations
from pathlib import Path
from openpyxl import Workbook
from offline_equity_extensions import apply_offline_research_extensions

def _synthetic_workbook():
    wb=Workbook(); ws=wb.active; ws.title="Company Data"
    ws["A8"]="Current Price"; ws["B8"]=100.0; ws["A9"]="Shares Outstanding"; ws["B9"]=10.0; ws["A12"]="Cash"; ws["B12"]=20.0; ws["A13"]="Debt"; ws["B13"]=5.0
    h=wb.create_sheet("Historical Financials"); h["A3"]="Metric"
    for i,y in enumerate([2022,2023,2024,2025],2): h.cell(3,i,y)
    for r,(lab,vals) in {4:("Revenue",[100,110,125,140]),9:("Operating Income",[20,23,28,32]),11:("Net Income",[15,17,20,24]),12:("Diluted EPS",[1.5,1.7,2.0,2.4]),14:("Operating Cash Flow",[25,28,31,35]),15:("Capital Expenditures",[-8,-9,-10,-11]),16:("Diluted Shares",[10.5,10.3,10.1,10.0])}.items():
        h.cell(r,1,lab)
        for c,v in enumerate(vals,2): h.cell(r,c,v)
    s=wb.create_sheet("Three-Case Scenarios")
    for i,y in enumerate([2026,2027,2028],14): s.cell(10,i,y)
    s["A11"]="Revenue"; s["N11"]=155; s["O11"]=170; s["P11"]=185
    s["A12"]="Free Cash Flow"; s["N12"]=28; s["O12"]=31; s["P12"]=35
    e=wb.create_sheet("Expectations & Consensus"); e.append(["Metric","Year","Consensus","Model"]); e.append(["Revenue",2026,150,155]); e.append(["EPS",2026,2.6,2.8])
    d=wb.create_sheet("Decision View"); d.append(["Current Market Price",100]); d.append(["Base DCF Fair Value",125]); d.append(["Base DCF Upside",.25]); d.append(["Market-Implied 10Y FCF CAGR",.08]); d.append(["MODEL VIEW","POTENTIALLY ATTRACTIVE"])
    p=wb.create_sheet("Peer Comps")
    for c,hdr in enumerate(["Name","Ticker","Forward P/E","EV/Revenue","EV/EBITDA","Revenue Growth","Operating Margin","ROE"],1): p.cell(3,c,hdr)
    p.cell(3,16,"Peer Type")
    for c,v in enumerate(["Demo","DEMO",20,5,15,.12,.23,.18],1): p.cell(4,c,v)
    p.cell(4,16,"Target classification")
    seg=wb.create_sheet("Segment Analysis"); seg.append(["Segment",2024,2025]); seg.append(["Core",80,90]); seg.append(["Growth",30,40])
    q=wb.create_sheet("Data Quality"); q.append(["Control","Status"]); q.append(["Core","PASS"])
    return wb

def test_offline_extensions_create_all_sheets_and_histories(tmp_path: Path):
    wb=_synthetic_workbook(); result=apply_offline_research_extensions(wb,"DEMO",{},tmp_path)
    expected={"Forecast Accountability","Earnings & Revisions","Capital Allocation","Valuation History","SOTP","Catalyst Timeline"}
    assert expected.issubset(set(wb.sheetnames))
    assert all(v["status"]=="PASS" for v in result.values())
    rdir=tmp_path/"research_data"/"DEMO"
    for name in ["forecast_snapshots.csv","estimate_revision_history.csv","capital_allocation_history.csv","valuation_history.csv","catalyst_timeline.csv"]: assert (rdir/name).exists()

def test_second_run_builds_revision_and_valuation_history(tmp_path: Path):
    import pandas as pd
    wb=_synthetic_workbook(); apply_offline_research_extensions(wb,"DEMO",{},tmp_path)
    wb["Expectations & Consensus"]["C2"]=165; wb["Decision View"]["B2"]=135
    apply_offline_research_extensions(wb,"DEMO",{},tmp_path)
    rev=pd.read_csv(tmp_path/"research_data"/"DEMO"/"estimate_revision_history.csv")
    val=pd.read_csv(tmp_path/"research_data"/"DEMO"/"valuation_history.csv")
    assert len(rev)>=4 and len(val)>=2 and rev["ConsensusRevision"].notna().any()
