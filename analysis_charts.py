"""Create a native Excel Analysis Charts sheet from existing model outputs."""

from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.utils import get_column_letter

NAVY="17365D"; BLUE="2F75B5"; WHITE="FFFFFF"; LIGHT="F5F9FC"; PALE_BLUE="D9EAF7"; GREEN="008000"; GREY="666666"
FMT_PCT='0.0%;[Red](0.0%);-'; FMT_PRICE='$#,##0.00;[Red]($#,##0.00);-'; FMT_BN='#,##0.0;[Red](#,##0.0);-'; FMT_MULT='0.0x;[Red](0.0x);-'


def _fill(c): return PatternFill("solid",fgColor=c)
def _num(v): return float(v) if isinstance(v,(int,float)) else None

def _section(ws,rng,title):
    ws.merge_cells(rng); c=ws[rng.split(":")[0]]; c.value=title; c.fill=_fill(NAVY); c.font=Font(bold=True,color=WHITE,size=11)

def _header(ws,row,start,end):
    for c in range(start,end+1):
        x=ws.cell(row,c); x.fill=_fill(BLUE); x.font=Font(bold=True,color=WHITE); x.alignment=Alignment(horizontal="center",wrap_text=True)

def _card(ws,tr,vr,title,formula,fmt):
    ws.merge_cells(tr); ws.merge_cells(vr); t=ws[tr.split(":")[0]]; v=ws[vr.split(":")[0]]
    t.value=title; t.fill=_fill(BLUE); t.font=Font(bold=True,color=WHITE); t.alignment=Alignment(horizontal="center")
    v.value=formula; v.fill=_fill(LIGHT); v.font=Font(bold=True,color=GREEN,size=15); v.alignment=Alignment(horizontal="center",vertical="center"); v.number_format=fmt

def _find_row(ws,label):
    for r in range(1,ws.max_row+1):
        if str(ws.cell(r,1).value or "").strip()==label: return r
    return None

def _business_rows(wb):
    if "Segment Analysis" not in wb.sheetnames: return []
    s=wb["Segment Analysis"]; r=_find_row(s,"Revenue by Business Line")
    if not r: return []
    header=r+1; latest=4
    for c in range(1,min(s.max_column,15)+1):
        x=str(s.cell(header,c).value or "").lower()
        if x in {"latest","2025"} or "latest revenue" in x: latest=c; break
    out=[]
    for rr in range(header+1,min(s.max_row,header+20)+1):
        n=s.cell(rr,1).value; v=_num(s.cell(rr,latest).value)
        if n in (None,""):
            if out: break
            continue
        if v is not None: out.append((str(n),v))
    return out[:10]

def _margin_rows(wb):
    if "Segment Analysis" not in wb.sheetnames: return []
    s=wb["Segment Analysis"]; r=_find_row(s,"Reported Operating Segments")
    if not r: return []
    header=r+1; cols=[]
    for c in range(1,min(s.max_column,18)+1):
        x=str(s.cell(header,c).value or "")
        if "Margin" in x and "Δ" not in x: cols.append(c)
    if len(cols)<2: return []
    cols=cols[-3:]; out=[]
    for rr in range(header+1,min(s.max_row,header+12)+1):
        n=s.cell(rr,1).value
        if n in (None,""):
            if out: break
            continue
        vals=[_num(s.cell(rr,c).value) for c in cols]
        if sum(v is not None for v in vals)>=2 and all(v is None or -1<=v<=1 for v in vals): out.append((str(n),vals))
    return out[:5]


def ensure_analysis_charts(wb,ticker):
    required={"Company Data","Historical Financials","Three-Case Scenarios","DCF"}
    if not required.issubset(wb.sheetnames): return None
    if "Analysis Charts" in wb.sheetnames: wb.remove(wb["Analysis Charts"])
    ws=wb.create_sheet("Analysis Charts"); ws.sheet_view.showGridLines=False
    try:
        wb._sheets.remove(ws); wb._sheets.insert(min(2,len(wb._sheets)),ws)
    except Exception: pass

    ws.merge_cells("A1:P2"); ws["A1"]=f"{ticker} — Analysis Charts"; ws["A1"].fill=_fill(NAVY); ws["A1"].font=Font(bold=True,color=WHITE,size=18)
    ws.merge_cells("A3:P3"); ws["A3"]="Automated Monte Carlo, scenario, stress, history, segment and valuation visualizations."; ws["A3"].font=Font(italic=True,color=GREY)
    for c in range(1,17): ws.column_dimensions[get_column_letter(c)].width=11

    company=wb["Company Data"]; hist=wb["Historical Financials"]; sc=wb["Three-Case Scenarios"]; dcf=wb["DCF"]; adv=wb["Advanced Analytics"] if "Advanced Analytics" in wb.sheetnames else None
    price=_num(company["B8"].value)

    # Helpers: X:Y Monte Carlo, AA:AB scenarios, AD:AE stress, AG:AK history, AM:AN business, AP:AS margins, AU:AV P/E.
    mc_end=2
    if adv:
        ws["X2"]="Value / Share"; ws["Y2"]="Frequency"
        for sr in range(33,53):
            v=_num(adv.cell(sr,18).value); f=_num(adv.cell(sr,19).value)
            if v is not None and f is not None:
                mc_end+=1; ws.cell(mc_end,24,v); ws.cell(mc_end,25,f); ws.cell(mc_end,24).number_format=FMT_PRICE

    ws["AA2"]="Scenario"; ws["AB2"]="Value / Share"
    for r,(n,v) in enumerate([("Bear",_num(sc["B39"].value)),("Base",_num(sc["C39"].value)),("Bull",_num(sc["D39"].value)),("Probability Weighted",_num(sc["E39"].value)),("Current Price",price)],3):
        ws.cell(r,27,n); ws.cell(r,28,v); ws.cell(r,28).number_format=FMT_PRICE

    ws["AD2"]="Stress Test"; ws["AE2"]="Value / Share"; stress_end=2
    for sr in range(48,59):
        n=sc.cell(sr,1).value; v=_num(sc.cell(sr,7).value)
        if n not in (None,"") and v is not None:
            stress_end+=1; ws.cell(stress_end,30,str(n)); ws.cell(stress_end,31,v); ws.cell(stress_end,31).number_format=FMT_PRICE

    for c,x in enumerate(["Year","Revenue","Free Cash Flow","Operating Margin","FCF Margin"],33): ws.cell(2,c,x)
    hist_end=2
    for c in range(2,8):
        y=hist.cell(3,c).value; rev=_num(hist.cell(4,c).value)
        if y is None or rev is None: continue
        hist_end+=1; vals=[y,rev,_num(hist.cell(16,c).value),_num(hist.cell(10,c).value),_num(hist.cell(17,c).value)]
        for cc,v in enumerate(vals,33): ws.cell(hist_end,cc,v)
        ws.cell(hist_end,34).number_format=FMT_BN; ws.cell(hist_end,35).number_format=FMT_BN; ws.cell(hist_end,36).number_format=FMT_PCT; ws.cell(hist_end,37).number_format=FMT_PCT

    ws["AM2"]="Business Line"; ws["AN2"]="Latest Revenue"; business_end=2
    for n,v in _business_rows(wb): business_end+=1; ws.cell(business_end,39,n); ws.cell(business_end,40,v); ws.cell(business_end,40).number_format=FMT_BN

    ws["AP2"]="Segment"; ws["AQ2"]="Year -2 Margin"; ws["AR2"]="Year -1 Margin"; ws["AS2"]="Latest Margin"; margin_end=2
    for n,vals in _margin_rows(wb):
        margin_end+=1; ws.cell(margin_end,42,n)
        for cc,v in enumerate(vals[-3:],43): ws.cell(margin_end,cc,v); ws.cell(margin_end,cc).number_format=FMT_PCT

    pe_end=2
    if adv:
        ws["AU2"]="Year"; ws["AV2"]="Year-End P/E"
        for sr in range(7,13):
            y=adv.cell(sr,1).value; pe=_num(adv.cell(sr,4).value)
            if y is not None and pe is not None: pe_end+=1; ws.cell(pe_end,47,y); ws.cell(pe_end,48,pe); ws.cell(pe_end,48).number_format=FMT_MULT

    # Monte Carlo
    _section(ws,"A5:H5","Monte Carlo Valuation")
    if mc_end>=4:
        ch=BarChart(); ch.type="col"; ch.style=10; ch.title="Monte Carlo Valuation Distribution — 5,000 Simulations"; ch.height=8; ch.width=13.5; ch.legend=None
        ch.add_data(Reference(ws,min_col=25,min_row=2,max_row=mc_end),titles_from_data=True); ch.set_categories(Reference(ws,min_col=24,min_row=3,max_row=mc_end)); ch.x_axis.numFmt="$0"; ws.add_chart(ch,"A7")
    else: ws["A7"]="Monte Carlo histogram unavailable in this run."

    # Scenario
    _section(ws,"I5:P5","Scenario Valuation"); ch=BarChart(); ch.type="col"; ch.style=10; ch.title="Bear / Base / Bull vs Current Price"; ch.height=8; ch.width=13.5; ch.legend=None
    ch.add_data(Reference(ws,min_col=28,min_row=2,max_row=7),titles_from_data=True); ch.set_categories(Reference(ws,min_col=27,min_row=3,max_row=7)); ch.y_axis.numFmt="$0"; ws.add_chart(ch,"I7")

    # Stress
    _section(ws,"A27:H27","Stress Testing")
    if stress_end>=4:
        ch=BarChart(); ch.type="bar"; ch.style=10; ch.title="Stress-Test Intrinsic Value / Share"; ch.height=9; ch.width=13.5; ch.legend=None
        ch.add_data(Reference(ws,min_col=31,min_row=2,max_row=stress_end),titles_from_data=True); ch.set_categories(Reference(ws,min_col=30,min_row=3,max_row=stress_end)); ch.x_axis.numFmt="$0"; ws.add_chart(ch,"A29")

    # History
    _section(ws,"I27:P27","Historical Financial Performance")
    if hist_end>=4:
        ch=LineChart(); ch.style=10; ch.title="Revenue & Free Cash Flow"; ch.height=9; ch.width=13.5; ch.legend.position="b"
        ch.add_data(Reference(ws,min_col=34,max_col=35,min_row=2,max_row=hist_end),titles_from_data=True); ch.set_categories(Reference(ws,min_col=33,min_row=3,max_row=hist_end)); ws.add_chart(ch,"I29")

    # Business mix
    _section(ws,"A51:H51","Business Mix")
    if business_end>=4:
        ch=BarChart(); ch.type="bar"; ch.style=10; ch.title="Latest Revenue by Business Line"; ch.height=8.5; ch.width=13.5; ch.legend=None
        ch.add_data(Reference(ws,min_col=40,min_row=2,max_row=business_end),titles_from_data=True); ch.set_categories(Reference(ws,min_col=39,min_row=3,max_row=business_end)); ch.x_axis.numFmt="$0"; ws.add_chart(ch,"A53")
    else: ws["A53"]="No disclosed business-line revenue table was available."

    # Segment margins
    _section(ws,"I51:P51","Segment Profitability")
    if margin_end>=4:
        ch=BarChart(); ch.type="col"; ch.style=10; ch.title="Latest Segment Operating Margins"; ch.height=8.5; ch.width=13.5; ch.legend=None
        ch.add_data(Reference(ws,min_col=45,min_row=2,max_row=margin_end),titles_from_data=True); ch.set_categories(Reference(ws,min_col=42,min_row=3,max_row=margin_end)); ch.y_axis.numFmt="0%"; ws.add_chart(ch,"I53")
    else: ws["I53"]="No comparable segment-margin history was available."

    # Historical P/E
    _section(ws,"A73:H73","Historical Valuation")
    if pe_end>=4:
        ch=LineChart(); ch.style=10; ch.title="Year-End P/E"; ch.height=8; ch.width=13.5; ch.legend=None
        ch.add_data(Reference(ws,min_col=48,min_row=2,max_row=pe_end),titles_from_data=True); ch.set_categories(Reference(ws,min_col=47,min_row=3,max_row=pe_end)); ch.y_axis.numFmt="0.0x"; ws.add_chart(ch,"A75")
    else: ws["A75"]="Historical valuation series unavailable."

    # DCF sensitivity heatmap
    _section(ws,"I73:P73","DCF Sensitivity"); tgr=[.02,.025,.03,.035,.04]; ws["I75"]="WACC / TGR"
    for c,v in enumerate(tgr,10): ws.cell(75,c,v); ws.cell(75,c).number_format=FMT_PCT
    for rr,sr in enumerate(range(22,27),76):
        ws.cell(rr,9,_num(dcf.cell(sr,1).value)); ws.cell(rr,9).number_format=FMT_PCT
        for cc,sc_col in enumerate(range(2,7),10): ws.cell(rr,cc,_num(dcf.cell(sr,sc_col).value)); ws.cell(rr,cc).number_format=FMT_PRICE
    _header(ws,75,9,14)
    for r in range(76,81): ws.cell(r,9).fill=_fill(PALE_BLUE); ws.cell(r,9).font=Font(bold=True)
    ws.conditional_formatting.add("J76:N80",ColorScaleRule(start_type="min",start_color="F8696B",mid_type="percentile",mid_value=50,mid_color="FFEB84",end_type="max",end_color="63BE7B"))
    ws.merge_cells("I83:P87"); ws["I83"]="Monte Carlo shows the valuation distribution; scenarios and stress tests are deterministic cases; the DCF heatmap shows sensitivity to WACC and terminal growth."; ws["I83"].fill=_fill(LIGHT); ws["I83"].alignment=Alignment(wrap_text=True,vertical="top")

    if adv:
        _card(ws,"A95:D95","A96:D99","Monte Carlo P10","='Advanced Analytics'!J33",FMT_PRICE)
        _card(ws,"E95:H95","E96:H99","Monte Carlo Median","='Advanced Analytics'!J35",FMT_PRICE)
        _card(ws,"I95:L95","I96:L99","Probability > Current Price","='Advanced Analytics'!J38",FMT_PCT)
        _card(ws,"M95:P95","M96:P99","Reverse DCF Implied FCF CAGR","='Advanced Analytics'!B38",FMT_PCT)

    for c in range(24,49): ws.column_dimensions[get_column_letter(c)].hidden=True
    ws.freeze_panes="A5"
    return ws
