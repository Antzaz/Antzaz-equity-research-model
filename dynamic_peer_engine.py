"""Dynamic, sector-safe comparable-company selection.

The target company's sector and industry are read from yfinance Ticker.info. Candidate
companies are discovered dynamically from yfinance Industry/Sector domain objects and
EquityQuery screening. Exact-industry companies are preferred; same-sector companies
may fill remaining slots. Cross-sector candidates are always rejected.

Hard-coded lists are only a last-resort discovery seed and are still validated against
the target sector before inclusion. A failed discovery clears stale peers rather than
silently retaining template companies.
"""

import math
try:
    import yfinance as yf
    from yfinance import EquityQuery
except Exception:
    yf = None
    EquityQuery = None

from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

NAVY="17365D"; BLUE="2F75B5"; WHITE="FFFFFF"; GREY="666666"
INPUT_BLUE="0000FF"; LINK_GREEN="008000"; PALE_GREEN="E2F0D9"; GOLD="FFF2CC"
FMT_PCT='0.0%;[Red](0.0%);-'; FMT_MULT='0.0x;[Red](0.0x);-'
THIN=Side(style="thin",color="D9E1F2")

FALLBACK={
    "Industrials":["ETN","PH","CMI","EMR","ROK","ITW","AME","IR","CAT","HON","DE","GE"],
    "Healthcare":["UNH","ELV","CI","HUM","CNC","MOH","JNJ","LLY","ABBV","MRK","TMO","ABT"],
    "Technology":["MSFT","AAPL","NVDA","AVGO","ORCL","CRM","ADBE","AMD","QCOM","NOW","INTU","TXN"],
    "Communication Services":["GOOGL","META","NFLX","DIS","CMCSA","TMUS","VZ","T","SPOT","PINS"],
    "Consumer Cyclical":["AMZN","TSLA","HD","MCD","LOW","BKNG","TJX","NKE","SBUX","MELI"],
    "Consumer Defensive":["WMT","COST","PG","KO","PEP","PM","TGT","CL","MDLZ"],
    "Financial Services":["JPM","BAC","WFC","C","GS","MS","USB","PNC","BLK","SCHW","AXP"],
    "Energy":["XOM","CVX","COP","EOG","SLB","MPC","PSX","OXY","VLO"],
    "Basic Materials":["LIN","APD","SHW","FCX","NEM","NUE","DOW","ECL"],
    "Real Estate":["PLD","AMT","EQIX","WELL","SPG","O","DLR","PSA"],
    "Utilities":["NEE","SO","DUK","CEG","AEP","SRE","D","EXC"],
}

COUNTRY_TO_REGION={
    "United States":"US","Canada":"CA","Finland":"FI","Sweden":"SE","Norway":"NO","Denmark":"DK",
    "Germany":"DE","France":"FR","United Kingdom":"GB","Switzerland":"CH","Netherlands":"NL","Japan":"JP",
    "Australia":"AU","China":"CN","Hong Kong":"HK","India":"IN","Singapore":"SG","Italy":"IT","Spain":"ES",
}

def _fill(c): return PatternFill("solid",fgColor=c)
def _num(v,default=None):
    try:
        if isinstance(v,bool): return default
        return float(v)
    except Exception: return default

def _info(symbol):
    if yf is None: return {}
    try: return yf.Ticker(symbol).info or {}
    except Exception: return {}

def _classification(wb,ticker,target):
    sector=target.get("sector"); industry=target.get("industry"); mc=_num(target.get("marketCap")); country=target.get("country")
    if "Company Data" in wb.sheetnames:
        d=wb["Company Data"]
        sector=sector or d["B6"].value; industry=industry or d["B7"].value
        if not mc:
            bn=_num(d["B10"].value); mc=bn*1e9 if bn else None
    return str(sector or "Unknown"),str(industry or "Unknown"),mc,COUNTRY_TO_REGION.get(str(country or ""),"US")

def _symbols_from_frame(obj):
    out=[]
    if obj is None: return out
    try:
        cols=[str(c).lower() for c in obj.columns]
        for key in ("symbol","ticker"):
            if key in cols:
                col=obj.columns[cols.index(key)]
                out.extend(str(x).upper() for x in obj[col].tolist() if x)
        for x in list(obj.index):
            s=str(x).upper().strip()
            if 1<=len(s)<=12 and any(ch.isalpha() for ch in s): out.append(s)
    except Exception: pass
    return out

def _symbols_from_screen(result):
    if not isinstance(result,dict): return []
    rows=result.get("quotes") or result.get("results") or []
    out=[]
    if isinstance(rows,list):
        for row in rows:
            if isinstance(row,dict):
                s=row.get("symbol") or row.get("ticker")
                if s: out.append(str(s).upper())
    return out

def _dynamic_candidates(target,sector,industry,region):
    if yf is None: return []
    out=[]
    industry_key=target.get("industryKey")
    sector_key=target.get("sectorKey")
    if industry_key:
        try:
            dom=yf.Industry(industry_key,region=region)
            for attr in ("top_companies","top_performing_companies","top_growth_companies"):
                out.extend(_symbols_from_frame(getattr(dom,attr,None)))
        except Exception: pass
    if EquityQuery is not None and sector!="Unknown" and industry!="Unknown":
        try:
            q=EquityQuery('and',[
                EquityQuery('eq',['region',region.lower()]),
                EquityQuery('eq',['sector',sector]),
                EquityQuery('eq',['industry',industry]),
            ])
            out.extend(_symbols_from_screen(yf.screen(q,size=100,sortField='intradaymarketcap',sortAsc=False)))
        except Exception: pass
    if sector_key:
        try: out.extend(_symbols_from_frame(yf.Sector(sector_key,region=region).top_companies))
        except Exception: pass
    if EquityQuery is not None and sector!="Unknown":
        try:
            q=EquityQuery('and',[
                EquityQuery('eq',['region',region.lower()]),
                EquityQuery('eq',['sector',sector]),
            ])
            out.extend(_symbols_from_screen(yf.screen(q,size=100,sortField='intradaymarketcap',sortAsc=False)))
        except Exception: pass
    # Last-resort discovery seed. Validation below still enforces same sector.
    out.extend(FALLBACK.get(sector,[]))
    seen=[]
    for s in out:
        s=str(s).upper().strip()
        if s and s not in seen: seen.append(s)
    return seen

def _rank(target_sector,target_industry,target_mc,symbol,info):
    sector=str(info.get("sector") or "")
    if sector!=target_sector: return None
    industry=str(info.get("industry") or "")
    class_penalty=0 if industry==target_industry else 100
    mc=_num(info.get("marketCap")); size_penalty=abs(math.log(mc/target_mc)) if target_mc and mc and mc>0 and target_mc>0 else 4
    coverage=sum(info.get(k) not in (None,"") for k in ("forwardPE","enterpriseToRevenue","enterpriseToEbitda","revenueGrowth","operatingMargins","returnOnEquity"))
    return class_penalty+size_penalty+(6-coverage)*.25

def select_dynamic_peers(wb,ticker,count=5):
    target=_info(ticker); sector,industry,target_mc,region=_classification(wb,ticker,target)
    ranked=[]
    for symbol in _dynamic_candidates(target,sector,industry,region):
        if symbol==ticker.upper(): continue
        info=_info(symbol); score=_rank(sector,industry,target_mc,symbol,info)
        if score is not None: ranked.append((score,symbol,info))
    ranked.sort(key=lambda x:x[0])
    return target,sector,industry,[(s,i) for _,s,i in ranked[:count]]

def _metric_row(symbol,info,sector,industry,method):
    return [
        info.get("longName") or symbol,symbol,info.get("forwardPE"),info.get("enterpriseToRevenue"),
        info.get("enterpriseToEbitda"),info.get("revenueGrowth"),info.get("operatingMargins"),info.get("returnOnEquity"),
        info.get("sector") or sector,info.get("industry") or industry,method,f"https://finance.yahoo.com/quote/{symbol}/",
    ]

def _header(ws,row,start,end):
    for c in range(start,end+1):
        x=ws.cell(row,c); x.fill=_fill(BLUE); x.font=Font(bold=True,color=WHITE); x.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True); x.border=Border(bottom=THIN)

def _repair_comparative(wb,ticker,sector,industry,peer_count):
    if "Comparative Analysis" not in wb.sheetnames: return
    ws=wb["Comparative Analysis"]; ws["A1"]=f"Comparative Analysis — {ticker} vs {industry if industry!='Unknown' else sector} Peers"
    specs=[("Forward P/E","C","Lower","Lower is better"),("EV/Revenue","D","Lower","Lower is better"),("EV/EBITDA","E","Lower","Lower is better"),("Revenue Growth","F","Higher","Higher is better"),("Operating Margin","G","Higher","Higher is better"),("ROE","H","Higher","Higher is better")]
    last=4+max(1,peer_count)
    for r,(label,col,direction,note) in enumerate(specs,4):
        ws.cell(r,1,label); ws.cell(r,2,f"='Peer Comps'!{col}4"); ws.cell(r,3,f'=IFERROR(MEDIAN(\'Peer Comps\'!{col}5:{col}{last}),"")'); ws.cell(r,4,f'=IFERROR(B{r}/C{r}-1,"")'); ws.cell(r,5,direction)
        ws.cell(r,6,f'=IF(D{r}="","",IF(D{r}{"<" if direction=="Lower" else ">"}0,"{"Attractive" if direction=="Lower" else "Better"}","{"Premium" if direction=="Lower" else "Worse"}"))'); ws.cell(r,7,note)
        ws.cell(r,2).font=Font(color=LINK_GREEN); ws.cell(r,3).font=Font(color=LINK_GREEN); ws.cell(r,4).font=Font(color="000000")
        for c in (2,3): ws.cell(r,c).number_format=FMT_MULT if r<=6 else FMT_PCT
        ws.cell(r,4).number_format=FMT_PCT
    # Do not manufacture relative price targets by mixing forward peer multiples with trailing target-company denominators.
    ws["A13"]="Method"; ws["B13"]="Implied Value / Share"; ws["C13"]="Comment"
    notes=[
        ("Peer Forward P/E",None,"Not calculated unless matching-period forward EPS is available for the target company."),
        ("Peer EV/Revenue",None,"Not used as a standalone price target without comparable forward revenue, margins and capital intensity."),
        ("Peer EV/EBITDA",None,"Not used as a standalone price target until comparable forward EBITDA is available."),
    ]
    for r,row in enumerate(notes,14):
        for c,v in enumerate(row,1): ws.cell(r,c,v)
        ws.cell(r,3).alignment=Alignment(wrap_text=True)
    ws["A18"]="Method note"; ws["B18"]="Rows 4–9 are like-for-like current public metric comparisons. Cross-sector peers are prohibited and unsupported price-target conversions remain blank."
    ws["B18"].alignment=Alignment(wrap_text=True)

def ensure_dynamic_peer_comps(wb,ticker,count=5):
    if "Peer Comps" not in wb.sheetnames: return []
    target,sector,industry,peers=select_dynamic_peers(wb,ticker,count); ws=wb["Peer Comps"]
    for row in ws.iter_rows(min_row=1,max_row=max(12,ws.max_row),min_col=1,max_col=12):
        for cell in row: cell.value=None
    for c in range(1,13): ws.cell(1,c).fill=_fill(NAVY); ws.cell(2,c).fill=_fill(NAVY)
    ws["A1"]=f"{industry if industry!='Unknown' else sector} — Industry / Sector Peer Comps"; ws["A1"].font=Font(bold=True,color=WHITE,size=18)
    ws["A2"]="Target classification is detected automatically from the company. Exact-industry peers are preferred; same-sector fallback only. Cross-sector peers are excluded."
    ws["A2"].font=Font(italic=True,color=GREY); ws["A2"].alignment=Alignment(wrap_text=True)
    headers=["Company","Ticker","Forward P/E","EV/Revenue","EV/EBITDA","Revenue Growth","Operating Margin","ROE","Sector","Industry","Discovery","Source URL"]
    for c,v in enumerate(headers,1): ws.cell(3,c,v)
    _header(ws,3,1,12)
    rows=[_metric_row(ticker,target,sector,industry,"Target classification")]
    for symbol,info in peers:
        method="Exact industry" if str(info.get("industry") or "")==industry else "Same-sector fallback"
        rows.append(_metric_row(symbol,info,sector,industry,method))
    for r,row in enumerate(rows,4):
        for c,v in enumerate(row,1): ws.cell(r,c,v)
        for c in range(3,11): ws.cell(r,c).font=Font(color=INPUT_BLUE)
        ws.cell(r,12).font=Font(color=LINK_GREEN)
        for c in (3,4,5): ws.cell(r,c).number_format=FMT_MULT
        for c in (6,7,8): ws.cell(r,c).number_format=FMT_PCT
    if not peers:
        ws["A5"]="REVIEW — no validated same-sector peers were returned. Stale template peers were removed rather than reused."; ws["A5"].fill=_fill(GOLD); ws["A5"].font=Font(color=INPUT_BLUE,bold=True)
    ws["A11"]="Method"; ws["B11"]="yfinance Ticker.info classification + Industry/Sector domain discovery + EquityQuery screen; every peer is revalidated to the target sector."
    widths={"A":31,"B":10,"C":13,"D":13,"E":13,"F":15,"G":16,"H":13,"I":18,"J":28,"K":20,"L":48}
    for col,w in widths.items(): ws.column_dimensions[col].width=w
    ws.freeze_panes="A4"
    _repair_comparative(wb,ticker,sector,industry,len(peers))
    return peers
