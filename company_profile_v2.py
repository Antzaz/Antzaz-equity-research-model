from __future__ import annotations

"""Enrich Company Data with an investor-useful business description and main products/services.

The product table is intentionally descriptive rather than financial: revenue is shown only in
Segment Analysis when the issuer discloses it. Curated issuer-backed rows are used for companies
where the project has verified current product taxonomies; all other tickers fall back to reported
segment/business-line names plus the company profile summary.
"""

from openpyxl.styles import Alignment, Font, PatternFill

NAVY="17365D"; BLUE="2F75B5"; WHITE="FFFFFF"; LIGHT="F5F9FC"; GREY="666666"

CURATED_PRODUCTS={
    "GEV":{
        "description":"GE Vernova supplies equipment, services and software used to generate, transfer, orchestrate, convert and store electricity across Power, Wind and Electrification.",
        "rows":[
            ("Power","Heavy-duty and aeroderivative gas turbines","Gas-fired generation equipment plus installed-base services and parts.","https://www.gevernova.com/investors"),
            ("Power","Steam, nuclear and hydro power equipment & services","Generation technologies and lifecycle services, including nuclear and hydro platforms.","https://www.gevernova.com/investors/annual-report"),
            ("Wind","Onshore wind turbines & services","Utility-scale onshore turbines, installed-base service and lifecycle support.","https://www.gevernova.com/investors"),
            ("Wind","Offshore wind turbines","Large-scale offshore wind generation equipment and related services.","https://www.gevernova.com/investors"),
            ("Electrification","Transformers, switchgear, HVDC/substations & synchronous condensers","Grid equipment used to transmit, stabilize and modernize electricity networks.","https://www.gevernova.com/investors/annual-report/ceo-letter"),
            ("Electrification","Power conversion, storage and grid software","Power electronics, storage and software/digital orchestration used across modern grids.","https://www.gevernova.com/investors"),
        ],
    },
    "CEG":{
        "description":"Constellation combines large-scale electricity generation with wholesale/retail energy supply and customer energy solutions. Its fleet includes nuclear and, following Calpine, substantial natural-gas and geothermal generation.",
        "rows":[
            ("Generation","Nuclear electricity generation","Largest U.S. nuclear fleet; 24/7 emissions-free generation and plant operations.","https://www.constellationenergy.com/work/generation/nuclear.html"),
            ("Generation / Calpine","Natural-gas generation","Dispatchable gas-fired generation, including the Calpine fleet acquired in 2026.","https://www.constellationenergy.com/our-work/what-we-do/generation.html"),
            ("Generation","Hydro, wind, solar and geothermal generation","Diversified zero-/low-carbon generation portfolio alongside nuclear and gas.","https://www.constellationenergy.com/our-work/what-we-do/generation.html"),
            ("Retail / Commercial","Electricity supply","Fixed, index and managed electricity procurement solutions for residential, commercial and public-sector customers.","https://www.constellationenergy.com/work/commercial.html"),
            ("Retail / Commercial","Natural-gas supply","Customized natural-gas procurement and risk-management strategies.","https://www.constellationenergy.com/work/commercial.html"),
            ("Customer Solutions","Energy efficiency, offsite renewables & hourly carbon-free matching","Energy-management and clean-energy products designed to lower cost and support sustainability goals.","https://www.constellationenergy.com/work/commercial.html"),
            ("Generation Services","Constellation Generation Solutions & PowerLabs","Nuclear outage/maintenance/technical services plus calibration, testing and quality services.","https://www.constellationenergy.com/our-work/what-we-do/generation.html"),
        ],
    },
}

IGNORE_SEGMENT_LABELS={
    "segment","business line","metric","total","company total","income statement","balance sheet",
    "cash flow statement","reported segments","segment analysis","manual segment input",
}


def _fill(color): return PatternFill("solid",fgColor=color)


def _segment_fallback(wb,limit=6):
    if "Segment Analysis" not in wb.sheetnames: return []
    ws=wb["Segment Analysis"]; out=[]; seen=set()
    for r in range(1,ws.max_row+1):
        name=str(ws.cell(r,1).value or "").strip()
        if not name or name.lower() in IGNORE_SEGMENT_LABELS or name.lower().startswith(("source","note","method")):
            continue
        numeric=sum(isinstance(ws.cell(r,c).value,(int,float)) for c in range(2,min(ws.max_column,9)+1))
        if numeric==0: continue
        key=name.lower()
        if key in seen: continue
        seen.add(key)
        source=""
        for c in range(ws.max_column,1,-1):
            v=ws.cell(r,c).value
            if isinstance(v,str) and v.startswith("http"):
                source=v; break
        out.append(("Reported business / segment",name,"Issuer-disclosed business line; see Segment Analysis for available economics.",source))
        if len(out)>=limit: break
    return out


def _summary(info):
    text=str((info or {}).get("longBusinessSummary") or "").strip()
    if not text: return ""
    # Two sentences is enough context for Company Data; detailed narrative belongs elsewhere.
    parts=[x.strip() for x in text.replace("\n"," ").split(". ") if x.strip()]
    return ". ".join(parts[:2])[:900].rstrip(". ")+"."


def _clear_area(ws):
    for r in range(17,46):
        for c in range(1,5):
            ws.cell(r,c).value=None
            ws.cell(r,c).fill=PatternFill(fill_type=None)
            ws.cell(r,c).font=Font()
            ws.cell(r,c).alignment=Alignment()


def enrich_company_data(wb,ticker,info=None):
    if "Company Data" not in wb.sheetnames: return {"products":0,"source":"missing"}
    ws=wb["Company Data"]; t=str(ticker).upper().strip(); _clear_area(ws)
    curated=CURATED_PRODUCTS.get(t)
    description=(curated or {}).get("description") or _summary(info)
    rows=list((curated or {}).get("rows") or [])
    if not rows: rows=_segment_fallback(wb)
    website=str((info or {}).get("website") or "")

    for c in range(1,5):
        ws.cell(17,c).fill=_fill(NAVY); ws.cell(17,c).font=Font(bold=True,color=WHITE)
    ws.cell(17,1,"Business Overview & Main Products / Services")
    ws.merge_cells(start_row=17,start_column=1,end_row=17,end_column=4)

    ws.cell(18,1,"Business Description"); ws.cell(18,1).font=Font(bold=True)
    ws.cell(18,2,description or "No concise business description was verified automatically.")
    ws.merge_cells(start_row=18,start_column=2,end_row=18,end_column=3)
    ws.cell(18,2).alignment=Alignment(wrap_text=True,vertical="top")
    ws.cell(18,4,(curated or {}).get("rows",[[None,None,None,website]])[0][3] if curated else website)
    ws.cell(18,4).font=Font(color="008000",underline="single")

    headers=["Business / Segment","Main Product or Service","What It Does / Why It Matters","Primary Source"]
    for c,v in enumerate(headers,1):
        ws.cell(20,c,v); ws.cell(20,c).fill=_fill(BLUE); ws.cell(20,c).font=Font(bold=True,color=WHITE); ws.cell(20,c).alignment=Alignment(horizontal="center",wrap_text=True)
    if not rows:
        rows=[("Business profile","See business description above","Product-level taxonomy was not reliably available; Segment Analysis remains the financial source of truth.",website)]
    for r_idx,item in enumerate(rows[:10],21):
        for c,v in enumerate(item,1):
            ws.cell(r_idx,c,v); ws.cell(r_idx,c).alignment=Alignment(wrap_text=True,vertical="top")
        if item[3]: ws.cell(r_idx,4).font=Font(color="008000",underline="single")

    ws.cell(32,1,"Research rule"); ws.cell(32,1).font=Font(bold=True,color=GREY)
    ws.cell(32,2,"Products/services are descriptive. Do not infer standalone revenue unless the issuer discloses it; use Segment Analysis for reported economics.")
    ws.merge_cells(start_row=32,start_column=2,end_row=32,end_column=4); ws.cell(32,2).font=Font(italic=True,color=GREY); ws.cell(32,2).alignment=Alignment(wrap_text=True)
    ws.column_dimensions["A"].width=max(ws.column_dimensions["A"].width or 0,28)
    ws.column_dimensions["B"].width=max(ws.column_dimensions["B"].width or 0,42)
    ws.column_dimensions["C"].width=max(ws.column_dimensions["C"].width or 0,62)
    ws.column_dimensions["D"].width=max(ws.column_dimensions["D"].width or 0,55)
    ws.row_dimensions[18].height=48
    return {"products":len(rows),"source":"curated" if curated else "segment/profile fallback"}
