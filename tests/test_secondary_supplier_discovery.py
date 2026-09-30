from supply_chain_intelligence import _secondary_supplier_records


def test_secondary_supplier_fallback_parses_sec_derived_supplier_section(monkeypatch):
    sample=("Supply chain from 10-K filings Suppliers 3 "
            "Micron Technology, Inc. SK Hynix Inc. Taiwan Semiconductor Manufacturing Company Limited "
            "Names it as a customer 2 Example Customer Inc.")
    monkeypatch.setattr("supply_chain_intelligence._fetch_text",lambda *args,**kwargs:sample)
    rows=_secondary_supplier_records("TEST",20)
    names=[r.supplier for r in rows]
    assert any("Micron" in x for x in names)
    assert any("Hynix" in x for x in names)
    assert all(r.evidence_status=="Secondary — SEC-derived" for r in rows)
