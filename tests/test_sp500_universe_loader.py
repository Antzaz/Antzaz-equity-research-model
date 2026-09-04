from __future__ import annotations

import io

import pandas as pd
import pytest

from machine_learning import history_sources as hs


def _csv(n=500):
    df=pd.DataFrame({
        "Symbol":[f"T{i:03d}" for i in range(n)],
        "Security":[f"Company {i}" for i in range(n)],
        "GICS Sector":["Information Technology"]*n,
        "GICS Sub-Industry":["Application Software"]*n,
        "CIK":[1000000+i for i in range(n)],
    })
    return df.to_csv(index=False)


class _Response:
    def __init__(self,text="",status=200,content_type="text/csv",payload=None):
        self.text=text
        self.status_code=status
        self.headers={"Content-Type":content_type}
        self._payload=payload
    def raise_for_status(self):
        if self.status_code>=400:
            raise RuntimeError(f"HTTP {self.status_code}")
    def json(self):
        if self._payload is None:
            raise ValueError("no json payload")
        return self._payload


def test_large_universe_uses_github_when_wikipedia_fails(monkeypatch,tmp_path):
    monkeypatch.setattr(hs,"SP500_CACHE",tmp_path/"sp500.csv")
    csv_text=_csv(500)

    def fake_get(url,**kwargs):
        if url==hs.SP500_WIKI_URL:
            raise RuntimeError("blocked")
        if url==hs.SP500_GITHUB_RAW_URL:
            return _Response(csv_text)
        raise AssertionError(f"unexpected URL {url}")

    monkeypatch.setattr(hs.requests,"get",fake_get)
    rows=hs.current_sp500_universe(500)
    assert len(rows)==500
    assert rows[0]["symbol"]=="T000"
    assert rows[-1]["symbol"]=="T499"
    assert hs.SP500_CACHE.exists()


def test_cached_broad_universe_survives_network_failure(monkeypatch,tmp_path):
    monkeypatch.setattr(hs,"SP500_CACHE",tmp_path/"sp500.csv")
    cached=pd.DataFrame({
        "symbol":[f"T{i:03d}" for i in range(500)],
        "name":[f"Company {i}" for i in range(500)],
        "sector":["Industrials"]*500,
        "industry":["Machinery"]*500,
        "cik":[2000000+i for i in range(500)],
        "universe_source":["seed"]*500,
    })
    cached.to_csv(hs.SP500_CACHE,index=False)
    monkeypatch.setattr(hs.requests,"get",lambda *a,**k: (_ for _ in ()).throw(RuntimeError("offline")))
    rows=hs.current_sp500_universe(500)
    assert len(rows)==500
    assert rows[0]["universe_source"]=="Cached S&P 500 constituent snapshot"


def test_large_request_never_silently_shrinks_to_builtin(monkeypatch,tmp_path):
    monkeypatch.setattr(hs,"SP500_CACHE",tmp_path/"missing.csv")
    monkeypatch.setattr(hs.requests,"get",lambda *a,**k: (_ for _ in ()).throw(RuntimeError("offline")))
    with pytest.raises(RuntimeError,match="Refusing to silently fall back"):
        hs.current_sp500_universe(500)
