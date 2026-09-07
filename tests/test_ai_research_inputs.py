from __future__ import annotations

from pathlib import Path
import json
import tempfile

from institutional_research.src.ai_research_inputs import suggested_ai_config


def test_ai_research_suggestion_is_nonzero_only_with_dated_ai_evidence():
    with tempfile.TemporaryDirectory() as td:
        root = Path(td)
        base = root / "research_data" / "AAA"
        base.mkdir(parents=True)
        payload = {
            "snapshots": [
                {
                    "captured_at": "2026-08-01T00:00:00Z",
                    "kpis": [
                        {
                            "kpi": "Generative AI paid usage",
                            "signal": "Strong",
                            "investment_read_through": "AI monetization is becoming material.",
                            "data_type": "PUBLIC-EVIDENCE",
                        },
                        {
                            "kpi": "AI data center capacity",
                            "signal": "Positive",
                            "investment_read_through": "Inference demand supports capacity growth.",
                            "data_type": "PUBLIC-EVIDENCE",
                        },
                    ],
                }
            ]
        }
        (base / "kpi_history.json").write_text(json.dumps(payload), encoding="utf-8")

        out = suggested_ai_config(root, ["AAA", "BBB"]).set_index("Ticker")
        assert out.loc["AAA", "AIExposure"] > 0
        assert out.loc["AAA", "ResearchEvidenceCount"] >= 2
        assert out.loc["AAA", "EvidenceConfidence"] > out.loc["BBB", "EvidenceConfidence"]
        assert out.loc["BBB", "AIExposure"] == 0.0
        assert out.loc["AAA", "AISkillConfidence"] == 0.25


def test_ai_suggestion_is_a_sensitivity_proxy_not_directional_return_forecast():
    with tempfile.TemporaryDirectory() as td:
        root = Path(td)
        base = root / "research_data" / "AAA"
        base.mkdir(parents=True)
        payload = {
            "snapshots": [
                {
                    "captured_at": "2026-08-01T00:00:00Z",
                    "kpis": [
                        {
                            "kpi": "AI inference economics",
                            "signal": "Key Risk",
                            "investment_read_through": "AI economics are under pressure.",
                        }
                    ],
                }
            ]
        }
        (base / "kpi_history.json").write_text(json.dumps(payload), encoding="utf-8")
        row = suggested_ai_config(root, ["AAA"]).iloc[0]
        # Negative AI evidence can still imply material AI sensitivity; scenario direction is separate.
        assert row["AIExposure"] > 0
        assert 0 <= row["EvidenceConfidence"] <= 1
