from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
APP = ROOT / "institutional_research" / "app.py"


def test_streamlit_launcher_exists_and_registers_all_pages():
    text = APP.read_text(encoding="utf-8")
    expected = [
        "dashboard.py",
        "pages/1_Alpha_Analysis.py",
        "pages/2_Portfolio_Optimization.py",
        "pages/3_Model_Learning.py",
        "pages/4_AI_Optionality.py",
    ]
    for relative in expected:
        assert relative in text
        assert (ROOT / "institutional_research" / relative).exists()


def test_ai_optionality_is_explicitly_visible_in_navigation():
    text = APP.read_text(encoding="utf-8")
    assert "AI Optionality & Uncertainty" in text
    assert "st.navigation" in text
