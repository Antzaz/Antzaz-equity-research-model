from automation.sec_preflight import validate_user_agent


def test_sec_user_agent_requires_contact_email():
    ok, note = validate_user_agent("EquityResearchBot/1.0")
    assert ok is False
    assert "email" in note.lower()


def test_sec_user_agent_accepts_descriptive_identity_with_email():
    ok, note = validate_user_agent("EquityResearchBot research@example.com")
    assert ok is True
    assert note == "configured"


def test_sec_user_agent_rejects_missing_value():
    ok, note = validate_user_agent("")
    assert ok is False
    assert "not configured" in note.lower()
