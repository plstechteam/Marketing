"""Tests for the pure helpers in main.py. Run: python tests/test_main.py"""
import os
import sys
from datetime import date, datetime

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
import main  # noqa: E402

P = main.PACIFIC


def test_month_windows_cover_the_range_without_gaps():
    start = P.localize(datetime(2026, 1, 1))
    end = P.localize(datetime(2026, 9, 23, 10, 30))
    w = main.month_windows(start, end)
    assert len(w) == 9
    assert w[0][0] == start and w[-1][1] == end
    for (_, a_end), (b_start, _) in zip(w, w[1:]):
        assert a_end == b_start
    # March crosses DST; the cut is still local midnight on the 1st.
    assert w[3][0] == P.localize(datetime(2026, 4, 1))


def test_month_windows_december_rolls_over():
    start = P.localize(datetime(2026, 12, 1))
    end = P.localize(datetime(2027, 1, 5))
    assert main.month_windows(start, end)[0][1] == P.localize(datetime(2027, 1, 1))


def test_label_value_maps_single_and_multi_and_keeps_unknown():
    labels = {"VDS": "Agency", "Scorpion": "SEO: Scorpion"}
    assert main.label_value("VDS", labels) == "Agency"
    assert main.label_value("VDS;Scorpion", labels) == "Agency;SEO: Scorpion"
    assert main.label_value("Gone", labels) == "Gone"
    assert main.label_value("", labels) is None
    assert main.label_value(None, labels) is None


def test_datetime_goes_to_bogota():
    assert main.parse_hubspot_datetime("2026-01-01T03:00:00.000Z") == datetime(2025, 12, 31, 22, 0)
    assert main.parse_hubspot_datetime(None) is None
    assert main.parse_hubspot_date("2026-02-03") == date(2026, 2, 3)


def test_build_deals_frame_headers_and_lookups():
    props = ["hs_object_id", "dealname", "pipeline", "dealstage", "hubspot_owner_id",
             "createdate", "lead___source", "ro_review__final_decision_",
             "hs_v2_date_entered_5411633"]
    defs = {
        "dealname": {"label": "Deal Name", "type": "string"},
        "pipeline": {"label": "Pipeline", "type": "enumeration"},
        "dealstage": {"label": "Deal Stage", "type": "enumeration"},
        "createdate": {"label": "Create Date", "type": "datetime"},
        "lead___source": {"label": "Lead - Source", "type": "enumeration",
                          "options": [{"value": "VDS", "label": "Agency"}]},
        "ro_review__final_decision_": {"label": "RO Review (Final Decision)", "type": "enumeration",
                                       "options": [{"value": "Sign Up - Pre-Lit (OPT IN OEM)",
                                                    "label": "Pre-Lit (OPT IN OEM)"}]},
        "hs_v2_date_entered_5411633": {"label": 'Date entered "Intake (Lemon Law)"', "type": "datetime"},
    }
    deals = [{"id": "1", "properties": {
        "hs_object_id": "1", "dealname": "Doe, Jane", "pipeline": "default", "dealstage": "5411633",
        "hubspot_owner_id": "77", "createdate": "2026-03-01T12:00:00Z", "lead___source": "VDS",
        "ro_review__final_decision_": "Sign Up - Pre-Lit (OPT IN OEM)",
        "hs_v2_date_entered_5411633": None}}]
    df = main.build_deals_frame(deals, props, defs, {"5411633": "Intake"}, {"77": "Ann Lee"}, "Lemon Law")
    assert list(df.columns) == [
        "Record ID", "Deal Name", "Pipeline", "Deal Stage ID", "Deal Stage", "Deal Owner ID",
        "Deal Owner", "Create Date", "Lead - Source", "RO Review (Final Decision)",
        'Date entered "Intake (Lemon Law)"']
    row = df.iloc[0]
    assert row["Pipeline"] == "Lemon Law"
    assert row["Deal Stage ID"] == "5411633" and row["Deal Stage"] == "Intake"
    assert row["Deal Owner ID"] == "77" and row["Deal Owner"] == "Ann Lee"
    assert row["Lead - Source"] == "Agency"
    assert row["RO Review (Final Decision)"] == "Pre-Lit (OPT IN OEM)"
    assert row["Create Date"] == datetime(2026, 3, 1, 7, 0)


def test_stage_date_properties_falls_back_to_exited_then_skips():
    defs = {"hs_v2_date_entered_1": {}, "hs_v2_date_exited_1": {}, "hs_v2_date_exited_2": {}}
    cols, skipped = main.stage_date_properties(["1", "2", "3"], defs)
    assert cols == ["hs_v2_date_entered_1", "hs_v2_date_exited_2"]
    assert skipped == ["3"]


def test_unique_headers():
    assert main.unique_headers([("Close Out", "a"), ("Close Out", "b"), ("X", "c")]) == \
        ["Close Out (a)", "Close Out (b)", "X"]


if __name__ == "__main__":
    tests = [v for k, v in dict(globals()).items() if k.startswith("test_")]
    for t in tests:
        t()
        print(f"ok  {t.__name__}")
    print(f"{len(tests)} passed")
