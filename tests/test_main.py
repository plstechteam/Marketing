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
    # People dropdowns: an archived owner has no option but is in the owner map.
    people = {"77": "Ann Lee"}
    assert main.label_value("77", people, {"99": "Old Owner"}) == "Ann Lee"
    assert main.label_value("99", people, {"99": "Old Owner"}) == "Old Owner"
    assert main.label_value("12", people, {"99": "Old Owner"}) == "12"


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


def test_add_stage_attributes_puts_order_and_closed_on_each_row():
    import pandas as pd
    df = pd.DataFrame({"Record ID": ["1", "2", "3"], "Deal Stage ID": ["a", "b", "zz"],
                       "Deal Stage": ["Intake", "Settled", "zz"], "X": [1, 2, 3]})
    stages = [{"id": "a", "label": "Intake", "displayOrder": 0},
              {"id": "b", "label": "Close Out", "displayOrder": 9}]
    out = main.add_stage_attributes(df, stages)
    assert list(out.columns) == ["Record ID", "Deal Stage ID", "Deal Stage",
                                 "Deal Stage Order", "Deal Stage Is Closed", "X"]
    assert out["Deal Stage Order"].tolist()[:2] == [0, 9]
    assert out["Deal Stage Is Closed"].tolist()[:2] == [False, True]
    assert pd.isna(out["Deal Stage Order"].iloc[2])


def test_close_date_for_picks_the_field_that_matches_how_it_closed():
    from datetime import date
    settled = {"date___settled": "2026-05-04", "date___dropped": "2026-01-01",
               "hs_v2_date_entered_current_stage": "2026-06-01T12:00:00Z"}
    assert main.close_date_for(settled, "Settled - Pre Lit", True) == (date(2026, 5, 4), "date___settled")
    closed_out = {"date___dropped": "2026-02-03"}
    assert main.close_date_for(closed_out, "Close Out", True) == (date(2026, 2, 3), "date___dropped")
    after_retained = {"date__intake_sign_up_close_out": "2026-03-01"}
    assert main.close_date_for(after_retained, "Retained - Client Dropped", True) == \
        (date(2026, 3, 1), "date__intake_sign_up_close_out")
    # Closed with the firm's field blank: never left undated.
    bare = {"hs_v2_date_entered_current_stage": "2026-06-01T12:00:00Z"}
    assert main.close_date_for(bare, "Settled - Lit", True) == \
        (date(2026, 6, 1), "hs_v2_date_entered_current_stage")
    # Open: no close date even when a date field happens to be filled.
    assert main.close_date_for(settled, "Intake", False) == (None, None)


def test_closed_stages_include_close_out_and_fail_when_renamed():
    assert "Close Out" in main.CLOSED_STAGE_LABELS
    everything = [{"label": l} for l in main.CLOSED_STAGE_LABELS]
    main.check_closed_stages(everything)          # no exit
    try:
        main.check_closed_stages(everything[1:])
    except SystemExit:
        pass
    else:
        raise AssertionError("a missing closed stage must stop the run")


def test_unique_headers():
    assert main.unique_headers([("Close Out", "a"), ("Close Out", "b"), ("X", "c")]) == \
        ["Close Out (a)", "Close Out (b)", "X"]


if __name__ == "__main__":
    tests = [v for k, v in dict(globals()).items() if k.startswith("test_")]
    for t in tests:
        t()
        print(f"ok  {t.__name__}")
    print(f"{len(tests)} passed")
