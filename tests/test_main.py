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


def test_stage_date_properties_falls_back_to_exited_then_skips():
    defs = {"hs_v2_date_entered_1": {}, "hs_v2_date_exited_1": {}, "hs_v2_date_exited_2": {}}
    cols, skipped = main.stage_date_properties(["1", "2", "3"], defs)
    assert cols == ["hs_v2_date_entered_1", "hs_v2_date_exited_2"]
    assert skipped == ["3"]


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
    referred = {"date___referred_out": "2026-04-09"}
    assert main.close_date_for(referred, "Referred Out - Complete", True) == \
        (date(2026, 4, 9), "date___referred_out")
    assert "Referred Out - Complete" in main.CLOSED_STAGE_LABELS
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


def test_parse_number():
    assert main.parse_number("5500") == 5500.0
    assert main.parse_number("12.5") == 12.5
    assert main.parse_number("") is None
    assert main.parse_number("n/a") is None


def test_ab1755_maps_stored_manufacturer_values():
    assert main.ab1755_for("General Motors LLC") == "Opt In"
    assert main.ab1755_for("Toyota Motor Sales, U.S.A., Inc. / Lexus") == "Opt Out"
    assert main.ab1755_for("Winnebago Industries, Inc.") == "Not on list"
    assert main.ab1755_for(None) is None
    assert not (main.AB1755_OPT_IN & main.AB1755_OPT_OUT)


def test_sheet_lookups_order_and_headers():
    props = ["hs_object_id", "dealname", "pipeline", "dealstage", "hubspot_owner_id",
             "createdate", "lead___source", "s__manufacturer", "ro_review__final_decision_",
             "hs_v2_date_entered_5411633", "some_new_property"]
    defs = {
        "dealname": {"label": "Deal Name", "type": "string"},
        "pipeline": {"label": "Pipeline", "type": "enumeration"},
        "dealstage": {"label": "Deal Stage", "type": "enumeration"},
        "createdate": {"label": "Create Date", "type": "datetime"},
        "lead___source": {"label": "Lead - Source", "type": "enumeration",
                          "options": [{"value": "VDS", "label": "Agency"}]},
        "s__manufacturer": {"label": "Manufacturer", "type": "enumeration",
                            "options": [{"value": "Polestar", "label": "POLESTAR AUTOMOTIVE USA, INC."}]},
        "ro_review__final_decision_": {"label": "RO Review (Final Decision)", "type": "enumeration",
                                       "options": [{"value": "Sign Up - Pre-Lit (OPT IN OEM)",
                                                    "label": "Pre-Lit (OPT IN OEM)"}]},
        "hs_v2_date_entered_5411633": {"label": 'Date entered "Intake (Lemon Law)"', "type": "datetime"},
        "some_new_property": {"label": "Some New Property", "type": "string"},
        "date___dropped": {"label": "Date - Closed Out", "type": "date"},
    }
    deals = [{"id": "1", "properties": {
        "hs_object_id": "1", "dealname": "Doe, Jane", "pipeline": "default", "dealstage": "5411635",
        "hubspot_owner_id": "77", "createdate": "2026-03-01T12:00:00Z", "lead___source": "VDS",
        "s__manufacturer": "Polestar", "ro_review__final_decision_": "Sign Up - Pre-Lit (OPT IN OEM)",
        "hs_v2_date_entered_5411633": None, "some_new_property": "x", "date___dropped": "2026-03-05"}}]
    stages = [{"id": "5411633", "label": "Intake", "displayOrder": 0},
              {"id": "5411635", "label": "Close Out", "displayOrder": 12}]
    df = main.build_deals_frame(deals, props, defs, {s["id"]: s["label"] for s in stages},
                                {"77": "Ann Lee"}, "Lemon Law")
    df = main.add_stage_attributes(df, stages)
    df = main.add_close_date(df, deals, stages)
    df["last_refresh"] = datetime(2026, 9, 23)
    sheet = main.to_sheet(df, defs, ["hs_v2_date_entered_5411633"])
    assert list(sheet.columns) == [
        "Record ID", "Deal Name", "Create Date",
        "Deal Stage", "Deal Stage ID", "Deal Stage Order", "Deal Stage Is Closed",
        "Close Date", "Close Date Source",
        "Lead - Source",
        "Manufacturer", main.AB1755_HEADER, "RO Review (Final Decision)",
        "Deal Owner", "Deal Owner ID",
        'Date entered "Intake (Lemon Law)"',
        "Some New Property",                  # unplaced property: kept, before Audit
        "Pipeline", "Last Refresh"]
    row = sheet.iloc[0]
    assert row["Pipeline"] == "Lemon Law"
    assert row["Deal Stage"] == "Close Out" and row["Deal Stage ID"] == "5411635"
    assert row["Deal Stage Order"] == 12 and row["Deal Stage Is Closed"] == True  # noqa: E712
    assert row["Close Date"] == date(2026, 3, 5)
    assert row["Close Date Source"] == "Date - Closed Out"   # label, not internal name
    assert row["Deal Owner"] == "Ann Lee" and row["Deal Owner ID"] == "77"
    assert row["Lead - Source"] == "Agency"
    assert row["Manufacturer"] == "POLESTAR AUTOMOTIVE USA, INC."
    assert row[main.AB1755_HEADER] == "Opt Out"
    assert row["Create Date"] == datetime(2026, 3, 1, 7, 0)


def test_every_base_property_has_a_place_in_the_layout():
    placed = {k for _, keys in main.COLUMN_GROUPS for k in keys}
    unplaced = [p for p in main.BASE_PROPERTIES if p not in placed]
    assert not unplaced, f"add to COLUMN_GROUPS: {unplaced}"
    assert set(main.FIXED_HEADERS) <= placed


def test_durations_are_calendar_days_and_blank_when_a_date_is_missing():
    import pandas as pd
    assert main.days_between(datetime(2026, 1, 1, 23, 0), date(2026, 1, 3)) == 2
    assert main.days_between(date(2026, 1, 3), datetime(2026, 1, 1, 1, 0)) == -2
    assert main.days_between(None, date(2026, 1, 3)) is None
    df = pd.DataFrame({
        "createdate": [datetime(2026, 1, 1, 9), datetime(2026, 2, 1, 9)],
        "date___ro_review": [date(2026, 1, 11), None],
        "hs_v2_date_exited_5792630": [datetime(2026, 1, 21, 15), None],
        "date___settled": [date(2026, 5, 1), None],
    })
    out = main.add_durations(df)
    assert out["days_created_to_ro_review"].tolist()[0] == 10
    assert out["days_ro_review_to_file_set_up"].tolist()[0] == 10
    assert out["days_created_to_file_set_up"].tolist()[0] == 20
    assert out["days_created_to_settled"].tolist()[0] == 120
    assert out["days_created_to_ro_review"].isna().tolist() == [False, True]


def test_headers_are_english_ascii():
    for header in main.FIXED_HEADERS.values():
        assert header.isascii(), header


def test_unique_headers():
    assert main.unique_headers([("Close Out", "a"), ("Close Out", "b"), ("X", "c")]) == \
        ["Close Out (a)", "Close Out (b)", "X"]


if __name__ == "__main__":
    tests = [v for k, v in dict(globals()).items() if k.startswith("test_")]
    for t in tests:
        t()
        print(f"ok  {t.__name__}")
    print(f"{len(tests)} passed")
