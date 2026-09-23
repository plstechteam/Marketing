"""
Marketing report — GitHub Actions version
Pulls every Lemon Law deal created since 1 January of the current year from
HubSpot and writes it, raw, to Marketing.xlsx in SharePoint via Microsoft
Graph. The workbook is the source for a Power BI funnel report; every funnel
calculation lives in Power BI, none of it here.
"""

import io
import os
import sys
import time
from datetime import date, datetime, timezone

import pandas as pd
import pytz
import requests
import xlsxwriter

# ======================================================
# CONFIGURATION
# ======================================================
# Same drive as the Monthly Settlement Report, same folder. Overridable so a
# run can be aimed at a test copy; the defaults are the live file.
# `or` and not a get() default: the workflow passes an empty string when an
# input is left blank, and an empty string would win over the default.
DRIVE_ID = (os.environ.get("SHAREPOINT_DRIVE_ID") or
            "b!Czt_67dx0EK0ERl-AWfwrlF4EcaopHRAlgnXnFrpZj6U8X8uYgT7QrymPap7uNOr")
FILE_PATH = (os.environ.get("SHAREPOINT_FILE_PATH") or
             "Data Inventory/PLS/Requests/Marketing.xlsx")

DRY_RUN = os.environ.get("DRY_RUN", "").strip().lower() in ("1", "true", "yes")

PIPELINE_ID = "default"          # Lemon Law. Employment Law is out of scope.

PACIFIC = pytz.timezone("America/Los_Angeles")
# Every date and time in the sheet is California time — deal dates, the
# year's cut-off and Last Refresh alike — so a deal created at 11 PM on
# 31 December in Los Angeles is a December deal, as the firm sees it.
DEAL_TZ = PACIFIC

# One sheet, and every row stands on its own: stage and owner come as names
# (with their IDs beside them), and the stage's order and closed flag ride on
# the row, so Power BI needs no lookup table to read or sort a deal.
DEALS_SHEET = "Deals"

# The deal columns, in sheet order. Headers come from HubSpot's own labels.
# The "Date entered <stage>" columns are not listed here: they are built from
# the pipeline's stages at run time, so a stage added in HubSpot gets its
# column without a code change — see stage_date_properties.
BASE_PROPERTIES = [
    "hs_object_id",
    "dealname",
    "pipeline",
    "dealstage",
    "hubspot_owner_id",
    # People on the case. HubSpot stores these as dropdowns whose values are
    # owner ids; they go out as names (see label_value's fallback).
    "n5__retainer_representative",  # Intake - Case Supervisor
    "senior_case_supervisor",       # Senior Case Supervisor
    "handling_attorney",            # Legal - Handling Attorney
    "supervising_attorney",         # Settlement Attorney
    "createdate",
    "hs_lastmodifieddate",
    "hs_v2_date_entered_current_stage",
    "case_category",                # Lit / Pre-Lit
    "date___settled",               # Date - Settled — confirms the settlement
    "date___dropped",               # Date - Closed Out
    "date__intake_sign_up_close_out",   # Date - Close Out After Retained
    "date___referred_out",          # Date - Referred Out
    "date___ro_review",             # Date - RO Review
    "hs_v2_date_exited_5792630",    # Date exited "New File Set Up - Doc Collection"
    "total_settled_attorneys_fees_and_cost",
    "net_attorney_fees",
    # Vehicle. Manufacturer is the full legal name (s__manufacturer), not
    # the short code, so it can be matched to the opt-in / opt-out list.
    "s__manufacturer",
    "vehicle___year",
    "c__vehicle___model__new_test_",    # Vehicle - Model
    "lead___source",
    "lead___source__group_",
    "hs_analytics_source",
    "hs_analytics_source_data_1",
    "hs_object_source_label",
    # Channel detail. Lead - Source (Group) has no calls or forms bucket, so
    # these are what split Prospect into inbound call / outbound call /
    # website form / PPC / social.
    "aircall_entry_number",         # inbound call — the Aircall line it came in on
    "auto_dialer_call_type",        # outbound auto-dialer (Crexendo) call type
    "tf__utm_source",               # Typeform UTMs — the deal-level utm_* are unused
    "tf__utm_medium",
    "tf__utm_campaign",
    "gclid",                        # Google Ads click id — PPC
    "drop_reason",                  # Close Out Reason — detail for Closed Lost
    "ro_review__final_decision_",   # Opt In / Opt Out split
    "deal_stage___sub_phase",
    "legal_pipeline",
]
# Left out on purpose, measured on every 2026 Lemon Law deal (18,111) in
# September 2026: closedate, intake_outcome, class_action, utm_source,
# utm_medium, utm_campaign, lead_generation_form and hs_form_id were blank on
# every one. HubSpot's own closedate is blank on every Lemon Law deal ever
# (none of the 231,000+), which is why the sheet's Close Date is taken from
# the firm's own date fields instead — see close_date_for. Add a property back
# here if it starts being used.

# The stages where a case is over, by label. HubSpot's own "closed" flag
# cannot be used: it marks only the three Settled stages, so Close Out — where
# 93% of this year's deals end — would read as open and carry no close date.
# Labels, not ids, as in the Settlement repo: ids are opaque, and a renamed
# stage should fail loudly (see check_closed_stages) rather than be silently
# reclassified. This is the one place to change the definition.
CLOSED_STAGE_LABELS = {
    # won
    "Settled - Lit",
    "Settled - Pre Lit",
    "Settled - Referred Out",
    # lost
    "Close Out",
    "Retained - Drop Client",
    "Retained - Client Dropped",
    # referred out and finished
    "Referred Out - Complete",
}

# AB 1755: which manufacturers opted in to California's new lemon law
# procedure and which stayed out, from the firm's published list (September
# 2026). Keyed on the STORED value of s__manufacturer — the full legal name —
# not its label, so relabelling a dropdown option does not unmap it.
# Brands the list names separately that HubSpot files under a parent:
# Genesis -> Hyundai Motor America, Infiniti -> Nissan North America,
# Mercedes -> Mercedes-Benz USA, Toyota/Lexus -> one value. Isuzu has no
# manufacturer value in HubSpot. A manufacturer not listed here comes out as
# "Not on list" — never guessed from a sister brand (Bentley and Lamborghini
# are VW group but are not on the list).
AB1755_OPT_IN = {
    "FCA US LLC",
    "Ford Motor Company",
    "General Motors LLC",
    "Hyundai Motor America",                  # Hyundai, Genesis
    "Jaguar Land Rover North America, LLC",   # JLRNA
    "Kia America, Inc.",
    "Maserati North America, Inc.",
    "Mercedes-Benz USA, LLC",
    "Mitsubishi Motors North America, INC.",
    "Nissan North America, Inc.",             # Nissan, Infiniti
    "Subaru of America, Inc.",
    "VinFast Auto, LLC",
}
AB1755_OPT_OUT = {
    "Aston Martin Lagonda of North America, Inc.",
    "BMW of North America, LLC",
    "American Honda Motor Co., Inc.",
    "Lucid Group, Inc.",
    "Mazda Motor of America, Inc.",
    "McLaren Automotive, Inc.",
    "Polestar",
    "Porsche Cars North America, Inc.",
    "Rivian Automotive",
    "TESLA MOTORS, INC.",
    "Toyota Motor Sales, U.S.A., Inc. / Lexus",
    "Volkswagen Group of America, Inc.",
    "Volvo Car USA LLC",
}
AB1755_HEADER = "AB 1755 (Manufacturer)"


def ab1755_for(manufacturer):
    """Opt In / Opt Out / Not on list for a stored manufacturer value."""
    if not manufacturer:
        return None
    if manufacturer in AB1755_OPT_IN:
        return "Opt In"
    if manufacturer in AB1755_OPT_OUT:
        return "Opt Out"
    return "Not on list"


# Search stops paging at 10,000 results with no error — it simply stops
# returning `after`. The year is well past that, so it is pulled a calendar
# month at a time, and a month that reaches the ceiling aborts the run rather
# than writing a truncated sheet.
SEARCH_PAGE_SIZE = 200
SEARCH_API_MAX_RESULTS = 10_000

# ======================================================
# HTTP SESSION — connection reuse, timeouts and retries
# ======================================================
HTTP_TIMEOUT   = (10, 120)
UPLOAD_TIMEOUT = (10, 300)
MAX_ATTEMPTS   = 4
RETRY_STATUS   = {429, 500, 502, 503, 504}
# 423: someone has the workbook open in Excel. Clears on a human timescale.
LOCKED_STATUS      = 423
LOCKED_RETRY_DELAY = 60

session = requests.Session()


def _retry_delay(response, attempt):
    """Honour Retry-After when present, otherwise exponential backoff."""
    if response is not None:
        if response.status_code == LOCKED_STATUS:
            return LOCKED_RETRY_DELAY
        retry_after = response.headers.get("Retry-After")
        if retry_after:
            try:
                return max(1.0, min(float(retry_after), 60.0))
            except ValueError:
                pass
    return min(2 ** attempt, 30)


def request_with_retry(method, url, *, timeout=HTTP_TIMEOUT, **kwargs):
    """Send a request, retrying rate limits, 5xx, 423 and network errors."""
    response = None
    reason = None
    for attempt in range(1, MAX_ATTEMPTS + 1):
        try:
            response = session.request(method, url, timeout=timeout, **kwargs)
        except requests.RequestException as exc:
            response = None
            reason = f"network error: {exc}"
        else:
            if (response.status_code not in RETRY_STATUS
                    and response.status_code != LOCKED_STATUS):
                return response
            reason = f"HTTP {response.status_code}"

        if attempt == MAX_ATTEMPTS:
            break
        wait = _retry_delay(response, attempt)
        print(f"  {method} failed ({reason}) — attempt {attempt}/{MAX_ATTEMPTS}, retrying in {wait:.0f}s")
        time.sleep(wait)

    if response is None:
        print(f"ERROR: {method} {url.split('?')[0]} failed after {MAX_ATTEMPTS} attempts — {reason}")
        sys.exit(1)
    return response


def fail(message):
    print(f"ERROR: {message}")
    sys.exit(1)


# ======================================================
# PURE HELPERS — no I/O, covered by tests/test_main.py
# ======================================================
def month_windows(start, end):
    """Split [start, end) into calendar-month windows, in start's timezone.

    Both bounds are aware (pytz) datetimes. Months are cut on local midnight
    of the 1st, so a window never straddles a month in that timezone.
    """
    tz = start.tzinfo
    windows = []
    cursor = start
    while cursor < end:
        if cursor.month == 12:
            nxt = tz.localize(datetime(cursor.year + 1, 1, 1))
        else:
            nxt = tz.localize(datetime(cursor.year, cursor.month + 1, 1))
        windows.append((cursor, min(nxt, end)))
        cursor = nxt
    return windows


def to_epoch_ms(dt):
    return str(int(dt.timestamp() * 1000))


def option_labels(definition):
    """value -> label for an enumeration property, {} for anything else."""
    return {o["value"]: o["label"] for o in definition.get("options") or []}


def label_value(raw, labels, fallback=None):
    """Map an enumeration's stored value(s) to what HubSpot's UI shows.

    Multi-select values arrive ';'-joined. A value with no option is looked up
    in `fallback` — the owner map, because the people dropdowns (case
    supervisor, attorneys) store owner ids and drop the option when someone
    is archived — and otherwise kept as stored rather than blanked.
    """
    if raw is None or raw == "":
        return None
    fallback = fallback or {}
    if not labels and not fallback:
        return raw
    return ";".join(labels.get(v) or fallback.get(v, v) for v in str(raw).split(";"))


def parse_hubspot_datetime(raw):
    """HubSpot datetime (ISO string, UTC) -> naive datetime in DEAL_TZ."""
    if raw is None or raw == "":
        return None
    ts = pd.Timestamp(raw)
    if ts.tzinfo is None:
        ts = ts.tz_localize("UTC")
    return ts.tz_convert(DEAL_TZ).tz_localize(None).to_pydatetime()


def parse_hubspot_date(raw):
    """HubSpot date property ('YYYY-MM-DD' or midnight-UTC ISO) -> date."""
    if raw is None or raw == "":
        return None
    return pd.Timestamp(str(raw)[:10]).date()


def parse_number(raw):
    """HubSpot number (sent as a string) -> float, None if blank or unreadable."""
    if raw is None or raw == "":
        return None
    try:
        return float(raw)
    except (TypeError, ValueError):
        return None


def stage_date_properties(stage_ids, definitions):
    """Pick one date column per stage: when the deal entered it, else when it
    exited it, else none.

    HubSpot does not create a "date entered" property for every stage — in
    this portal HOLD, Retained - Client Dropped, Settled - Referred Out and
    TEST have none, and HOLD has only "date exited". A column for a property
    that does not exist would be blank on every row and read as missing data,
    so those stages get the exited date where there is one and no column
    otherwise. The deals themselves are unaffected: their stage is in
    Deal Stage and the date they got there in Date entered current stage.
    """
    columns, skipped = [], []
    for sid in stage_ids:
        for kind in ("entered", "exited"):
            name = f"hs_v2_date_{kind}_{sid}"
            if name in definitions:
                columns.append(name)
                break
        else:
            skipped.append(sid)
    return columns, skipped


def unique_headers(names):
    """Make headers unique by appending the internal name to any repeat."""
    seen = {}
    for label, _ in names:
        seen[label] = seen.get(label, 0) + 1
    return [label if seen[label] == 1 else f"{label} ({name})" for label, name in names]


def build_deals_frame(deals, properties, definitions, stage_labels, owners, pipeline_label):
    """Raw deals -> one row per deal, columns keyed by HubSpot internal name.

    Only lookups happen here (stage id -> name, owner id -> name, enumeration
    value -> label) and a timezone conversion on dates. Columns stay keyed by
    internal name until to_sheet orders and labels them, so nothing
    downstream depends on a label HubSpot could rename.
    """
    records = []
    for deal in deals:
        p = deal.get("properties", {})
        row = {}
        for name in properties:
            raw = p.get(name)
            d = definitions.get(name, {})
            if name == "pipeline":
                row[name] = pipeline_label if raw == PIPELINE_ID else raw
            elif name == "dealstage":
                row["dealstage__id"] = raw
                row[name] = stage_labels.get(raw, raw)
            elif name == "hubspot_owner_id":
                row["hubspot_owner_id__id"] = raw
                row[name] = owners.get(str(raw), raw) if raw else None
            elif d.get("type") == "datetime":
                row[name] = parse_hubspot_datetime(raw)
            elif d.get("type") == "date":
                row[name] = parse_hubspot_date(raw)
            elif name == "s__manufacturer":
                row[name] = label_value(raw, option_labels(d), owners)
                row["s__manufacturer__ab1755"] = ab1755_for(raw)
            elif d.get("type") == "number":
                row[name] = parse_number(raw)
            elif d.get("type") == "enumeration":
                row[name] = label_value(raw, option_labels(d), owners)
            else:
                row[name] = raw if raw != "" else None
        records.append(row)
    return pd.DataFrame(records)


# Cycle times, in calendar days, computed on every run from that run's
# dates — so a date corrected in HubSpot corrects its duration on the next
# run. (key, header, from date, to date). "File Set Up" is the day the deal
# left New File Set Up - Doc Collection.
DURATIONS = [
    ("days_created_to_ro_review", "Days: Created to RO Review",
     "createdate", "date___ro_review"),
    ("days_ro_review_to_file_set_up", "Days: RO Review to File Set Up",
     "date___ro_review", "hs_v2_date_exited_5792630"),
    ("days_created_to_file_set_up", "Days: Created to File Set Up",
     "createdate", "hs_v2_date_exited_5792630"),
    ("days_created_to_settled", "Days: Created to Settled",
     "createdate", "date___settled"),
]


def days_between(start, end):
    """Whole calendar days from start to end (date or datetime, both in the
    sheet's timezone); None if either is missing. A negative result is kept:
    it means the dates in HubSpot are out of order, which is worth seeing."""
    if start is None or end is None or pd.isna(start) or pd.isna(end):
        return None
    to_date = lambda v: v.date() if isinstance(v, datetime) else v  # noqa: E731
    return (to_date(end) - to_date(start)).days


def add_durations(df):
    for key, _, start, end in DURATIONS:
        if start in df.columns and end in df.columns:
            values = [days_between(a, b) for a, b in zip(df[start], df[end])]
        else:
            values = [None] * len(df)
        df[key] = pd.array(values, dtype="Int64")
    return df


# Headers for the columns this script adds or renames; everything else is
# headed by its HubSpot label.
FIXED_HEADERS = {
    "hs_object_id": "Record ID",
    "dealstage__id": "Deal Stage ID",
    "dealstage__order": "Deal Stage Order",
    "dealstage__closed": "Deal Stage Is Closed",
    "close_date": "Close Date",
    "close_date_source": "Close Date Source",
    "s__manufacturer__ab1755": AB1755_HEADER,
    "hubspot_owner_id": "Deal Owner",
    "hubspot_owner_id__id": "Deal Owner ID",
    "last_refresh": "Last Refresh",
    **{key: header for key, header, _, _ in DURATIONS},
}

# Sheet layout, left to right in the order the funnel reads: who the deal is,
# where it stands now, where it came from, what vehicle and AB 1755 side it
# is on, the milestones on the way, how it ended, who is on it, and the full
# stage history. "STAGE_HISTORY" expands to one date column per pipeline
# stage, in pipeline order.
COLUMN_GROUPS = [
    ("Deal", ["hs_object_id", "dealname", "legal_pipeline", "createdate"]),
    ("Current stage", ["dealstage", "dealstage__id", "dealstage__order", "dealstage__closed",
                       "hs_v2_date_entered_current_stage", "close_date", "close_date_source"]),
    ("Source / channel", ["lead___source", "lead___source__group_", "hs_analytics_source",
                          "hs_analytics_source_data_1", "hs_object_source_label",
                          "aircall_entry_number", "auto_dialer_call_type",
                          "tf__utm_source", "tf__utm_medium", "tf__utm_campaign", "gclid"]),
    ("Vehicle / AB 1755", ["s__manufacturer", "s__manufacturer__ab1755",
                           "ro_review__final_decision_", "vehicle___year",
                           "c__vehicle___model__new_test_"]),
    ("Milestones", ["date___ro_review", "hs_v2_date_exited_5792630", "date___referred_out"]),
    ("Outcome", ["case_category", "date___settled", "total_settled_attorneys_fees_and_cost",
                 "net_attorney_fees", "drop_reason", "date___dropped",
                 "date__intake_sign_up_close_out", "deal_stage___sub_phase"]),
    ("Cycle time (days)", [key for key, _, _, _ in DURATIONS]),
    ("People", ["hubspot_owner_id", "hubspot_owner_id__id", "n5__retainer_representative",
                "senior_case_supervisor", "handling_attorney", "supervising_attorney"]),
    ("Stage history", ["STAGE_HISTORY"]),
    ("Audit", ["pipeline", "hs_lastmodifieddate", "last_refresh"]),
]


def column_order(columns, stage_columns):
    """Internal column keys in sheet order. A column no group names (a
    property added to BASE_PROPERTIES but not placed) goes just before
    Audit rather than being dropped."""
    ordered = []
    for group, keys in COLUMN_GROUPS:
        if group == "Audit":
            placed = set(ordered) | set(keys)
            ordered += [c for c in columns if c not in placed]
        for key in keys:
            if key == "STAGE_HISTORY":
                ordered += [c for c in stage_columns if c in columns]
            elif key in columns:
                ordered.append(key)
    return list(dict.fromkeys(ordered))


def header_for(key, definitions):
    return FIXED_HEADERS.get(key) or (definitions.get(key, {}).get("label") or key).strip()


def to_sheet(df, definitions, stage_columns):
    """Order the columns by COLUMN_GROUPS and head them with labels. Close
    Date Source is written as the label of the field used, so the row says
    "Date - Settled" rather than an internal name."""
    df = df[column_order(list(df.columns), stage_columns)].copy()
    if "close_date_source" in df.columns:
        df["close_date_source"] = df["close_date_source"].map(
            lambda k: header_for(k, definitions) if k else None)
    df.columns = unique_headers([(header_for(c, definitions), c) for c in df.columns])
    return df


# ======================================================
# HUBSPOT
# ======================================================
def hubspot_get(path, headers, **params):
    r = request_with_retry("GET", f"https://api.hubapi.com{path}", headers=headers, params=params)
    if r.status_code != 200:
        fail(f"GET {path}: {r.status_code} — {r.text}")
    return r.json()


def fetch_pipeline(headers):
    data = hubspot_get(f"/crm/v3/pipelines/deals/{PIPELINE_ID}", headers)
    stages = sorted(data.get("stages", []), key=lambda s: s.get("displayOrder", 0))
    if not stages:
        fail(f"pipeline '{PIPELINE_ID}' returned no stages")
    return data.get("label", PIPELINE_ID), stages


def fetch_property_definitions(names, headers):
    r = request_with_retry(
        "POST", "https://api.hubapi.com/crm/v3/properties/deals/batch/read",
        headers=headers,
        json={"archived": False, "inputs": [{"name": n} for n in names]},
    )
    if r.status_code == 403:
        fail("property definitions: 403 — the HubSpot token needs the "
             "crm.schemas.deals.read scope (headers and dropdown labels come from it).")
    if r.status_code not in (200, 207):
        fail(f"property definitions: {r.status_code} — {r.text}")
    definitions = {p["name"]: p for p in r.json().get("results", [])}
    # Stage date properties are probed, not assumed — stage_date_properties
    # reports the ones that are absent — so only the fixed list warns here.
    missing = [n for n in names if n not in definitions
               and not (n.startswith("hs_v2_date_") and n.rsplit("_", 1)[-1].isdigit())]
    if missing:
        # Kept as blank columns: a renamed or archived property should show up
        # as an empty column and a warning, not kill the report.
        print(f"WARNING: properties not found in HubSpot: {', '.join(missing)}")
    return definitions


def fetch_owners(headers):
    owners = {}
    for archived in ("false", "true"):
        after = None
        while True:
            params = {"limit": 500, "archived": archived}
            if after:
                params["after"] = after
            data = hubspot_get("/crm/v3/owners", headers, **params)
            for o in data.get("results", []):
                owners[str(o["id"])] = {
                    "name": f"{o.get('firstName', '')} {o.get('lastName', '')}".strip(),
                    "email": o.get("email"),
                    "archived": archived == "true",
                }
            after = data.get("paging", {}).get("next", {}).get("after")
            if not after:
                break
    return owners


def fetch_deals(properties, windows, headers):
    url = "https://api.hubapi.com/crm/v3/objects/deals/search"
    deals = {}
    for start, end in windows:
        payload = {
            "properties": properties,
            "limit": SEARCH_PAGE_SIZE,
            "sorts": [{"propertyName": "hs_object_id", "direction": "ASCENDING"}],
            "filterGroups": [{"filters": [
                {"propertyName": "pipeline", "operator": "EQ", "value": PIPELINE_ID},
                {"propertyName": "createdate", "operator": "GTE", "value": to_epoch_ms(start)},
                {"propertyName": "createdate", "operator": "LT", "value": to_epoch_ms(end)},
            ]}],
        }
        window_rows = 0
        total = None
        after = None
        while True:
            if after:
                payload["after"] = after
            r = request_with_retry("POST", url, headers=headers, json=payload)
            if r.status_code != 200:
                fail(f"deal search: {r.status_code} — {r.text}")
            data = r.json()
            total = data.get("total", total)
            for deal in data.get("results", []):
                deals[deal["id"]] = deal
                window_rows += 1
            after = data.get("paging", {}).get("next", {}).get("after")
            if not after:
                break

        label = start.strftime("%Y-%m")
        print(f"  {label}: {window_rows} deals (HubSpot total {total})")
        if window_rows >= SEARCH_API_MAX_RESULTS or (total or 0) >= SEARCH_API_MAX_RESULTS:
            fail(f"{label} reached the {SEARCH_API_MAX_RESULTS:,}-result search ceiling — "
                 "the pull would be truncated. Split the windows further.")
        if total is not None and window_rows < total:
            fail(f"{label}: downloaded {window_rows} of {total} — incomplete pull, nothing written.")
    return list(deals.values())


# ======================================================
# MICROSOFT GRAPH
# ======================================================
def graph_token(tenant_id, client_id, client_secret):
    r = request_with_retry(
        "POST", f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token",
        data={
            "grant_type": "client_credentials",
            "client_id": client_id,
            "client_secret": client_secret,
            "scope": "https://graph.microsoft.com/.default",
        },
    )
    if r.status_code != 200:
        fail(f"Graph token: {r.status_code} — {r.text}")
    return r.json()["access_token"]


def check_destination(token):
    """Confirm the target folder exists and say whether the file does yet."""
    headers = {"Authorization": f"Bearer {token}"}
    folder = FILE_PATH.rsplit("/", 1)[0]
    base = f"https://graph.microsoft.com/v1.0/drives/{DRIVE_ID}/root:/"
    r = request_with_retry("GET", base + folder, headers=headers)
    if r.status_code != 200:
        fail(f"SharePoint folder '{folder}': {r.status_code} — {r.text}")
    print(f"SharePoint folder OK: {folder}")
    r = request_with_retry("GET", base + FILE_PATH, headers=headers)
    if r.status_code == 200:
        print(f"'{FILE_PATH}' exists ({r.json().get('size', 0) / 1024:.1f} KB) — a real run overwrites it")
    elif r.status_code == 404:
        print(f"'{FILE_PATH}' does not exist yet — the first real run creates it")
    else:
        fail(f"SharePoint file check: {r.status_code} — {r.text}")


def check_stored(item, expected_size, started):
    """Did SharePoint keep this run's workbook? (None if yes, else why not.)

    Not an exact size match: SharePoint writes its own metadata into Office
    files as it stores them, so the stored file is a few KB larger than the
    one uploaded (measured: 3,003,318 bytes stored for 2,994,000 sent). What
    is checked instead is that the stored file was modified by this upload
    and is about the size sent — a stale or truncated file fails both.
    """
    stored = item.get("size")
    if not isinstance(stored, int) or not (0.95 * expected_size <= stored <= 1.10 * expected_size + 64 * 1024):
        return f"SharePoint stored {stored} bytes for a {expected_size}-byte workbook"
    modified = item.get("lastModifiedDateTime")
    if not modified:
        return "SharePoint returned no modification time"
    if pd.Timestamp(modified) < pd.Timestamp(started) - pd.Timedelta(minutes=2):
        return f"SharePoint's copy was last modified {modified}, before this upload"
    return None


def upload(workbook_bytes, token, expected_size):
    url = f"https://graph.microsoft.com/v1.0/drives/{DRIVE_ID}/root:/{FILE_PATH}:/content"
    started = datetime.now(timezone.utc)
    r = request_with_retry(
        "PUT", url,
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        },
        data=workbook_bytes,
        timeout=UPLOAD_TIMEOUT,
    )
    if r.status_code not in (200, 201):
        fail(f"upload: {r.status_code} — {r.text}")
    # Graph answers with the stored file; make sure it is this run's.
    problem = check_stored(r.json() or {}, expected_size, started)
    if problem:
        fail(f"upload: {problem}")
    return r.status_code


# ======================================================
# WORKBOOK
# ======================================================
def add_stage_attributes(df, stages):
    """The stage's pipeline order and closed flag on every row, so the funnel
    can be sorted and filtered without a lookup table."""
    info = {s["id"]: s for s in stages}
    ids = df["dealstage__id"]
    df["dealstage__order"] = ids.map(lambda i: info[i].get("displayOrder") if i in info else None)
    df["dealstage__closed"] = ids.map(
        lambda i: info[i]["label"] in CLOSED_STAGE_LABELS if i in info else None)
    return df


def close_date_for(props, stage_label, is_closed):
    """(Close Date, where it came from) for one deal; (None, None) if open.

    HubSpot never sets closedate in this pipeline, so a closed deal is dated
    from the firm's own fields, the one that matches how it closed:

      Settled stages   Date - Settled, the date that confirms the settlement
      Referred Out     Date - Referred Out
      any other closed Date - Closed Out, then Date - Close Out After Retained

    and, for a closed deal whose field is blank, the day it entered its
    current stage — HubSpot stamps that on every move, so a closed row is
    never left without a date. The second column names the field used, so a
    row dated by the fallback can be told apart from one dated by the firm.
    """
    if not is_closed:
        return None, None
    if stage_label.lower().startswith("settled"):
        preferred = ["date___settled"]
    elif stage_label.lower().startswith("referred out"):
        preferred = ["date___referred_out"]
    else:
        preferred = ["date___dropped", "date__intake_sign_up_close_out"]
    for name in preferred:
        if props.get(name):
            return parse_hubspot_date(props[name]), name
    entered = parse_hubspot_datetime(props.get("hs_v2_date_entered_current_stage"))
    if entered:
        return entered.date(), "hs_v2_date_entered_current_stage"
    return None, None


def check_closed_stages(stages):
    """Every closed label must still exist in the pipeline, or the run stops:
    a renamed stage would otherwise quietly become "open" and lose its dates."""
    missing = CLOSED_STAGE_LABELS - {s["label"] for s in stages}
    if missing:
        fail(f"closed stage(s) not in the pipeline any more: {', '.join(sorted(missing))} "
             "— update CLOSED_STAGE_LABELS in main.py.")


def add_close_date(df, deals, stages):
    """Close Date and Close Date Source for every row.

    Relies on df having one row per deal in the order of `deals`, which is
    how build_deals_frame builds it.
    """
    info = {s["id"]: s for s in stages}
    dates, sources = [], []
    for deal in deals:
        p = deal.get("properties", {})
        stage = info.get(p.get("dealstage"), {})
        closed = stage.get("label") in CLOSED_STAGE_LABELS
        d, src = close_date_for(p, stage.get("label", ""), closed)
        dates.append(d)
        sources.append(src)
    df["close_date"] = dates
    df["close_date_source"] = sources
    return df


def build_workbook(df_deals):
    """Write the sheet with xlsxwriter, row by row.

    About 4x faster than pandas + openpyxl on this sheet (18,000 x 67: ~6s
    against ~25s, measured) and a smaller file. Row by row is what lets
    constant_memory stream each row out as it is written — do NOT swap this
    for pandas' to_excel with constant_memory: pandas writes column by column,
    and in that mode xlsxwriter silently drops everything but the first
    column (measured: a 100 KB file).

    Every text value goes through write_string, so a value that starts with
    "=" stays text and is never evaluated as a formula.
    """
    buf = io.BytesIO()
    wb = xlsxwriter.Workbook(buf, {"in_memory": True, "constant_memory": True})
    ws = wb.add_worksheet(DEALS_SHEET)
    fmt_datetime = wb.add_format({"num_format": "yyyy-mm-dd hh:mm:ss"})
    fmt_date = wb.add_format({"num_format": "yyyy-mm-dd"})
    ws.write_row(0, 0, [str(c) for c in df_deals.columns])
    values = df_deals.astype(object).where(df_deals.notna(), None)
    for r, row in enumerate(values.itertuples(index=False, name=None), start=1):
        for c, v in enumerate(row):
            if v is None:
                continue
            if isinstance(v, datetime):
                ws.write_datetime(r, c, v, fmt_datetime)
            elif isinstance(v, date):
                ws.write_datetime(r, c, datetime(v.year, v.month, v.day), fmt_date)
            elif isinstance(v, bool):
                ws.write_boolean(r, c, v)
            elif isinstance(v, (int, float)):
                ws.write_number(r, c, v)
            else:
                ws.write_string(r, c, str(v))
    wb.close()
    return buf.getvalue()


def main():
    env = {k: os.environ.get(k) for k in
           ("HUBSPOT_TOKEN", "AZURE_TENANT_ID", "AZURE_CLIENT_ID", "AZURE_CLIENT_SECRET")}
    missing = [k for k, v in env.items() if not v]
    if missing:
        fail(f"Missing environment variables: {', '.join(missing)}")
    print("All credentials loaded OK")
    if DRY_RUN:
        print("DRY RUN: nothing will be written to SharePoint")

    now_pacific = datetime.now(timezone.utc).astimezone(PACIFIC)
    print(f"Run started: {now_pacific.strftime('%Y-%m-%d %I:%M %p %Z')}")

    # The window moves on its own: 1 January of the year the run happens in.
    # It is cut on the same California clock the deal dates are written in,
    # so the sheet's first Create Date is 1 January and every month in Power
    # BI is complete.
    now_deal_tz = now_pacific.astimezone(DEAL_TZ)
    year_start = DEAL_TZ.localize(datetime(now_deal_tz.year, 1, 1))
    windows = month_windows(year_start, now_deal_tz)

    hs_headers = {
        "Authorization": f"Bearer {env['HUBSPOT_TOKEN']}",
        "Content-Type": "application/json",
        "Accept": "application/json",
    }

    pipeline_label, stages = fetch_pipeline(hs_headers)
    print(f"Pipeline '{pipeline_label}': {len(stages)} stages")
    check_closed_stages(stages)
    stage_labels = {s["id"]: s["label"] for s in stages}
    stage_candidates = [f"hs_v2_date_{kind}_{s['id']}"
                        for s in stages for kind in ("entered", "exited")]
    definitions = fetch_property_definitions(
        list(dict.fromkeys(BASE_PROPERTIES + stage_candidates)), hs_headers)
    stage_columns, skipped = stage_date_properties([s["id"] for s in stages], definitions)
    # A stage date already in the fixed list (New File Set Up's exit date)
    # must not come out as a second, identical column.
    stage_columns = [c for c in stage_columns if c not in BASE_PROPERTIES]
    if skipped:
        print(f"Stages with no entered/exited date property in HubSpot (no column): "
              f"{', '.join(stage_labels[i] for i in skipped)}")
    properties = BASE_PROPERTIES + stage_columns
    owners = fetch_owners(hs_headers)
    print(f"Owners loaded: {len(owners)}")

    print(f"Pulling deals created {year_start:%Y-%m-%d} to now, one month at a time:")
    t0 = time.monotonic()
    deals = fetch_deals(properties, windows, hs_headers)
    print(f"Total deals: {len(deals)} ({time.monotonic() - t0:.0f}s)")
    if not deals:
        fail("HubSpot returned no deals — refusing to write an empty report.")

    # One row per deal created this year, every one dated. The windows already
    # guarantee it; checking the rows themselves means a filter that ever
    # stops meaning what we think fails here instead of shipping.
    created = [parse_hubspot_datetime(d.get("properties", {}).get("createdate")) for d in deals]
    undated = sum(c is None for c in created)
    outside = sum(c is not None and not (year_start.replace(tzinfo=None) <= c <= now_deal_tz.replace(tzinfo=None))
                  for c in created)
    if undated or outside:
        fail(f"{undated} deals without a create date and {outside} created outside "
             f"{year_start:%Y-%m-%d} to now — nothing written.")
    print(f"Every deal created {year_start:%Y-%m-%d} or later, all {len(deals)} dated")

    owner_names = {k: v["name"] for k, v in owners.items()}
    df_deals = build_deals_frame(deals, properties, definitions, stage_labels,
                                 owner_names, pipeline_label)
    df_deals = add_stage_attributes(df_deals, stages)
    df_deals = add_close_date(df_deals, deals, stages)
    df_deals = add_durations(df_deals)
    refreshed = now_pacific.replace(tzinfo=None, microsecond=0)
    df_deals["last_refresh"] = refreshed
    df_deals = to_sheet(df_deals, definitions, stage_columns)

    t0 = time.monotonic()
    workbook = build_workbook(df_deals)
    size_kb = len(workbook) / 1024
    print(f"Workbook built: {size_kb:.0f} KB ({time.monotonic() - t0:.0f}s)")

    token = graph_token(env["AZURE_TENANT_ID"], env["AZURE_CLIENT_ID"], env["AZURE_CLIENT_SECRET"])
    print("Graph token obtained OK")

    if DRY_RUN:
        # A dry run still proves Graph can see the destination, so the first
        # real run is not where a wrong drive, folder or permission shows up.
        check_destination(token)
        # Counts only — deal names are client data and must not reach CI logs.
        print(f"DRY RUN: would upload {size_kb:.1f} KB to {FILE_PATH}")
        print(f"DRY RUN: one sheet '{DEALS_SHEET}', {len(df_deals)} rows x {len(df_deals.columns)} columns")
        # Fill rate per column — counts, no client data — so a dry run shows
        # which columns HubSpot actually populates.
        filled = df_deals.notna().sum()
        closed = df_deals["Deal Stage Is Closed"] == True  # noqa: E712

        print(f"DRY RUN: closed deals {int(closed.sum())}, with Close Date "
              f"{int(df_deals.loc[closed, 'Close Date'].notna().sum())}; by source:")
        for src, n in df_deals.loc[closed, "Close Date Source"].value_counts(dropna=False).items():
            print(f"    {n:>6}  {src}")
        print("DRY RUN: cycle times (days) — rows, median, negative:")
        for _, header, _, _ in DURATIONS:
            col = df_deals[header].dropna()
            median = f"{col.median():.0f}" if len(col) else "-"
            print(f"    {header}: {len(col)} rows, median {median}, negative {int((col < 0).sum())}")
        print("DRY RUN: AB 1755 by manufacturer:")
        for v, n in df_deals[AB1755_HEADER].value_counts(dropna=False).items():
            print(f"    {n:>6}  {v}")
        print("DRY RUN: rows with a value, per column:")
        for col, n in filled.items():
            print(f"    {n:>6}  {col}")
        return

    status = upload(workbook, token, len(workbook))
    print(f"File {'created' if status == 201 else 'uploaded'} ({size_kb:.1f} KB): {FILE_PATH}")
    print(f"Rows written: {len(df_deals)} x {len(df_deals.columns)} columns")
    print(f"Last Refresh: {refreshed:%B %d, %Y at %I:%M %p} Pacific")


if __name__ == "__main__":
    main()
