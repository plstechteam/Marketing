"""
Marketing report — GitHub Actions version
Pulls every Lemon Law deal ever created (the firm's HubSpot history starts
in April 2021) and writes it to the Deals tab of Marketing.xlsx in
SharePoint via Microsoft Graph. The workbook is the source for a Power BI funnel report; every funnel
calculation lives in Power BI, none of it here.
"""

import gzip
import io
import json
import os
import sys
import threading
import time
import zipfile
from concurrent.futures import ThreadPoolExecutor
from datetime import date, datetime, timezone

import pandas as pd
import pytz
import requests
import xlsxwriter

import splice

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
# One-off: let a run change a Deals column other tabs read (e.g. to move
# columns back under formulas that were built for them). Never left on.
# Read every call again instead of trusting the calls cache (it is rebuilt).
REBUILD_CALLS_CACHE = os.environ.get("REBUILD_CALLS_CACHE", "").strip().lower() in ("1", "true", "yes")
ACCEPT_COLUMN_CHANGES = os.environ.get("ACCEPT_COLUMN_CHANGES", "").strip().lower() in ("1", "true", "yes")

PIPELINE_ID = "default"          # Lemon Law. Employment Law is out of scope.

# Lemon Law - State values kept. A handful of deals in the Lemon Law pipeline
# are marked "Employment Law" (12 in September 2026); they are not Lemon Law
# cases and are dropped, as the Monthly Settlement Report drops them. Stored
# values, not labels.
ALLOWED_STATES = {"LEMON LAW (CA)", "LEMON LAW (WA)"}

# The whole history is pulled — marketing is compared year against year — so
# the pull starts at the firm's first Lemon Law deal in HubSpot (18 April
# 2021) and runs to now. Monthly windows from HISTORY_START; one extra window
# below it, back to HISTORY_FLOOR, catches any deal imported with an older
# create date, so "every deal" never depends on that first date staying true.
HISTORY_START_YEAR = 2021
HISTORY_FLOOR_YEAR = 2000

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
    "date___inquiry_qualified",     # Date - Inquiry Qualified: intake, before 2026
    "date___ro_review",             # Date - RO Review
    "date___retained",              # Date - Retainer Signed
    "hs_v2_date_exited_5792630",    # left New File Set Up = ready for Legal
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
    """Opt In / Opt Out / Not on list for a stored manufacturer value. A deal
    with no manufacturer is "Not on list" too — the column is never blank."""
    if not manufacturer:
        return "Not on list"
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

# HubSpot allows a private app a number of requests per rolling 10 seconds
# (100 to 190 by tier; the first response says which, in its
# X-HubSpot-RateLimit-* headers) and its search endpoints 5 per second on top
# of that. Every HubSpot request, from any thread, waits for a slot on the
# matching pacer so parallel reads stay under both; a 429 still gets retried.
HUBSPOT_HOST         = "api.hubapi.com"
HUBSPOT_MAX_PER_SEC  = 8     # until the rate-limit headers say otherwise
HUBSPOT_SEARCH_PER_SEC = 4
RATE_LIMIT_SHARE     = 0.8   # of the advertised limit — headroom for retries
RATE_LIMIT_ATTEMPTS  = 8     # 429s clear once the 10-second window rolls
RATE_LIMIT_DELAY     = 10

session = requests.Session()
session.mount("https://", requests.adapters.HTTPAdapter(pool_maxsize=32))


class Pacer:
    """Hands out request start times no closer than 1/rate apart, across threads."""

    def __init__(self, per_sec):
        self.per_sec = per_sec
        self._lock = threading.Lock()
        self._next = 0.0

    def wait(self):
        with self._lock:
            now = time.monotonic()
            start = max(now, self._next)
            self._next = start + 1.0 / self.per_sec
        if start > now:
            time.sleep(start - now)


HUBSPOT_PACER = Pacer(HUBSPOT_MAX_PER_SEC)
SEARCH_PACER = Pacer(HUBSPOT_SEARCH_PER_SEC)
_limit_read = [False]


def _adopt_rate_limit(response):
    """Set the general pacer from the first X-HubSpot-RateLimit-* headers."""
    if _limit_read[0]:
        return
    try:
        cap = int(response.headers["X-HubSpot-RateLimit-Max"])
        interval = int(response.headers["X-HubSpot-RateLimit-Interval-Milliseconds"]) / 1000
    except (KeyError, ValueError):
        return
    _limit_read[0] = True
    if cap > 0 and interval > 0:
        HUBSPOT_PACER.per_sec = max(1.0, cap / interval * RATE_LIMIT_SHARE)
        print(f"HubSpot rate limit: {cap} per {interval:.0f}s — pacing at "
              f"{HUBSPOT_PACER.per_sec:.1f} requests/s")


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
    attempt = 0
    while True:
        attempt += 1
        hubspot = HUBSPOT_HOST in url
        if hubspot:
            (SEARCH_PACER if url.rstrip("/").endswith("/search") else HUBSPOT_PACER).wait()
        try:
            response = session.request(method, url, timeout=timeout, **kwargs)
        except requests.RequestException as exc:
            response = None
            reason = f"network error: {exc}"
        else:
            if hubspot:
                _adopt_rate_limit(response)
            if (response.status_code not in RETRY_STATUS
                    and response.status_code != LOCKED_STATUS):
                return response
            reason = f"HTTP {response.status_code}"

        limit = RATE_LIMIT_ATTEMPTS if response is not None and response.status_code == 429 else MAX_ATTEMPTS
        if attempt >= limit:
            break
        wait = _retry_delay(response, attempt)
        if response is not None and response.status_code == 429:
            wait = max(wait, RATE_LIMIT_DELAY)
        print(f"  {method} failed ({reason}) — attempt {attempt}/{limit}, retrying in {wait:.0f}s")
        time.sleep(wait)

    if response is None:
        print(f"ERROR: {method} {url.split('?')[0]} failed after {attempt} attempts — {reason}")
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


# The milestones the cycle times run between. Intake is the day the deal
# entered the Intake stage; Ready for Legal is the day it left New File Set
# Up - Doc Collection, i.e. the signed case handed to Legal.
INTAKE_ENTERED = "hs_v2_date_entered_5411633"
READY_FOR_LEGAL = "hs_v2_date_exited_5792630"

# The Intake date the sheet reports. HubSpot's own stamp for entering the
# Intake stage is the one to use, but it only covers deals that passed
# through the stage since it was set up that way; before 2026 the firm
# recorded intake in "Date - Inquiry Qualified" instead, and that property
# was never back-filled into the stage stamp (40,567 deals created before
# 2026 carry it; none created in 2026 do). So the stage stamp always wins,
# and Inquiry Qualified only fills the rows where it is blank.
INTAKE_LEGACY = "date___inquiry_qualified"
INTAKE_DATE = "intake_date"

# Cycle times, in calendar days, computed on every run from that run's
# dates — so a date corrected in HubSpot corrects its duration on the next
# run. (key, header, from date, to date).
DURATIONS = [
    ("days_created_to_intake", "Days: Created to Intake",
     "createdate", INTAKE_DATE),
    ("days_created_to_ro_review", "Days: Created to RO Review",
     "createdate", "date___ro_review"),
    ("days_intake_to_signed", "Days: Intake to Retainer Signed",
     INTAKE_DATE, "date___retained"),
    ("days_ro_review_to_signed", "Days: RO Review to Retainer Signed",
     "date___ro_review", "date___retained"),
    ("days_signed_to_ready_for_legal", "Days: Retainer Signed to Ready for Legal",
     "date___retained", READY_FOR_LEGAL),
    ("days_ro_review_to_ready_for_legal", "Days: RO Review to Ready for Legal",
     "date___ro_review", READY_FOR_LEGAL),
    ("days_intake_to_ready_for_legal", "Days: Intake to Ready for Legal",
     INTAKE_DATE, READY_FOR_LEGAL),
    ("days_created_to_ready_for_legal", "Days: Created to Ready for Legal",
     "createdate", READY_FOR_LEGAL),
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


def inbound_call_for(dealname, aircall_entry_number):
    """Yes when the deal was created by an inbound call to an Aircall line.

    An inbound Aircall call creates the deal through the "[2026] Deal
    Creation" automation under the name "Aircall new contact, +1…", and
    AirCall Entry Number holds the firm's line the call came in on. The name
    is usually replaced once intake knows the client, so either signal is
    enough. Outbound cannot be read from the deal; see probe_calls_access.
    """
    name = str(dealname or "").strip().lower()
    if name.startswith("aircall new contact") or aircall_entry_number:
        return "Yes"
    return "No"


def add_inbound_call(df):
    names = df["dealname"] if "dealname" in df.columns else pd.Series(None, index=df.index)
    lines = df["aircall_entry_number"] if "aircall_entry_number" in df.columns \
        else pd.Series(None, index=df.index)
    df["created_by_inbound_call"] = [
        inbound_call_for(n, l if l is not None and not pd.isna(l) else None)
        for n, l in zip(names.tolist(), lines.tolist())]
    return df


def add_intake_date(df):
    """Date - Intake: the Intake stage stamp, or Inquiry Qualified where the
    stamp is blank — never the other way round — plus where it came from."""
    new = df[INTAKE_ENTERED] if INTAKE_ENTERED in df.columns else pd.Series(None, index=df.index)
    old = df[INTAKE_LEGACY] if INTAKE_LEGACY in df.columns else pd.Series(None, index=df.index)
    # Value by value, so each keeps its own type — a stamp stays a datetime,
    # an Inquiry Qualified stays a date — instead of pandas coercing the lot.
    dates, sources = [], []
    for stamp, legacy in zip(new.tolist(), old.tolist()):
        if stamp is not None and not pd.isna(stamp):
            dates.append(stamp.to_pydatetime() if isinstance(stamp, pd.Timestamp) else stamp)
            sources.append(INTAKE_ENTERED)
        elif legacy is not None and not pd.isna(legacy):
            dates.append(legacy)
            sources.append(INTAKE_LEGACY)
        else:
            dates.append(None)
            sources.append(None)
    df[INTAKE_DATE] = pd.Series(dates, index=df.index, dtype=object)
    df["intake_date_source"] = pd.Series(sources, index=df.index, dtype=object)
    return df


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
    # Fixed rather than HubSpot's label: the dry-run summary groups by it.
    "dealstage": "Deal Stage",
    "dealstage__id": "Deal Stage ID",
    "dealstage__order": "Deal Stage Order",
    "dealstage__closed": "Is Closed",
    "is_settled": "Is Settled",
    "close_date": "Close Date",
    "close_date_source": "Close Date Source",
    "s__manufacturer__ab1755": AB1755_HEADER,
    "hubspot_owner_id": "Deal Owner",
    "hubspot_owner_id__id": "Deal Owner ID",
    "last_refresh": "Last Refresh",
    INTAKE_DATE: "Date - Intake",
    "created_by_inbound_call": "Created by Inbound Call",
    "first_call_direction": "First Call Direction",
    "first_call_date": "First Call Date",
    "intake_date_source": "Intake Date Source",
    READY_FOR_LEGAL: "Date - Ready for Legal (Exited File Set Up)",
    **{key: header for key, header, _, _ in DURATIONS},
}

# Sheet layout, left to right in the order the funnel reads: who the deal is,
# where it stands now, where it came from, what vehicle and AB 1755 side it
# is on, the milestones on the way, how it ended, who is on it, and the full
# stage history. "STAGE_HISTORY" expands to one date column per pipeline
# stage, in pipeline order.
#
# Columns A..AI are frozen: the Maz tab reads Deals by letter (AC Date -
# Intake, AE Date - Retainer Signed, AF Ready for Legal, AI Date - Settled,
# X Manufacturer, Y AB 1755, plus A, C, D, E). Anything new goes in "Added
# later", at the far right.
COLUMN_GROUPS = [
    ("Deal", ["hs_object_id", "dealname", "legal_pipeline", "createdate"]),
    ("Current stage", ["dealstage", "dealstage__id", "dealstage__order", "dealstage__closed", "is_settled",
                       "hs_v2_date_entered_current_stage", "close_date", "close_date_source"]),
    ("Source / channel", ["lead___source", "lead___source__group_", "hs_analytics_source",
                          "hs_analytics_source_data_1", "hs_object_source_label",
                          "aircall_entry_number", "auto_dialer_call_type",
                          "tf__utm_source", "tf__utm_medium", "tf__utm_campaign", "gclid"]),
    ("Vehicle / AB 1755", ["s__manufacturer", "s__manufacturer__ab1755",
                           "ro_review__final_decision_", "vehicle___year",
                           "c__vehicle___model__new_test_"]),
    # In the order a case moves: intake, RO review, retainer signed, handed
    # to Legal. Date - Intake is the stage stamp back-filled with Inquiry
    # Qualified (see add_intake_date); both originals stay for audit — the
    # stamp in Stage history, Inquiry Qualified here.
    ("Milestones", [INTAKE_DATE, "date___ro_review",
                    "date___retained", READY_FOR_LEGAL, "date___referred_out"]),
    ("Outcome", ["case_category", "date___settled", "total_settled_attorneys_fees_and_cost",
                 "net_attorney_fees", "drop_reason", "date___dropped",
                 "date__intake_sign_up_close_out", "deal_stage___sub_phase"]),
    ("Cycle time (days)", [key for key, _, _, _ in DURATIONS]),
    ("People", ["hubspot_owner_id", "hubspot_owner_id__id", "n5__retainer_representative",
                "senior_case_supervisor", "handling_attorney", "supervising_attorney"]),
    ("Stage history", ["STAGE_HISTORY"]),
    ("Audit", ["pipeline", "hs_lastmodifieddate", "last_refresh"]),
    # Columns added after people started building on the workbook (formulas
    # on Maz read Deals by column letter). New columns go HERE, at the far
    # right, never in the groups above — inserting in the middle shifts every
    # letter after it. The run also refuses to write if a column another tab
    # reads would change header (see splice.shifted_references).
    ("Added later", ["created_by_inbound_call", "first_call_direction", "first_call_date",
                     "intake_date_source", INTAKE_LEGACY]),
]


def column_order(columns, stage_columns):
    """Internal column keys in sheet order. A column no group names (a
    property added to BASE_PROPERTIES but not placed) goes at the far right,
    after everything else, so it can never shift a column other tabs read."""
    ordered = []
    for group, keys in COLUMN_GROUPS:
        for key in keys:
            if key == "STAGE_HISTORY":
                ordered += [c for c in stage_columns if c in columns]
            elif key in columns:
                ordered.append(key)
    ordered = list(dict.fromkeys(ordered))
    return ordered + [c for c in columns if c not in set(ordered)]


def header_for(key, definitions):
    return FIXED_HEADERS.get(key) or (definitions.get(key, {}).get("label") or key).strip()


def to_sheet(df, definitions, stage_columns):
    """Order the columns by COLUMN_GROUPS and head them with labels. Close
    Date Source is written as the label of the field used, so the row says
    "Date - Settled" rather than an internal name."""
    df = df[column_order(list(df.columns), stage_columns)].copy()
    for source in ("close_date_source", "intake_date_source"):
        if source in df.columns:
            df[source] = df[source].map(lambda k: header_for(k, definitions) if k else None)
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


def probe_calls_access(headers):
    """Can this token read HubSpot calls? Reported only, never fatal.

    Telling an inbound deal from an outbound one needs the direction of the
    first call on the deal, which lives on the Calls object. Whether the
    private app's token may read it depends on its scopes, so every run
    says so in the log.
    """
    r = request_with_retry("GET", "https://api.hubapi.com/crm/v3/objects/calls",
                           headers=headers, params={"limit": 1, "properties": "hs_call_direction"})
    if r.status_code == 200:
        sample = (r.json().get("results") or [{}])[0].get("properties", {})
        print(f"Calls API: readable (sample hs_call_direction: {sample.get('hs_call_direction')!r})")
    else:
        detail = ""
        try:
            detail = r.json().get("message", "")
        except ValueError:
            pass
        print(f"Calls API: not readable ({r.status_code}) {detail[:200]}")


CALL_READ_WORKERS = 10   # parallel reads; the pacers set the actual rate
CALLS_CACHE_VERSION = 1


def earliest_call(calls):
    """The earliest of some calls, each a tuple ending in its timestamp;
    calls without a timestamp only count when none has one."""
    dated = [c for c in calls if c[-1]]
    if not dated:
        return calls[0] if calls else (None, None)
    return min(dated, key=lambda c: pd.Timestamp(c[-1]))


def encode_calls_cache(seen, first):
    """gzip JSON: every call id already read (sorted, stored as gaps, which
    keeps 1.5 million ids to a few MB) and each deal's first call as
    [call id, direction, timestamp]. Ids only — no names or client data."""
    ids = sorted(seen)
    gaps = [b - a for a, b in zip([0] + ids[:-1], ids)]
    body = {"version": CALLS_CACHE_VERSION, "seen_gaps": gaps,
            "first": {d: list(v) for d, v in first.items()}}
    return gzip.compress(json.dumps(body, separators=(",", ":")).encode(), compresslevel=6)


def decode_calls_cache(raw):
    """(seen, first) from encode_calls_cache's bytes, or None if unusable."""
    try:
        body = json.loads(gzip.decompress(raw))
        if body.get("version") != CALLS_CACHE_VERSION:
            return None
        seen, total = set(), 0
        for gap in body["seen_gaps"]:
            total += gap
            seen.add(total)
        first = {str(d): (int(v[0]), v[1], v[2]) for d, v in body["first"].items()}
        return seen, first
    except (OSError, ValueError, KeyError, TypeError, IndexError):
        return None


def plan_call_reads(deal_calls, seen, first):
    """{deal: [call ids to read]} — only what the cache cannot answer.

    A deal with a cached first call that is still linked to it needs only
    its calls not read before; any other deal with calls (new, or its cached
    first call unlinked or deleted) has all of its calls read."""
    plan = {}
    for deal, ids in deal_calls.items():
        cached = first.get(deal)
        if cached is not None and cached[0] in ids:
            need = [c for c in ids if c not in seen]
        else:
            need = list(ids)
        if need:
            plan[deal] = need
    return plan


def merge_first_calls(deal_calls, seen, first, plan, read):
    """(first calls {deal: (id, direction, timestamp)}, seen ids to cache).

    Each deal's first call is the earliest of its cached first call (when
    still linked) and the calls just read for it."""
    merged = {}
    for deal, ids in deal_calls.items():
        cached = first.get(deal)
        candidates = [cached] if cached is not None and cached[0] in ids else []
        candidates += [(c, *read[c]) for c in plan.get(deal, []) if c in read]
        if candidates:
            merged[deal] = earliest_call(candidates)
    linked = {c for ids in deal_calls.values() for c in ids}
    return merged, (seen | set(read)) & linked


def _parallel(fn, items, label):
    """fn over items on CALL_READ_WORKERS threads, in order; a RuntimeError
    from any of them fails the run."""
    out = []
    with ThreadPoolExecutor(max_workers=CALL_READ_WORKERS) as pool:
        try:
            for n, part in enumerate(pool.map(fn, items), 1):
                out.append(part)
                if n % 2000 == 0:
                    print(f"  {label}: {n} of {len(items)}")
        except RuntimeError as exc:
            fail(str(exc))
    return out


def fetch_deal_calls(deal_ids, headers):
    """{deal id: [call ids]} for every deal with calls, 1,000 deals per request."""
    base = "https://api.hubapi.com"

    def read(chunk):
        r = request_with_retry("POST", f"{base}/crm/v4/associations/deals/calls/batch/read",
                               headers=headers, json={"inputs": [{"id": d} for d in chunk]})
        if r.status_code not in (200, 207):
            raise RuntimeError(f"deal→call associations: {r.status_code} — {r.text[:300]}")
        out = {}
        for item in r.json().get("results", []):
            deal = str(item["from"]["id"])
            ids = [int(t["toObjectId"]) for t in item.get("to", [])]
            after = (item.get("paging") or {}).get("next", {}).get("after")
            while after:   # a deal with more calls than one page
                rr = request_with_retry(
                    "GET", f"{base}/crm/v4/objects/deals/{deal}/associations/calls",
                    headers=headers, params={"limit": 500, "after": after})
                if rr.status_code != 200:
                    raise RuntimeError(f"deal→call associations (next page): "
                                       f"{rr.status_code} — {rr.text[:300]}")
                body = rr.json()
                ids += [int(t["toObjectId"]) for t in body.get("results", [])]
                after = (body.get("paging") or {}).get("next", {}).get("after")
            if ids:
                out[deal] = ids
        return out

    chunks = [deal_ids[i:i + 1000] for i in range(0, len(deal_ids), 1000)]
    deal_calls = {}
    for part in _parallel(read, chunks, "association batches"):
        deal_calls.update(part)
    return deal_calls


def read_calls(call_ids, headers):
    """{call id: (direction, timestamp)}, 100 calls per request."""
    url = "https://api.hubapi.com/crm/v3/objects/calls/batch/read"

    def read(batch):
        r = request_with_retry("POST", url, headers=headers,
                               json={"inputs": [{"id": str(c)} for c in batch],
                                     "properties": ["hs_call_direction", "hs_timestamp", "hs_createdate"]})
        if r.status_code not in (200, 207):
            raise RuntimeError(f"calls batch read: {r.status_code} — {r.text[:300]}")
        out = {}
        for c in r.json().get("results", []):
            p = c.get("properties", {})
            out[int(c["id"])] = (p.get("hs_call_direction"), p.get("hs_timestamp") or p.get("hs_createdate"))
        return out

    ids = sorted(call_ids)
    calls = {}
    for part in _parallel(read, [ids[i:i + 100] for i in range(0, len(ids), 100)], "call batches"):
        calls.update(part)
    return calls


def fetch_first_calls(deal_ids, headers, cache):
    """({deal id: (direction, timestamp)} of each deal's earliest call, the
    new cache as (seen, first)).

    Aircall logs every call on HubSpot's Calls object with hs_call_direction
    (INBOUND / OUTBOUND). Call ids do not follow call time — the lowest id was
    the earliest call for only 2 of 3 deals — so the earliest call has to
    come from the timestamps of every call on the deal. Reading all ~1.5
    million takes half an hour, so the calls already read, and each deal's
    first call, are kept in a cache next to the workbook: a run reads the
    associations (all of them, every run) and then only the calls it has not
    seen. Without a cache (first run, or a rebuild) every call is read.
    """
    seen, first = cache if cache else (set(), {})
    deal_calls = fetch_deal_calls(deal_ids, headers)
    plan = plan_call_reads(deal_calls, seen, first)
    wanted = {c for ids in plan.values() for c in ids}
    linked = sum(len(ids) for ids in deal_calls.values())
    print(f"  {linked} calls linked to {len(deal_calls)} deals; "
          f"{len(wanted)} not in the cache — reading them")
    read = read_calls(wanted, headers) if wanted else {}
    merged, seen = merge_first_calls(deal_calls, seen, first, plan, read)
    return {d: (v[1], v[2]) for d, v in merged.items()}, (seen, merged)


def add_first_call(df, deals, first_calls):
    """First Call Direction (Inbound / Outbound / Unknown / No calls) and First
    Call Date, one per row, in the order of `deals`."""
    names = {"INBOUND": "Inbound", "OUTBOUND": "Outbound"}
    directions, dates = [], []
    for deal in deals:
        hit = first_calls.get(str(deal["id"]))
        if hit is None:
            directions.append("No calls")
            dates.append(None)
        else:
            direction, ts = hit
            directions.append(names.get(direction, "Unknown"))
            dates.append(parse_hubspot_datetime(ts))
    df["first_call_direction"] = directions
    df["first_call_date"] = pd.Series(dates, index=df.index, dtype=object)
    return df


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


def window_label(start, end):
    """A readable name for a window: the month when it is exactly one,
    otherwise the span (the floor window, a split half, the current month)."""
    if start.year < HISTORY_START_YEAR:
        return f"before {end:%Y-%m-%d}"
    tz = start.tzinfo
    if start.day == 1 and start.hour == 0 and start.minute == 0:
        nxt = tz.localize(datetime(start.year + start.month // 12, start.month % 12 + 1, 1))
        if end == nxt:
            return start.strftime("%Y-%m")
    return f"{start:%Y-%m-%d %H:%M}..{end:%Y-%m-%d %H:%M}"


DEAL_SEARCH_WORKERS = 4   # months pulled at once; SEARCH_PACER sets the rate


def fetch_window(properties, start, end, headers):
    """Every deal created in [start, end), paged to the end.

    A window whose HubSpot total reaches the search API's 10,000-result
    ceiling is split in half and each half pulled on its own, as often as it
    takes — busy months in 2023 run past 5,000 deals, and one twice that
    would otherwise be silently truncated.

    Properties HubSpot returns empty are dropped as they arrive: with the
    full history (~230,000 deals) keeping every null would cost gigabytes.
    """
    url = "https://api.hubapi.com/crm/v3/objects/deals/search"
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
    rows = []
    total = None
    after = None
    while True:
        if after:
            payload["after"] = after
        r = request_with_retry("POST", url, headers=headers, json=payload)
        if r.status_code != 200:
            raise RuntimeError(f"deal search: {r.status_code} — {r.text}")
        data = r.json()
        total = data.get("total", total)
        if total is not None and total >= SEARCH_API_MAX_RESULTS and after is None:
            if end - start < pd.Timedelta(minutes=1):
                raise RuntimeError(f"{window_label(start, end)}: {total} deals in under a minute — "
                                   "cannot split further; nothing written.")
            middle = start + (end - start) / 2
            print(f"  {window_label(start, end)}: {total} deals — splitting in two")
            return (fetch_window(properties, start, middle, headers)
                    + fetch_window(properties, middle, end, headers))
        for deal in data.get("results", []):
            props = {k: v for k, v in deal.get("properties", {}).items() if v not in (None, "")}
            rows.append({"id": deal["id"], "properties": props})
        after = data.get("paging", {}).get("next", {}).get("after")
        if not after:
            break

    print(f"  {window_label(start, end)}: {len(rows)} deals (HubSpot total {total})")
    if len(rows) >= SEARCH_API_MAX_RESULTS:
        raise RuntimeError(f"{window_label(start, end)} reached the {SEARCH_API_MAX_RESULTS:,}-result "
                           "search ceiling — the pull would be truncated; nothing written.")
    if total is not None and len(rows) < total:
        raise RuntimeError(f"{window_label(start, end)}: downloaded {len(rows)} of {total} — "
                           "incomplete pull, nothing written.")
    return rows


def fetch_deals(properties, windows, headers):
    """Every deal created in the windows, several windows at a time, in
    window order (then by id), each deal once."""
    deals = {}
    with ThreadPoolExecutor(max_workers=DEAL_SEARCH_WORKERS) as pool:
        try:
            parts = list(pool.map(lambda w: fetch_window(properties, w[0], w[1], headers), windows))
        except RuntimeError as exc:
            fail(str(exc))
    for rows in parts:
        for deal in rows:
            deals[deal["id"]] = deal
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


def fetch_existing(token):
    """(bytes, eTag) of the workbook in SharePoint, or (None, None) if it does
    not exist yet. Also proves the folder is reachable."""
    headers = {"Authorization": f"Bearer {token}"}
    base = f"https://graph.microsoft.com/v1.0/drives/{DRIVE_ID}/root:/"
    folder = FILE_PATH.rsplit("/", 1)[0]
    r = request_with_retry("GET", base + folder, headers=headers)
    if r.status_code != 200:
        fail(f"SharePoint folder '{folder}': {r.status_code} — {r.text}")
    r = request_with_retry("GET", base + FILE_PATH, headers=headers)
    if r.status_code == 404:
        return None, None
    if r.status_code != 200:
        fail(f"SharePoint file check: {r.status_code} — {r.text}")
    etag = r.json().get("eTag")
    r = request_with_retry("GET", base + FILE_PATH + ":/content", headers=headers,
                           timeout=UPLOAD_TIMEOUT)
    if r.status_code != 200:
        fail(f"SharePoint download: {r.status_code} — {r.text}")
    return r.content, etag


def calls_cache_path():
    """The calls cache sits next to the workbook: Marketing_calls_cache.json.gz."""
    return FILE_PATH.rsplit(".", 1)[0] + "_calls_cache.json.gz"


def fetch_calls_cache(token):
    """(seen, first) from SharePoint, or None when missing or unreadable."""
    url = f"https://graph.microsoft.com/v1.0/drives/{DRIVE_ID}/root:/{calls_cache_path()}:/content"
    r = request_with_retry("GET", url, headers={"Authorization": f"Bearer {token}"},
                           timeout=UPLOAD_TIMEOUT)
    if r.status_code == 404:
        print("Calls cache: none yet — every call will be read (about 30 minutes, once)")
        return None
    if r.status_code != 200:
        print(f"Calls cache: could not read it ({r.status_code}) — reading every call")
        return None
    cache = decode_calls_cache(r.content)
    if cache is None:
        print("Calls cache: unreadable or an old format — reading every call")
    else:
        print(f"Calls cache: {len(cache[0])} calls and {len(cache[1])} deals' first calls "
              f"({len(r.content) / 1024:.0f} KB)")
    return cache


def upload_calls_cache(token, seen, first):
    raw = encode_calls_cache(seen, first)
    url = f"https://graph.microsoft.com/v1.0/drives/{DRIVE_ID}/root:/{calls_cache_path()}:/content"
    r = request_with_retry("PUT", url, headers={"Authorization": f"Bearer {token}",
                                                "Content-Type": "application/gzip"},
                           data=raw, timeout=UPLOAD_TIMEOUT)
    if r.status_code not in (200, 201):
        # Not fatal: the next run just reads more calls.
        print(f"Calls cache: upload failed ({r.status_code}) — the next run reads more calls")
        return
    print(f"Calls cache saved: {len(seen)} calls, {len(first)} deals ({len(raw) / 1024:.0f} KB)")


def other_sheet_parts(xlsx_bytes, report):
    """{part: bytes} for every worksheet part except the one this job owns."""
    with zipfile.ZipFile(io.BytesIO(xlsx_bytes)) as z:
        return {n: z.read(n) for n in z.namelist()
                if n.startswith("xl/worksheets/") and n.endswith(".xml")
                and n != report["sheet_part"]}


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


def upload(workbook_bytes, token, expected_size, etag=None):
    url = f"https://graph.microsoft.com/v1.0/drives/{DRIVE_ID}/root:/{FILE_PATH}:/content"
    started = datetime.now(timezone.utc)
    r = request_with_retry(
        "PUT", url,
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            # Only overwrite the version this run read. If someone saved the
            # file in between (a change to Maz, say), SharePoint answers 412
            # and their work stands; the next run starts from it.
            **({"If-Match": etag} if etag else {}),
        },
        data=workbook_bytes,
        timeout=UPLOAD_TIMEOUT,
    )
    if r.status_code == 412:
        fail("upload: the workbook changed in SharePoint while this run was working — "
             "nothing overwritten; the next run will pick up the new version.")
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
    # A deal with a settlement date is settled, whatever stage it sits in:
    # settled cases routinely stay in Retained - Lit / Retained - Pre Lit
    # (all 31 of September 2026's settlements created this year did), and
    # reading the stage alone showed them as open with no close date.
    settled = df["date___settled"].notna() if "date___settled" in df.columns \
        else pd.Series(False, index=df.index)
    df["is_settled"] = settled
    df["dealstage__closed"] = ids.map(
        lambda i: info[i]["label"] in CLOSED_STAGE_LABELS if i in info else False) | settled
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
    # Date - Settled is what confirms a settlement, so it dates the close
    # whenever it is filled — also for a settled deal still in a Retained
    # stage.
    if props.get("date___settled"):
        return parse_hubspot_date(props["date___settled"]), "date___settled"
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
        closed = stage.get("label") in CLOSED_STAGE_LABELS or bool(p.get("date___settled"))
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


def preview_other_tabs(xlsx_bytes, refs, max_row=12):
    """Dry run only: the labels and formulas at the top of each tab that reads
    Deals, so a log shows what those formulas were built to read. Prints
    literal text and formulas only — never a formula's cached result, which
    can be client data."""
    import openpyxl
    sheets = sorted({w.split("!")[0] for ws in refs.values() for w in ws if "!" in w})
    wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes), read_only=True, data_only=False)
    for name in sheets:
        if name not in wb.sheetnames:
            continue
        print(f"DRY RUN: top of '{name}' (labels and formulas only):")
        for row in wb[name].iter_rows(min_row=1, max_row=max_row):
            for cell in row:
                v = cell.value
                if isinstance(v, str) and v.strip():
                    print(f"    {name}!{cell.coordinate}: {v[:120]}")
    wb.close()


def main():
    env = {k: os.environ.get(k) for k in
           ("HUBSPOT_TOKEN", "AZURE_TENANT_ID", "AZURE_CLIENT_ID", "AZURE_CLIENT_SECRET")}
    missing = [k for k, v in env.items() if not v]
    if missing:
        fail(f"Missing environment variables: {', '.join(missing)}")
    print("All credentials loaded OK")
    if DRY_RUN:
        print("DRY RUN: Marketing.xlsx will not be written (the calls cache still is)")

    now_pacific = datetime.now(timezone.utc).astimezone(PACIFIC)
    print(f"Run started: {now_pacific.strftime('%Y-%m-%d %I:%M %p %Z')}")

    # Every Lemon Law deal ever created, cut into California-time months.
    now_deal_tz = now_pacific.astimezone(DEAL_TZ)
    history_floor = DEAL_TZ.localize(datetime(HISTORY_FLOOR_YEAR, 1, 1))
    history_start = DEAL_TZ.localize(datetime(HISTORY_START_YEAR, 1, 1))
    windows = [(history_floor, history_start)] + month_windows(history_start, now_deal_tz)

    hs_headers = {
        "Authorization": f"Bearer {env['HUBSPOT_TOKEN']}",
        "Content-Type": "application/json",
        "Accept": "application/json",
    }

    # SharePoint downloads (the 80 MB workbook, the calls cache) run in the
    # background while HubSpot is read.
    token = graph_token(env["AZURE_TENANT_ID"], env["AZURE_CLIENT_ID"], env["AZURE_CLIENT_SECRET"])
    print("Graph token obtained OK")
    background = ThreadPoolExecutor(max_workers=2)
    existing_job = background.submit(fetch_existing, token)
    if REBUILD_CALLS_CACHE:
        print("REBUILD_CALLS_CACHE is set: ignoring the calls cache, reading every call")
        cache_job = None
    else:
        cache_job = background.submit(fetch_calls_cache, token)

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
    probe_calls_access(hs_headers)
    owners = fetch_owners(hs_headers)
    print(f"Owners loaded: {len(owners)}")

    print("Pulling every Lemon Law deal ever created, one month at a time:")
    t0 = time.monotonic()
    deals = fetch_deals(properties, windows, hs_headers)
    print(f"Total deals: {len(deals)} ({time.monotonic() - t0:.0f}s)")
    if not deals:
        fail("HubSpot returned no deals — refusing to write an empty report.")

    # California and Washington only — see ALLOWED_STATES.
    kept, other = [], []
    for d in deals:
        state = d.get("properties", {}).get("legal_pipeline")
        if state in ALLOWED_STATES:
            kept.append(d)
        else:
            other.append(state or "(blank)")
    dropped = pd.Series(other, dtype=object).value_counts()
    if len(dropped):
        print(f"Dropped {int(dropped.sum())} deals outside California / Washington: "
              + ", ".join(f"{v}: {n}" for v, n in dropped.items()))
    deals = kept

    # One row per deal, every one dated, none outside the pulled range. The windows already
    # guarantee it; checking the rows themselves means a filter that ever
    # stops meaning what we think fails here instead of shipping.
    created = [parse_hubspot_datetime(d.get("properties", {}).get("createdate")) for d in deals]
    undated = sum(c is None for c in created)
    outside = sum(c is not None and not (history_floor.replace(tzinfo=None) <= c <= now_deal_tz.replace(tzinfo=None))
                  for c in created)
    if undated or outside:
        fail(f"{undated} deals without a create date and {outside} created outside "
             f"{history_floor:%Y-%m-%d} to now — nothing written.")
    by_year = pd.Series([c.year for c in created]).value_counts().sort_index()
    print(f"All {len(deals)} deals dated; by year created: "
          + ", ".join(f"{y}: {n}" for y, n in by_year.items()))

    t0 = time.monotonic()
    cache = cache_job.result() if cache_job else None
    first_calls, (seen, first) = fetch_first_calls([d["id"] for d in deals], hs_headers, cache)
    print(f"First calls: {len(first_calls)} of {len(deals)} deals have a call "
          f"({time.monotonic() - t0:.0f}s)")
    upload_calls_cache(token, seen, first)

    owner_names = {k: v["name"] for k, v in owners.items()}
    df_deals = build_deals_frame(deals, properties, definitions, stage_labels,
                                 owner_names, pipeline_label)
    df_deals = add_stage_attributes(df_deals, stages)
    df_deals = add_close_date(df_deals, deals, stages)
    df_deals = add_inbound_call(df_deals)
    df_deals = add_first_call(df_deals, deals, first_calls)
    df_deals = add_intake_date(df_deals)
    df_deals = add_durations(df_deals)
    refreshed = now_pacific.replace(tzinfo=None, microsecond=0)
    df_deals["last_refresh"] = refreshed
    df_deals = to_sheet(df_deals, definitions, stage_columns)

    # The workbook is shared: people keep their own tabs next to Deals (Maz
    # was the first). Only the Deals sheet is replaced; every other part of
    # the file is copied byte for byte — see splice.py.
    t0 = time.monotonic()
    existing, etag = existing_job.result()
    background.shutdown()
    if existing is None:
        print(f"'{FILE_PATH}' does not exist yet — building it from scratch")
        workbook, report = build_workbook(df_deals), None
    else:
        try:
            workbook, report = splice.replace_sheet(existing, DEALS_SHEET, df_deals)
        except splice.SpliceError as exc:
            fail(f"cannot update '{DEALS_SHEET}' safely: {exc} — nothing written")
        if not splice.untouched_parts_identical(existing, workbook, report):
            fail("a part other than Deals would change — nothing written")
        # Formulas in other tabs read Deals by column letter; make sure every
        # column they read keeps its header.
        refs = splice.referenced_columns(existing, DEALS_SHEET)
        old_headers = splice.header_row(existing, DEALS_SHEET)
        if refs:
            print(f"Other tabs read {len(refs)} Deals column(s): " + "; ".join(
                f"{col} ({old_headers.get(col, '?')}) <- {len(w)} formula(s), e.g. {', '.join(w[:3])}"
                for col, w in sorted(refs.items(), key=lambda kv: splice.column_index(kv[0]))))
        if DRY_RUN and refs:
            preview_other_tabs(existing, refs)
        shifts = splice.shifted_references(existing, workbook, DEALS_SHEET)
        if shifts:
            for col, where, before, after in shifts:
                print(f"  COLUMN SHIFT {col}: '{before}' -> '{after}' — read by {', '.join(where[:5])}")
            if ACCEPT_COLUMN_CHANGES:
                print("ACCEPT_COLUMN_CHANGES is set: writing the new layout on purpose.")
            elif DRY_RUN:
                print("DRY RUN: a real run would stop here to protect those formulas.")
            else:
                fail("a column other tabs read would change — nothing written. Put new columns "
                     "at the far right (the 'Added later' group) instead.")
        print(f"Sheets in the workbook: {', '.join(report['sheets'])}")
        print(f"Replacing only '{DEALS_SHEET}' ({report['sheet_part']}); "
              f"{len(report['untouched'])} other parts copied unchanged")
    size_kb = len(workbook) / 1024
    print(f"Workbook built: {size_kb:.0f} KB ({time.monotonic() - t0:.0f}s)")

    if DRY_RUN:
        # Counts only — deal names are client data and must not reach CI logs.
        print(f"DRY RUN: would upload {size_kb:.1f} KB to {FILE_PATH}")
        print(f"DRY RUN: one sheet '{DEALS_SHEET}', {len(df_deals)} rows x {len(df_deals.columns)} columns")
        # Fill rate per column — counts, no client data — so a dry run shows
        # which columns HubSpot actually populates.
        filled = df_deals.notna().sum()
        closed = df_deals["Is Closed"] == True  # noqa: E712
        settled = df_deals["Is Settled"] == True  # noqa: E712
        print(f"DRY RUN: settled deals (Date - Settled filled) {int(settled.sum())}; by current stage:")
        for stage, n in df_deals.loc[settled, "Deal Stage"].value_counts().items():
            print(f"    {n:>6}  {stage}")

        print(f"DRY RUN: closed deals {int(closed.sum())}, with Close Date "
              f"{int(df_deals.loc[closed, 'Close Date'].notna().sum())}; by source:")
        for src, n in df_deals.loc[closed, "Close Date Source"].value_counts(dropna=False).items():
            print(f"    {n:>6}  {src}")
        print("DRY RUN: cycle times (days) — rows, median, negative:")
        for _, header, _, _ in DURATIONS:
            col = df_deals[header].dropna()
            median = f"{col.median():.0f}" if len(col) else "-"
            print(f"    {header}: {len(col)} rows, median {median}, negative {int((col < 0).sum())}")
        print("DRY RUN: First Call Direction:")
        for v, n in df_deals["First Call Direction"].value_counts().items():
            print(f"    {n:>6}  {v}")
        print("DRY RUN: Created by Inbound Call:")
        for v, n in df_deals["Created by Inbound Call"].value_counts().items():
            print(f"    {n:>6}  {v}")
        print("DRY RUN: Date - Intake by source:")
        for v, n in df_deals["Intake Date Source"].value_counts(dropna=False).items():
            print(f"    {n:>6}  {v}")
        print("DRY RUN: AB 1755 by manufacturer:")
        for v, n in df_deals[AB1755_HEADER].value_counts(dropna=False).items():
            print(f"    {n:>6}  {v}")
        print("DRY RUN: rows with a value, per column:")
        for col, n in filled.items():
            print(f"    {n:>6}  {col}")
        return

    status = upload(workbook, token, len(workbook), etag)
    print(f"File {'created' if status == 201 else 'uploaded'} ({size_kb:.1f} KB): {FILE_PATH}")

    # Read it back and make sure every other tab arrived exactly as it was.
    # This cannot undo an upload, but it turns a silent loss into a failed
    # run — and SharePoint's version history can restore the previous file.
    if report is not None:
        stored, _ = fetch_existing(token)
        if other_sheet_parts(stored, report) != other_sheet_parts(existing, report):
            fail("the other tabs in SharePoint's copy differ from before the upload — "
                 "check the file and restore the previous version from its version history")
        print("Verified: every other tab in SharePoint is unchanged")
    print(f"Rows written: {len(df_deals)} x {len(df_deals.columns)} columns")
    print(f"Last Refresh: {refreshed:%B %d, %Y at %I:%M %p} Pacific")


if __name__ == "__main__":
    main()
