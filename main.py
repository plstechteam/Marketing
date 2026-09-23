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
from datetime import datetime, timezone

import pandas as pd
import pytz
import requests

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
# Deal dates go out in Bogota time, as in the Monthly Settlement Report, so
# the two workbooks agree on which day a deal belongs to.
DEAL_TZ = pytz.timezone("America/Bogota")

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
    "createdate",
    "hs_lastmodifieddate",
    "hs_v2_date_entered_current_stage",
    "case_category",                # Lit / Pre-Lit
    "date___settled",               # Date - Settled — confirms the settlement
    "date___dropped",               # Date - Closed Out
    "date__intake_sign_up_close_out",   # Date - Close Out After Retained
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


def label_value(raw, labels):
    """Map an enumeration's stored value(s) to what HubSpot's UI shows.

    Multi-select values arrive ';'-joined. A value with no option (an option
    since deleted) is kept as stored rather than blanked.
    """
    if raw is None or raw == "":
        return None
    if not labels:
        return raw
    return ";".join(labels.get(v, v) for v in str(raw).split(";"))


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
    """Raw deals -> one row per deal, HubSpot labels as headers.

    Only lookups happen here (stage id -> name, owner id -> name, enumeration
    value -> label) and a timezone conversion on dates. No derived columns.
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
            elif d.get("type") == "enumeration":
                row[name] = label_value(raw, option_labels(d))
            else:
                row[name] = raw if raw != "" else None
        records.append(row)

    columns = []
    for name in properties:
        label = (definitions.get(name, {}).get("label") or name).strip()
        if name == "dealstage":
            columns.append(("dealstage__id", "Deal Stage ID"))
        if name == "hubspot_owner_id":
            columns.append(("hubspot_owner_id__id", "Deal Owner ID"))
            label = "Deal Owner"
        if name == "hs_object_id":
            label = "Record ID"
        columns.append((name, label))

    df = pd.DataFrame(records, columns=[c for c, _ in columns])
    df.columns = unique_headers([(label, c) for c, label in columns])
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


def upload(workbook_bytes, token):
    url = f"https://graph.microsoft.com/v1.0/drives/{DRIVE_ID}/root:/{FILE_PATH}:/content"
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
    return r.status_code


# ======================================================
# WORKBOOK
# ======================================================
def add_stage_attributes(df, stages):
    """Put the stage's pipeline order and closed flag on every row, right
    after Deal Stage, so the funnel can be sorted without a lookup table."""
    info = {s["id"]: s for s in stages}
    ids = df["Deal Stage ID"]
    # Positioned off "Deal Stage ID", which this script names, not off the
    # stage column whose header is HubSpot's label and could be renamed.
    at = df.columns.get_loc("Deal Stage ID") + 2
    df.insert(at, "Deal Stage Order",
              ids.map(lambda i: info[i].get("displayOrder") if i in info else None))
    df.insert(at + 1, "Deal Stage Is Closed",
              ids.map(lambda i: str((info[i].get("metadata") or {}).get("isClosed", "")).lower() == "true"
                      if i in info else None))
    return df


def close_date_for(props, stage_label, is_closed):
    """(Close Date, where it came from) for one deal; (None, None) if open.

    HubSpot never sets closedate in this pipeline, so a closed deal is dated
    from the firm's own fields, the one that matches how it closed:

      Settled stages   Date - Settled, the date that confirms the settlement
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
    else:
        preferred = ["date___dropped", "date__intake_sign_up_close_out"]
    for name in preferred:
        if props.get(name):
            return parse_hubspot_date(props[name]), name
    entered = parse_hubspot_datetime(props.get("hs_v2_date_entered_current_stage"))
    if entered:
        return entered.date(), "hs_v2_date_entered_current_stage"
    return None, None


def add_close_date(df, deals, stages):
    """Close Date and Close Date Source, right after Deal Stage Is Closed.

    Relies on df having one row per deal in the order of `deals`, which is
    how build_deals_frame builds it.
    """
    info = {s["id"]: s for s in stages}
    dates, sources = [], []
    for deal in deals:
        p = deal.get("properties", {})
        stage = info.get(p.get("dealstage"), {})
        closed = str((stage.get("metadata") or {}).get("isClosed", "")).lower() == "true"
        d, src = close_date_for(p, stage.get("label", ""), closed)
        dates.append(d)
        sources.append(src)
    at = df.columns.get_loc("Deal Stage Is Closed") + 1
    df.insert(at, "Close Date", dates)
    df.insert(at + 1, "Close Date Source", sources)
    return df


def build_workbook(df_deals):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl",
                        datetime_format="yyyy-mm-dd hh:mm:ss",
                        date_format="yyyy-mm-dd") as writer:
        df_deals.to_excel(writer, sheet_name=DEALS_SHEET, index=False)
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
    # It is cut on the same Bogota clock the deal dates are written in, so the
    # sheet's first Create Date is 1 January and every month in Power BI is
    # complete — a Pacific cut would drop the deals created in the last hours
    # of 31 December Pacific, which the sheet shows as 1 January.
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
    stage_labels = {s["id"]: s["label"] for s in stages}
    stage_candidates = [f"hs_v2_date_{kind}_{s['id']}"
                        for s in stages for kind in ("entered", "exited")]
    definitions = fetch_property_definitions(BASE_PROPERTIES + stage_candidates, hs_headers)
    stage_columns, skipped = stage_date_properties([s["id"] for s in stages], definitions)
    if skipped:
        print(f"Stages with no entered/exited date property in HubSpot (no column): "
              f"{', '.join(stage_labels[i] for i in skipped)}")
    properties = BASE_PROPERTIES + stage_columns
    owners = fetch_owners(hs_headers)
    print(f"Owners loaded: {len(owners)}")

    print(f"Pulling deals created {year_start:%Y-%m-%d} to now, one month at a time:")
    deals = fetch_deals(properties, windows, hs_headers)
    print(f"Total deals: {len(deals)}")
    if not deals:
        fail("HubSpot returned no deals — refusing to write an empty report.")

    owner_names = {k: v["name"] for k, v in owners.items()}
    df_deals = build_deals_frame(deals, properties, definitions, stage_labels,
                                 owner_names, pipeline_label)
    df_deals = add_stage_attributes(df_deals, stages)
    df_deals = add_close_date(df_deals, deals, stages)
    closed_stages = [s["label"] for s in stages
                     if str((s.get("metadata") or {}).get("isClosed", "")).lower() == "true"]
    print(f"Stages HubSpot marks closed: {', '.join(closed_stages) or 'none'}")
    refreshed = now_pacific.replace(tzinfo=None, microsecond=0)
    df_deals["Last Refresh"] = refreshed

    workbook = build_workbook(df_deals)
    size_kb = len(workbook) / 1024

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
        print("DRY RUN: rows with a value, per column:")
        for col, n in filled.items():
            print(f"    {n:>6}  {col}")
        return

    status = upload(workbook, token)
    print(f"File {'created' if status == 201 else 'uploaded'} ({size_kb:.1f} KB): {FILE_PATH}")
    print(f"Rows written: {len(df_deals)} x {len(df_deals.columns)} columns")
    print(f"Last Refresh: {refreshed:%B %d, %Y at %I:%M %p} Pacific")


if __name__ == "__main__":
    main()
