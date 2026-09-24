# Marketing

Pulls **every Lemon Law deal ever created** from HubSpot — the firm's history
there starts on 18 April 2021, ~232,000 deals by September 2026 — and writes
it to the **Deals** tab of `Marketing.xlsx` in SharePoint
(`Data Inventory/PLS/Requests/`, next to the Monthly Settlement Report) via
the Microsoft Graph API. The workbook is the source for the Power BI
marketing funnel, compared year against year, and for questions such as
cases settled this year that were created in earlier years.

**Almost nothing is calculated in the workbook.** The funnel buckets
(Prospect, Lead, MQL, SQL, Closed Won, Closed Lost) and every rate live in
Power BI. The exceptions are per-deal values: the Close Date, the AB 1755
side and the four cycle times, all recomputed on every run. The script only does lookups — stage id → name, owner
id → name, dropdown value → the label HubSpot shows — and converts dates to
California time (`America/Los_Angeles`).

**California and Washington only.** Deals in the Lemon Law pipeline whose
Lemon Law - State is anything else ("Employment Law" — 12 in September 2026)
are dropped, as in the Monthly Settlement Report; the log counts them.

## How it runs

Same pattern as the Monthly Settlement Report: the workflow is
**dispatch-only**, and an external cron calls the `workflow_dispatch` API.
Do not add a `schedule:` trigger — it would run on top of the cron.

Each run rebuilds the whole file from HubSpot and overwrites it. There is no
upsert: the pull is the complete year, so a deal merged, deleted or moved off
the pipeline simply stops appearing.

The pull covers the whole history: monthly windows from January 2021 to
now, plus one window before 2021 that catches any deal imported with an
older create date. A month HubSpot reports at 10,000 deals or more (the
search API's ceiling) is split in half until every piece fits.

**Run it weekly.** A full-history run takes several minutes, almost all of
it the HubSpot pull, so the external cron is set to once a week.

## Setup

Repository secrets (**Settings → Secrets and variables → Actions**), the same
four the Settlement repo uses:

| Secret | What it is |
|---|---|
| `HUBSPOT_TOKEN` | HubSpot private app token — `crm.objects.deals.read`, `crm.schemas.deals.read`, `crm.objects.owners.read` |
| `AZURE_TENANT_ID` | Entra ID tenant of the SharePoint site |
| `AZURE_CLIENT_ID` | App registration client ID |
| `AZURE_CLIENT_SECRET` | App registration client secret (`Files.ReadWrite.All`, admin consent) |

Optional environment variables:

| Variable | Default |
|---|---|
| `SHAREPOINT_DRIVE_ID` | the production drive (same as the Settlement report) |
| `SHAREPOINT_FILE_PATH` | `Data Inventory/PLS/Requests/Marketing.xlsx` — the workflow's `file_path` input |
| `DRY_RUN` | unset — `true` runs everything but the upload |

The first real run creates the file; later runs overwrite it.

## The workbook

**One tab, `Deals`, one row per Lemon Law deal ever created — every one of
them — and every row stands on its own** (no lookup tabs). Headers are
HubSpot's property labels, except the columns this script adds.

### Column layout

Grouped left to right in the order the funnel reads (`COLUMN_GROUPS` in
`main.py` — the one place to move a column):

| Group | Columns |
|---|---|
| **Deal** | Record ID, Deal Name, Lemon Law - State, Create Date |
| **Current stage** | Deal Stage, Deal Stage ID, Deal Stage Order, Is Closed, Is Settled, Date entered current stage, Close Date, Close Date Source |
| **Source / channel** | Lead - Source, Lead - Source (Group), Original Traffic Source, Original Traffic Source Drill-Down 1, Record source, AirCall Entry Number, Auto Dialer Call Type, TF: UTM Source / Medium / Campaign, GCLID |
| **Vehicle / AB 1755** | Manufacturer, AB 1755 (Manufacturer), RO Review (Final Decision), Vehicle - Year, Vehicle - Model |
| **Milestones** | Date - Intake, Date - RO Review, Date - Retainer Signed, Date - Ready for Legal (Exited File Set Up), Date - Referred Out |
| **Outcome** | Case Category, Date - Settled, Total Settled Attorneys Fees and Cost, Net Attorney Fees, Close Out Reason, Date - Closed Out, Date - Close Out After Retained, Legal Sub Phase |
| **Cycle time (days)** | Created to Intake, Created to RO Review, Intake to Retainer Signed, RO Review to Retainer Signed, Retainer Signed to Ready for Legal, RO Review to Ready for Legal, Intake to Ready for Legal, Created to Ready for Legal, Created to Settled |
| **People** | Deal Owner, Deal Owner ID, Intake - Case Supervisor, Senior Case Supervisor, Legal - Handling Attorney, Settlement Attorney |
| **Stage history** | Date entered "<stage>" for every stage, in pipeline order |
| **Audit** | Pipeline, Last Modified Date, Last Refresh |
| **Added later** | Created by Inbound Call, First Call Direction, First Call Date, Intake Date Source, Date - Inquiry Qualified |

**New columns go at the far right** (the "Added later" group), never in the
middle: formulas on other tabs (Maz) read Deals by column letter, and a
column inserted in the middle shifts every letter after it. A property added
to `BASE_PROPERTIES` without a place in `COLUMN_GROUPS` also lands at the far
right (and a test fails until it is placed).

**Columns A..AI are frozen.** Maz reads A Record ID, C Lemon Law - State,
D Create Date, E Deal Stage, X Manufacturer, Y AB 1755, AC Date - Intake,
AE Date - Retainer Signed, AF Date - Ready for Legal and AI Date - Settled;
a test fails if any of them moves. (PRs #8 and #9 had inserted three columns
mid-sheet, which slid Maz onto the wrong columns; they now sit in "Added
later".)

**Formula guard.** Every run lists the Deals columns other tabs read (cell
formulas, defined names, charts) and **refuses to write** if the header under
any of them would change — the previous file stays as it was, and the log
names the column, the header change and the formulas that read it. A
deliberate move (such as putting columns back under Maz) needs a one-off
run with the `accept_column_changes` input ticked.

### Column notes

- **Is Settled** — true when Date - Settled is filled, whatever the stage.
  Date - Settled is what confirms a settlement: settled cases routinely stay
  in Retained - Lit / Retained - Pre Lit (all 31 of September 2026's
  settlements created this year did), so the stage alone misses them.
- **Is Closed** — true when Is Settled, or when the stage is one of Settled -
  Lit, Settled - Pre Lit, Settled - Referred Out, Close Out, Retained - Drop
  Client, Retained - Client Dropped and Referred Out - Complete
  (`CLOSED_STAGE_LABELS`). HubSpot's own closed flag is not used: it marks
  only the Settled stages. A closed label missing from the pipeline stops the
  run.
- **Close Date / Close Date Source** — filled for every closed deal, blank for
  open ones. HubSpot's own Close Date is never set in this pipeline, so it
  comes from Date - Settled whenever it is filled (any stage), Date -
  Referred Out (Referred Out - Complete) or Date - Closed Out then Date - Close Out After Retained
  (the rest), and the date the deal entered its current stage when that field
  is blank. Close Date Source names the field used, by its label.
- **AB 1755 (Manufacturer)** — `Opt In`, `Opt Out` or `Not on list` (also for
  deals with no manufacturer, so it is never blank), from the
  published AB 1755 list (`AB1755_OPT_IN` / `AB1755_OPT_OUT`, keyed on the
  stored manufacturer value). Genesis files under Hyundai, Infiniti under
  Nissan; Isuzu has no value in HubSpot. AB 1755 is California law — filter
  Lemon Law - State to California for the split. Checked against RO Review
  (Final Decision) on 2026 deals: 2 of 485 disagree.
- **Created by Inbound Call** — Yes / No: Yes when the deal was created by an
  inbound call to an Aircall line (named "Aircall new contact, +1…" by the
  deal-creation automation, or carrying an AirCall Entry Number; the name is
  usually replaced after intake). Outbound cannot be read from the deal — it
  needs the direction of the deal's first call on HubSpot's Calls object;
  every run logs whether the token can read it ("Calls API: …").
- **First Call Direction** — Inbound / Outbound / Unknown / No calls: the
  direction of the deal's first call, as Aircall logged it on HubSpot's Calls
  object (`hs_call_direction`), and **First Call Date**. The first call is the
  call with the earliest timestamp among all calls associated with the deal
  (call ids do not follow call time, so every call is read). Unknown means the
  first call has no direction. Outbound means the firm made the first
  contact — typical of form, mailer and PPC leads the dialer works — not that
  a deal was created by an outbound call. About 3 of 4 deals have calls
  associated with the deal itself; the rest read "No calls".
  Reading the associations and every call adds a large part of the run time;
  all HubSpot requests are paced under its 10-second rate limit.
- **Manufacturer** is the full legal name, e.g. "General Motors LLC".
- **People** columns are names; archived people resolve through the owner list.
- **Stage history** — HubSpot has no "date entered" property for four stages:
  HOLD gets Date exited "HOLD" instead; Retained - Client Dropped, Settled -
  Referred Out and TEST get no column. Their deals are still in the sheet.
- There is no "New File Set Up - Intake" stage in this pipeline; the file
  set-up stage is "New File Set Up - Doc Collection".
- **Milestones**, in the order a case moves:
  1. **Intake** — **Date - Intake**: HubSpot's stamp for entering the Intake
     stage, and — only where that stamp is blank — **Date - Inquiry
     Qualified**, where the firm recorded intake before 2026 (that property
     was never back-filled into the stamp: 40,567 deals created before 2026
     carry it, none created in 2026). The stamp always wins. **Intake Date
     Source** says which one each row used; both originals stay in the sheet
     (the stamp in Stage history, Inquiry Qualified in Added later).
     Checked on samples: Inquiry Qualified is on or before Sign Up for 99% of
     2021 and 100% of 2023 signed cases, 85% of 2025's
  2. **RO Review** — Date - RO Review
  3. **Retainer Signed** — Date - Retainer Signed (`date___retained`)
  4. **Ready for Legal** — the day the deal left New File Set Up - Doc
     Collection (`hs_v2_date_exited_5792630`)
- **Cycle time (days)** — whole calendar days between two of the row's own
  dates (`DURATIONS` in `main.py`; the Intake ones use Date - Intake): Created to Intake, Created to RO Review,
  Intake to Retainer Signed, RO Review to Retainer Signed, Retainer Signed to
  Ready for Legal, RO Review to Ready for Legal, Intake to Ready for Legal,
  Created to Ready for Legal, Created to Settled.

  Blank when either date is missing. A negative value is kept — it means the
  dates in HubSpot are out of order — and the dry run counts them. Recomputed
  from scratch every run, so a corrected date corrects its duration.
- All headers are in English.
- Fees and cycle times are written as numbers; every date and time,
  Last Refresh included, in California time.

**Left out because HubSpot never fills them** on this year's Lemon Law deals
(measured on all 18,111 in September 2026): Close Date, Intake Outcome, Class
Action, the deal-level utm_source / utm_medium / utm_campaign, Lead Generation
Form and Form ID.

A dry run prints how many rows have a value in each column, the Close Date
sources and the AB 1755 split (counts only).

Size: ~232,000 rows and ~50–60 MB in September 2026, growing ~60,000 rows a
year.

## Run time

A run takes about a minute end to end, measured in September 2026 on
~18,000 deals:

| Step | Time |
|---|---|
| Runner setup (Python, cached dependencies, tests) | ~13s |
| HubSpot pull, 9 monthly windows, ~90 requests | ~30s |
| Building the workbook (xlsxwriter, row by row) | ~6s |
| Graph token + upload | ~2s |

The log prints the pull and build times on every run. The workbook used to
take ~25s with pandas + openpyxl; xlsxwriter cut it to ~6s and the file
shrank. The pull is bound by HubSpot's search rate limit (a few requests a
second, shared with the Settlement job), so it grows with the year —
roughly 1.5s per 1,000 deals.

A full rewrite is kept on purpose rather than updating only changed deals:
at this size it costs seconds, and it is what guarantees the file matches
HubSpot exactly on every run.

## Other tabs in the workbook are never touched

The workbook is shared: people keep their own tabs next to Deals (the first
was **Maz**). A run replaces **only the Deals sheet** and copies every other
part of the file byte for byte — other tabs, their formulas, formatting,
charts, pivots and shared strings (`splice.py`).

- The file is downloaded, the Deals worksheet part is swapped, and the run
  checks every other part is byte-identical **before** uploading.
- The upload carries the version it read (`If-Match`). If someone saved the
  file in the meantime — a change to Maz, say — SharePoint refuses it and
  their work stands; the next run starts from their version.
- After the upload the file is downloaded again and every other tab is
  compared with what was there before; any difference fails the run
  (SharePoint's version history can restore the previous file).
- The file is not opened and re-saved with a spreadsheet library: openpyxl
  and similar drop charts and images they cannot read. Formulas that read
  Deals recalculate when the file is next opened (`fullCalcOnLoad`).
- If Deals is missing or has something attached to it (a table, a chart on
  the Deals tab itself), the run stops without writing rather than guess.
- **Do not put your own content on the Deals tab** — it is rebuilt every
  run. Build on another tab and point formulas or pivots at Deals.

## Every run is a full refresh

Nothing is carried over from the previous file. Each run pulls **every** deal
ever created and **every** column from HubSpot again, and rewrites the
Deals tab completely. So whatever changed in HubSpot since the last run — a stage move, a
close, a new settlement date or fee, a reassigned attorney, a corrected
manufacturer, a deal merged or deleted — is in the next file. There is no
"only recent changes" mode to fall out of step.

What makes sure a run either lands complete or changes nothing:

- Each month's pull must return exactly the count HubSpot reports, and none
  may reach the 10,000-result ceiling.
- Every row must have a create date, and none may fall outside the pulled
  range.
- After the upload, SharePoint's copy must have been modified by this upload
  and be about the size sent (not exact: SharePoint writes a few KB of its
  own metadata into Office files); otherwise the run fails.
- Any of these failing stops the run **before or without** a partial write,
  and the previous file stays as it was — stale, never truncated.

## Failure behaviour

- HubSpot search stops paging at 10,000 results without an error, so the year
  is pulled **a calendar month at a time**. A month that reaches the ceiling,
  or that downloads fewer deals than HubSpot reports, aborts the run
  **without writing** rather than publishing a truncated file.
- Zero deals aborts without writing.
- 429, 5xx and network errors are retried 4 times with backoff, honouring
  `Retry-After`. 423 (someone has the workbook open) is retried every 60s.
- A property renamed or archived in HubSpot comes out as an empty column plus
  a warning, not a failed run.
- Logs carry counts only — deal names are client data.

## Tests

```
python tests/test_main.py
python tests/test_splice.py
```
