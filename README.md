# Marketing

Pulls every **Lemon Law** deal created since **1 January of the current year**
from HubSpot and writes it, raw, to `Marketing.xlsx` in SharePoint
(`Data Inventory/PLS/Requests/`, next to the Monthly Settlement Report) via
the Microsoft Graph API. The workbook is the source for the Power BI
marketing funnel.

**Almost nothing is calculated in the workbook.** The funnel buckets
(Prospect, Lead, MQL, SQL, Closed Won, Closed Lost) and every rate live in
Power BI. The exceptions are per-deal values: the Close Date, the AB 1755
side and the four cycle times, all recomputed on every run. The script only does lookups — stage id → name, owner
id → name, dropdown value → the label HubSpot shows — and converts dates to
`America/Bogota`, as the Settlement report does.

## How it runs

Same pattern as the Monthly Settlement Report: the workflow is
**dispatch-only**, and an external cron calls the `workflow_dispatch` API.
Do not add a `schedule:` trigger — it would run on top of the cron.

Each run rebuilds the whole file from HubSpot and overwrites it. There is no
upsert: the pull is the complete year, so a deal merged, deleted or moved off
the pipeline simply stops appearing.

The window moves on its own: on 1 January the file starts over with the new
year.

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

**One sheet, `Deals`, one row per Lemon Law deal created this year — every
one of them — and every row stands on its own** (no lookup tabs). Headers are
HubSpot's property labels, except the columns this script adds.

### Column layout

Grouped left to right in the order the funnel reads (`COLUMN_GROUPS` in
`main.py` — the one place to move a column):

| Group | Columns |
|---|---|
| **Deal** | Record ID, Deal Name, Lemon Law - State, Create Date |
| **Current stage** | Deal Stage, Deal Stage ID, Deal Stage Order, Deal Stage Is Closed, Date entered current stage, Close Date, Close Date Source |
| **Source / channel** | Lead - Source, Lead - Source (Group), Original Traffic Source, Original Traffic Source Drill-Down 1, Record source, AirCall Entry Number, Auto Dialer Call Type, TF: UTM Source / Medium / Campaign, GCLID |
| **Vehicle / AB 1755** | Manufacturer, AB 1755 (Manufacturer), RO Review (Final Decision), Vehicle - Year, Vehicle - Model |
| **Milestones** | Date - RO Review, Date exited "New File Set Up - Doc Collection", Date - Referred Out |
| **Outcome** | Case Category, Date - Settled, Total Settled Attorneys Fees and Cost, Net Attorney Fees, Close Out Reason, Date - Closed Out, Date - Close Out After Retained, Legal Sub Phase |
| **Cycle time (days)** | Days: Created to RO Review, Days: RO Review to File Set Up, Days: Created to File Set Up, Days: Created to Settled |
| **People** | Deal Owner, Deal Owner ID, Intake - Case Supervisor, Senior Case Supervisor, Legal - Handling Attorney, Settlement Attorney |
| **Stage history** | Date entered "<stage>" for every stage, in pipeline order |
| **Audit** | Pipeline, Last Modified Date, Last Refresh |

A property added to `BASE_PROPERTIES` without a place in `COLUMN_GROUPS`
still comes out, just before Audit — and a test fails until it is placed.

### Column notes

- **Deal Stage Is Closed** — true for Settled - Lit, Settled - Pre Lit,
  Settled - Referred Out, Close Out, Retained - Drop Client, Retained - Client
  Dropped and Referred Out - Complete (`CLOSED_STAGE_LABELS`). HubSpot's own
  closed flag is not used: it marks only the Settled stages. A closed label
  missing from the pipeline stops the run.
- **Close Date / Close Date Source** — filled for every closed deal, blank for
  open ones. HubSpot's own Close Date is never set in this pipeline, so it
  comes from Date - Settled (Settled stages), Date - Referred Out (Referred
  Out - Complete) or Date - Closed Out then Date - Close Out After Retained
  (the rest), and the date the deal entered its current stage when that field
  is blank. Close Date Source names the field used, by its label.
- **AB 1755 (Manufacturer)** — `Opt In`, `Opt Out` or `Not on list`, from the
  published AB 1755 list (`AB1755_OPT_IN` / `AB1755_OPT_OUT`, keyed on the
  stored manufacturer value). Genesis files under Hyundai, Infiniti under
  Nissan; Isuzu has no value in HubSpot. AB 1755 is California law — filter
  Lemon Law - State to California for the split. Checked against RO Review
  (Final Decision) on 2026 deals: 2 of 485 disagree.
- **Manufacturer** is the full legal name, e.g. "General Motors LLC".
- **People** columns are names; archived people resolve through the owner list.
- **Stage history** — HubSpot has no "date entered" property for four stages:
  HOLD gets Date exited "HOLD" instead; Retained - Client Dropped, Settled -
  Referred Out and TEST get no column. Their deals are still in the sheet.
- There is no "New File Set Up - Intake" stage in this pipeline; the file
  set-up stage is "New File Set Up - Doc Collection".
- **Cycle time (days)** — whole calendar days between two of the row's own
  dates (`DURATIONS` in `main.py`):
  - Created to RO Review: Create Date -> Date - RO Review
  - RO Review to File Set Up: Date - RO Review -> Date exited "New File Set
    Up - Doc Collection"
  - Created to File Set Up: Create Date -> Date exited "New File Set Up - Doc
    Collection"
  - Created to Settled: Create Date -> Date - Settled

  Blank when either date is missing. A negative value is kept — it means the
  dates in HubSpot are out of order — and the dry run counts them. Recomputed
  from scratch every run, so a corrected date corrects its duration.
- All headers are in English.
- Fees and cycle times are written as numbers; dates in Bogota time; Last
  Refresh in Pacific.

**Left out because HubSpot never fills them** on this year's Lemon Law deals
(measured on all 18,111 in September 2026): Close Date, Intake Outcome, Class
Action, the deal-level utm_source / utm_medium / utm_campaign, Lead Generation
Form and Form ID.

A dry run prints how many rows have a value in each column, the Close Date
sources and the AB 1755 split (counts only).

Size: ~18,000 rows and ~5 MB in September 2026, ~7 MB by year end.

## Every run is a full refresh

Nothing is carried over from the previous file. Each run pulls **every** deal
created this year and **every** column from HubSpot again, and overwrites the
workbook. So whatever changed in HubSpot since the last run — a stage move, a
close, a new settlement date or fee, a reassigned attorney, a corrected
manufacturer, a deal merged or deleted — is in the next file. There is no
"only recent changes" mode to fall out of step.

What makes sure a run either lands complete or changes nothing:

- Each month's pull must return exactly the count HubSpot reports, and none
  may reach the 10,000-result ceiling.
- Every row must have a create date inside 1 January .. now.
- After the upload, the size SharePoint reports storing must equal the
  workbook the run built; otherwise the run fails.
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
```
