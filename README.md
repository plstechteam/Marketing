# Marketing

Pulls every **Lemon Law** deal created since **1 January of the current year**
from HubSpot and writes it, raw, to `Marketing.xlsx` in SharePoint
(`Data Inventory/PLS/Requests/`, next to the Monthly Settlement Report) via
the Microsoft Graph API. The workbook is the source for the Power BI
marketing funnel.

**Nothing is calculated in the workbook.** The funnel buckets (Prospect,
Lead, MQL, SQL, Closed Won, Closed Lost), the opt-in / opt-out split and every
rate live in Power BI. The script only does lookups — stage id → name, owner
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

**One sheet, `Deals`, one row per deal, and every row stands on its own** —
no lookup tabs. Stage and owner come as names with their IDs beside them, and
the stage's pipeline order and closed flag are on the row. Headers are
HubSpot's property labels.

Columns:

- Record ID, Deal Name, Pipeline
- Deal Stage ID, Deal Stage, **Deal Stage Order** (pipeline order, for sorting
  the funnel), **Deal Stage Is Closed**
- Deal Owner ID, Deal Owner
- Create Date, Close Date, Last Modified Date, Date entered current stage
- Lead - Source, Lead - Source (Group), Original Traffic Source, Original
  Traffic Source Drill-Down 1, Record source
- **Channel detail** — what splits Prospect by channel, since Lead - Source
  (Group) has no calls or forms bucket:
  - AirCall Entry Number — inbound call, and the line it came in on
  - Auto Dialer Call Type — outbound auto-dialer (Crexendo) call
  - Lead Generation Form, Form ID — website / Facebook / Typeform forms
  - utm_source, utm_medium, utm_campaign and the Typeform `TF: UTM` trio
  - GCLID — Google Ads click (PPC)
- **Close Out Reason** (`drop_reason`) — the detail for Closed Lost
- **RO Review (Final Decision)** (`ro_review__final_decision_`) — the opt-in /
  opt-out split: `Lit - Opt In`, `Lit - Opt Out`, `Pre-Lit (OPT IN OEM)`,
  `Pre-Lit (OPT OUT OEM)`, plus the non-opt values (`Sign Up - Pre-Lit`,
  `Sign Up - Lit (AB1755)`, `Sign Up - Pre-Lit (GM)`, …)
- Legal Sub Phase, Intake Outcome, Class Action, Lemon Law - State
- **Date entered "<stage>"** — one column per stage of the pipeline, built
  from the pipeline at run time so a new stage gets its column without a code
  change. HubSpot has no "date entered" property for four stages: **HOLD**
  gets **Date exited "HOLD"** instead; Retained - Client Dropped, Settled -
  Referred Out and TEST have neither and get no column. Their deals are still
  in the sheet — Deal Stage says where they are and Date entered current
  stage says since when
- Last Refresh (Pacific)

A dry run prints how many rows have a value in each column (counts only).

Size: ~18,000 rows and ~3.5–4 MB in September 2026, ~5–6 MB by year end.

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
