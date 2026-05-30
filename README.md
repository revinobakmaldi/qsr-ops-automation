# Report Ops Automation

Two automations:

1. Export Power BI reports to PDF with correct filter/slicer state and upload to SharePoint.
2. Find the generated PDFs in SharePoint and send each report link to the mapped Teams/email recipient.

## Architecture

```
export_powerbi_to_sharepoint.py  (Python orchestrator)
  → validates date slicer values exist in the dataset (pre-flight check)
  → splits jobs into N worker chunks
  → each worker calls export_report_pdf.js (Node.js + Puppeteer)
  → Node.js embeds the report, sets date slicers + value filters, generates PDF
  → Python uploads each PDF to SharePoint immediately as Node.js finishes it
```

The ExportTo REST API path still exists as a fallback (no Node.js required) but is not used for reports with HTML visuals or date slicers with Edit Interactions.

## Setup

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
cp .env.example .env
cp config/reports.example.yaml config/reports.yaml
```

Fill `.env` and `config/reports.yaml`.

## Export key format

Each generated PDF has a unique key:

```
{report_key}.{export_group_key}.{value_key}
```

Example: `operation.regional.regional_2`, `operation.store.f103`.

When an export group uses `values_from`, the script queries the Power BI semantic model for distinct column values and generates one PDF per value.

## Running the export

```bash
# Full run — all reports for yesterday's business date
python3 scripts/export_powerbi_to_sharepoint.py \
  --node-script scripts/capture_slicer_state.js \
  --workers 4

# Specific date
python3 scripts/export_powerbi_to_sharepoint.py \
  --business-date 2026-05-25 \
  --node-script scripts/capture_slicer_state.js \
  --workers 4

# Test a single report (use --export-key, never modify reports.yaml)
python3 scripts/export_powerbi_to_sharepoint.py \
  --business-date 2026-05-25 \
  --export-key operation.regional.regional_2 \
  --node-script scripts/capture_slicer_state.js \
  --workers 1
```

### All CLI flags

| Flag | Default | Description |
|---|---|---|
| `--config` | `config/reports.yaml` | Config file path |
| `--business-date` | yesterday | Date for daily/weekly/monthly slicer values (YYYY-MM-DD) |
| `--run-date` | today | Date used in output filenames |
| `--export-key` | *(all)* | Limit to specific key(s), repeatable |
| `--export-key-file` | — | File with one export key per line (use for retry) |
| `--on-missing-date` | `fail` | What to do if slicer date not in dataset: `fail` or `latest` |
| `--workers` | `1` | Parallel browser sessions. **Max 2** — 3+ concurrent sessions overwhelm Power BI's dataset query pipeline, causing charts to render empty. |
| `--render-wait` | `4.0` | Seconds to wait after network idle before PDF capture |
| `--max-jobs` | `0` | Stop after N jobs (0 = unlimited) |
| `--filter-mode` | `all` | Debug filter mode: `all`, `none`, `date-only`, etc. |
| `--node-script` | — | Path to `capture_slicer_state.js` (required for Puppeteer path) |
| `--list-values` | — | Print discovered column values and exit |
| `--list-bookmarks` | — | Print saved bookmarks and exit |

### Date slicer validation (pre-flight)

Before any export runs, each date slicer value is checked against the live dataset:

```bash
# If May 26 data isn't ready yet:
# ValueError: Date slicer value(s) not found in dataset.
#   Mapping_Date_RFID/DateOfBusiness = '2026-05-26T00:00:00'
# Rerun with --on-missing-date latest to use the most recent available value instead.

# Use the most recent available date instead of failing:
python3 scripts/export_powerbi_to_sharepoint.py \
  --business-date 2026-05-26 \
  --on-missing-date latest \
  --node-script scripts/capture_slicer_state.js \
  --workers 4
# WARNING: Mapping_Date_RFID/DateOfBusiness = '2026-05-26T00:00:00' not in dataset. Using latest: '2026-05-25T00:00:00'
```

### Retrying failed jobs

When jobs fail, their export keys are saved to `failed_YYYYMMDD_HHMMSS.txt`. The terminal prints the exact retry command:

```bash
# Retry only the failed jobs
python3 scripts/export_powerbi_to_sharepoint.py \
  --business-date 2026-05-25 \
  --export-key-file failed_20260525_003142.txt \
  --node-script scripts/capture_slicer_state.js \
  --workers 2
```

Each PDF is uploaded to SharePoint immediately after it's generated, so a crash on job N does not lose the already-uploaded jobs 1..(N-1).

## Delivery

```bash
python3 scripts/send_sharepoint_pdfs_to_teams.py --config config/reports.yaml
python3 scripts/send_sharepoint_pdfs_to_teams.py --config config/reports.yaml --dry-run
```

```yaml
report_delivery:
  - export_key: "operation.regional.regional_2"
    channel: "email"
    recipient_upn: "person@company.com"
    subject: "{report_name} PDF is ready"
    message: "Hi, your {report_name} PDF is ready: <a href=\"{web_url}\">{file_name}</a>"
```

## Azure permissions

Power BI and SharePoint can use different app registrations.

**Power BI** (app credentials):
- Power BI REST API application permissions
- Tenant setting: allow service principals to use Power BI APIs

**SharePoint** (app credentials):
- `Sites.ReadWrite.All` Microsoft Graph application permission

**Email** (app credentials):
- `Mail.Send` Microsoft Graph application permission
- Set `EMAIL_TENANT_ID`, `EMAIL_CLIENT_ID`, `EMAIL_CLIENT_SECRET`, `EMAIL_SENDER_UPN`
- Falls back to `SHAREPOINT_*` credentials if `EMAIL_*` are not set

**Teams** (delegated device-code auth):
- `Chat.ReadWrite`, `ChatMessage.Send`, `User.Read`, `Files.Read.All`
- Signed-in account must match `TEAMS_SENDER_UPN`

`.env` variables:
```
POWERBI_TENANT_ID / POWERBI_CLIENT_ID / POWERBI_CLIENT_SECRET
SHAREPOINT_TENANT_ID / SHAREPOINT_CLIENT_ID / SHAREPOINT_CLIENT_SECRET
# AZURE_TENANT_ID / AZURE_CLIENT_ID / AZURE_CLIENT_SECRET  (legacy fallback for both)
```

## Known limitations & lessons learned

### Power BI ExportTo API does NOT work for these cases

| Situation | What happens | Fix |
|---|---|---|
| Report has HTML content visuals | ExportTo renders "This visual does not support exporting" | Use Puppeteer path |
| Date filters are slicers with Edit Interactions | `reportLevelFilters` bypasses edit interactions — wrong visuals filtered | Use Puppeteer `setSlicerState()` |
| Slicer has a saved default selection | API filter + slicer default = empty intersection = No Data | Clear slicer defaults in the report, or use Puppeteer |
| `defaultBookmark.state` constructed manually | ExportTo rejects all manually built states — format is SDK-only | Use `capture_slicer_state.js` to capture via the JS SDK |

### Filter rules (Puppeteer path)

- **Value filters** (region, area, store): use `report.setFilters()` at report level.
- **Date slicers**: use `setSlicerState()` — matches slicers by `getSlicerState().targets` (table/column), not by display name which is a GUID.
- **Date value format**: pass the full datetime **without** a `Z` suffix — `"2026-05-25T00:00:00"`. Power BI stores dates as local-midnight-in-UTC. The browser (WIB = UTC+7) converts the timezone-naive string to the correct UTC value. Do NOT strip the time (date-only = UTC midnight = 7 hours off = no rows matched). Do NOT add `Z` (bypasses browser conversion = also wrong).
- **Text slicer columns** (e.g. `WeekMonthYear`): pass the exact string value — `"Week 3 May 2026"`. No datetime conversion needed.
- **Slicer setup**: navigate to each page and call `setSlicerState()` — slicers on inactive pages are not accessible.

### PDF generation rules

- **Side whitespace**: use `report.setZoom(containerWidth / nativeW)` from `page.defaultSize` to fill the container width. `LayoutType.Custom + DisplayOption.FitToWidth` via `updateSettings` silently does nothing.
- **Per-page height**: each report page has a different `defaultSize.height`. Use a per-job viewport resize to exactly match each page's height — using one fixed height causes Power BI to vertically centre shorter pages, adding whitespace at the top.
- **Slicer loop order**: the slicer setup loop ends on the last page. Navigate away from that page before generating any PDF on the same page — `setActive()` on the already-active page does not fire a rendered event.
- **Streaming upload**: Node.js writes each job result to stdout immediately on completion. Python reads line-by-line via `Popen` and uploads per result — a crash on job N does not lose already-uploaded jobs.
- **When to export**: use `page.waitForNetworkIdle({ idleTime: 1500 })` to wait for all data queries to finish, then wait `renderWait` seconds for the browser to paint. Do NOT rely on Power BI's `rendered` event alone — it fires when layout is done, before data queries complete. Do NOT rely on `.powerbi-spinner` DOM checks — actual spinners live in shadow DOM and are invisible to `querySelectorAll`. Do NOT use screenshot comparison — empty loading charts look identical between frames.
- **Parallel worker limit**: max `--workers 2`. More than 2 concurrent Power BI sessions overwhelm the dataset query pipeline — queries compete, data arrives empty, charts render blank. Confirmed by comparing PDF file sizes (521KB broken vs 863KB correct). This is a Power BI service limit, not a code issue.

### Mistakes made and fixed (don't repeat these)

| Mistake | What happened | Correct approach |
|---|---|---|
| Used `updateSettings(LayoutType.Custom, DisplayOption.FitToWidth)` | Silently does nothing | `report.setZoom(containerWidth / nativeW)` from `page.defaultSize` |
| Stripped time from datetime slicer values (`T00:00:00` → date only) | Power BI treated as UTC midnight — 7 hours off, no rows matched | Keep full datetime without `Z`; browser converts timezone correctly |
| Used display label format for daily/monthly slicers (`"25-May-26"`, `"April 2026"`) | Slicer showed label visually but label ≠ internal datetime value — no filtering | Use `business_date_datetime` / `month_datetime` keys — Power BI auto-formats the display |
| Assumed weekly slicer was broken because daily/monthly were | Weekly worked fine — it's a text column, exact string match is correct | Distinguish text columns (exact match) from date columns (datetime with browser conversion) |
| Used TOPN(3) to check if a date value exists in the dataset | Only returned the 3 most recent dates — target date not in top 3 = false negative | Use a FILTER DAX query for existence check; use TOPN separately only for the latest-value fallback |
| Used `--workers 4` for the full run | 4 concurrent Power BI embed sessions overwhelm the dataset query pipeline — data queries starve each other, charts render empty. Tried 6+ different render-wait strategies (debounce, spinner DOM check, screenshot diff, network idle) before isolating root cause by comparing PDF file sizes (521KB broken vs 863KB correct). 2 workers confirmed reliable. | Max `--workers 2`. Use `waitForNetworkIdle` to detect true data load completion. Never debug render issues without checking file size first — it immediately tells you if content loaded. |
| Tried `.powerbi-spinner` DOM check to detect loading visuals | The `powerbi-spinner` elements found by `querySelectorAll` are permanent hidden placeholders (display:none). Actual animated spinners live in Power BI's shadow DOM — invisible to querySelectorAll, so the check always passes immediately and exports too early. | Don't rely on DOM spinner checks for Power BI — use `waitForNetworkIdle` instead. |
| Tried screenshot comparison (compare consecutive screenshots until identical) | Empty loading chart areas look the same in consecutive screenshots — "stable" is indistinguishable from "never loaded". The approach detects stability but not completeness. | Screenshot comparison can't distinguish empty-loading from empty-static. Use network idle. |
| Store page had header whitespace | Each page has a different `defaultSize.height`. Single viewport height caused Power BI to vertically centre shorter pages. Wrongly diagnosed 3 times before root cause found by logging `defaultSize` per page | Expand viewport to tallest page once at startup; resize per-job to each page's exact scaled height |
| Added `key_prefix` to `values_from` groups | Redundant names: `region_regional_2`, `area_jakarta_1` | Remove `key_prefix`; value key derived from column value; `level_key` already provides context |
| Added hardcoded `values` list to a group with `values_from` | Blocked dynamic discovery — only that one value was ever exported | Never add static `values` to a group that uses `values_from`; use `--export-key` to test one value |
| Put `config/reports.yaml` in `.gitignore` | Config changes were never committed — lost on next clone | Never gitignore the config; use `--export-key` for targeted testing |
| Added a revert commit instead of removing with `git reset --hard` | Extra noise in git history | `git reset --hard HEAD~N && git push --force` to cleanly remove bad commits |
| Ran `export_report_pdf.js` directly with a hand-crafted JSON | Bypassed Python orchestration — slicers, validation, filename templating all skipped | Always test via `export_powerbi_to_sharepoint.py --export-key` |

## Microsoft API references

- Power BI `ExportTo` starts an async export job and supports `reportLevelFilters`.
- Power BI dataset `executeQueries` runs DAX queries for column value discovery and slicer validation.
- Graph drive upload uses `PUT /drives/{drive-id}/root:/{path}:/content` for files up to 250 MB.
- Email delivery uses `POST /users/{sender}/sendMail`.
- Teams messages use `POST /chats/{chat-id}/messages`; one-on-one chats via `POST /chats`.
