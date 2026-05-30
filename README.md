# Report Ops Automation

Two automations are included:

1. Export Power BI reports with configured filter state to PDF, rename them, and upload them to a SharePoint document library folder.
2. Find the generated PDFs in SharePoint and send each report link to the mapped Teams recipient.

## Setup

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
cp .env.example .env
cp config/reports.example.yaml config/reports.yaml
```

Fill `.env` and `config/reports.yaml`.

For reports with multiple pages and value maps, use `export_groups` in `config/reports.yaml`.
Each generated PDF gets this key:

```text
{report_key}.{export_group_key}.{value_key}
```

Example: `sales_daily.regional.west`.

When an export group has `values_from`, the script queries the Power BI semantic model for distinct values from that table/column and generates one PDF per value.

## Azure permissions

The export script uses app credentials. Power BI and SharePoint can use different app registrations:

- Power BI REST API application permissions needed for report export.
- Microsoft Graph application permission such as `Sites.ReadWrite.All` for SharePoint upload.
- Power BI tenant settings must allow service principals to use Power BI APIs, and report export requires supported capacity/licensing.

Set these in `.env`:

- `POWERBI_TENANT_ID`, `POWERBI_CLIENT_ID`, `POWERBI_CLIENT_SECRET`
- `SHAREPOINT_TENANT_ID`, `SHAREPOINT_CLIENT_ID`, `SHAREPOINT_CLIENT_SECRET`

The old shared names `AZURE_TENANT_ID`, `AZURE_CLIENT_ID`, and `AZURE_CLIENT_SECRET` are still accepted as fallbacks.

Email delivery uses Microsoft Graph application permissions:

- Add `Mail.Send` application permission to the email app registration and grant admin consent.
- Set `EMAIL_TENANT_ID`, `EMAIL_CLIENT_ID`, `EMAIL_CLIENT_SECRET`, and `EMAIL_SENDER_UPN`.
- If `EMAIL_*` credentials are omitted, the script falls back to `SHAREPOINT_*`, but that app must also have `Mail.Send`.

Teams delivery uses delegated device-code auth because normal Teams chat message sending is a user action:

- Microsoft Graph delegated permissions such as `Chat.ReadWrite`, `ChatMessage.Send`, `User.Read`, and SharePoint read permissions such as `Files.Read.All` or `Sites.Read.All`.
- The signed-in account must be the same user as `TEAMS_SENDER_UPN`.

## Run

```bash
python scripts/export_powerbi_to_sharepoint.py --config config/reports.yaml
python scripts/send_sharepoint_pdfs_to_teams.py --config config/reports.yaml
```

Use `--dry-run` on the delivery script to print planned deliveries without sending messages.

Delivery mapping supports `channel: email` or `channel: teams`. Email is the default.

```yaml
report_delivery:
  - export_key: "operation.regional.region_jakarta"
    channel: "email"
    recipient_upn: "person@company.com"
    subject: "{report_name} PDF is ready"
    message: "Hi, your {report_name} PDF is ready: <a href=\"{web_url}\">{file_name}</a>"
```

## Known limitations & lessons learned

### Power BI ExportTo API does NOT work for these cases

| Situation | What happens | Fix |
|---|---|---|
| Report has HTML content visuals | ExportTo renders "This visual does not support exporting" in PDF | Use Puppeteer PDF path (`export_report_pdf.js`) |
| Date filters are slicers with Edit Interactions | `reportLevelFilters` bypasses edit interactions — wrong visuals get filtered | Use Puppeteer `setSlicerState()` which respects interactions |
| Slicer has a saved default selection | API filter + slicer default = empty intersection = No Data | Clear slicer defaults in the report, or use Puppeteer |
| `defaultBookmark.state` constructed manually | ExportTo rejects all manually built states — format is SDK-only, undocumented | Use `capture_slicer_state.js` to capture a real state via the JS SDK |

### Filter rules

- **Filter pane fields** (e.g., regional filter): use `reportLevelFilters` — works correctly.
- **Slicer fields** (e.g., date slicers with edit interactions): use Puppeteer `setSlicerState()` — only way to respect edit interactions.
- **Date values sent to `setSlicerState`**: send the full datetime string **without** a timezone suffix (`2026-05-25T00:00:00`). Power BI stores dates as local-midnight-in-UTC (e.g. May 25 00:00 WIB = `2026-05-24T17:00:00.000Z`). The browser (WIB = UTC+7) converts the timezone-naive string to the correct UTC value automatically. Do NOT strip the time component — passing a date-only string (`2026-05-25`) makes Power BI treat it as UTC midnight, which is 7 hours off and matches no rows. Do NOT add a `Z` suffix — that bypasses the browser conversion and also lands on the wrong UTC value.
- **Finding the right slicer**: match by `getSlicerState().targets` (table/column), not by visual name — the JS SDK `visual.name` is a GUID, not the Selection pane display name.
- **Non-active pages**: call `report.setPage(pageName)` before setting slicers on that page — inactive page visuals throw errors.

### When to use Puppeteer vs ExportTo

| Report has... | Use |
|---|---|
| HTML content visuals | Puppeteer (`export_report_pdf.js`) |
| Date slicers with edit interactions | Puppeteer |
| Only filter-pane filters, no HTML visuals | ExportTo API (faster, no Node.js needed) |

### Mistakes made and fixed (don't repeat these)

| Mistake | What happened | Correct approach |
|---|---|---|
| Used `updateSettings(LayoutType.Custom, DisplayOption.FitToWidth)` to fix side whitespace | API call silently does nothing — whitespace remained | Use `report.setZoom(containerWidth / nativeW)` calculated from `page.defaultSize` |
| Stripped time component from datetime slicer values (`"2026-05-25T00:00:00"` → `"2026-05-25"`) | Power BI treated the date-only string as UTC midnight — 7 hours off from local midnight, matched no rows | Keep the full datetime without `Z` suffix; the browser (UTC+7) converts it to the correct UTC value automatically |
| Changed daily/monthly slicer `value` keys to display label format (`"25-May-26"`, `"April 2026"`) | Slicer showed the right label visually but the display string doesn't match the internal datetime column value — no filtering | Use `business_date_datetime` / `month_datetime` keys (which produce `"2026-05-25T00:00:00"`) — Power BI auto-formats the display |
| Identified weekly slicer broken; tried same datetime fix | Weekly already worked — it uses a text column (`WeekMonthYear`) so exact string match is correct. Only date columns need the datetime format | Distinguish text columns (exact string match) from date columns (datetime value, let browser convert timezone) |
| Store page had header whitespace that other levels didn't | Slicer setup loop always ends on "Operation - Store". When a store job ran next, `setActive()` on the already-active page didn't fire a rendered event — page captured mid-update | Detect if target is already the active page; navigate to another page first to force a genuine page transition before navigating back |
| Added `key_prefix: "region"/"area"/"store"` to `values_from` groups | Generated redundant names: `region_regional_2`, `area_jakarta_1`, `store_12345` — `level_key` already provides context | Remove `key_prefix`; the value key is derived cleanly from the column value |
| Added hardcoded `values` list with only `regional_2` to the regional group | Blocked `values_from` dynamic discovery — only Regional 2 was ever exported | Never add a static `values` list to a group that uses `values_from`. Use `--export-key` to test a single value |
| Put `config/reports.yaml` in `.gitignore` during a test session | Config changes (slicer formats, key fixes) were never committed — lost on the next clone | Never gitignore the config. Use `--export-key` flag to limit jobs during testing instead of modifying the config |
| Ran the same store (F103) twice in an "other stores" test | Wasted a test run | Check which stores were already tested before picking the next one |
| Set `--workers` default to 4 in argparse | User didn't want that — needed two extra commits to revert (one to change, one to revert) | Discuss default values before changing them. Use `git reset --hard` to remove bad commits rather than adding a revert commit |
| Created a manual JSON config and ran `export_report_pdf.js` directly to test | Bypassed the Python orchestration script — slicers, filter modes, filename templating all skipped | Always use `export_powerbi_to_sharepoint.py` with `--export-key` and `--business-date` for testing |

## Microsoft API references

- Power BI `ExportTo` starts an async export job and supports `reportLevelFilters`.
- Graph drive upload uses `PUT /drives/{drive-id}/root:/{path}:/content` for files up to 250 MB.
- Graph drive folder listing uses `GET /drives/{drive-id}/root:/{path}:/children`.
- Email delivery uses `POST /users/{sender}/sendMail`.
- Teams messages use `POST /chats/{chat-id}/messages`; one-on-one chats can be created with `POST /chats`.
