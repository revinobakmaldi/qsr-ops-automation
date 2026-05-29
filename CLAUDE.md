# Report Ops Automation — Codebase Notes

## Architecture

```
Python export script (export_powerbi_to_sharepoint.py)
  → calls export_report_pdf.js (Node.js + Puppeteer) per report
  → Node.js embeds the PBI report, sets slicers + filters, generates PDF
  → Python uploads PDF bytes to SharePoint
```

Fallback path (no HTML visuals, no slicer edit interactions): use ExportTo API directly with `defaultBookmark.state` captured by `capture_slicer_state.js`.

---

## Power BI Export — Lessons Learned

### What does NOT work and why

**`reportLevelFilters` ignores slicer edit interactions**
Adding date filters via `reportLevelFilters` applies to ALL visuals unconditionally — it bypasses the report's "Edit Interactions" settings. Visuals that should not be date-filtered get filtered anyway. Wrong output.

**Slicer saved default + API filter = No Data**
When a report slicer has a saved default selection (e.g., REGIONAL 1), adding a `reportLevelFilters` for a different value (e.g., REGIONAL 2) produces an empty intersection → No Data in the PDF. Must clear slicer defaults in the report first.

**Manual bookmark state construction is rejected by ExportTo**
The `defaultBookmark.state` format is opaque — generated only by the Power BI JavaScript SDK. Tried every base64/JSON/URL-encoded variant. ExportTo API rejects all manually constructed states with `InvalidRequest`.

**REST API `/visuals` and `/bookmarks` endpoints return 404 for service principals**
Both endpoints require user-delegated tokens, not app-only (service principal) auth.

**Selection pane visual names ≠ JS SDK `visual.name`**
Renaming slicers in Power BI Desktop (e.g., "Daily Slicer") changes the display name only. The JS SDK's `visual.name` is a GUID. Match slicers by `getSlicerState().targets` (table/column) instead of by name.

**ExportTo API cannot export HTML content visuals**
The `htmlContent` custom visual type is on Microsoft's export blocklist. It renders as "This visual does not support exporting" in the PDF. No API flag bypasses this — hard platform limitation.

**Puppeteer `setSlicerState` on inactive pages throws errors**
Only the currently active page's visuals are accessible in the embedded report. Must call `report.setPage(pageName)` before interacting with that page's slicers.

**Browser UTC timezone conversion breaks date slicer values**
Passing `"2026-05-25T00:00:00"` to `setSlicerState` causes the browser (running in UTC+7) to convert to `"2026-05-24T17:00:00.000Z"`. Send date-only strings (`"2026-05-25"`) to avoid the shift.

---

### What works

**Puppeteer PDF generation (`export_report_pdf.js`)**
- Loads the embedded report in headless Chrome
- Navigates to each page and calls `setSlicerState()` on date slicers, matched by table/column via `getSlicerState().targets`
- Applies value filters (regional/area/store) via `report.setFilters()`
- Waits for `rendered` event + 4s buffer for HTML visuals to finish
- Generates PDF via `page.pdf()` — HTML visuals render correctly

**Regional/area/store filter via `reportLevelFilters`**
These fields live in the filter pane (not slicers), so `reportLevelFilters` works correctly for them.

**Date slicer values must be date-only strings**
Strip `T00:00:00` before passing to `setSlicerState` to avoid UTC conversion.
