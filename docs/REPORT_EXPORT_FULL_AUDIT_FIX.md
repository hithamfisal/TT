# Report Export Full Audit Fix

This update reviews and fixes the Reports Center export flow, especially KPI and Performance card exports.

Updated files:
- src/pages/Home.tsx
- docs/REPORT_EXPORT_FULL_AUDIT_FIX.md

Fixes:
- KPI Cards PNG/PDF now captures a cloned visible export stage instead of a placeholder/icon or hidden tab content.
- Performance KPI Cards PNG/PDF now captures the full gauge-card grid cleanly.
- Export temporarily opens the required tab, waits for render, captures the content, then returns to the previous tab.
- Card export now returns success/failure so report history is added only after a real export.
- Quick action report cards now generate the report immediately instead of only changing the selected report type.
- Generated Reports History file names now include file extensions such as .png, .pdf, .xlsx, .ppt.
- Existing XLSX/PDF/PPT report exports remain in the Reports Center.
- Chart-level PNG export remains separate.
- Section-header PNG export remains removed.
