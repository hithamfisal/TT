# Report Center KPI / Performance Export Options

This update centralizes KPI and Performance KPI card exports inside the Report Export Center.

Updated files:
- src/pages/Home.tsx

Changes:
- Added KPI Cards Export report type with PNG and PDF formats.
- Added Performance KPI Cards Export report type with PNG and PDF formats.
- Added quick report cards for:
  - KPI Cards PNG
  - KPI Cards PDF
  - Performance Cards PNG
  - Performance Cards PDF
- Removed KPI/Performance card export buttons from their tabs.
- Export generation temporarily renders the required tab, captures the full card grid, then returns to the Reports tab.
- XLSX/report exports remain centralized in the Report Export Center.
