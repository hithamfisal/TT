# Report Center XLSX Deduplication and Section PNG Removal

Updated files:
- src/pages/Home.tsx

Changes:
- Centralized XLSX report actions in the Reports tab / Report Export Center.
- Removed duplicated XLSX buttons from operational tabs.
- Removed section-header PNG export buttons from all section headers.
- Kept chart-level PNG export buttons intact.
- Kept section title/header and Top Nav controls intact.
- Added RCA / SLA Workbook quick action in the Report Export Center.

Notes:
- XLSX exports are now managed from the Reports tab to avoid duplicated export locations.
- Section-level PNG buttons were removed as requested.
- Chart PNG exports remain available on individual chart cards.
