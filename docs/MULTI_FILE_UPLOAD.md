# Multi-file Upload from Browse Menu

This update allows selecting more than one workbook from the browse menu.

Updated files:
- src/pages/Home.tsx

Behavior:
- Main upload input now supports multiple files.
- Drag/drop upload now supports multiple files.
- New workbook(s) action can replace the current dashboard with multiple selected files.
- Add regions action can add multiple selected files to the current dashboard.
- Selected files are parsed using the existing workbook parsing logic.
- Parsed workbooks are merged using the same ticket aggregation/site-order behavior used by Add Region.

Preserved:
- Excel parsing rules
- ticket aggregation
- RCA logic
- KPI formulas
- exports
- section navigation/collapse behavior
