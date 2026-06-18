# Report Center Shared Filter Row Fix

Updated the Reports tab layout:

- Moved “Report Management Center” title and subtitle into the section header.
- Removed the duplicate title/subtitle from the report content area.
- Put Filtered Records and Output Format on one row.
- Put Performance Month and Performance Region on one row.
- Month and Region are shared filters for all report types.
- Switching report type does not reset or reload Month/Region selections.

Updated files:
- src/pages/Home.tsx
- src/index.css
