# Dashboard Sheet Matching Todo

- [x] Inventory every KPI card and chart on the workbook Dashboard sheet.
- [x] Confirm which workbook Dashboard metrics should use global unique TT counting and which should use affected-site exposure counting.
- [x] Preserve local saved-session restoration from browser storage.
- [x] Add or revise web KPI cards to match the workbook Dashboard sheet.
- [x] Add or revise web charts to match the workbook Dashboard sheet.
- [x] Validate TypeScript and production build.
- [ ] Save a new checkpoint and deliver the updated dashboard version.

# ZIP Delivery Todo

- [x] Create updated ZIP archive after latest validated dashboard changes and deliver it to the user.
- [x] Package the current dashboard source code into a downloadable ZIP archive.
- [x] Verify the archive was created successfully.
- [x] Send the ZIP file to the user.

# Deployment Guide Todo

- [ ] Create a deployment guide for the Follow-Up Sheets dashboard.
- [ ] Include local setup, build, preview, and Manus publishing steps.
- [ ] Include optional external hosting notes and compatibility cautions.
- [ ] Deliver the guide as a Markdown document.

# Chart Label and Card Revision Todo

- [ ] Add clear visible data labels to all dashboard charts where supported.
- [ ] Replace the Service Impact table/list with a chart.
- [ ] Show full site/category names on the horizontal bar chart axis.
- [ ] Remove the Total Site Effected KPI card from the dashboard cards.
- [ ] Validate TypeScript and production build after the changes.
- [ ] Save a new checkpoint and deliver the updated dashboard version.

# Month Slicer Guidance Todo

- [x] Recommend the best Excel/source-data approach for filtering TT records by opening month.
- [x] Explain how the same month slicer should be implemented in the web dashboard.
- [x] Include practical sorting and unique-TT-counting considerations.

# Opening Month Slicer Implementation Todo

- [x] Add month key and label helpers based on TT opening / Observation Date.
- [x] Extend dashboard filter state with an Opening Month slicer.
- [x] Populate month options from unique ticket opening months in chronological order.
- [x] Apply the month filter before KPI, chart, and register calculations.
- [x] Validate TypeScript after the implementation.
- [x] Send the modified Home.tsx file to the user.

# Excel Opening Month Slicer Implementation Todo

- [x] Inspect the uploaded Follow-Up Sheets workbook and identify the Tickets_Data opening date column.
- [x] Add Opening Month Key and Opening Month helper fields to Tickets_Data.
- [x] Add a workbook-facing Opening Month slicer/control area using unique chronological months.
- [x] Preserve the original workbook structure as much as possible and save a modified copy.
- [x] Validate the modified workbook can be opened and inspected.
- [x] Deliver the modified Excel workbook and updated Home.tsx file.

# Opening Month Defect Fix Todo

- [ ] Replace static Excel Opening Month helper values with formulas based on Observation Date.
- [ ] Audit the web dashboard data flow to find metrics and charts still using unfiltered data.
- [ ] Apply the selected Opening Month before all KPI, chart, export, and register calculations.
- [ ] Validate the workbook formulas and web dashboard build.
- [ ] Deliver corrected Home.tsx and Excel workbook files.
