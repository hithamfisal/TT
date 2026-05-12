# Dashboard Sheet Matching Todo

- [x] Inventory every KPI card and chart on the workbook Dashboard sheet.
- [x] Confirm which workbook Dashboard metrics should use global unique TT counting and which should use affected-site exposure counting.
- [x] Preserve local saved-session restoration from browser storage.
- [x] Add or revise web KPI cards to match the workbook Dashboard sheet.
- [x] Add or revise web charts to match the workbook Dashboard sheet.
- [x] Validate TypeScript and production build.
- [ ] Save a new checkpoint and deliver the updated dashboard version.

# ZIP Delivery Todo

- [x] Package the latest distinct-report dashboard update as a ZIP archive for user download.
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

# Distinct Report Update Todo

- [ ] Inspect the current distinct report export and workbook report outputs.
- [ ] Update the distinct report columns to match the requested structure: #, Site ID, Site Name, Managed Resource, Severity, Issues, Observation Date, Observation Time, Recovery Date, Recovery Time, Escalated for L3 Support Date, Escalated for L3 Support Time, Total Duration Days/Hours, TT, Status, Escalated to, Comments-Feedback.
- [ ] Ensure the site output uses all affected sites from the affected sites column instead of only the primary site.
- [ ] Validate the modified dashboard source and generated Excel report artifact.
- [ ] Deliver the corrected modified files to the user.

# Excel Opening Month Dropdown Guidance Todo

- [x] Explain what the `OpeningMonthList` named range means in Excel data validation.
- [x] Explain how to create or verify the named range for the month dropdown source.
- [x] Explain how to make the Dashboard formulas respond to the selected month.
- [x] Provide practical formula examples for KPI cards and filtered distinct TT reports.

# Monthly TT Export Filter Todo

- [ ] Inspect the current web dashboard month filter and TT export functions.
- [ ] Add a dedicated export-month filter for TT monthly reports.
- [ ] Apply the monthly export inclusion logic: Observation Date is selected month, Recovery Date is selected month, or Status is pending.
- [ ] Ensure CSV and Excel exports use the same filtered TT report rows.
- [ ] Validate TypeScript and production build after implementation.
- [ ] Save a checkpoint and deliver the updated dashboard version.

# Separate Monthly TT Export Filter Clarification Todo

- [x] Keep the existing TT opened-month dashboard filter unchanged.
- [x] Add a separate monthly report export filter control.
- [x] Use the separate export filter only for CSV and Excel monthly TT report exports.
- [x] Apply monthly export logic as Observation Date in selected month OR Recovery Date in selected month OR Status pending.

# Latest ZIP Delivery Todo

- [x] Package the latest dashboard version with the separate monthly TT export filter as a lightweight ZIP archive.
- [x] Exclude dependency folders, build output, logs, and git metadata from the archive.
- [x] Verify the ZIP exists before delivering it to the user.

# Monthly Export Interval Overlap Refinement Todo

- [x] Keep the existing dashboard Opening Month filter unchanged.
- [x] Refine the separate Report Month export logic to include TTs where Observation Date is in the selected month.
- [x] Refine the separate Report Month export logic to include TTs where Recovery Date is in the selected month.
- [x] Refine the separate Report Month export logic to include TTs where Status is pending.
- [x] Add interval-overlap logic so TTs are included when the Observation-to-Recovery period covers any day in the selected month.
- [ ] Save a new checkpoint for the validated interval-overlap update.
