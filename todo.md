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
- [x] Save a new checkpoint for the validated interval-overlap update.

# Interval-Overlap Update ZIP Delivery Todo

- [x] Package the latest validated dashboard source with the interval-overlap monthly export refinement as a lightweight ZIP archive.
- [x] Exclude dependency folders, build output, logs, and git metadata from the archive.
- [x] Verify the ZIP archive integrity before delivery.
- [x] Deliver the ZIP archive to the user.

# TT 16262 April 2026 Export Investigation Todo

- [x] Inspect the current Report Month export inclusion logic for date parsing and interval overlap.
- [x] Locate available workbook data containing TT 16262.
- [x] Check TT 16262 Observation Date, Recovery Date, Status, and parsed date values against April 2026.
- [x] Determine that TT 16262 matches April 2026 by interval overlap but can be excluded when the dashboard Opening Month filter is active because the monthly export was based on `filteredTickets`.
- [x] Apply the fix so the monthly export ignores the dashboard Opening Month slicer while preserving the other dashboard filters.

# TT 16262 Fix ZIP Delivery Todo

- [x] Package the latest dashboard source after the TT 16262 monthly export fix.
- [x] Exclude dependency folders, build output, logs, and git metadata from the archive.
- [x] Verify the ZIP archive exists and can be listed.
- [x] Deliver the ZIP archive to the user.

# Header Logo Layout Revision Todo

- [x] Locate existing National Grid and Nasco logo assets or embedded logo references from the project workspace; no reusable local logo files were found, so embedded wordmark-style logos were created in the header.
- [x] Update the hero/header area to match the provided reference composition with a compact title, top-right actions, and export panel.
- [x] Embed the National Grid and Nasco logos in a professional header arrangement.
- [x] Validate the production build after the header update.
- [ ] Save a new checkpoint and deliver the updated version and ZIP archive.

# Header Update ZIP Delivery Todo

- [ ] Confirm the latest header-logo ZIP archive is available.
- [ ] Attach the ZIP archive for the user.

# Hero Cleanup and ZIP Delivery Todo

- [x] Remove the descriptive paragraph below the Follow-Up Sheets Dashboard title.
- [x] Remove the buttons/chips below the title.
- [ ] Validate the production build after the hero cleanup.
- [ ] Create a fresh lightweight ZIP archive.
- [x] Deliver the updated ZIP archive to the user.

# RCA Column Replacement and ZIP Delivery Todo

- [x] Inspect current ticket parsing and report/export column mappings for Comments.
- [x] Replace Comments with the new RCA column from `Tickets_Data` in the full report layout.
- [x] Replace Comments with RCA in the monthly CSV and Excel exports.
- [x] Validate the production build after the RCA update.
- [x] Create a fresh lightweight ZIP archive after all requested changes.
- [x] Deliver the updated ZIP archive to the user.

# RCA Measures Advisory Todo

- [x] Recommend practical RCA-based measures for each TT.
- [x] Group RCA categories into operational families for clearer management reporting.
- [x] Suggest dashboard KPIs, charts, and export fields that can use the RCA column.

# RCA KPI, Formula, and Dashboard Enhancement Todo

- [ ] Inspect the uploaded Excel workbook sheets, table headers, and existing report structure.
- [ ] Add RCA Family mapping using corrected categories, including DC Charger Faulty under Power & Environment and Media Converter Faulty under Fiber & Physical.
- [ ] Add derived Excel columns for RCA Family, Preventable / Non-preventable, Responsible Team, and Recommended Action.
- [ ] Add Excel formulas or summary cells for Top RCA by TT Count, Top RCA by Downtime, Highest MTTR RCA, Repeated RCA Sites, RCA not Provided %, and Preventable RCA %.
- [ ] Update the web dashboard parser and data model with the derived RCA fields.
- [ ] Add RCA KPI cards and management summaries to the web page.
- [ ] Add RCA-derived columns to monthly CSV/Excel exports and the full report layout.
- [ ] Widen the monthly export filter box so the Report Month selector and export buttons are clear and readable.
- [ ] Validate the production build and workbook integrity.
- [ ] Package and deliver the updated workbook and latest source ZIP.
