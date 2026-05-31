# Section Navigation / Continue Previous Workbook / PNG Export

Added dashboard navigation controls similar to the CDR project workflow.

## Added

- Top section navigation buttons.
- Expand All / Collapse All button.
- Section control bar above each major section.
- Per-section PNG export.
- Return-to-top button for every section.
- Per-chart PNG export buttons.
- Continue previous workbook option using browser localStorage.

## Notes

The browser cannot permanently keep the original Excel file path for security reasons. The implementation saves the parsed dashboard snapshot locally and restores the dashboard state without requiring a new upload.
