# KPI Export and Report History Location Fix

This update fixes KPI/Performance card export and improves report history.

## Fixed
- KPI Cards PNG/PDF export now clones only the actual KPI cards into a dedicated export stage.
- Performance KPI Cards PNG/PDF export now clones only the actual gauge cards into a dedicated export stage.
- The card export no longer captures UI icons or empty placeholders.
- Card export now waits for the selected tab to render before capture.
- Report history now includes a Download Location column.
- For browser-generated KPI/Performance card exports, history includes a temporary download link for the generated Blob.
- For other exports, history shows the browser download destination note because the browser does not expose local file paths to web apps.

## Browser limitation
A web dashboard cannot read the user's real Downloads folder path or create a direct link to a local downloaded file path. The dashboard can only show the filename and a browser download location note. For exports generated as Blob URLs, the table can provide a temporary session link.
