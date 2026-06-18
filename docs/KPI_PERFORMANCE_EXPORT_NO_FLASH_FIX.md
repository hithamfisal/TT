# KPI / Performance Export No-Flash Fix

This update fixes KPI and Performance KPI card exports from the Report Export Center.

Root cause:
- The previous implementation temporarily switched tabs to KPI or Performance KPI, waited for the DOM, and then tried to capture visible card elements.
- This caused screen flashing, sometimes opened the KPI tab and returned, and sometimes failed to export because the target elements were not available/rendered at the capture moment.

Fix:
- KPI card PNG/PDF export now draws the KPI card report directly to a canvas from the existing dashboard data.
- Performance KPI PNG/PDF export now draws the Performance gauge report directly to a canvas from the existing performance KPI data.
- No tab switching is required.
- No screen flashing is required.
- No hidden DOM capture is required.
- The export history receives the real generated filename and a temporary browser download link where possible.

Browser limitation:
- Browsers cannot reveal the real local Downloads folder path.
- The history shows `Browser Downloads / filename.ext` and temporary download links for Blob-based exports.
