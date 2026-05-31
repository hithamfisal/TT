# Section Navigation / Default Collapse / Chart PNG Fix

This update adjusts the dashboard section navigation behavior.

## Changes

- All dashboard sections now start collapsed by default after workbook load.
- Top navigation buttons scroll directly to each section control header, not the hidden section body.
- Section `Top Nav` buttons scroll back to the sticky dashboard navigation bar.
- Collapsed section headers now stay close together with reduced spacing.
- Chart PNG export detection was expanded to cover all Recharts charts, including charts inside glass cards and deep-dive panels.
- Section-level PNG export remains available from every section header.

## Files updated

- `src/pages/Home.tsx`
- `src/index.css`

## Test checklist

1. Upload or continue a workbook.
2. Confirm all sections are collapsed by default.
3. Click each top navigation button and confirm it scrolls to that section header.
4. Expand a section and confirm chart PNG buttons appear on all charts.
5. Click `Top Nav` from any section and confirm it returns to the dashboard navigation bar.
6. Use Expand All / Collapse All and confirm no large blank gaps remain.
