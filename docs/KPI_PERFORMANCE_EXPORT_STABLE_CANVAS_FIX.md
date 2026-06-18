# KPI / Performance Export Stable Canvas Fix

This update fixes KPI Cards and Performance KPI Cards export actions from the Reports tab.

Changes:
- KPI Cards PNG/PDF now generates from a stable canvas renderer.
- Performance KPI Cards PNG/PDF now generates from a stable canvas renderer.
- Export no longer depends on hidden tab DOM capture.
- Export no longer switches tabs or flashes the screen.
- PNG and PDF actions should trigger downloads immediately.
- Output styling was adjusted to be closer to the dashboard card layout with compact card grids.

Updated files:
- src/pages/Home.tsx
