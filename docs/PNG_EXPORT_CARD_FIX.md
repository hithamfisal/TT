# PNG Export Card Fix

Fixed KPI and Performance KPI PNG export.

Problem:
- The export fallback was capturing the PNG icon SVG instead of the KPI cards.
- This created a large placeholder/icon image rather than the actual dashboard cards.

Fix:
- Chart SVG fallback now only targets Recharts SVG surfaces, not UI icon SVGs.
- KPI card export now uses a dedicated canvas renderer for KPI card grids.
- Performance gauge export now uses a dedicated canvas renderer for the gauge cards.
- Section-header PNG export remains removed.
- Chart-level PNG export remains available.
