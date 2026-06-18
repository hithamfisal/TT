# KPI / Performance Export Match Tab Layout Fix

This update changes KPI and Performance KPI PNG/PDF exports to capture a cloned copy of the real tab card layout instead of drawing a separate canvas template.

Updated:
- src/pages/Home.tsx

Result:
- KPI Cards export matches the KPI tab card design.
- Performance KPI export matches the Performance KPI tab gauge-card design.
- Export is done offscreen to avoid visible tab switching/flashing.
- Report history behavior is preserved.
