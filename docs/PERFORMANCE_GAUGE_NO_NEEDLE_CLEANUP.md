# Performance Gauge Cleanup

This update adjusts the Performance KPI gauge cards based on the latest design feedback.

Updated files:
- src/pages/Home.tsx
- src/index.css

Changes:
- Removed the gauge needle.
- Removed the small sparkline under each KPI.
- Enlarged the half-circle gauge arc.
- Kept the icon/value/caption inside the card, with the larger arc surrounding the KPI content more cleanly.
- Preserved all KPI formulas, values, section controls, collapse/expand, and PNG export behavior.
