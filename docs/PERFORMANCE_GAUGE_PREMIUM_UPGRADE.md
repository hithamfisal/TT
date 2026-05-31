# Performance KPI Premium Gauge Upgrade

This update enhances the top Performance KPI cards with:

- subtle load animation
- needle-style gauges instead of ring gauges
- mini sparklines under each KPI
- one central configuration block for colors and thresholds

Updated files:
- src/pages/Home.tsx
- src/index.css

Key implementation notes:
- `PERFORMANCE_GAUGE_CONFIG` is the single configurable block for KPI colors, thresholds, scales, labels, and sparkline logic.
- Gauge status labels are derived from the threshold config.
- Needle position is animated on load.
- Sparkline stroke is animated on load.
- Existing KPI values and dashboard behavior are preserved.
