# Refactor Step 3 — Executive Insights and Network Health Score

This step adds a new executive analytics layer without changing existing workbook parsing, KPI formulas, exports, filters, charts, or visual theme.

## Added

- `src/lib/ticketAnalytics.ts`
  - `calculateExecutiveInsights()`
  - Network Health Score calculation
  - Executive summary text
  - Insight card data
  - High-risk site ranking

## Updated

- `src/types/dashboard.ts`
  - `NetworkHealthScore`
  - `ExecutiveInsightCard`
  - `HighRiskSiteRow`
  - `ExecutiveInsights`

- `src/pages/Home.tsx`
  - imports `calculateExecutiveInsights`
  - computes `executiveInsights` with `useMemo`
  - renders an Executive Insights section before the charts

## Preserved

- Excel upload and parsing behavior
- Ticket aggregation behavior
- RCA logic
- Performance KPI formulas
- Export behavior
- Current dark premium dashboard style

## Formula Notes

The Network Health Score starts from 100 and subtracts penalties for:

- total downtime
- affected sites
- critical tickets
- open/pending tickets
- missing RCA percentage
- average reliability reduction

The High Risk Sites score uses weighted factors:

- ticket count: 20%
- downtime hours: 30%
- service-impact tickets: 20%
- critical tickets: 15%
- missing RCA count: 10%
- reliability gap: 5%

These formulas are readable and intentionally centralized in `src/lib/ticketAnalytics.ts` so they can be adjusted later without touching the UI.
