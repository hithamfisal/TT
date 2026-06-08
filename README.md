# DMR Ticketing Dashboard

## Duration Calculation Rules

Durations are parsed by `src/lib/dateUtils.ts`.

Supported duration formats include:

```text
0 days 2 hrs 15 mins
0 d, 2 h, 15 m
2 hrs
02:15
2.5
```

When a duration text cannot be parsed, analytics fall back to the calculated resolution hours from observation date/time to recovery date/time.

Performance downtime uses one contribution per `TT + RF Site` pair, so duplicated rows for the same ticket/site do not inflate downtime.

For a selected month, downtime is calculated by outage overlap with that month:

```text
siteDownHours = sum(overlapHours between outage window and selected month)
```

For `All`, downtime uses the parsed or calculated ticket duration, still counted once per `TT + RF Site`.

Site downtime is capped to the selected period hours:

```text
siteDownHours = min(calculatedSiteDownHours, selectedPeriodHours)
```


## Formula Helpers

### Clamp

`clamp(value, min, max)` keeps a calculated number inside an allowed range:

```text
clamp(value, min, max) = Math.max(min, Math.min(max, value))
```

Examples:

```text
clamp(25, 0, 20) = 20
clamp(-5, 0, 20) = 0
clamp(12, 0, 20) = 12
```

The dashboard uses `clamp` to stop one metric from overpowering the full score. For example, the downtime penalty can never exceed 20 points even if downtime is very high.

### Region + RF Site Key

Several analytics use a region-aware site key:

```text
analyticsSiteKey = normalizedRegion + RF Site ID
```

Region normalization:

```text
EOA and NEOA -> EOA
SOA          -> SOA
COA          -> COA
WOA          -> WOA
```

This prevents `RF SITE 28` in SOA from being merged with `RF SITE 28` in EOA.

## High Risk Sites Ranking

The High Risk Sites Ranking is calculated in `src/lib/ticketAnalytics.ts`.

Each RF Site receives a risk score from `0` to `100` using normalized site metrics:

```text
Risk Score =
  (ticketCount / maxTicketCount) * 20
+ (downtimeHours / maxDowntimeHours) * 30
+ (serviceImpactCount / maxServiceImpactCount) * 20
+ (criticalCount / maxCriticalCount) * 15
+ (missingRcaCount / maxMissingRcaCount) * 10
+ ((100 - reliability) / 100) * 5
```

The score is rounded to one decimal and clamped between `0` and `100`.

### Inputs

- `ticketCount`: number of tickets linked to the RF Site.
- `downtimeHours`: performance `Sites Down, hrs` when available; otherwise ticket duration total. It is capped to the selected period hours.
- `serviceImpactCount`: tickets where service impact includes `Service Impact`.
- `criticalCount`: tickets with severity equal to `Critical`.
- `missingRcaCount`: tickets with missing or not-provided RCA.
- `reliability`: performance reliability percentage for the RF Site.

### Risk Levels

```text
Critical: score >= 75
High:     score >= 50
Medium:   score >= 25
Low:      score < 25
```

The downtime shown in the ranking table is intentionally aligned with the reliability value by using the performance calculation whenever it exists. A site's displayed downtime cannot exceed the total hours in the selected month range.

The first column is `Region`. Region is part of the internal site key, so the same RF Site number in different regions is kept as separate rows instead of being joined.

## Repeated Offender Sites

The Repeated Offender Sites table is calculated in `src/lib/ticketAnalytics.ts`.

It uses grouped TT records and then expands each TT into its actual RF Site rows. This prevents one primary site name from being copied to every RF Site in a multi-site TT.

### Inclusion Rule

An RF Site appears in the table when:

```text
tickets > 1 OR downtimeHours > 0
```

### Site Metrics

```text
tickets = number of grouped TT records linked to the RF Site
```

```text
downtimeHours = performance Sites Down, hrs when available
```

If performance downtime is not available, the fallback is:

```text
downtimeHours = sum of ticket duration hours assigned to the RF Site
```

For multi-site TTs, the ticket duration is divided across the RF Sites in that TT:

```text
perSiteDowntime = ticketDurationHours / numberOfRfSitesInTheTT
```

Then:

```text
siteDowntimeHours = sum(perSiteDowntime for the RF Site)
```

### Ranking Order

Repeated offender sites are sorted by:

```text
1. downtimeHours descending
2. tickets descending
3. siteId ascending
```

The table shows the top 10 sites after this ranking.

The first column is `Region`. Region is part of the internal site key, so repeated offender rows remain one region per RF Site row.

## Recommended Management Actions

Recommended Management Actions are calculated in `src/lib/ticketAnalytics.ts`.

The four recommendations are generated from:

```text
1. Most frequent RCA family by ticket count
2. Region with the highest missing RCA count
3. Pending tickets older than 7 days
4. Top repeated offender RF Site
```

The top repeated site action uses the corrected repeated offender calculation, including de-duplicated RF Site rows and performance downtime when available.

## RCA Downtime Charts And Cards

The following visuals use RCA duration aggregation:

- `Top RCA Families by Downtime`
- `Top 10 RCA by total downtime (hrs)`
- `Highest MTTR by RCA`
- `Top RCA / Downtime`
- `Highest MTTR RCA`

Each grouped ticket contributes once per RCA:

```text
rcaTicketKey = TT + RCA
```

The duration contribution is:

```text
durationHours = min(parsedOrCalculatedTicketDuration, selectedPeriodHours)
```

Then:

```text
RCA total downtime = sum(durationHours for the RCA)
RCA MTTR = RCA total downtime / numberOfTicketsWithPositiveDuration
```

This prevents duplicate rows for the same TT/RCA from inflating RCA downtime and prevents a single ticket from exceeding the selected reporting period.

### Highest MTTR By RCA

The `Highest MTTR by RCA` chart and `Highest MTTR RCA` card use the same de-duplicated RCA duration source:

```text
rcaTicketKey = TT + RCA
```

Only tickets with positive duration are included:

```text
positiveDurationTickets = count(TT + RCA where durationHours > 0)
```

The MTTR value for each RCA is:

```text
RCA MTTR = sum(durationHours for RCA) / positiveDurationTickets
```

The chart is sorted by:

```text
1. RCA MTTR descending
2. RCA name ascending
```

The `Highest MTTR RCA` card displays the first RCA after this sorting.

## Performance KPIs And Gauges

Performance KPIs and gauges are calculated from `PerfRow` records in `src/pages/Home.tsx`.

Each performance row represents one RF Site in the selected region/month scope.

### Site Availability

For each RF Site:

```text
siteDownHours = calculated/capped Sites Down, hrs
siteAvailableHours = selectedPeriodHours - siteDownHours
```

The selected period is:

```text
selectedPeriodHours = total hours in selected month(s)
```

For `All`, the selected period is the sum of all loaded months.

### Total Availability

```text
totalAvail = sum(siteAvailableHours)
totalDown = sum(siteDownHours)
totalHours = totalAvail + totalDown
```

### % Availability

```text
% Availability = (totalAvail / totalHours) * 100
```

### Affected Sites

```text
Affected Sites = count(unique RF Site where siteDownHours > 0)
```

### Non-Affected Sites

```text
Non-Affected Sites = count(unique RF Site) - Affected Sites
```

### Total Down

```text
Total Down = sum(siteDownHours)
```

Displayed as:

```text
days, hours, minutes
```

### MTTR

In the performance KPI cards/gauges:

```text
MTTR = totalDown / Affected Sites
```

This is the average downtime per affected RF Site in the selected scope.

### MTBF

```text
MTBF = totalHours / totalDown
```

If there is no downtime, MTBF is blank.

### MTTF

```text
MTTF = MTBF + MTTR
```

If MTBF or MTTR is unavailable, MTTF is blank.

### Gauge Direction

Higher is better:

```text
% Availability
MTBF
MTTF
Non-Affected Sites
```

Lower is better:

```text
MTTR
Affected Sites
Total Down
```

The gauges use these KPI values directly and convert them to a visual progress scale based on the gauge thresholds in `PERFORMANCE_GAUGE_CONFIG`.

## Dashboard-Wide Calculation Reference

Most ticket calculations start from grouped TT records:

```text
groupedTicket = one distinct TT number
primaryRow = earliest observation row for that TT
rows = all source rows with the same TT
```

Global dashboard filters affect the analytics tabs except the report export cards. Report export cards use their own month and region dropdowns.

## KPI Cards

### Total TT's

```text
Total TT's = count(distinct TT after filters)
```

### Closed / Pending / Status Cards

```text
Status count = count(primaryRow where Status = selected status)
```

### Critical / Severity Cards

```text
Severity count = count(primaryRow where Severity = selected severity)
```

### Region Cards

```text
Region count = count(primaryRow where Region = selected region)
```

### Service Impact

```text
Service Impact count = count(primaryRow where Service Impaction Status contains "Service Impact")
```

### RCA Updated

```text
RCA Updated = count(primaryRow where Action exists OR RCA is not missing)
```

### RCA Not Provided

```text
RCA Not Provided = count(primaryRow where RCA is blank, missing, N/A, not provided, or equivalent)
```

### Preventable Events

```text
Preventable Events = count(primaryRow where RCA family/preventability rule = Preventable)
```

### No. Of Sites

```text
No. Of Sites = count(unique RF Site from performance rows after filters)
```

### Affected Sites

```text
Affected Sites = count(unique RF Site where performance siteDownHours > 0)
```

### Non-Affected Sites

```text
Non-Affected Sites = count(unique RF Site) - Affected Sites
```

## Tickets Data Table

The Tickets Data Table is based on distinct TT groups.

```text
one displayed row = one distinct TT
```

For multi-site TTs:

```text
Site ID cell = all RF Site IDs for that TT
Site Name cell = all site names for that TT
```

The primary row fields, such as severity, status, dates, owner, RCA, and action, come from the earliest observation row in that TT group unless a table-specific export template uses all rows.

## Monthly Tickets Export

Monthly ticket export filters by:

```text
selected export month(s)
selected export region(s)
```

It does not use the global tab filters.

Template routing:

```text
EOA + NEOA -> EOA template
SOA        -> SOA template
COA        -> COA template
WOA        -> WOA template
```

If multiple template regions are selected, the dashboard exports multiple files instead of merging them.

## Overview And Trend Charts

### Tickets Per Month / Week

When no opening month filter is selected:

```text
grain = month
opened = count(primaryRow by openingMonthKey)
resolved = count(primaryRow where Status contains "resolved" by recovery month)
```

When one or more opening months are selected:

```text
grain = week
opened = count(primaryRow by observation week)
resolved = count(primaryRow where Status contains "resolved" by recovery week)
```

### Status Distribution

```text
Status Distribution = count(primaryRow by Status)
```

### Severity Distribution

```text
Severity Distribution = count(primaryRow by Severity)
```

### Region Distribution

```text
Region Distribution = count(primaryRow by Region)
```

### Impact Distribution

```text
Impact Distribution = count(primaryRow by Service Impaction Status)
```

### Escalation Distribution

```text
Escalation Distribution = count(primaryRow by Escalation Level)
```

### Top Sites By Tickets

```text
Top Sites = count(distinct TT linked to each RF Site)
```

Sorted by:

```text
1. ticket count descending
2. site label ascending
```

## Executive Insights

### Network Health Score

The health score starts at `100` and subtracts penalties:

```text
Health Score =
100
- downtimePenalty
- affectedSitesPenalty
- criticalPenalty
- pendingPenalty
- missingRcaPenalty
- reliabilityPenalty
```

Penalty formulas:

```text
downtimePenalty = clamp(totalDowntime / 20, 0, 20)
affectedSitesPenalty = clamp(affectedSites * 1.5, 0, 15)
criticalPenalty = clamp(criticalTickets * 4, 0, 20)
pendingPenalty = clamp(pendingTickets * 3, 0, 15)
missingRcaPenalty = clamp((missingRca / totalTickets) * 100 * 0.12, 0, 12)
reliabilityPenalty = clamp(100 - avgReliability, 0, 18)
```

Health status:

```text
Excellent: score >= 90
Good:      score >= 75
Warning:   score >= 55
Critical:  score < 55
```

### Executive Cards

```text
Service Impact Rate = serviceImpactCount / totalTickets * 100
RCA Completion Rate = (totalTickets - missingRca) / totalTickets * 100
Preventable Events % = preventableCount / totalTickets * 100
Average Reliability = average(performance reliability values)
```

## Operational Quality & Follow-Up Priorities

This section is calculated in `src/lib/ticketAnalytics.ts` using the current filtered scope.

### Avg FRT

```text
Avg FRT = average(first reply/escalation hours)
```

If a source FRT field exists, it is used. Otherwise:

```text
FRT = Escalated for L3 Support Date/Time - Observation Date/Time
```

### Avg Response

```text
Avg Response = average(responseHours)
```

If a source response time exists, it is used. Otherwise it uses the same fallback as FRT:

```text
Response = Escalated for L3 Support Date/Time - Observation Date/Time
```

### Avg Resolution

```text
Avg Resolution = average(resolutionHours)
```

Resolution fallback:

```text
resolutionHours = parsed duration
```

If duration is not available:

```text
resolutionHours = Recovery Date/Time - Observation Date/Time
```

### SLA Breaches

Targets:

```text
FRT target        = 1 hour
Response target   = 4 hours
Resolution target = 24 hours
```

Breaches:

```text
FRT Breaches = count(primaryRow where frtHours > 1)
Response Breaches = count(primaryRow where responseHours > 4)
Resolution Breaches = count(primaryRow where resolutionHours > 24)
```

### Pending Aging Buckets

Only pending/open tickets are included:

```text
pendingAgeHours = currentDateTime - Observation Date/Time
```

Buckets:

```text
0-24h
1-3d
3-7d
7d+
```

### Preventability By Tickets

```text
Preventability By Tickets = count(primaryRow by preventability rule)
```

The preventability rule is derived from RCA classification.

### Preventability By Downtime

```text
Preventability By Downtime = sum(durationHours by preventability rule)
```

Duration uses the safe parsed/calculated duration rules.

### RCA Family Deep-Dive

For each RCA family:

```text
tickets = count(distinct TT + RCA family)
downtimeHours = sum(capped durationHours for TT + RCA family)
missingRca = count(missing RCA records in that family)
preventableTickets = count(preventable records in that family)
serviceImpactTickets = count(service-impact records in that family)
```

Sorted by:

```text
1. downtimeHours descending
2. tickets descending
```

### Missing RCA By Region

```text
Missing RCA By Region = count(primaryRow with missing RCA by Region)
percentage = regionMissingRca / totalTickets * 100
```

### Responsible Team Load

```text
Responsible Team Load = count(primaryRow by responsible team)
percentage = teamCount / totalTickets * 100
```

Responsible team is derived from RCA family rules when not explicitly provided.

## RCA / SLA Charts

### Top RCA Families By Downtime

```text
RCA family downtime = sum(capped durationHours for distinct TT + RCA family)
```

Sorted by:

```text
1. downtimeHours descending
2. tickets descending
```

### Top 10 RCA By Total Downtime

```text
RCA downtime = sum(capped durationHours for distinct TT + RCA)
```

Sorted by:

```text
1. RCA downtime descending
2. RCA name ascending
```

### Highest MTTR By RCA

```text
RCA MTTR = RCA total downtime / positiveDurationTickets
```

Sorted by:

```text
1. RCA MTTR descending
2. RCA name ascending
```

### RCA Family Over Time

```text
RCA Family Over Time = count(primaryRow by opening month and RCA family)
```

## Performance Tables And Charts

### Performance Table

Each row is one RF Site in the selected performance scope.

```text
Site Availability, Hrs = selectedPeriodHours - Sites Down, hrs
Site Availability, days = Site Availability, Hrs converted to d/h/m
DMR Reliability = Site Availability, Hrs / selectedPeriodHours * 100
Sites Down, hrs = calculated/capped downtime for that RF Site
```

### Site Availability Chart

```text
chart value = Site Availability, Hrs
```

### Site Downtime Chart

```text
chart value = Sites Down, hrs
```

### Downtime Status Labels

```text
0 hrs       -> No Downtime
> 0 to 8    -> Moderate
> 8 to 24   -> High
> 24        -> Critical
```

## Export Center

Report exports use their own month and region controls:

```text
Ticket exports use exportMonths + exportRegions
Performance exports use perfMonths + perfRegions
```

The global dashboard filters do not affect the report export cards.

## Calculation Audit Checklist

The README documents the dashboard calculations for these areas:

```text
Duration parsing and downtime capping
Grouped/distinct TT logic
Monthly tickets export logic
Monthly performance export logic
Dashboard KPI cards
Performance KPI cards and gauges
Performance table calculations
Site availability chart
Site downtime chart
Network Health Score
High Risk Sites Ranking
Repeated Offender Sites
Recommended Management Actions
RCA downtime charts and cards
Highest MTTR by RCA
Operational Quality & Follow-Up Priorities
SLA averages and breach buckets
Pending aging buckets
Preventability by tickets and downtime
RCA family deep-dive
Missing RCA by region
Responsible team load
Overview charts
Trend charts
Tickets data table
Report export filter behavior
```

If a new dashboard card, table, chart, or export is added, add its formula here at the same time as the code change.
