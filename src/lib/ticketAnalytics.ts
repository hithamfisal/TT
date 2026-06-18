import type {
  ExecutiveInsights,
  ExecutiveInsightCard,
  NetworkHealthScore,
  PerfRow,
  HighRiskSiteRow,
  DeepDiveAnalytics,
  RcaDeepDiveRow,
  SlaBucket,
  TicketAggregate,
  TicketRecord,
} from "../types/dashboard";

import {
  average,
  clean,
  combineDateTime,
  isRfSiteId,
  normalizeSiteId,
  openingMonthKey,
  parseDurationHours,
  totalHoursInMonth,
} from "./dateUtils";
import {
  getPreventability,
  getRcaFamily,
  getRecommendedAction,
  getResponsibleTeam,
  rcaNotProvided,
} from "./rcaRules";

export type ExecutiveInsightsInput = {
  tickets: TicketAggregate[];
  performanceRows: PerfRow[];
};

const clamp = (value: number, min = 0, max = 100): number =>
  Math.max(min, Math.min(max, value));

const round1 = (value: number): number => Math.round(value * 10) / 10;
const round2 = (value: number): number => Math.round(value * 100) / 100;

function pctNumber(value: number, total: number): number {
  if (!total) return 0;
  return (value / total) * 100;
}

function formatPct(value: number): string {
  return `${round1(value)}%`;
}

function normalizeAnalyticsRegion(region: string): string {
  const value = clean(region).toUpperCase();
  if (value === "EOA" || value === "NEOA") return "EOA";
  if (value === "SOA") return "SOA";
  if (value === "COA") return "COA";
  if (value === "WOA") return "WOA";
  return clean(region);
}

function inferAnalyticsRegionFromText(value: string): string {
  const text = clean(value).toUpperCase();
  if (/\b(?:EOA|NEOA)\b/.test(text)) return "EOA";
  if (/\bSOA\b/.test(text)) return "SOA";
  if (/\bCOA\b/.test(text)) return "COA";
  if (/\bWOA\b/.test(text)) return "WOA";
  return "";
}

function primaryAnalyticsRegion(regions: string[], siteName = ""): string {
  const inferred = inferAnalyticsRegionFromText(siteName);
  if (inferred) return inferred;
  const normalized = regions.map(normalizeAnalyticsRegion).find(Boolean);
  return normalized || "-";
}

function analyticsSiteKey(region: string, siteId: string): string {
  const normalizedRegion = normalizeAnalyticsRegion(region);
  return `${normalizedRegion || "UNKNOWN"}||${normalizeSiteId(clean(siteId))}`;
}

function buildPerformanceSiteMap(rows: PerfRow[]): Map<string, PerfRow> {
  const map = new Map<string, PerfRow>();
  rows.forEach((row) => {
    const region =
      inferAnalyticsRegionFromText(row.siteName) ||
      inferAnalyticsRegionFromText(row.displayName) ||
      inferAnalyticsRegionFromText(row.sourceLabel);
    const key = analyticsSiteKey(region, row.siteId);
    map.set(key, row);

    const fallbackKey = analyticsSiteKey("", row.siteId);
    if (!map.has(fallbackKey)) map.set(fallbackKey, row);
  });
  return map;
}

function getPerformanceSite(
  perfBySite: Map<string, PerfRow>,
  siteId: string,
  region: string,
): PerfRow | undefined {
  return (
    perfBySite.get(analyticsSiteKey(region, siteId)) ??
    perfBySite.get(analyticsSiteKey("", siteId))
  );
}

function parseReliability(value: string | number | null | undefined): number {
  if (typeof value === "number" && Number.isFinite(value)) {
    return value > 1 ? value : value * 100;
  }
  const parsed = Number(
    String(value ?? "")
      .replace("%", "")
      .trim(),
  );
  return Number.isFinite(parsed) ? parsed : 100;
}

function ticketDurationHours(row: TicketRecord): number {
  const parsed = parseDurationHours(row.duration);
  if (parsed !== null && Number.isFinite(parsed)) return parsed;
  return typeof row.resolutionHours === "number" &&
    Number.isFinite(row.resolutionHours)
    ? row.resolutionHours
    : 0;
}

function countByName(values: string[]): Array<{ name: string; value: number }> {
  const map = new Map<string, number>();
  values.forEach((value) => {
    const key = clean(value) || "Unknown";
    map.set(key, (map.get(key) ?? 0) + 1);
  });
  return Array.from(map.entries())
    .map(([name, value]) => ({ name, value }))
    .sort((a, b) => b.value - a.value || a.name.localeCompare(b.name));
}

function riskLevel(score: number): HighRiskSiteRow["riskLevel"] {
  if (score >= 75) return "Critical";
  if (score >= 50) return "High";
  if (score >= 25) return "Medium";
  return "Low";
}

function healthStatus(score: number): NetworkHealthScore["status"] {
  if (score >= 90) return "Excellent";
  if (score >= 75) return "Good";
  if (score >= 55) return "Warning";
  return "Critical";
}

function buildHealthReason(
  reductions: Array<{ label: string; points: number }>,
): string {
  const main = reductions
    .filter((item) => item.points > 0)
    .sort((a, b) => b.points - a.points)[0];

  return main
    ? `${main.label} is the largest score reduction.`
    : "No major risk driver detected in the selected scope.";
}

export function calculateExecutiveInsights({
  tickets,
  performanceRows,
}: ExecutiveInsightsInput): ExecutiveInsights {
  const primaryRows = tickets.map((ticket) => ticket.primary);
  const totalTickets = tickets.length;
  const totalDowntime = round1(
    primaryRows.reduce((sum, row) => sum + ticketDurationHours(row), 0),
  );

  const affectedSiteKeys = new Set<string>();
  const affectedSiteNames = new Map<string, string>();
  const ticketSitePairs = (ticket: TicketAggregate) => {
    const pairs = new Map<
      string,
      { siteKey: string; siteId: string; siteName: string; regions: string[] }
    >();

    const addPair = (siteId: string, siteName: string, region: string) => {
      const normalizedRegion =
        inferAnalyticsRegionFromText(siteName) ||
        normalizeAnalyticsRegion(region);
      const siteKey = analyticsSiteKey(normalizedRegion, siteId);
      const current = pairs.get(siteKey) ?? {
        siteKey,
        siteId,
        siteName: "",
        regions: [],
      };
      if (!current.siteName && siteName) current.siteName = siteName;
      if (normalizedRegion && !current.regions.includes(normalizedRegion)) {
        current.regions.push(normalizedRegion);
      }
      pairs.set(siteKey, current);
    };

    ticket.rows.forEach((row) => {
      const siteId = normalizeSiteId(clean(row.siteId));
      if (!isRfSiteId(siteId)) return;
      addPair(siteId, clean(row.siteName), row.region);
    });

    if (!pairs.size) {
      const primarySiteId = normalizeSiteId(clean(ticket.primary.siteId));
      if (isRfSiteId(primarySiteId)) {
        addPair(
          primarySiteId,
          clean(ticket.primary.siteName),
          ticket.primary.region,
        );
      }
    }

    ticket.siteIds.forEach((site) => {
      const siteId = normalizeSiteId(clean(site));
      const siteKey = analyticsSiteKey(ticket.primary.region, siteId);
      if (isRfSiteId(siteId) && !pairs.has(siteKey))
        addPair(siteId, "", ticket.primary.region);
    });

    return Array.from(pairs.values());
  };

  tickets.forEach((ticket) => {
    ticketSitePairs(ticket).forEach(({ siteKey, siteName }) => {
      affectedSiteKeys.add(siteKey);
      if (siteName && !affectedSiteNames.has(siteKey)) {
        affectedSiteNames.set(siteKey, siteName);
      }
    });
  });

  const criticalTickets = primaryRows.filter(
    (row) => clean(row.severity).toLowerCase() === "critical",
  ).length;
  const pendingTickets = primaryRows.filter((row) => {
    const status = clean(row.status).toLowerCase();
    return (
      status === "pending" || status === "open" || status.includes("pending")
    );
  }).length;
  const missingRca = primaryRows.filter((row) =>
    rcaNotProvided(row.rca),
  ).length;
  const serviceImpactCount = primaryRows.filter((row) =>
    clean(row.impact).toLowerCase().includes("service impact"),
  ).length;
  const preventableCount = primaryRows.filter(
    (row) =>
      (row.preventability || getPreventability(row.rca)) === "Preventable",
  ).length;

  const reliabilityValues = performanceRows
    .map((row) => parseReliability(row.reliability))
    .filter((value) => Number.isFinite(value));
  const avgReliability = reliabilityValues.length
    ? reliabilityValues.reduce((sum, value) => sum + value, 0) /
      reliabilityValues.length
    : 100;

  const affectedSites = affectedSiteKeys.size;
  const rcaFamilies = countByName(
    primaryRows.map((row) => row.rcaFamily || getRcaFamily(row.rca)),
  );
  const topRcaFamily = rcaFamilies[0] ?? { name: "N/A", value: 0 };

  const downtimeBySite = new Map<
    string,
    { siteId: string; siteName: string; downtime: number }
  >();
  const siteRisk = new Map<
    string,
    {
      siteId: string;
      regions: string[];
      siteName: string;
      ticketCount: number;
      downtimeHours: number;
      serviceImpactCount: number;
      criticalCount: number;
      missingRcaCount: number;
      rcaValues: string[];
    }
  >();

  tickets.forEach((ticket) => {
    const sites = ticketSitePairs(ticket);
    const duration = ticketDurationHours(ticket.primary);
    const perSiteDowntime = sites.length ? duration / sites.length : duration;
    const isServiceImpact = clean(ticket.primary.impact)
      .toLowerCase()
      .includes("service impact");
    const isCritical =
      clean(ticket.primary.severity).toLowerCase() === "critical";
    const isMissingRca = rcaNotProvided(ticket.primary.rca);
    const rcaFamily =
      ticket.primary.rcaFamily || getRcaFamily(ticket.primary.rca);

    sites.forEach(({ siteKey, siteId, siteName: pairedSiteName, regions }) => {
      const siteName = pairedSiteName || affectedSiteNames.get(siteKey) || "";

      const downtimeCurrent = downtimeBySite.get(siteKey) ?? {
        siteId,
        siteName,
        downtime: 0,
      };
      downtimeCurrent.downtime += perSiteDowntime;
      downtimeBySite.set(siteKey, downtimeCurrent);

      const current = siteRisk.get(siteKey) ?? {
        siteId,
        regions: [],
        siteName,
        ticketCount: 0,
        downtimeHours: 0,
        serviceImpactCount: 0,
        criticalCount: 0,
        missingRcaCount: 0,
        rcaValues: [],
      };

      current.ticketCount += 1;
      regions.forEach((region) => {
        if (region && !current.regions.includes(region))
          current.regions.push(region);
      });
      current.downtimeHours += perSiteDowntime;
      if (isServiceImpact) current.serviceImpactCount += 1;
      if (isCritical) current.criticalCount += 1;
      if (isMissingRca) current.missingRcaCount += 1;
      current.rcaValues.push(rcaFamily);
      siteRisk.set(siteKey, current);
    });
  });

  const perfBySite = buildPerformanceSiteMap(performanceRows);

  const performanceDowntimeBySite = performanceRows
    .filter((row) => isRfSiteId(row.siteId))
    .map((row) => ({
      siteId: normalizeSiteId(clean(row.siteId)),
      siteName: clean(row.siteName),
      downtime: row.sitesDownHours,
    }));
  const downtimeSource = performanceDowntimeBySite.length
    ? performanceDowntimeBySite
    : Array.from(downtimeBySite.values());

  const highestDowntime = downtimeSource.sort(
    (a, b) => b.downtime - a.downtime || a.siteId.localeCompare(b.siteId),
  )[0];

  const worstReliability = performanceRows
    .filter((row) => clean(row.siteId))
    .map((row) => ({
      siteId: clean(row.siteId),
      siteName: clean(row.siteName),
      reliability: parseReliability(row.reliability),
    }))
    .sort(
      (a, b) =>
        a.reliability - b.reliability || a.siteId.localeCompare(b.siteId),
    )[0];

  const maxTickets = Math.max(
    1,
    ...Array.from(siteRisk.values()).map((row) => row.ticketCount),
  );
  const maxDowntime = Math.max(
    1,
    ...Array.from(siteRisk.values()).map((row) => {
      const region = primaryAnalyticsRegion(row.regions, row.siteName);
      const perf = getPerformanceSite(perfBySite, row.siteId, region);
      return perf ? perf.sitesDownHours : row.downtimeHours;
    }),
  );
  const maxServiceImpact = Math.max(
    1,
    ...Array.from(siteRisk.values()).map((row) => row.serviceImpactCount),
  );
  const maxCritical = Math.max(
    1,
    ...Array.from(siteRisk.values()).map((row) => row.criticalCount),
  );
  const maxMissingRca = Math.max(
    1,
    ...Array.from(siteRisk.values()).map((row) => row.missingRcaCount),
  );

  // High Risk Sites Ranking formula:
  // score = ticket share * 20 + downtime share * 30 + service impact share * 20
  //       + critical severity share * 15 + missing RCA share * 10
  //       + reliability loss share * 5.
  // Downtime uses performance "Sites Down, hrs" when available, so it stays
  // aligned with the reliability shown in the same table.
  const highRiskSites: HighRiskSiteRow[] = Array.from(siteRisk.values())
    .map((site) => {
      const region = primaryAnalyticsRegion(site.regions, site.siteName);
      const perf = getPerformanceSite(perfBySite, site.siteId, region);
      const reliability = perf ? parseReliability(perf.reliability) : 100;
      const downtimeHours = perf ? perf.sitesDownHours : site.downtimeHours;
      const topRca = countByName(site.rcaValues)[0]?.name ?? "N/A";
      const riskScore = round1(
        clamp(
          (site.ticketCount / maxTickets) * 20 +
            (downtimeHours / maxDowntime) * 30 +
            (site.serviceImpactCount / maxServiceImpact) * 20 +
            (site.criticalCount / maxCritical) * 15 +
            (site.missingRcaCount / maxMissingRca) * 10 +
            ((100 - reliability) / 100) * 5,
        ),
      );

      return {
        rank: 0,
        region,
        siteId: site.siteId,
        siteName: site.siteName || perf?.siteName || "",
        ticketCount: site.ticketCount,
        downtimeHours: round1(downtimeHours),
        reliability: round2(reliability),
        topRca,
        riskScore,
        riskLevel: riskLevel(riskScore),
      };
    })
    .sort(
      (a, b) =>
        b.riskScore - a.riskScore ||
        b.downtimeHours - a.downtimeHours ||
        a.siteId.localeCompare(b.siteId),
    )
    .map((row, index) => ({ ...row, rank: index + 1 }))
    .slice(0, 10);

  const downtimePenalty = clamp(totalDowntime / 20, 0, 20);
  const affectedSitesPenalty = clamp(affectedSites * 1.5, 0, 15);
  const criticalPenalty = clamp(criticalTickets * 4, 0, 20);
  const pendingPenalty = clamp(pendingTickets * 3, 0, 15);
  const missingRcaPenalty = clamp(
    pctNumber(missingRca, totalTickets) * 0.12,
    0,
    12,
  );
  const reliabilityPenalty = clamp(100 - avgReliability, 0, 18);

  const scoreReductions = [
    { label: "Total downtime", points: downtimePenalty },
    { label: "Affected sites", points: affectedSitesPenalty },
    { label: "Critical tickets", points: criticalPenalty },
    { label: "Open/Pending tickets", points: pendingPenalty },
    { label: "Missing RCA", points: missingRcaPenalty },
    { label: "Reliability percentage", points: reliabilityPenalty },
  ];

  const score = Math.round(
    clamp(
      100 -
        downtimePenalty -
        affectedSitesPenalty -
        criticalPenalty -
        pendingPenalty -
        missingRcaPenalty -
        reliabilityPenalty,
    ),
  );

  const healthScore: NetworkHealthScore = {
    score,
    status: healthStatus(score),
    mainReason: buildHealthReason(scoreReductions),
    reductions: scoreReductions.map((item) => ({
      ...item,
      points: round1(item.points),
    })),
  };

  const followUpSites = highRiskSites.filter(
    (site) => site.riskLevel !== "Low",
  ).length;

  const summaryText = `During the selected period, the network recorded ${totalTickets.toLocaleString()} tickets, ${serviceImpactCount.toLocaleString()} service-impact events, ${affectedSites.toLocaleString()} affected sites, and ${round1(totalDowntime).toLocaleString()} hours of downtime. The leading RCA family was ${topRcaFamily.name}, while ${followUpSites.toLocaleString()} sites require follow-up.`;

  const rcaCompletionRate = 100 - pctNumber(missingRca, totalTickets);
  const serviceImpactRate = pctNumber(serviceImpactCount, totalTickets);
  const preventableRate = pctNumber(preventableCount, totalTickets);

  const cards: ExecutiveInsightCard[] = [
    {
      label: "Highest Downtime Site",
      value: highestDowntime?.siteId ?? "N/A",
      note: highestDowntime
        ? `${highestDowntime.siteName || "No site name"} · ${round1(highestDowntime.downtime)} hrs`
        : "No downtime in selected scope",
      tone: "#f59e0b",
    },
    {
      label: "Worst Reliability Site",
      value: worstReliability?.siteId ?? "N/A",
      note: worstReliability
        ? `${worstReliability.siteName || "No site name"} · ${round2(worstReliability.reliability)}%`
        : "No performance rows available",
      tone: "#ef4444",
    },
    {
      label: "Top RCA Family",
      value: topRcaFamily.name,
      note: `${topRcaFamily.value.toLocaleString()} tickets`,
      tone: "#22d3ee",
    },
    {
      label: "RCA Completion Rate",
      value: formatPct(rcaCompletionRate),
      note: `${missingRca.toLocaleString()} missing RCA`,
      tone: "#34d399",
    },
    {
      label: "Service Impact %",
      value: formatPct(serviceImpactRate),
      note: `${serviceImpactCount.toLocaleString()} service-impact tickets`,
      tone: "#fb7185",
    },
    {
      label: "Open/Pending Tickets",
      value: pendingTickets.toLocaleString(),
      note: "Tickets requiring closure follow-up",
      tone: "#f97316",
    },
    {
      label: "Affected Sites",
      value: affectedSites.toLocaleString(),
      note: "Unique affected sites in scope",
      tone: "#a78bfa",
    },
    {
      label: "Preventable Events %",
      value: formatPct(preventableRate),
      note: `${preventableCount.toLocaleString()} preventable tickets`,
      tone: "#38bdf8",
    },
  ];

  return {
    healthScore,
    summaryText,
    cards,
    highRiskSites,
    totals: {
      totalTickets,
      serviceImpactCount,
      affectedSites,
      totalDowntime,
      criticalTickets,
      pendingTickets,
      missingRca,
      rcaCompletionRate: round1(rcaCompletionRate),
      serviceImpactRate: round1(serviceImpactRate),
      preventableRate: round1(preventableRate),
      avgReliability: round2(avgReliability),
    },
  };
}

export type DeepDiveAnalyticsInput = {
  tickets: TicketAggregate[];
  performanceRows: PerfRow[];
  now?: Date;
};

const SLA_TARGETS = {
  frtHours: 1,
  responseHours: 4,
  resolutionHours: 24,
};

function bucketPendingAge(hours: number): SlaBucket["name"] {
  if (hours <= 24) return "0–24h";
  if (hours <= 72) return "1–3d";
  if (hours <= 168) return "3–7d";
  return "7d+";
}

export function calculateDeepDiveAnalytics({
  tickets,
  performanceRows,
  now = new Date(),
}: DeepDiveAnalyticsInput): DeepDiveAnalytics {
  const primaryRows = tickets.map((ticket) => ticket.primary);
  const totalTickets = Math.max(1, primaryRows.length);
  const perfBySite = buildPerformanceSiteMap(performanceRows);

  const familyMap = new Map<string, RcaDeepDiveRow>();
  const missingRcaByRegionMap = new Map<string, number>();
  const teamLoadMap = new Map<string, number>();
  const preventabilityCountMap = new Map<
    string,
    { count: number; downtime: number }
  >();
  const siteMap = new Map<
    string,
    {
      siteId: string;
      regions: string[];
      siteName: string;
      tickets: number;
      downtimeHours: number;
      rcas: string[];
    }
  >();
  const durationMonthKeys = Array.from(
    new Set(
      primaryRows
        .map(
          (row) => row.openingMonthKey || openingMonthKey(row.observationDate),
        )
        .filter((key) => key && key !== "Unknown"),
    ),
  );
  const selectedPeriodHours = durationMonthKeys.length
    ? durationMonthKeys.reduce((sum, key) => sum + totalHoursInMonth(key), 0)
    : 24 * 30;
  const countedFamilyTickets = new Set<string>();

  tickets.forEach((ticket) => {
    const row = ticket.primary;
    const rcaFamily = row.rcaFamily || getRcaFamily(row.rca);
    const downtime = Math.min(ticketDurationHours(row), selectedPeriodHours);
    const missingRca = rcaNotProvided(row.rca);
    const preventability = row.preventability || getPreventability(row.rca);
    const responsibleTeam =
      row.responsibleTeam || getResponsibleTeam(rcaFamily);
    const recommendedAction =
      row.recommendedAction || getRecommendedAction(rcaFamily);
    const isServiceImpact = clean(row.impact)
      .toLowerCase()
      .includes("service impact");

    const family = familyMap.get(rcaFamily) ?? {
      family: rcaFamily,
      tickets: 0,
      downtimeHours: 0,
      missingRca: 0,
      preventableTickets: 0,
      serviceImpactTickets: 0,
      responsibleTeam,
      recommendedAction,
    };

    const familyTicketKey = `${clean(row.tt) || row.rowNo}||${rcaFamily}`;
    if (!countedFamilyTickets.has(familyTicketKey)) {
      countedFamilyTickets.add(familyTicketKey);
      family.tickets += 1;
      family.downtimeHours += downtime;
      if (missingRca) family.missingRca += 1;
      if (preventability === "Preventable") family.preventableTickets += 1;
      if (isServiceImpact) family.serviceImpactTickets += 1;
    }
    familyMap.set(rcaFamily, family);

    if (missingRca) {
      const region = clean(row.region) || "Unknown";
      missingRcaByRegionMap.set(
        region,
        (missingRcaByRegionMap.get(region) ?? 0) + 1,
      );
    }

    teamLoadMap.set(
      responsibleTeam,
      (teamLoadMap.get(responsibleTeam) ?? 0) + 1,
    );

    const preventabilityItem = preventabilityCountMap.get(preventability) ?? {
      count: 0,
      downtime: 0,
    };
    preventabilityItem.count += 1;
    preventabilityItem.downtime += downtime;
    preventabilityCountMap.set(preventability, preventabilityItem);

    const sitePairs = new Map<
      string,
      { siteKey: string; siteId: string; siteName: string; regions: string[] }
    >();
    const addSitePair = (siteId: string, siteName: string, region: string) => {
      const normalizedRegion =
        inferAnalyticsRegionFromText(siteName) ||
        normalizeAnalyticsRegion(region);
      const siteKey = analyticsSiteKey(normalizedRegion, siteId);
      const current = sitePairs.get(siteKey) ?? {
        siteKey,
        siteId,
        siteName: "",
        regions: [],
      };
      if (!current.siteName && siteName) current.siteName = siteName;
      if (normalizedRegion && !current.regions.includes(normalizedRegion)) {
        current.regions.push(normalizedRegion);
      }
      sitePairs.set(siteKey, current);
    };

    ticket.rows.forEach((siteRow) => {
      const siteId = normalizeSiteId(clean(siteRow.siteId));
      if (!isRfSiteId(siteId)) return;
      addSitePair(siteId, clean(siteRow.siteName), siteRow.region);
    });
    if (!sitePairs.size) {
      const siteId = normalizeSiteId(clean(row.siteId));
      if (isRfSiteId(siteId))
        addSitePair(siteId, clean(row.siteName), row.region);
    }
    const perSiteDowntime = sitePairs.size
      ? downtime / sitePairs.size
      : downtime;

    sitePairs.forEach(({ siteKey, siteId, siteName, regions }) => {
      const region = primaryAnalyticsRegion(regions, siteName);
      const perf = getPerformanceSite(perfBySite, siteId, region);
      const current = siteMap.get(siteKey) ?? {
        siteId,
        regions: [],
        siteName: siteName || perf?.siteName || "",
        tickets: 0,
        downtimeHours: 0,
        rcas: [],
      };
      current.tickets += 1;
      regions.forEach((region) => {
        if (region && !current.regions.includes(region))
          current.regions.push(region);
      });
      current.downtimeHours += perSiteDowntime;
      current.rcas.push(rcaFamily);
      siteMap.set(siteKey, current);
    });
  });

  const rcaFamilyDeepDive = Array.from(familyMap.values())
    .map((row) => ({
      ...row,
      downtimeHours: round1(row.downtimeHours),
    }))
    .sort((a, b) => b.downtimeHours - a.downtimeHours || b.tickets - a.tickets)
    .slice(0, 10);

  const preventabilityByCount = Array.from(preventabilityCountMap.entries())
    .map(([name, value]) => ({
      name,
      value: value.count,
      percentage: round1((value.count / totalTickets) * 100),
    }))
    .sort((a, b) => b.value - a.value);

  const totalDowntime = Array.from(preventabilityCountMap.values()).reduce(
    (sum, item) => sum + item.downtime,
    0,
  );
  const preventabilityByDowntime = Array.from(preventabilityCountMap.entries())
    .map(([name, value]) => ({
      name,
      value: round1(value.downtime),
      downtimeHours: round1(value.downtime),
      percentage: totalDowntime
        ? round1((value.downtime / totalDowntime) * 100)
        : 0,
    }))
    .sort((a, b) => b.value - a.value);

  const missingRcaByRegion = Array.from(missingRcaByRegionMap.entries())
    .map(([name, value]) => ({
      name,
      value,
      percentage: round1((value / totalTickets) * 100),
    }))
    .sort((a, b) => b.value - a.value)
    .slice(0, 8);

  const responsibleTeamLoad = Array.from(teamLoadMap.entries())
    .map(([name, value]) => ({
      name,
      value,
      percentage: round1((value / totalTickets) * 100),
    }))
    .sort((a, b) => b.value - a.value)
    .slice(0, 8);

  const pendingAgeCounts = new Map<SlaBucket["name"], number>([
    ["0–24h", 0],
    ["1–3d", 0],
    ["3–7d", 0],
    ["7d+", 0],
  ]);
  primaryRows
    .filter(
      (row) =>
        clean(row.status).toLowerCase().includes("pending") ||
        clean(row.status).toLowerCase() === "open",
    )
    .forEach((row) => {
      const opened = combineDateTime(row.observationDate, row.observationTime);
      if (!opened) return;
      const ageHours = (now.getTime() - opened.getTime()) / 36e5;
      if (!Number.isFinite(ageHours) || ageHours < 0) return;
      const bucket = bucketPendingAge(ageHours);
      pendingAgeCounts.set(bucket, (pendingAgeCounts.get(bucket) ?? 0) + 1);
    });

  const slaSummary = {
    avgFrtHours: round1(average(primaryRows.map((row) => row.frtHours))),
    avgResponseHours: round1(
      average(primaryRows.map((row) => row.responseHours)),
    ),
    avgResolutionHours: round1(
      average(primaryRows.map((row) => row.resolutionHours)),
    ),
    frtBreaches: primaryRows.filter(
      (row) =>
        typeof row.frtHours === "number" && row.frtHours > SLA_TARGETS.frtHours,
    ).length,
    responseBreaches: primaryRows.filter(
      (row) =>
        typeof row.responseHours === "number" &&
        row.responseHours > SLA_TARGETS.responseHours,
    ).length,
    resolutionBreaches: primaryRows.filter(
      (row) =>
        typeof row.resolutionHours === "number" &&
        row.resolutionHours > SLA_TARGETS.resolutionHours,
    ).length,
    pendingAgingBuckets: Array.from(pendingAgeCounts.entries()).map(
      ([name, value]) => ({
        name,
        value,
        label: `${value} tickets`,
      }),
    ),
  };

  // Repeated Offender Sites formula:
  // include RF Sites with more than one ticket or any repeated ticket downtime.
  // Ticket downtime remains the repeated-offender downtime; performance fields
  // are joined by Region + RF Site ID so availability, performance downtime,
  // and reliability are shown beside the repeated TT pattern without merging
  // same-number RF Sites from different regions.
  const repeatedOffenderSites = Array.from(siteMap.values())
    .map((site) => {
      const region = primaryAnalyticsRegion(site.regions, site.siteName);
      const perf = getPerformanceSite(perfBySite, site.siteId, region);
      return {
        region,
        siteId: site.siteId,
        siteName: site.siteName || perf?.siteName || "",
        tickets: site.tickets,
        downtimeHours: round1(site.downtimeHours),
        performanceAvailabilityHours: perf ? round1(perf.availHours) : null,
        performanceDowntimeHours: perf ? round1(perf.sitesDownHours) : null,
        reliability: perf ? round2(parseReliability(perf.reliability)) : null,
        topRca: countByName(site.rcas)[0]?.name ?? "N/A",
      };
    })
    .filter((site) => site.tickets > 1 || site.downtimeHours > 0)
    .sort(
      (a, b) =>
        (b.performanceDowntimeHours ?? b.downtimeHours) -
          (a.performanceDowntimeHours ?? a.downtimeHours) ||
        b.tickets - a.tickets ||
        a.siteId.localeCompare(b.siteId),
    )
    .slice(0, 10);

  const topRcaFamilyByTickets = [...rcaFamilyDeepDive].sort(
    (a, b) =>
      b.tickets - a.tickets || b.serviceImpactTickets - a.serviceImpactTickets,
  )[0];
  const missingRcaLeader = missingRcaByRegion[0];
  const longPendingCount =
    slaSummary.pendingAgingBuckets.find((bucket) => bucket.name === "7d+")
      ?.value ?? 0;
  const topRepeatedSite = repeatedOffenderSites[0];

  // Recommended Management Actions:
  // 1) address the most frequent RCA family,
  // 2) close the largest RCA completion gap,
  // 3) escalate long-aging pending tickets,
  // 4) create an action plan for the top repeated/high-downtime RF Site.
  const recommendations = [
    topRcaFamilyByTickets
      ? `Prioritize ${topRcaFamilyByTickets.family}; it appears in ${topRcaFamilyByTickets.tickets.toLocaleString()} tickets, including ${topRcaFamilyByTickets.serviceImpactTickets.toLocaleString()} service-impact tickets. Recommended action: ${topRcaFamilyByTickets.recommendedAction}`
      : "No dominant RCA family detected in the selected scope.",
    missingRcaLeader
      ? `Close RCA gaps in ${missingRcaLeader.name}; it has ${missingRcaLeader.value.toLocaleString()} missing RCA records (${missingRcaLeader.percentage}% of selected tickets).`
      : "RCA completion is acceptable in the selected scope.",
    longPendingCount
      ? `Escalate ${longPendingCount.toLocaleString()} pending tickets older than 7 days and assign accountable owners.`
      : "No critical long-aging pending bucket detected.",
    topRepeatedSite
      ? `Create a recovery action plan for ${topRepeatedSite.siteId}${topRepeatedSite.siteName ? ` - ${topRepeatedSite.siteName}` : ""}; it has ${topRepeatedSite.tickets.toLocaleString()} repeated tickets, ${topRepeatedSite.downtimeHours.toLocaleString()} downtime hours, and top RCA family ${topRepeatedSite.topRca}.`
      : "No repeated offender site detected in the selected scope.",
  ];

  return {
    rcaFamilyDeepDive,
    preventabilityByCount,
    preventabilityByDowntime,
    missingRcaByRegion,
    responsibleTeamLoad,
    slaSummary,
    repeatedOffenderSites,
    recommendations,
  };
}
