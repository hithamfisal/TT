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
} from "../types/dashboard";

import { average, clean, combineDateTime, normalizeSiteId, parseDurationHours } from "./dateUtils";
import { getPreventability, getRcaFamily, getRecommendedAction, getResponsibleTeam, rcaNotProvided } from "./rcaRules";

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

function parseReliability(value: string | number | null | undefined): number {
  if (typeof value === "number" && Number.isFinite(value)) {
    return value > 1 ? value : value * 100;
  }
  const parsed = Number(String(value ?? "").replace("%", "").trim());
  return Number.isFinite(parsed) ? parsed : 100;
}

function countByName(values: string[]): Array<{ name: string; value: number }> {
  const map = new Map<string, number>();
  values.forEach(value => {
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

function buildHealthReason(reductions: Array<{ label: string; points: number }>): string {
  const main = reductions
    .filter(item => item.points > 0)
    .sort((a, b) => b.points - a.points)[0];

  return main
    ? `${main.label} is the largest score reduction.`
    : "No major risk driver detected in the selected scope.";
}

export function calculateExecutiveInsights({
  tickets,
  performanceRows,
}: ExecutiveInsightsInput): ExecutiveInsights {
  const primaryRows = tickets.map(ticket => ticket.primary);
  const totalTickets = tickets.length;
  const totalDowntime = round1(
    primaryRows.reduce((sum, row) => sum + (parseDurationHours(row.duration) ?? 0), 0)
  );

  const affectedSiteIds = new Set<string>();
  const affectedSiteNames = new Map<string, string>();

  tickets.forEach(ticket => {
    const sites = ticket.siteIds.size
      ? Array.from(ticket.siteIds)
      : [ticket.primary.siteId || "Blank"];

    sites.forEach(site => {
      const siteId = normalizeSiteId(clean(site));
      if (!siteId || siteId === "Blank") return;
      affectedSiteIds.add(siteId);
      if (ticket.primary.siteName && !affectedSiteNames.has(siteId)) {
        affectedSiteNames.set(siteId, ticket.primary.siteName);
      }
    });
  });

  const criticalTickets = primaryRows.filter(
    row => clean(row.severity).toLowerCase() === "critical"
  ).length;
  const pendingTickets = primaryRows.filter(row => {
    const status = clean(row.status).toLowerCase();
    return status === "pending" || status === "open" || status.includes("pending");
  }).length;
  const missingRca = primaryRows.filter(row => rcaNotProvided(row.rca)).length;
  const serviceImpactCount = primaryRows.filter(row =>
    clean(row.impact).toLowerCase().includes("service impact")
  ).length;
  const preventableCount = primaryRows.filter(row =>
    (row.preventability || getPreventability(row.rca)) === "Preventable"
  ).length;

  const reliabilityValues = performanceRows
    .map(row => parseReliability(row.reliability))
    .filter(value => Number.isFinite(value));
  const avgReliability = reliabilityValues.length
    ? reliabilityValues.reduce((sum, value) => sum + value, 0) / reliabilityValues.length
    : 100;

  const affectedSites = affectedSiteIds.size;
  const rcaFamilies = countByName(
    primaryRows.map(row => row.rcaFamily || getRcaFamily(row.rca))
  );
  const topRcaFamily = rcaFamilies[0] ?? { name: "N/A", value: 0 };

  const downtimeBySite = new Map<string, { siteId: string; siteName: string; downtime: number }>();
  const siteRisk = new Map<
    string,
    {
      siteId: string;
      siteName: string;
      ticketCount: number;
      downtimeHours: number;
      serviceImpactCount: number;
      criticalCount: number;
      missingRcaCount: number;
      rcaValues: string[];
    }
  >();

  tickets.forEach(ticket => {
    const sites = ticket.siteIds.size
      ? Array.from(ticket.siteIds)
      : [ticket.primary.siteId || "Blank"];
    const duration = parseDurationHours(ticket.primary.duration) ?? 0;
    const perSiteDowntime = sites.length ? duration / sites.length : duration;
    const isServiceImpact = clean(ticket.primary.impact)
      .toLowerCase()
      .includes("service impact");
    const isCritical = clean(ticket.primary.severity).toLowerCase() === "critical";
    const isMissingRca = rcaNotProvided(ticket.primary.rca);
    const rcaFamily = ticket.primary.rcaFamily || getRcaFamily(ticket.primary.rca);

    sites.forEach(site => {
      const siteId = normalizeSiteId(clean(site || ticket.primary.siteId || "Blank"));
      if (!siteId || siteId === "Blank") return;
      const siteName = ticket.primary.siteName || affectedSiteNames.get(siteId) || "";

      const downtimeCurrent = downtimeBySite.get(siteId) ?? {
        siteId,
        siteName,
        downtime: 0,
      };
      downtimeCurrent.downtime += perSiteDowntime;
      downtimeBySite.set(siteId, downtimeCurrent);

      const current = siteRisk.get(siteId) ?? {
        siteId,
        siteName,
        ticketCount: 0,
        downtimeHours: 0,
        serviceImpactCount: 0,
        criticalCount: 0,
        missingRcaCount: 0,
        rcaValues: [],
      };

      current.ticketCount += 1;
      current.downtimeHours += perSiteDowntime;
      if (isServiceImpact) current.serviceImpactCount += 1;
      if (isCritical) current.criticalCount += 1;
      if (isMissingRca) current.missingRcaCount += 1;
      current.rcaValues.push(rcaFamily);
      siteRisk.set(siteId, current);
    });
  });

  const perfBySite = new Map(
    performanceRows.map(row => [normalizeSiteId(clean(row.siteId)), row])
  );

  const highestDowntime = Array.from(downtimeBySite.values()).sort(
    (a, b) => b.downtime - a.downtime || a.siteId.localeCompare(b.siteId)
  )[0];

  const worstReliability = performanceRows
    .filter(row => clean(row.siteId))
    .map(row => ({
      siteId: clean(row.siteId),
      siteName: clean(row.siteName),
      reliability: parseReliability(row.reliability),
    }))
    .sort((a, b) => a.reliability - b.reliability || a.siteId.localeCompare(b.siteId))[0];

  const maxTickets = Math.max(1, ...Array.from(siteRisk.values()).map(row => row.ticketCount));
  const maxDowntime = Math.max(1, ...Array.from(siteRisk.values()).map(row => row.downtimeHours));
  const maxServiceImpact = Math.max(
    1,
    ...Array.from(siteRisk.values()).map(row => row.serviceImpactCount)
  );
  const maxCritical = Math.max(1, ...Array.from(siteRisk.values()).map(row => row.criticalCount));
  const maxMissingRca = Math.max(
    1,
    ...Array.from(siteRisk.values()).map(row => row.missingRcaCount)
  );

  const highRiskSites: HighRiskSiteRow[] = Array.from(siteRisk.values())
    .map(site => {
      const perf = perfBySite.get(site.siteId);
      const reliability = perf ? parseReliability(perf.reliability) : 100;
      const topRca = countByName(site.rcaValues)[0]?.name ?? "N/A";
      const riskScore = round1(
        clamp(
          (site.ticketCount / maxTickets) * 20 +
            (site.downtimeHours / maxDowntime) * 30 +
            (site.serviceImpactCount / maxServiceImpact) * 20 +
            (site.criticalCount / maxCritical) * 15 +
            (site.missingRcaCount / maxMissingRca) * 10 +
            ((100 - reliability) / 100) * 5
        )
      );

      return {
        rank: 0,
        siteId: site.siteId,
        siteName: site.siteName || perf?.siteName || "",
        ticketCount: site.ticketCount,
        downtimeHours: round1(site.downtimeHours),
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
        a.siteId.localeCompare(b.siteId)
    )
    .map((row, index) => ({ ...row, rank: index + 1 }))
    .slice(0, 10);

  const downtimePenalty = clamp(totalDowntime / 20, 0, 20);
  const affectedSitesPenalty = clamp(affectedSites * 1.5, 0, 15);
  const criticalPenalty = clamp(criticalTickets * 4, 0, 20);
  const pendingPenalty = clamp(pendingTickets * 3, 0, 15);
  const missingRcaPenalty = clamp(pctNumber(missingRca, totalTickets) * 0.12, 0, 12);
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
        reliabilityPenalty
    )
  );

  const healthScore: NetworkHealthScore = {
    score,
    status: healthStatus(score),
    mainReason: buildHealthReason(scoreReductions),
    reductions: scoreReductions.map(item => ({ ...item, points: round1(item.points) })),
  };

  const followUpSites = highRiskSites.filter(site => site.riskLevel !== "Low").length;

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
  const primaryRows = tickets.map(ticket => ticket.primary);
  const totalTickets = Math.max(1, primaryRows.length);
  const perfBySite = new Map(
    performanceRows.map(row => [normalizeSiteId(clean(row.siteId)), row])
  );

  const familyMap = new Map<string, RcaDeepDiveRow>();
  const missingRcaByRegionMap = new Map<string, number>();
  const teamLoadMap = new Map<string, number>();
  const preventabilityCountMap = new Map<string, { count: number; downtime: number }>();
  const siteMap = new Map<
    string,
    { siteId: string; siteName: string; tickets: number; downtimeHours: number; rcas: string[] }
  >();

  primaryRows.forEach(row => {
    const rcaFamily = row.rcaFamily || getRcaFamily(row.rca);
    const downtime = parseDurationHours(row.duration) ?? 0;
    const missingRca = rcaNotProvided(row.rca);
    const preventability = row.preventability || getPreventability(row.rca);
    const responsibleTeam = row.responsibleTeam || getResponsibleTeam(rcaFamily);
    const recommendedAction = row.recommendedAction || getRecommendedAction(rcaFamily);
    const isServiceImpact = clean(row.impact).toLowerCase().includes("service impact");

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

    family.tickets += 1;
    family.downtimeHours += downtime;
    if (missingRca) family.missingRca += 1;
    if (preventability === "Preventable") family.preventableTickets += 1;
    if (isServiceImpact) family.serviceImpactTickets += 1;
    familyMap.set(rcaFamily, family);

    if (missingRca) {
      const region = clean(row.region) || "Unknown";
      missingRcaByRegionMap.set(region, (missingRcaByRegionMap.get(region) ?? 0) + 1);
    }

    teamLoadMap.set(responsibleTeam, (teamLoadMap.get(responsibleTeam) ?? 0) + 1);

    const preventabilityItem = preventabilityCountMap.get(preventability) ?? {
      count: 0,
      downtime: 0,
    };
    preventabilityItem.count += 1;
    preventabilityItem.downtime += downtime;
    preventabilityCountMap.set(preventability, preventabilityItem);

    const siteId = normalizeSiteId(clean(row.siteId));
    if (siteId) {
      const current = siteMap.get(siteId) ?? {
        siteId,
        siteName: clean(row.siteName) || perfBySite.get(siteId)?.siteName || "",
        tickets: 0,
        downtimeHours: 0,
        rcas: [],
      };
      current.tickets += 1;
      current.downtimeHours += downtime;
      current.rcas.push(rcaFamily);
      siteMap.set(siteId, current);
    }
  });

  const rcaFamilyDeepDive = Array.from(familyMap.values())
    .map(row => ({
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
    0
  );
  const preventabilityByDowntime = Array.from(preventabilityCountMap.entries())
    .map(([name, value]) => ({
      name,
      value: round1(value.downtime),
      downtimeHours: round1(value.downtime),
      percentage: totalDowntime ? round1((value.downtime / totalDowntime) * 100) : 0,
    }))
    .sort((a, b) => b.value - a.value);

  const missingRcaByRegion = Array.from(missingRcaByRegionMap.entries())
    .map(([name, value]) => ({ name, value, percentage: round1((value / totalTickets) * 100) }))
    .sort((a, b) => b.value - a.value)
    .slice(0, 8);

  const responsibleTeamLoad = Array.from(teamLoadMap.entries())
    .map(([name, value]) => ({ name, value, percentage: round1((value / totalTickets) * 100) }))
    .sort((a, b) => b.value - a.value)
    .slice(0, 8);

  const pendingAgeCounts = new Map<SlaBucket["name"], number>([
    ["0–24h", 0],
    ["1–3d", 0],
    ["3–7d", 0],
    ["7d+", 0],
  ]);
  primaryRows
    .filter(row => clean(row.status).toLowerCase().includes("pending") || clean(row.status).toLowerCase() === "open")
    .forEach(row => {
      const opened = combineDateTime(row.observationDate, row.observationTime);
      if (!opened) return;
      const ageHours = (now.getTime() - opened.getTime()) / 36e5;
      if (!Number.isFinite(ageHours) || ageHours < 0) return;
      const bucket = bucketPendingAge(ageHours);
      pendingAgeCounts.set(bucket, (pendingAgeCounts.get(bucket) ?? 0) + 1);
    });

  const slaSummary = {
    avgFrtHours: round1(average(primaryRows.map(row => row.frtHours))),
    avgResponseHours: round1(average(primaryRows.map(row => row.responseHours))),
    avgResolutionHours: round1(average(primaryRows.map(row => row.resolutionHours))),
    frtBreaches: primaryRows.filter(row => typeof row.frtHours === "number" && row.frtHours > SLA_TARGETS.frtHours).length,
    responseBreaches: primaryRows.filter(row => typeof row.responseHours === "number" && row.responseHours > SLA_TARGETS.responseHours).length,
    resolutionBreaches: primaryRows.filter(row => typeof row.resolutionHours === "number" && row.resolutionHours > SLA_TARGETS.resolutionHours).length,
    pendingAgingBuckets: Array.from(pendingAgeCounts.entries()).map(([name, value]) => ({
      name,
      value,
      label: `${value} tickets`,
    })),
  };

  const repeatedOffenderSites = Array.from(siteMap.values())
    .filter(site => site.tickets > 1 || site.downtimeHours > 0)
    .map(site => ({
      siteId: site.siteId,
      siteName: site.siteName,
      tickets: site.tickets,
      downtimeHours: round1(site.downtimeHours),
      topRca: countByName(site.rcas)[0]?.name ?? "N/A",
    }))
    .sort((a, b) => b.downtimeHours - a.downtimeHours || b.tickets - a.tickets)
    .slice(0, 10);

  const recommendations = [
    rcaFamilyDeepDive[0]
      ? `Prioritize ${rcaFamilyDeepDive[0].family}; it contributes ${rcaFamilyDeepDive[0].downtimeHours.toLocaleString()} downtime hours.`
      : "No dominant RCA family detected in the selected scope.",
    missingRcaByRegion[0]
      ? `Close RCA gaps in ${missingRcaByRegion[0].name}; it has ${missingRcaByRegion[0].value} missing RCA records.`
      : "RCA completion is acceptable in the selected scope.",
    slaSummary.pendingAgingBuckets.find(bucket => bucket.name === "7d+")?.value
      ? "Escalate pending tickets older than 7 days and assign accountable owners."
      : "No critical long-aging pending bucket detected.",
    repeatedOffenderSites[0]
      ? `Create a recovery action plan for ${repeatedOffenderSites[0].siteId}; it is the top repeated/high-downtime site.`
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
