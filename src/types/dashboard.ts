// src/types/dashboard.ts

export type TicketRecord = {
  rowNo: number;
  tt: string;
  siteId: string;
  siteName: string;
  managedResource: string;
  issue: string;
  severity: string;
  region: string;
  observationDate: string;
  observationTime: string;
  openingMonthKey: string;
  openingMonthLabel: string;
  recoveryDate: string;
  recoveryTime: string;
  duration: string;
  impact: string;
  escalatedTo: string;
  escalationLevel: string;
  escalatedForL3SupportDate: string;
  escalatedForL3SupportTime: string;
  frtHours: number | null;
  responseHours: number | null;
  resolutionHours: number | null;
  status: string;
  rca: string;
  rcaFamily: string;
  preventability: string;
  responsibleTeam: string;
  recommendedAction: string;
  actionTaken: string;
  sourceFile: string;
};

export type TicketAggregate = {
  tt: string;
  primary: TicketRecord;
  siteIds: Set<string>;
  siteNames: Set<string>;
  rows: TicketRecord[];
};

export type DashboardData = {
  fileName: string;
  sheetName: string;
  generatedAt: string;
  rows: TicketRecord[];
  uniqueTickets: TicketAggregate[];
  siteOrder: { siteId: string; siteName: string }[];
  rcaLookup: { action: string; rca: string }[];
};

export type Filters = {
  search: string;
  status: string[];
  severity: string[];
  region: string[];
  impact: string[];
  site: string[];
  openingMonth: string[];
  rcaFamily: string[];
};

export type PerfRow = {
  siteId: string;
  siteName: string;
  displayName: string;
  sourceLabel: string;
  perfKey: string;
  sitesDownHours: number;
  availHours: number;
  availDay: string;
  reliability: string;
  channelBusy: number;
  mwLinkPerf: number;
  ticketCount: number;
};

export type NetworkHealthScore = {
  score: number;
  status: "Excellent" | "Good" | "Warning" | "Critical";
  mainReason: string;
  reductions: Array<{ label: string; points: number }>;
};

export type ExecutiveInsightCard = {
  label: string;
  value: string;
  note: string;
  tone: string;
};

export type HighRiskSiteRow = {
  rank: number;
  region: string;
  siteId: string;
  siteName: string;
  ticketCount: number;
  downtimeHours: number;
  reliability: number;
  topRca: string;
  riskScore: number;
  riskLevel: "Low" | "Medium" | "High" | "Critical";
};

export type ExecutiveInsights = {
  healthScore: NetworkHealthScore;
  summaryText: string;
  cards: ExecutiveInsightCard[];
  highRiskSites: HighRiskSiteRow[];
  totals: {
    totalTickets: number;
    serviceImpactCount: number;
    affectedSites: number;
    totalDowntime: number;
    criticalTickets: number;
    pendingTickets: number;
    missingRca: number;
    rcaCompletionRate: number;
    serviceImpactRate: number;
    preventableRate: number;
    avgReliability: number;
  };
};

export type AnalyticsBucket = {
  name: string;
  value: number;
  downtimeHours?: number;
  percentage?: number;
};

export type RcaDeepDiveRow = {
  family: string;
  tickets: number;
  downtimeHours: number;
  missingRca: number;
  preventableTickets: number;
  serviceImpactTickets: number;
  responsibleTeam: string;
  recommendedAction: string;
};

export type SlaBucket = {
  name: string;
  value: number;
  label: string;
};

export type SlaSummary = {
  avgFrtHours: number;
  avgResponseHours: number;
  avgResolutionHours: number;
  frtBreaches: number;
  responseBreaches: number;
  resolutionBreaches: number;
  pendingAgingBuckets: SlaBucket[];
};

export type DeepDiveAnalytics = {
  rcaFamilyDeepDive: RcaDeepDiveRow[];
  preventabilityByCount: AnalyticsBucket[];
  preventabilityByDowntime: AnalyticsBucket[];
  missingRcaByRegion: AnalyticsBucket[];
  responsibleTeamLoad: AnalyticsBucket[];
  slaSummary: SlaSummary;
  repeatedOffenderSites: Array<{
    region: string;
    siteId: string;
    siteName: string;
    tickets: number;
    downtimeHours: number;
    performanceAvailabilityHours: number | null;
    performanceDowntimeHours: number | null;
    reliability: number | null;
    topRca: string;
  }>;
  recommendations: string[];
};
