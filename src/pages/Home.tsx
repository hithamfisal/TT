import PptxGenJS from "pptxgenjs";
import {
  type CSSProperties,
  type ReactNode,
  useEffect,
  useMemo,
  useRef,
  useState,
} from "react";
import { createPortal } from "react-dom";
import * as XLSX from "xlsx";
import jsPDF from "jspdf";
import nascoLogoSrc from "../assets/nascologo.png";
import { useTheme } from "../contexts/ThemeContext";

const THEME_IMAGES = {
  dark: "/dark-bg.png",
  light: "/light-bg.png",
} as const;

type DashboardTheme = keyof typeof THEME_IMAGES;

type DashboardSectionId =
  | "reports"
  | "input"
  | "performanceKpis"
  | "kpis"
  | "executive"
  | "deepDive"
  | "overviewCharts"
  | "trendCharts"
  | "ticketsTable";

type DashboardSectionDefinition = {
  id: DashboardSectionId;
  label: string;
  title: string;
  selector: string;
};

const DASHBOARD_SECTIONS: DashboardSectionDefinition[] = [
  {
    id: "input",
    label: "Input",
    title: "Manual Ticket Input",
    selector: "#section-manual-input",
  },
  {
    id: "performanceKpis",
    label: "KPI",
    title: "Performance KPIs",
    selector: "#section-kpis",
  },
  {
    id: "ticketsTable",
    label: "Tickets",
    title: "Tickets Data Table",
    selector: "#section-tickets-table",
  },
  {
    id: "overviewCharts",
    label: "Overview",
    title: "Operational Overview Charts",
    selector: "#section-overview-charts",
  },
  {
    id: "executive",
    label: "Executive",
    title: "Executive Insights",
    selector: "#section-executive",
  },

  {
    id: "trendCharts",
    label: "Trend Charts",
    title: "Trend & RCA Charts",
    selector: "#section-trend-charts",
  },
  {
    id: "reports",
    label: "Reports",
    title: "Report Management Center",
    selector: "#section-reports",
  },
];

const INITIAL_COLLAPSED_SECTIONS: Record<DashboardSectionId, boolean> = {
  reports: true,
  input: true,
  performanceKpis: true,
  kpis: true,
  executive: true,
  deepDive: true,
  overviewCharts: true,
  trendCharts: true,
  ticketsTable: true,
};

const SAVED_DASHBOARD_KEY = "followup-dashboard:last-workbook:v1";
const SAVED_MANUAL_TICKETS_KEY = "followup-dashboard:manual-tickets:v1";
const LOGIN_SESSION_KEY = "followup-dashboard:login-session:v1";
const DASHBOARD_LOGIN_USERNAME = "admin";
const DASHBOARD_LOGIN_PASSWORD = "DMR@2026";
const ALLOWED_UPLOAD_EXTENSIONS = [".xlsx", ".xls", ".csv"] as const;
const MAX_UPLOAD_FILE_SIZE_BYTES = 50 * 1024 * 1024;
const REQUIRED_UPLOAD_COLUMNS = ["TT", "Site ID", "Status", "Opening Month / Observation Date"] as const;
const SAFE_UPLOAD_ERROR_MESSAGE =
  "The uploaded file could not be processed. Please check the file format and required columns.";

const GOOGLE_REGION_LINKS = [
  { key: "eoaNeoa", label: "EOA&NEOA" },
  { key: "soa", label: "SOA" },
  { key: "coaWoa", label: "COA&WOA" },
] as const;

type GoogleRegionKey = (typeof GOOGLE_REGION_LINKS)[number]["key"];
type GoogleRegionLinks = Record<GoogleRegionKey, string>;
type GoogleRegionSaveLinks = Record<GoogleRegionKey, string>;
type GoogleRegionSelection = Record<GoogleRegionKey, boolean>;
type UploadedWorkbookSource = {
  fileName: string;
  workbook: XLSX.WorkBook;
  data: DashboardData;
};
type MicrosoftGraphConfig = {
  clientId: string;
  tenantId: string;
  scopes: string;
};


type ManagedReportType =
  | "tickets"
  | "performance"
  | "executive"
  | "quality"
  | "kpiCards"
  | "performanceCards";
type ManagedReportFormat = "xlsx" | "pdf" | "ppt" | "png";

type GeneratedReportItem = {
  id: string;
  fileName: string;
  reportType: string;
  generatedAt: string;
  format: ManagedReportFormat;
  records: number;
};

function getInitialLoginState() {
  if (typeof window === "undefined") return false;
  return window.sessionStorage.getItem(LOGIN_SESSION_KEY) === "authenticated";
}

const EMPTY_GOOGLE_REGION_LINKS: GoogleRegionLinks = {
  eoaNeoa: "",
  soa: "",
  coaWoa: "",
};

const EMPTY_GOOGLE_REGION_SAVE_LINKS: GoogleRegionSaveLinks = {
  eoaNeoa: "",
  soa: "",
  coaWoa: "",
};

const EMPTY_GOOGLE_REGION_SELECTION: GoogleRegionSelection = {
  eoaNeoa: true,
  soa: true,
  coaWoa: true,
};

const EMPTY_MICROSOFT_GRAPH_CONFIG: MicrosoftGraphConfig = {
  clientId: "",
  tenantId: "common",
  scopes: "User.Read Files.Read.All Sites.Read.All",
};

import {
  Activity,
  AlertTriangle,
  ArrowUp,
  BarChart3,
  CheckCircle2,
  ChevronDown,
  ChevronLeft,
  ChevronRight,
  CircleDot,
  Download,
  FileSpreadsheet,
  Filter,
  CloudOff,
  Home as HomeIcon,
  Layers3,
  ImageDown,
  Link as LinkIcon,
  LogOut,
  Maximize2,
  Minimize2,
  Moon,
  type LucideIcon,
  Network,
  Presentation,
  Printer,
  RefreshCw,
  Save,
  Search,
  ShieldAlert,
  Sun,
  UploadCloud,
} from "lucide-react";

import {
  Bar,
  BarChart,
  CartesianGrid,
  Cell,
  ComposedChart,
  LabelList,
  Legend,
  Line,
  LineChart,
  Pie,
  PieChart,
  ReferenceLine,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from "recharts";

import autoTable from "jspdf-autotable";

import type {
  DashboardData,
  DeepDiveAnalytics,
  ExecutiveInsights,
  Filters,
  PerfRow,
  TicketAggregate,
  TicketRecord,
} from "../types/dashboard";

import {
  average,
  clean,
  combineDateTime,
  coveredMonthKeys,
  dateKey,
  dateWithinMonth,
  formatDateDDMMYYYY,
  formatHours,
  formatMonthMMMMYYYY,
  getField,
  getRawField,
  hoursBetween,
  isPendingStatus,
  isRfSiteId,
  normalizeHeader,
  normalizeMonthKey,
  normalizeSiteId,
  openingMonthKey,
  openingMonthLabel,
  parseDateValue,
  parseDurationHours,
  recordDateMonthKey,
  resolveOpeningMonthKey,
  selectedMonthRange,
  ticketMatchesMonthlyExport,
  totalHoursInMonth,
  weekKey,
  weekLabel,
} from "../lib/dateUtils";

import {
  getPreventability,
  getRcaFamily,
  getRecommendedAction,
  getResponsibleTeam,
  rcaNotProvided,
  RCA_FAMILY_MAP,
} from "../lib/rcaRules";

import { groupTickets, parseRows } from "../lib/parseWorkbook";
import {
  calculateDeepDiveAnalytics,
  calculateExecutiveInsights,
} from "../lib/ticketAnalytics";

type ManualTicketDraft = {
  id: string;
  tt: string;
  siteId: string;
  siteName: string;
  customSite: boolean;
  managedResource: string;
  issue: string;
  severity: string;
  region: string;
  observationDate: string;
  observationTime: string;
  recoveryDate: string;
  recoveryTime: string;
  escalatedForL3SupportDate: string;
  escalatedForL3SupportTime: string;
  duration: string;
  durationHours: string;
  neDetail: string;
  impact: string;
  correlatedAlarms: string;
  escalatedTo: string;
  escalationLevel: string;
  status: string;
  commentsDate: string;
  commentsFeedback: string;
  maintenanceTeam: string;
  maintenanceContact: string;
  actionTaken: string;
  action: string;
  customRca: boolean;
  rca: string;
};

const SEVERITY_OPTIONS = ["Critical", "Major", "Minor"];
const REGION_OPTIONS = ["EOA", "NEOA", "SOA", "COA", "WOA"];
const STATUS_OPTIONS = ["Closed", "Resolved", "Pending"];
const IMPACT_OPTIONS = ["Service Impact", "Non Service Impact"];
const ESCALATION_LEVEL_OPTIONS = ["L1", "L2", "L3", "L4", "L5"];

const GOOGLE_TT_HISTORY_HEADERS = [
  "#",
  "TT",
  "Site ID",
  "Site Name",
  "Managed Resource",
  "Issues",
  "Severity",
  "Region",
  "Observation Date",
  "Observation Time",
  "Recovery Date",
  "Recovery Time",
  "Escalated for L3 Support Date",
  "Escalated for L3 Support Time",
  "Total Duration Days/Hours",
  "Duration (hrs)",
  "NE Detail/Impacted Object",
  "Service Impaction Status",
  "No. of correlated Alarms",
  "Escalated to",
  "Escalation Level",
  "Status",
  "Comments Date",
  "Comments-Feedback",
  "Maintenance person/Team",
  "Maintenance Contact Details",
  "Action Taken/RCA",
  "Action",
  "RCA",
];

const MANUAL_TICKET_EXPORT_HEADERS = [
  "SN.",
  "TT's",
  "Site ID",
  "Site Name",
  "Managed Resource ",
  "Issues",
  "Severity",
  "Region",
  "Observation Date",
  "Observation Time",
  "Recovery Date",
  "Recovery Time",
  "Escalated for L3 Support Date",
  "Escalated for L3 Support Time",
  "Total Duration/Days/Hours",
  "Duration (hrs)",
  "NE Detail/Impacted Object",
  "Service Impaction Status",
  "No. of correlated Alarms",
  "Escalated to ",
  "Escalation Level",
  "Status",
  "Comments Date",
  "Comments-Feedback",
  "Maintenance person/Team",
  "Maintenance Contact Details",
  "Action Taken/RCA",
  "Action",
  "RCA",
];

function createManualTicketDraft(): ManualTicketDraft {
  return {
    id: `${Date.now()}-${Math.random().toString(36).slice(2)}`,
    tt: "",
    siteId: "",
    siteName: "",
    customSite: false,
    managedResource: "",
    issue: "",
    severity: "Critical",
    region: "",
    observationDate: "",
    observationTime: "",
    recoveryDate: "",
    recoveryTime: "",
    escalatedForL3SupportDate: "",
    escalatedForL3SupportTime: "",
    duration: "",
    durationHours: "",
    neDetail: "",
    impact: "Service Impact",
    correlatedAlarms: "",
    escalatedTo: "",
    escalationLevel: "L1",
    status: "Pending",
    commentsDate: "",
    commentsFeedback: "",
    maintenanceTeam: "",
    maintenanceContact: "",
    actionTaken: "",
    action: "",
    customRca: false,
    rca: "",
  };
}

function normalizeManualDate(value: string): string {
  const text = clean(value);
  if (!text) return "";
  const iso = text.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (iso) return `${iso[3]}/${iso[2]}/${iso[1]}`;
  return text;
}

function normalizeManualTime(value: string): string {
  const digits = clean(value).replace(/\D/g, "").slice(0, 4);
  if (digits.length <= 2) return digits;
  return `${digits.slice(0, 2)}:${digits.slice(2)}`;
}

function isCompleteManualTime(value: string): boolean {
  const match = clean(value).match(/^(\d{2}):(\d{2})$/);
  if (!match) return false;
  const hour = Number(match[1]);
  const minute = Number(match[2]);
  return hour >= 0 && hour <= 23 && minute >= 0 && minute <= 59;
}

function manualDurationHours(draft: ManualTicketDraft): number | null {
  const observationDate = normalizeManualDate(draft.observationDate);
  const recoveryDate = normalizeManualDate(draft.recoveryDate);
  if (
    !observationDate ||
    !recoveryDate ||
    !isCompleteManualTime(draft.observationTime) ||
    !isCompleteManualTime(draft.recoveryTime)
  ) {
    return null;
  }
  const observedAt = combineDateTime(observationDate, draft.observationTime);
  const recoveredAt = combineDateTime(recoveryDate, draft.recoveryTime);
  const hours = hoursBetween(observedAt, recoveredAt);
  return hours !== null && Number.isFinite(hours) && hours >= 0 ? hours : null;
}

function formatManualDurationHours(draft: ManualTicketDraft): string {
  const hours = manualDurationHours(draft);
  return hours === null ? "" : String(Math.round(hours * 10) / 10);
}

function formatManualTotalDuration(draft: ManualTicketDraft): string {
  const hours = manualDurationHours(draft);
  if (hours === null) return "";
  const totalMinutes = Math.max(0, Math.round(hours * 60));
  const days = Math.floor(totalMinutes / 1440);
  const remainingAfterDays = totalMinutes % 1440;
  const hrs = Math.floor(remainingAfterDays / 60);
  const mins = remainingAfterDays % 60;
  return `${days} days ${hrs} hrs ${mins} mins`;
}
function manualDraftRegionKey(
  draft: ManualTicketDraft,
): GoogleRegionKey | null {
  const region = normalizeHeader(draft.region);
  if (region === "eoa" || region === "neoa") return "eoaNeoa";
  if (region === "soa") return "soa";
  if (region === "coa" || region === "woa") return "coaWoa";
  return null;
}

function manualDraftToSheetRow(draft: ManualTicketDraft, index: number) {
  const durationHours = formatManualDurationHours(draft);
  return {
    "SN.": index + 1,
    "TT's": clean(draft.tt),
    "Site ID": normalizeSiteId(clean(draft.siteId)),
    "Site Name": clean(draft.siteName),
    "Managed Resource ": clean(draft.managedResource),
    Issues: clean(draft.issue),
    Severity: clean(draft.severity),
    Region: clean(draft.region),
    "Observation Date": normalizeManualDate(draft.observationDate),
    "Observation Time": clean(draft.observationTime),
    "Recovery Date": normalizeManualDate(draft.recoveryDate),
    "Recovery Time": clean(draft.recoveryTime),
    "Escalated for L3 Support Date": normalizeManualDate(
      draft.escalatedForL3SupportDate,
    ),
    "Escalated for L3 Support Time": clean(draft.escalatedForL3SupportTime),
    "Total Duration/Days/Hours": formatManualTotalDuration(draft),
    "Duration (hrs)": durationHours,
    "NE Detail/Impacted Object": clean(draft.neDetail),
    "Service Impaction Status": clean(draft.impact),
    "No. of correlated Alarms": clean(draft.correlatedAlarms),
    "Escalated to ": clean(draft.escalatedTo),
    "Escalation Level": clean(draft.escalationLevel),
    Status: clean(draft.status),
    "Comments Date": normalizeManualDate(draft.commentsDate),
    "Comments-Feedback": clean(draft.commentsFeedback),
    "Maintenance person/Team": clean(draft.maintenanceTeam),
    "Maintenance Contact Details": clean(draft.maintenanceContact),
    "Action Taken/RCA": clean(draft.actionTaken),
    Action: clean(draft.action),
    RCA: clean(draft.rca),
  };
}
function manualDraftToTicketRecord(
  draft: ManualTicketDraft,
  index: number,
): TicketRecord | null {
  const tt = clean(draft.tt);
  const siteId = normalizeSiteId(clean(draft.siteId));
  const siteName = clean(draft.siteName);
  const issue = clean(draft.issue);
  if (!tt && !siteId && !siteName && !issue) return null;

  const observationDate = normalizeManualDate(draft.observationDate);
  const recoveryDate = normalizeManualDate(draft.recoveryDate);
  const l3Date = normalizeManualDate(draft.escalatedForL3SupportDate);
  const duration = formatManualTotalDuration(draft);
  const calculatedDuration = manualDurationHours(draft);
  const observedAt = combineDateTime(observationDate, draft.observationTime);
  const recoveredAt = combineDateTime(recoveryDate, draft.recoveryTime);
  const l3At = combineDateTime(l3Date, draft.escalatedForL3SupportTime);
  const rca = clean(draft.rca);
  const rcaFamily = getRcaFamily(rca);
  const openingKey = resolveOpeningMonthKey("", "", observationDate);

  return {
    rowNo: index + 1,
    tt,
    siteId,
    siteName,
    managedResource: clean(draft.managedResource),
    issue,
    severity: clean(draft.severity),
    region: clean(draft.region),
    observationDate,
    observationTime: clean(draft.observationTime),
    openingMonthKey: openingKey,
    openingMonthLabel: openingMonthLabel(openingKey),
    recoveryDate,
    recoveryTime: clean(draft.recoveryTime),
    duration,
    impact: clean(draft.impact),
    escalatedTo: clean(draft.escalatedTo),
    escalationLevel: clean(draft.escalationLevel),
    escalatedForL3SupportDate: l3Date,
    escalatedForL3SupportTime: clean(draft.escalatedForL3SupportTime),
    frtHours: hoursBetween(observedAt, l3At),
    responseHours: hoursBetween(observedAt, l3At),
    resolutionHours:
      calculatedDuration ??
      parseDurationHours(duration) ??
      hoursBetween(observedAt, recoveredAt),
    status: clean(draft.status),
    rca,
    rcaFamily,
    preventability: getPreventability(rca),
    responsibleTeam: getResponsibleTeam(rcaFamily),
    recommendedAction: getRecommendedAction(rcaFamily),
    actionTaken: clean(draft.action || draft.actionTaken),
    sourceFile: "Manual Input",
  };
}
function ticketDurationHours(row: TicketRecord): number {
  const parsed = parseDurationHours(row.duration);
  if (parsed !== null && Number.isFinite(parsed)) return parsed;
  return typeof row.resolutionHours === "number" &&
    Number.isFinite(row.resolutionHours)
    ? row.resolutionHours
    : 0;
}

function publicWorkbookUrl(path: string) {
  const cleanPath = path.startsWith("/") ? path.slice(1) : path;
  return `${window.location.origin}/${cleanPath}?v=${Date.now()}`;
}

type ZipFileMap = Record<string, Uint8Array>;
type ZipTextCodec = {
  strFromU8: (bytes: Uint8Array) => string;
  strToU8: (text: string) => Uint8Array;
};
type ExcelHorizontalAlign = "left" | "center";

const applyExcelCellStyle = (
  xml: string,
  ref: string,
  styleIndex: number,
): string => {
  const markerIdx = xml.indexOf(` r="${ref}"`);
  if (markerIdx === -1) return xml;
  const cStart = xml.lastIndexOf("<c", markerIdx);
  const tagClose = xml.indexOf(">", cStart);
  if (cStart === -1 || tagClose === -1) return xml;

  const openTag = xml.slice(cStart, tagClose + 1);
  const styledTag = openTag.includes(' s="')
    ? openTag.replace(/\s+s="[^"]*"/, ` s="${styleIndex}"`)
    : openTag.replace(/\/?>$/, (ending) => ` s="${styleIndex}"${ending}`);

  return xml.slice(0, cStart) + styledTag + xml.slice(tagClose + 1);
};

const getExcelCellStyle = (xml: string, ref: string): number => {
  const markerIdx = xml.indexOf(` r="${ref}"`);
  if (markerIdx === -1) return 0;
  const cStart = xml.lastIndexOf("<c", markerIdx);
  const tagClose = xml.indexOf(">", cStart);
  if (cStart === -1 || tagClose === -1) return 0;
  const styleMatch = xml.slice(cStart, tagClose + 1).match(/\s+s="(\d+)"/);
  return styleMatch ? Number(styleMatch[1]) : 0;
};

const ensureExcelAlignedStyle = (
  files: ZipFileMap,
  codec: ZipTextCodec,
  baseStyleIndex: number,
  horizontal: ExcelHorizontalAlign,
) => {
  const stylesKey = "xl/styles.xml";
  if (!files[stylesKey]) return baseStyleIndex;

  let stylesXml = codec.strFromU8(files[stylesKey]);
  const cellXfsStart = stylesXml.indexOf("<cellXfs");
  const cellXfsEnd = stylesXml.indexOf("</cellXfs>", cellXfsStart);
  if (cellXfsStart === -1 || cellXfsEnd === -1) return baseStyleIndex;

  const openEnd = stylesXml.indexOf(">", cellXfsStart);
  const cellXfsInner = stylesXml.slice(openEnd + 1, cellXfsEnd);
  const xfs = cellXfsInner.match(/<xf\b[^>]*(?:\/>|>[\s\S]*?<\/xf>)/g) ?? [];
  const baseXf =
    xfs[baseStyleIndex] ??
    xfs[0] ??
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>';
  const alignmentTag = `<alignment horizontal="${horizontal}" vertical="center" wrapText="1"/>`;
  const withApplyAlignment = baseXf.includes("applyAlignment=")
    ? baseXf.replace(/applyAlignment="[^"]*"/, 'applyAlignment="1"')
    : baseXf.replace(/^<xf\b/, '<xf applyAlignment="1"');
  const alignedXf = withApplyAlignment.endsWith("/>")
    ? withApplyAlignment.replace(/\/>$/, `>${alignmentTag}</xf>`)
    : withApplyAlignment
        .replace(
          /<alignment\b[^>]*\/>|<alignment\b[^>]*>[\s\S]*?<\/alignment>/g,
          "",
        )
        .replace("</xf>", `${alignmentTag}</xf>`);

  const existingIndex = xfs.findIndex((xf) => xf === alignedXf);
  if (existingIndex !== -1) return existingIndex;

  const nextIndex = xfs.length;
  const beforeClose = stylesXml.slice(0, cellXfsEnd);
  const afterClose = stylesXml.slice(cellXfsEnd);
  stylesXml = `${beforeClose}${alignedXf}${afterClose}`;
  stylesXml = stylesXml.replace(
    /(<cellXfs\b[^>]*\bcount=")\d+(")/,
    `$1${nextIndex + 1}$2`,
  );
  files[stylesKey] = codec.strToU8(stylesXml);
  return nextIndex;
};

const ensureExcelStyleVariant = (
  files: ZipFileMap,
  codec: ZipTextCodec,
  baseStyleIndex: number,
  options: {
    horizontal?: ExcelHorizontalAlign;
    fontId?: number;
    fillId?: number;
    borderId?: number;
    numFmtId?: number;
  },
) => {
  const stylesKey = "xl/styles.xml";
  if (!files[stylesKey]) return baseStyleIndex;

  let stylesXml = codec.strFromU8(files[stylesKey]);
  const cellXfsStart = stylesXml.indexOf("<cellXfs");
  const cellXfsEnd = stylesXml.indexOf("</cellXfs>", cellXfsStart);
  if (cellXfsStart === -1 || cellXfsEnd === -1) return baseStyleIndex;

  const openEnd = stylesXml.indexOf(">", cellXfsStart);
  const cellXfsInner = stylesXml.slice(openEnd + 1, cellXfsEnd);
  const xfs = cellXfsInner.match(/<xf\b[^>]*(?:\/>|>[\s\S]*?<\/xf>)/g) ?? [];
  const baseXf =
    xfs[baseStyleIndex] ??
    xfs[0] ??
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>';

  const setAttr = (xf: string, attr: string, value: number | string) =>
    xf.includes(`${attr}=`)
      ? xf.replace(new RegExp(`${attr}="[^"]*"`), `${attr}="${value}"`)
      : xf.replace(/^<xf\b/, `<xf ${attr}="${value}"`);

  let variant = baseXf;
  if (options.fontId !== undefined) {
    variant = setAttr(variant, "fontId", options.fontId);
    variant = variant.includes("applyFont=")
      ? variant.replace(/applyFont="[^"]*"/, 'applyFont="1"')
      : variant.replace(/^<xf\b/, '<xf applyFont="1"');
  }
  if (options.fillId !== undefined) {
    variant = setAttr(variant, "fillId", options.fillId);
    variant = variant.includes("applyFill=")
      ? variant.replace(/applyFill="[^"]*"/, 'applyFill="1"')
      : variant.replace(/^<xf\b/, '<xf applyFill="1"');
  }
  if (options.borderId !== undefined) {
    variant = setAttr(variant, "borderId", options.borderId);
    variant = variant.includes("applyBorder=")
      ? variant.replace(/applyBorder="[^"]*"/, 'applyBorder="1"')
      : variant.replace(/^<xf\b/, '<xf applyBorder="1"');
  }
  if (options.numFmtId !== undefined) {
    variant = setAttr(variant, "numFmtId", options.numFmtId);
    variant = variant.includes("applyNumberFormat=")
      ? variant.replace(/applyNumberFormat="[^"]*"/, 'applyNumberFormat="1"')
      : variant.replace(/^<xf\b/, '<xf applyNumberFormat="1"');
  }
  if (options.horizontal) {
    const alignmentTag = `<alignment horizontal="${options.horizontal}" vertical="center" wrapText="1"/>`;
    variant = variant.includes("applyAlignment=")
      ? variant.replace(/applyAlignment="[^"]*"/, 'applyAlignment="1"')
      : variant.replace(/^<xf\b/, '<xf applyAlignment="1"');
    variant = variant.endsWith("/>")
      ? variant.replace(/\/>$/, `>${alignmentTag}</xf>`)
      : variant
          .replace(
            /<alignment\b[^>]*\/>|<alignment\b[^>]*>[\s\S]*?<\/alignment>/g,
            "",
          )
          .replace("</xf>", `${alignmentTag}</xf>`);
  }

  const existingIndex = xfs.findIndex((xf) => xf === variant);
  if (existingIndex !== -1) return existingIndex;

  const nextIndex = xfs.length;
  stylesXml = `${stylesXml.slice(0, cellXfsEnd)}${variant}${stylesXml.slice(cellXfsEnd)}`;
  stylesXml = stylesXml.replace(
    /(<cellXfs\b[^>]*\bcount=")\d+(")/,
    `$1${nextIndex + 1}$2`,
  );
  files[stylesKey] = codec.strToU8(stylesXml);
  return nextIndex;
};

const ensureExcelSolidFill = (
  files: ZipFileMap,
  codec: ZipTextCodec,
  rgb: string,
) => {
  const stylesKey = "xl/styles.xml";
  if (!files[stylesKey]) return 0;

  let stylesXml = codec.strFromU8(files[stylesKey]);
  const fillsStart = stylesXml.indexOf("<fills");
  const fillsEnd = stylesXml.indexOf("</fills>", fillsStart);
  if (fillsStart === -1 || fillsEnd === -1) return 0;

  const normalized = rgb.replace(/^#/, "").toUpperCase();
  const fillXml = `<fill><patternFill patternType="solid"><fgColor rgb="${normalized}"/><bgColor rgb="${normalized}"/></patternFill></fill>`;
  const openEnd = stylesXml.indexOf(">", fillsStart);
  const fillsInner = stylesXml.slice(openEnd + 1, fillsEnd);
  const fills = fillsInner.match(/<fill>[\s\S]*?<\/fill>/g) ?? [];
  const existingIndex = fills.findIndex((fill) => fill === fillXml);
  if (existingIndex !== -1) return existingIndex;

  const nextIndex = fills.length;
  stylesXml = `${stylesXml.slice(0, fillsEnd)}${fillXml}${stylesXml.slice(fillsEnd)}`;
  stylesXml = stylesXml.replace(
    /(<fills\b[^>]*\bcount=")\d+(")/,
    `$1${nextIndex + 1}$2`,
  );
  files[stylesKey] = codec.strToU8(stylesXml);
  return nextIndex;
};

const ensureExcelWhiteFill = (files: ZipFileMap, codec: ZipTextCodec) =>
  ensureExcelSolidFill(files, codec, "FFFFFFFF");

const ensureExcelFont = (
  files: ZipFileMap,
  codec: ZipTextCodec,
  options: { size: number; name?: string; bold?: boolean },
) => {
  const stylesKey = "xl/styles.xml";
  if (!files[stylesKey]) return 0;

  let stylesXml = codec.strFromU8(files[stylesKey]);
  const fontsStart = stylesXml.indexOf("<fonts");
  const fontsEnd = stylesXml.indexOf("</fonts>", fontsStart);
  if (fontsStart === -1 || fontsEnd === -1) return 0;

  const fontXml =
    `<font>${options.bold ? "<b/>" : ""}<sz val="${options.size}"/><color theme="1"/>` +
    `<name val="${options.name ?? "Calibri"}"/><family val="2"/><scheme val="minor"/></font>`;
  const openEnd = stylesXml.indexOf(">", fontsStart);
  const fontsInner = stylesXml.slice(openEnd + 1, fontsEnd);
  const fonts = fontsInner.match(/<font>[\s\S]*?<\/font>/g) ?? [];
  const existingIndex = fonts.findIndex((font) => font === fontXml);
  if (existingIndex !== -1) return existingIndex;

  const nextIndex = fonts.length;
  stylesXml = `${stylesXml.slice(0, fontsEnd)}${fontXml}${stylesXml.slice(fontsEnd)}`;
  stylesXml = stylesXml.replace(
    /(<fonts\b[^>]*\bcount=")\d+(")/,
    `$1${nextIndex + 1}$2`,
  );
  files[stylesKey] = codec.strToU8(stylesXml);
  return nextIndex;
};

const ensureExcelThinBlackBorder = (files: ZipFileMap, codec: ZipTextCodec) => {
  const stylesKey = "xl/styles.xml";
  if (!files[stylesKey]) return 0;

  let stylesXml = codec.strFromU8(files[stylesKey]);
  const bordersStart = stylesXml.indexOf("<borders");
  const bordersEnd = stylesXml.indexOf("</borders>", bordersStart);
  if (bordersStart === -1 || bordersEnd === -1) return 0;

  const side =
    '<left style="thin"><color rgb="FF000000"/></left><right style="thin"><color rgb="FF000000"/></right><top style="thin"><color rgb="FF000000"/></top><bottom style="thin"><color rgb="FF000000"/></bottom><diagonal/>';
  const borderXml = `<border>${side}</border>`;
  const openEnd = stylesXml.indexOf(">", bordersStart);
  const bordersInner = stylesXml.slice(openEnd + 1, bordersEnd);
  const borders = bordersInner.match(/<border>[\s\S]*?<\/border>/g) ?? [];
  const existingIndex = borders.findIndex((border) => border === borderXml);
  if (existingIndex !== -1) return existingIndex;

  const nextIndex = borders.length;
  stylesXml = `${stylesXml.slice(0, bordersEnd)}${borderXml}${stylesXml.slice(bordersEnd)}`;
  stylesXml = stylesXml.replace(
    /(<borders\b[^>]*\bcount=")\d+(")/,
    `$1${nextIndex + 1}$2`,
  );
  files[stylesKey] = codec.strToU8(stylesXml);
  return nextIndex;
};

const ensureExcelTableStyle = (
  files: ZipFileMap,
  codec: ZipTextCodec,
  options: {
    fillRgb: string;
    numFmtId?: number;
  },
) => {
  const stylesKey = "xl/styles.xml";
  if (!files[stylesKey]) return 0;

  const fillId = ensureExcelSolidFill(files, codec, options.fillRgb);
  const fontId = ensureExcelFont(files, codec, { size: 12 });
  const borderId = ensureExcelThinBlackBorder(files, codec);
  let stylesXml = codec.strFromU8(files[stylesKey]);
  const cellXfsStart = stylesXml.indexOf("<cellXfs");
  const cellXfsEnd = stylesXml.indexOf("</cellXfs>", cellXfsStart);
  if (cellXfsStart === -1 || cellXfsEnd === -1) return 0;

  const openEnd = stylesXml.indexOf(">", cellXfsStart);
  const cellXfsInner = stylesXml.slice(openEnd + 1, cellXfsEnd);
  const xfs = cellXfsInner.match(/<xf\b[^>]*(?:\/>|>[\s\S]*?<\/xf>)/g) ?? [];
  const styleXml =
    `<xf numFmtId="${options.numFmtId ?? 0}" fontId="${fontId}" fillId="${fillId}" borderId="${borderId}" xfId="0" ` +
    `applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"` +
    `${options.numFmtId !== undefined ? ' applyNumberFormat="1"' : ""}>` +
    `<alignment horizontal="center" vertical="center" wrapText="1"/></xf>`;
  const existingIndex = xfs.findIndex((xf) => xf === styleXml);
  if (existingIndex !== -1) return existingIndex;

  const nextIndex = xfs.length;
  stylesXml = `${stylesXml.slice(0, cellXfsEnd)}${styleXml}${stylesXml.slice(cellXfsEnd)}`;
  stylesXml = stylesXml.replace(
    /(<cellXfs\b[^>]*\bcount=")\d+(")/,
    `$1${nextIndex + 1}$2`,
  );
  files[stylesKey] = codec.strToU8(stylesXml);
  return nextIndex;
};

// â”€â”€â”€ MODIFIED: plain body style â€” font size 20, thin black border, white fill â”€
const ensureExcelPlainTicketBodyStyle = (
  files: ZipFileMap,
  codec: ZipTextCodec,
) => {
  const stylesKey = "xl/styles.xml";
  if (!files[stylesKey]) return 0;

  // Build the components we need: a size-20 font, a white fill, a thin black border.
  const fontId = ensureExcelFont(files, codec, { size: 20 }); // â† font size 20
  const fillId = ensureExcelWhiteFill(files, codec); // â† no colour fill (white)
  const borderId = ensureExcelThinBlackBorder(files, codec); // â† border on all sides

  let stylesXml = codec.strFromU8(files[stylesKey]);
  const cellXfsStart = stylesXml.indexOf("<cellXfs");
  const cellXfsEnd = stylesXml.indexOf("</cellXfs>", cellXfsStart);
  if (cellXfsStart === -1 || cellXfsEnd === -1) return 0;

  const openEnd = stylesXml.indexOf(">", cellXfsStart);
  const cellXfsInner = stylesXml.slice(openEnd + 1, cellXfsEnd);
  const xfs = cellXfsInner.match(/<xf\b[^>]*(?:\/>|>[\s\S]*?<\/xf>)/g) ?? [];

  // xf: General format, size-20 font, white fill, thin black border, center+middle, wrap on
  const styleXml =
    `<xf numFmtId="0" fontId="${fontId}" fillId="${fillId}" borderId="${borderId}" xfId="0" ` +
    `applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">` +
    `<alignment horizontal="center" vertical="center" wrapText="1"/></xf>`;

  const existingIndex = xfs.findIndex((xf) => xf === styleXml);
  if (existingIndex !== -1) return existingIndex;

  const nextIndex = xfs.length;
  stylesXml = `${stylesXml.slice(0, cellXfsEnd)}${styleXml}${stylesXml.slice(cellXfsEnd)}`;
  stylesXml = stylesXml.replace(
    /(<cellXfs\b[^>]*\bcount=")\d+(")/,
    `$1${nextIndex + 1}$2`,
  );
  files[stylesKey] = codec.strToU8(stylesXml);
  return nextIndex;
};

const applyExcelColumnAlignment = (
  files: ZipFileMap,
  codec: ZipTextCodec,
  xml: string,
  refs: string[],
  horizontal: ExcelHorizontalAlign,
) => {
  const styleCache = new Map<number, number>();
  return refs.reduce((sheetXml, ref) => {
    const baseStyle = getExcelCellStyle(sheetXml, ref);
    const alignedStyle =
      styleCache.get(baseStyle) ??
      ensureExcelAlignedStyle(files, codec, baseStyle, horizontal);
    styleCache.set(baseStyle, alignedStyle);
    return applyExcelCellStyle(sheetXml, ref, alignedStyle);
  }, xml);
};

const applySheetCellAlignment = (
  sheet: XLSX.WorkSheet,
  ref: string,
  horizontal: ExcelHorizontalAlign,
) => {
  const cell = sheet[ref] as XLSX.CellObject & { s?: any };
  if (!cell) return;
  cell.s = {
    ...(cell.s ?? {}),
    alignment: {
      ...(cell.s?.alignment ?? {}),
      horizontal,
      vertical: "center",
      wrapText: true,
    },
  };
};

const ticketExportCenteredIndexes = new Set([3, 4, 6, 7, 8, 9]);
const perfExportLeftIndexes = new Set([1, 2]);

const PPT_SLIDE_W = 13.333;
const PPT_SLIDE_H = 7.5;

const safePptText = (value: unknown, maxLength = 900): string => {
  const text = value === null || value === undefined ? "" : String(value);
  return text
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/g, " ")
    .replace(/[\uFFFE\uFFFF]/g, " ")
    .replace(/[\uD800-\uDFFF]/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, maxLength);
};

const safePptTableRows = (rows: any[][]): any[][] =>
  rows.map((row) =>
    row.map((cell) => {
      if (typeof cell === "string" || typeof cell === "number") {
        return safePptText(cell, 350);
      }
      if (cell && typeof cell === "object" && "text" in cell) {
        return {
          ...cell,
          text: safePptText((cell as { text?: unknown }).text, 550),
        };
      }
      return cell;
    }),
  );

const safePptNumber = (value: unknown): number => {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? Math.max(parsed, 0) : 0;
};

const addPptShapeBarChart = (
  slide: any,
  rectShape: any,
  rows: { label: string; value: number }[],
  options: {
    x: number;
    y: number;
    w: number;
    h: number;
    barColor: string;
    bgColor: string;
    gridColor: string;
    labelColor: string;
    valueColor: string;
    emptyText: string;
  },
) => {
  const chartRows = rows
    .map((row) => ({
      label: safePptText(row.label || "N/A", 24),
      value: safePptNumber(row.value),
    }))
    .slice(0, 34);

  slide.addShape(rectShape, {
    x: options.x,
    y: options.y,
    w: options.w,
    h: options.h,
    fill: { color: options.bgColor, transparency: 4 },
    line: { color: options.gridColor, transparency: 15, width: 0.5 },
  });

  if (!chartRows.length) {
    slide.addText(options.emptyText, {
      x: options.x + 0.25,
      y: options.y + options.h / 2 - 0.15,
      w: options.w - 0.5,
      h: 0.3,
      align: "center",
      fontSize: 12,
      color: options.labelColor,
      fontFace: "Segoe UI",
    });
    return;
  }

  const plotX = options.x + 0.45;
  const plotY = options.y + 0.3;
  const plotW = options.w - 0.75;
  const plotH = options.h - 1.15;
  const labelY = plotY + plotH + 0.1;
  const maxValue = Math.max(...chartRows.map((row) => row.value), 1);

  for (let i = 0; i <= 4; i += 1) {
    const y = plotY + (plotH / 4) * i;
    slide.addShape(rectShape, {
      x: plotX,
      y,
      w: plotW,
      h: 0.005,
      fill: { color: options.gridColor, transparency: 35 },
      line: { color: options.gridColor, transparency: 40, width: 0 },
    });
  }

  const slotW = plotW / Math.max(chartRows.length, 1);
  const barW = Math.min(0.24, Math.max(0.05, slotW * 0.55));

  chartRows.forEach((row, index) => {
    const barH =
      row.value > 0 ? Math.max((row.value / maxValue) * plotH, 0.04) : 0;
    const barX = plotX + index * slotW + (slotW - barW) / 2;
    const barY = plotY + plotH - barH;

    if (barH > 0) {
      slide.addShape(rectShape, {
        x: barX,
        y: barY,
        w: barW,
        h: barH,
        fill: { color: options.barColor },
        line: { color: options.barColor, transparency: 20, width: 0 },
      });
    }

    if (chartRows.length <= 24 && row.value > 0) {
      slide.addText(String(Math.round(row.value * 10) / 10), {
        x: barX - 0.08,
        y: Math.max(plotY - 0.08, barY - 0.18),
        w: barW + 0.16,
        h: 0.15,
        align: "center",
        fontSize: 5.5,
        color: options.valueColor,
        fontFace: "Segoe UI",
        fit: "shrink",
      } as any);
    }

    slide.addText(row.label, {
      x: barX - 0.18,
      y: labelY,
      w: 0.52,
      h: 0.34,
      rotate: 45,
      fontSize: 5.5,
      color: options.labelColor,
      fontFace: "Segoe UI",
      fit: "shrink",
    } as any);
  });
};

const getPptCellTextAndOptions = (cell: any) => {
  if (cell && typeof cell === "object" && "text" in cell) {
    return {
      text: safePptText(cell.text, 550),
      options: cell.options || {},
    };
  }

  return {
    text: safePptText(cell, 350),
    options: {},
  };
};

const addPptShapeTable = (
  slide: any,
  rectShape: any,
  rows: any[][],
  options: {
    x: number;
    y: number;
    colW: number[];
    rowH: number;
    fontSize: number;
    fontFace?: string;
    borderColor: string;
    fillColor: string;
    textColor: string;
  },
) => {
  const safeRows = safePptTableRows(rows);
  let y = options.y;

  safeRows.forEach((row) => {
    let x = options.x;

    options.colW.forEach((colW, colIndex) => {
      const { text, options: cellOptions } = getPptCellTextAndOptions(
        row[colIndex],
      );
      const fillColor = cellOptions.fill?.color || options.fillColor;
      const fontSize =
        typeof cellOptions.fontSize === "number"
          ? cellOptions.fontSize
          : options.fontSize;

      slide.addShape(rectShape, {
        x,
        y,
        w: colW,
        h: options.rowH,
        fill: { color: fillColor },
        line: { color: options.borderColor, width: 0.4 },
      });
      slide.addText(text, {
        x: x + 0.04,
        y: y + 0.03,
        w: Math.max(colW - 0.08, 0.1),
        h: Math.max(options.rowH - 0.06, 0.1),
        fontSize,
        fontFace: options.fontFace || "Segoe UI",
        color: cellOptions.color || options.textColor,
        bold: Boolean(cellOptions.bold),
        italic: Boolean(cellOptions.italic),
        align: cellOptions.align || "left",
        valign: cellOptions.valign || "mid",
        fit: "shrink",
      } as any);

      x += colW;
    });

    y += options.rowH;
  });
};

// 1. Performance Data Exporter to PPT
const exportPerfPpt = (
  data: PerfRow[],
  monthLabel: string = "All",
  executiveInsights?: ExecutiveInsights,
) => {
  try {
    const pptx = new PptxGenJS();
    pptx.layout = "LAYOUT_WIDE";
    const BG = "0a1628",
      CARD_BG = "0f1f38",
      CYAN = "22d3ee",
      GREEN = "10b981",
      RED = "ef4444",
      AMBER = "f59e0b",
      MUTED = "94a3b8",
      WHITE = "f8fafc";
    const availabilityChartRows = (data || []).map((r) => ({
      label: r?.siteName || r?.siteId || "N/A",
      value: safePptNumber(r?.availHours),
    }));
    const downtimeChartRows = (data || []).map((r) => ({
      label: r?.siteName || r?.siteId || "N/A",
      value: safePptNumber(r?.sitesDownHours),
    }));

    const s1 = pptx.addSlide();
    s1.background = { color: BG };
    s1.addShape(pptx.ShapeType.rect, {
      x: 0,
      y: 0,
      w: PPT_SLIDE_W,
      h: 0.06,
      fill: { color: CYAN },
    });
    s1.addText("DMR Monthly Performance Report", {
      x: 0.7,
      y: 1.6,
      w: 11.6,
      fontSize: 36,
      bold: true,
      color: WHITE,
      fontFace: "Segoe UI",
    });
    s1.addText(`Month: ${monthLabel}`, {
      x: 0.7,
      y: 2.5,
      w: 6,
      fontSize: 16,
      color: CYAN,
      fontFace: "Segoe UI",
    });
    if (executiveInsights) {
      s1.addText(
        `Network Health Score: ${executiveInsights.healthScore.score} / 100 Â· ${executiveInsights.healthScore.status}`,
        {
          x: 0.7,
          y: 2.95,
          w: 9,
          fontSize: 15,
          bold: true,
          color: WHITE,
          fontFace: "Segoe UI",
        },
      );
      s1.addText(safePptText(executiveInsights.healthScore.mainReason, 180), {
        x: 0.7,
        y: 3.32,
        w: 9.5,
        fontSize: 11,
        color: MUTED,
        fontFace: "Segoe UI",
      });
    }

    const kpi = computePerfKPIs(data);
    const kpiSlide = pptx.addSlide();
    kpiSlide.background = { color: BG };
    kpiSlide.addShape(pptx.ShapeType.rect, {
      x: 0,
      y: 0,
      w: PPT_SLIDE_W,
      h: 0.06,
      fill: { color: CYAN },
    });
    kpiSlide.addText("KPI Summary", {
      x: 0.5,
      y: 0.18,
      w: 12,
      fontSize: 22,
      bold: true,
      color: WHITE,
      fontFace: "Segoe UI",
    });
    kpiSlide.addText(`Month: ${monthLabel}  Â·  ${data.length} sites`, {
      x: 0.5,
      y: 0.55,
      w: 12,
      fontSize: 11,
      color: MUTED,
      fontFace: "Segoe UI",
    });

    const kpiItems = [
      { label: "% Availability", value: kpi.pctAvailability, color: GREEN },
      { label: "MTTR", value: kpi.mttr, color: AMBER },
      { label: "MTBF", value: kpi.mtbf, color: CYAN },
      { label: "MTTF", value: kpi.mttf, color: CYAN },
      {
        label: "Affected Sites",
        value: String(kpi.affectedSites),
        color: kpi.affectedSites > 0 ? RED : GREEN,
      },
      {
        label: "Total Down Time",
        value: kpi.totalDownHrs,
        color: kpi.totalDownHrs === "0.0 hrs" ? GREEN : RED,
      },
    ];

    const cW = 4.0,
      cH = 2.5,
      cGapX = 0.18,
      cGapY = 0.22;
    const cStartX = 0.4,
      cStartY = 0.95;
    kpiItems.forEach((item, idx) => {
      const col = idx % 3;
      const row = Math.floor(idx / 3);
      const x = cStartX + col * (cW + cGapX);
      const y = cStartY + row * (cH + cGapY);
      kpiSlide.addShape(pptx.ShapeType.roundRect, {
        x,
        y,
        w: cW,
        h: cH,
        fill: { color: CARD_BG },
        line: { color: "1e3a5f", width: 0.5 },
      });
      kpiSlide.addShape(pptx.ShapeType.rect, {
        x,
        y,
        w: cW,
        h: 0.06,
        fill: { color: item.color },
      });
      kpiSlide.addText(item.label.toUpperCase(), {
        x: x + 0.18,
        y: y + 0.18,
        w: cW - 0.36,
        fontSize: 9,
        color: MUTED,
        fontFace: "Segoe UI",
        bold: false,
      });
      kpiSlide.addText(item.value, {
        x: x + 0.18,
        y: y + 0.6,
        w: cW - 0.36,
        fontSize: 28,
        bold: true,
        color: item.color,
        fontFace: "Segoe UI",
      });
      kpiSlide.addText(`Month: ${monthLabel}`, {
        x: x + 0.18,
        y: y + cH - 0.4,
        w: cW - 0.36,
        fontSize: 8,
        color: MUTED,
        fontFace: "Segoe UI",
      });
    });

    const itemsPerSlide = 18;
    const headerRow = [
      {
        text: "S No",
        options: {
          bold: true,
          color: "0a1628",
          fill: { color: "22d3ee" },
          align: "center",
          valign: "mid",
        },
      },
      {
        text: "Site Name",
        options: {
          bold: true,
          color: "0a1628",
          fill: { color: "22d3ee" },
          align: "center",
          valign: "mid",
        },
      },
      {
        text: "Availability (Hrs)",
        options: {
          bold: true,
          color: "0a1628",
          fill: { color: "22d3ee" },
          align: "center",
          valign: "mid",
        },
      },
      {
        text: "Down (Hrs)",
        options: {
          bold: true,
          color: "0a1628",
          fill: { color: "22d3ee" },
          align: "center",
          valign: "mid",
        },
      },
      {
        text: "Reliability",
        options: {
          bold: true,
          color: "0a1628",
          fill: { color: "22d3ee" },
          align: "center",
          valign: "mid",
        },
      },
    ];

    for (let i = 0; i < data.length; i += itemsPerSlide) {
      const slide = pptx.addSlide();
      slide.background = { color: BG };
      slide.addShape(pptx.ShapeType.rect, {
        x: 0,
        y: 0,
        w: PPT_SLIDE_W,
        h: 0.06,
        fill: { color: CYAN },
      });
      const titleSuffix = i === 0 ? "" : " (Contd.)";
      slide.addText("Performance Table" + titleSuffix, {
        x: 0.5,
        y: 0.18,
        w: 12,
        fontSize: 18,
        bold: true,
        color: WHITE,
        fontFace: "Segoe UI",
      });
      const chunkData = data.slice(i, i + itemsPerSlide);
      const tableRows = [
        headerRow,
        ...chunkData.map((r, index) => {
          const actualIdx = i + index;
          const relNum = parseFloat(r.reliability);
          const relColor = !r.sitesDownHours
            ? GREEN
            : relNum < 95
              ? RED
              : relNum < 99
                ? AMBER
                : GREEN;
          return [
            {
              text: String(actualIdx + 1),
              options: { color: MUTED, align: "center", valign: "mid" },
            },
            {
              text: r?.siteName || "N/A",
              options: {
                color: WHITE,
                bold: true,
                align: "left",
                valign: "mid",
              },
            },
            {
              text: String(r?.availHours || 0),
              options: { color: GREEN, align: "center", valign: "mid" },
            },
            {
              text: String(r?.sitesDownHours || 0),
              options: {
                color: r.sitesDownHours > 0 ? RED : MUTED,
                align: "center",
                valign: "mid",
              },
            },
            {
              text: r?.reliability || "100%",
              options: {
                color: relColor,
                bold: true,
                align: "center",
                valign: "mid",
              },
            },
          ];
        }),
      ];
      addPptShapeTable(slide, pptx.ShapeType.rect, tableRows as any, {
        x: 0.5,
        y: 0.65,
        fontSize: 10,
        fontFace: "Segoe UI",
        borderColor: "1e3a5f",
        fillColor: CARD_BG,
        textColor: WHITE,
        rowH: 0.28,
        colW: [0.6, 3.5, 2.2, 2.0, 1.7],
      });
    }

    const slide3 = pptx.addSlide();
    slide3.background = { color: BG };
    slide3.addShape(pptx.ShapeType.rect, {
      x: 0,
      y: 0,
      w: PPT_SLIDE_W,
      h: 0.06,
      fill: { color: CYAN },
    });
    slide3.addText("Site Availability - Hours", {
      x: 0.5,
      y: 0.18,
      w: 12,
      fontSize: 18,
      bold: true,
      color: WHITE,
      fontFace: "Segoe UI",
    });
    slide3.addText(`Month: ${monthLabel}`, {
      x: 0.5,
      y: 0.52,
      w: 6,
      fontSize: 11,
      color: MUTED,
      fontFace: "Segoe UI",
    });
    addPptShapeBarChart(slide3, pptx.ShapeType.rect, availabilityChartRows, {
      x: 0.4,
      y: 0.85,
      w: 12.2,
      h: 5.8,
      barColor: GREEN,
      bgColor: CARD_BG,
      gridColor: "1e3a5f",
      labelColor: MUTED,
      valueColor: WHITE,
      emptyText: "No availability data available for this selection.",
    });

    const slide4 = pptx.addSlide();
    slide4.background = { color: BG };
    slide4.addShape(pptx.ShapeType.rect, {
      x: 0,
      y: 0,
      w: PPT_SLIDE_W,
      h: 0.06,
      fill: { color: RED },
    });
    slide4.addText("Site Downtime - Hours", {
      x: 0.5,
      y: 0.18,
      w: 12,
      fontSize: 18,
      bold: true,
      color: WHITE,
      fontFace: "Segoe UI",
    });
    slide4.addText(
      `Month: ${monthLabel}  Â·  ${data.filter((r) => r.sitesDownHours > 0).length} of ${data.length} sites affected`,
      {
        x: 0.5,
        y: 0.52,
        w: 10,
        fontSize: 11,
        color: MUTED,
        fontFace: "Segoe UI",
      },
    );
    addPptShapeBarChart(slide4, pptx.ShapeType.rect, downtimeChartRows, {
      x: 0.4,
      y: 0.85,
      w: 12.2,
      h: 5.8,
      barColor: RED,
      bgColor: CARD_BG,
      gridColor: "1e3a5f",
      labelColor: MUTED,
      valueColor: WHITE,
      emptyText: "No downtime data available for this selection.",
    });

    pptx.writeFile({
      fileName: `Performance_Report_${monthLabel.replace(/\s+/g, "_")}.pptx`,
    });
  } catch (error) {
    console.error("Failed to generate Performance PPT:", error);
  }
};

// 2. Tickets Data Exporter to PPT
const exportTicketsPpt = (
  tickets: any[],
  monthLabel: string = "All",
  executiveInsights?: ExecutiveInsights,
  deepDive?: DeepDiveAnalytics,
) => {
  try {
    const pptx = new PptxGenJS();
    pptx.layout = "LAYOUT_WIDE";
    const BG = "0a1628",
      CARD_BG = "0f1f38",
      CYAN = "22d3ee",
      WHITE = "f8fafc",
      MUTED = "94a3b8";

    const slide1 = pptx.addSlide();
    slide1.background = { color: BG };
    slide1.addShape(pptx.ShapeType.rect, {
      x: 0,
      y: 0,
      w: PPT_SLIDE_W,
      h: 0.06,
      fill: { color: CYAN },
    });
    slide1.addText("DMR Monthly Tickets Report", {
      x: 0.5,
      y: 2.5,
      w: 12,
      fontSize: 36,
      bold: true,
      color: WHITE,
      fontFace: "Segoe UI",
    });
    slide1.addText(`Target Month: ${monthLabel}`, {
      x: 0.5,
      y: 3.2,
      w: 12,
      fontSize: 18,
      color: CYAN,
      fontFace: "Segoe UI",
    });
    slide1.addText(`Total Aggregated Tickets: ${tickets.length}`, {
      x: 0.5,
      y: 3.6,
      w: 12,
      fontSize: 14,
      color: MUTED,
      italic: true,
    });

    if (executiveInsights) {
      const execSlide = pptx.addSlide();
      execSlide.background = { color: BG };
      execSlide.addShape(pptx.ShapeType.rect, {
        x: 0,
        y: 0,
        w: PPT_SLIDE_W,
        h: 0.06,
        fill: { color: CYAN },
      });
      execSlide.addText("Executive Management Summary", {
        x: 0.55,
        y: 0.28,
        w: 12,
        fontSize: 24,
        bold: true,
        color: WHITE,
        fontFace: "Segoe UI",
      });
      execSlide.addText(safePptText(executiveInsights.summaryText, 700), {
        x: 0.55,
        y: 0.85,
        w: 7.8,
        h: 1.1,
        fontSize: 14,
        color: MUTED,
        breakLine: false,
        fit: "shrink",
      } as any);
      execSlide.addShape(pptx.ShapeType.roundRect, {
        x: 8.65,
        y: 0.82,
        w: 3.85,
        h: 1.5,
        fill: { color: CARD_BG, transparency: 3 },
        line: { color: CYAN, transparency: 15 },
      } as any);
      execSlide.addText("Network Health Score", {
        x: 8.9,
        y: 1.0,
        w: 3.4,
        fontSize: 10,
        color: MUTED,
        bold: true,
      });
      execSlide.addText(String(executiveInsights.healthScore.score), {
        x: 8.9,
        y: 1.25,
        w: 1.2,
        fontSize: 34,
        bold: true,
        color: CYAN,
      });
      execSlide.addText(executiveInsights.healthScore.status, {
        x: 10.15,
        y: 1.38,
        w: 1.9,
        fontSize: 16,
        bold: true,
        color: WHITE,
      });

      const insightRows = executiveInsights.cards.slice(0, 8).map((card) => [
        {
          text: card.label,
          options: { color: MUTED, fontSize: 9, bold: true },
        },
        {
          text: String(card.value),
          options: { color: WHITE, fontSize: 12, bold: true },
        },
        { text: card.note, options: { color: MUTED, fontSize: 8 } },
      ]);
      addPptShapeTable(
        execSlide,
        pptx.ShapeType.rect,
        insightRows as any,
        {
          x: 0.55,
          y: 2.25,
          colW: [3.1, 2.2, 6.7],
          rowH: 0.35,
          fontSize: 9,
          fontFace: "Segoe UI",
          borderColor: "1e3a5f",
          fillColor: CARD_BG,
          textColor: WHITE,
        } as any,
      );
    }

    if (executiveInsights?.highRiskSites.length) {
      const riskSlide = pptx.addSlide();
      riskSlide.background = { color: BG };
      riskSlide.addShape(pptx.ShapeType.rect, {
        x: 0,
        y: 0,
        w: PPT_SLIDE_W,
        h: 0.06,
        fill: { color: "ef4444" },
      });
      riskSlide.addText("High-Risk Sites & Follow-Up Priority", {
        x: 0.55,
        y: 0.25,
        w: 12,
        fontSize: 23,
        bold: true,
        color: WHITE,
      });
      const riskRows = [
        [
          "Rank",
          "Site ID",
          "Site Name",
          "Tickets",
          "Downtime",
          "Reliability",
          "Top RCA",
          "Risk",
        ].map((text) => ({
          text,
          options: {
            bold: true,
            color: "0a1628",
            fill: { color: "ef4444" },
            align: "center",
          },
        })),
        ...executiveInsights.highRiskSites.slice(0, 10).map((site) => [
          {
            text: String(site.rank),
            options: { color: MUTED, align: "center" },
          },
          { text: site.siteId, options: { color: CYAN, bold: true } },
          { text: site.siteName || "-", options: { color: WHITE } },
          {
            text: String(site.ticketCount),
            options: { color: WHITE, align: "center" },
          },
          {
            text: `${site.downtimeHours} hrs`,
            options: { color: WHITE, align: "center" },
          },
          {
            text: `${site.reliability.toFixed(2)}%`,
            options: { color: WHITE, align: "center" },
          },
          { text: site.topRca, options: { color: WHITE } },
          {
            text: `${site.riskLevel} ${site.riskScore}`,
            options: { color: "f59e0b", bold: true },
          },
        ]),
      ];
      addPptShapeTable(
        riskSlide,
        pptx.ShapeType.rect,
        riskRows as any,
        {
          x: 0.35,
          y: 0.9,
          fontSize: 8,
          rowH: 0.34,
          colW: [0.55, 1.2, 2.2, 0.8, 1.0, 1.0, 3.3, 1.6],
          fontFace: "Segoe UI",
          borderColor: "1e3a5f",
          fillColor: CARD_BG,
          textColor: WHITE,
        } as any,
      );
    }

    if (deepDive) {
      const deepSlide = pptx.addSlide();
      deepSlide.background = { color: BG };
      deepSlide.addShape(pptx.ShapeType.rect, {
        x: 0,
        y: 0,
        w: PPT_SLIDE_W,
        h: 0.06,
        fill: { color: "a78bfa" },
      });
      deepSlide.addText("RCA, Preventability & SLA Deep-Dive", {
        x: 0.55,
        y: 0.25,
        w: 12,
        fontSize: 23,
        bold: true,
        color: WHITE,
      });
      const rcaRows = [
        [
          "RCA Family",
          "Tickets",
          "Downtime",
          "Missing RCA",
          "Preventable",
          "Owner Team",
        ].map((text) => ({
          text,
          options: {
            bold: true,
            color: "0a1628",
            fill: { color: "a78bfa" },
            align: "center",
          },
        })),
        ...deepDive.rcaFamilyDeepDive.slice(0, 7).map((row) => [
          { text: row.family, options: { color: WHITE, bold: true } },
          {
            text: String(row.tickets),
            options: { color: WHITE, align: "center" },
          },
          {
            text: `${row.downtimeHours} hrs`,
            options: { color: WHITE, align: "center" },
          },
          {
            text: String(row.missingRca),
            options: {
              color: row.missingRca ? "f59e0b" : "34d399",
              align: "center",
            },
          },
          {
            text: String(row.preventableTickets),
            options: { color: WHITE, align: "center" },
          },
          { text: row.responsibleTeam, options: { color: MUTED } },
        ]),
      ];
      addPptShapeTable(
        deepSlide,
        pptx.ShapeType.rect,
        rcaRows as any,
        {
          x: 0.45,
          y: 0.85,
          fontSize: 8,
          rowH: 0.35,
          colW: [2.7, 0.8, 1.1, 1.1, 1.1, 5.6],
          fontFace: "Segoe UI",
          borderColor: "1e3a5f",
          fillColor: CARD_BG,
          textColor: WHITE,
        } as any,
      );
      deepSlide.addText("Recommended Management Actions", {
        x: 0.55,
        y: 4.05,
        w: 12,
        fontSize: 14,
        bold: true,
        color: CYAN,
      });
      deepDive.recommendations.forEach((item, idx) => {
        deepSlide.addText(safePptText(`${idx + 1}. ${item}`, 260), {
          x: 0.7,
          y: 4.4 + idx * 0.42,
          w: 11.5,
          fontSize: 10,
          color: WHITE,
          fit: "shrink",
        } as any);
      });
    }

    const itemsPerSlide = 16;
    const headerRow = [
      {
        text: "S No",
        options: {
          bold: true,
          color: "0a1628",
          fill: { color: CYAN },
          align: "center",
          valign: "mid",
        },
      },
      {
        text: "Ticket ID",
        options: {
          bold: true,
          color: "0a1628",
          fill: { color: CYAN },
          align: "center",
          valign: "mid",
        },
      },
      {
        text: "Site IDs",
        options: {
          bold: true,
          color: "0a1628",
          fill: { color: CYAN },
          align: "center",
          valign: "mid",
        },
      },
      {
        text: "Resource",
        options: {
          bold: true,
          color: "0a1628",
          fill: { color: CYAN },
          align: "center",
          valign: "mid",
        },
      },
      {
        text: "Sev",
        options: {
          bold: true,
          color: "0a1628",
          fill: { color: CYAN },
          align: "center",
          valign: "mid",
        },
      },
      {
        text: "Issue Description",
        options: {
          bold: true,
          color: "0a1628",
          fill: { color: CYAN },
          align: "center",
          valign: "mid",
        },
      },
      {
        text: "Status",
        options: {
          bold: true,
          color: "0a1628",
          fill: { color: CYAN },
          align: "center",
          valign: "mid",
        },
      },
    ];

    for (let i = 0; i < tickets.length; i += itemsPerSlide) {
      const slide = pptx.addSlide();
      slide.background = { color: BG };
      slide.addShape(pptx.ShapeType.rect, {
        x: 0,
        y: 0,
        w: PPT_SLIDE_W,
        h: 0.06,
        fill: { color: CYAN },
      });
      const titleSuffix = i === 0 ? "" : " (Contd.)";
      slide.addText(`Aggregated Ticket Details${titleSuffix}`, {
        x: 0.5,
        y: 0.2,
        w: 12,
        fontSize: 20,
        bold: true,
        color: WHITE,
        fontFace: "Segoe UI",
      });
      const chunk = tickets.slice(i, i + itemsPerSlide);
      const tableRows = [
        headerRow,
        ...chunk.map((t, idx) => {
          const row = t.primary || {};
          const sites = Array.from(t.siteIds || []).join(", ");
          return [
            {
              text: String(i + idx + 1),
              options: { color: MUTED, align: "center" },
            },
            { text: t.tt || "N/A", options: { color: CYAN, bold: true } },
            {
              text: sites.length > 30 ? sites.substring(0, 30) + "..." : sites,
              options: { color: WHITE, fontSize: 7 },
            },
            {
              text: row.managedResource || "N/A",
              options: { color: WHITE, align: "center", valign: "mid" },
            },
            {
              text: row.severity || "N/A",
              options: {
                color: row.severity === "Critical" ? "ef4444" : WHITE,
                bold: true,
                align: "center",
                valign: "mid",
              },
            },
            {
              text: row.issue || "N/A",
              options: { color: WHITE, fontSize: 8 },
            },
            {
              text: row.status || "N/A",
              options: {
                color: row.status === "Resolved" ? "10b981" : "f59e0b",
                bold: true,
              },
            },
          ];
        }),
      ];
      addPptShapeTable(slide, pptx.ShapeType.rect, tableRows as any, {
        x: 0.4,
        y: 0.7,
        fontSize: 9,
        fontFace: "Segoe UI",
        borderColor: "1e3a5f",
        fillColor: CARD_BG,
        textColor: WHITE,
        rowH: 0.35,
        colW: [0.6, 1.2, 1.8, 1.8, 0.8, 5.1, 1.2],
      });
      slide.addText(`Page ${Math.floor(i / itemsPerSlide) + 2}`, {
        x: 12,
        y: 7.1,
        w: 1,
        fontSize: 8,
        color: MUTED,
        align: "right",
      });
    }

    pptx.writeFile({
      fileName: `DMR_Tickets_Report_${monthLabel.replace(/\s+/g, "_")}.pptx`,
    });
  } catch (error) {
    console.error("Failed to generate Tickets PPT:", error);
    alert("Error generating PowerPoint. Please check the console.");
  }
};

const COLORS = [
  "#22d3ee",
  "#60a5fa",
  "#f59e0b",
  "#ef4444",
  "#34d399",
  "#a78bfa",
  "#f472b6",
  "#94a3b8",
];
const STATUS_COLORS: Record<string, string> = {
  Closed: "#34d399",
  Pending: "#f59e0b",
  Resolved: "#60a5fa",
  Open: "#ef4444",
};
const SEVERITY_COLORS: Record<string, string> = {
  Critical: "#ef4444",
  Major: "#f59e0b",
  Minor: "#22d3ee",
  Low: "#34d399",
};
const CHART_GRID_STROKE = "rgba(148, 163, 184, .14)";
const CHART_AXIS_STROKE = "#8ea4c2";
const CHART_LABEL_FILL = "#e5f0ff";
const CHART_TOOLTIP_STYLE: CSSProperties = {
  background:
    "linear-gradient(145deg, rgba(8, 17, 34, .98), rgba(15, 31, 56, .96))",
  border: "1px solid rgba(125, 211, 252, .34)",
  borderRadius: 10,
  color: "#e5f0ff",
  boxShadow: "0 18px 45px rgba(0, 0, 0, .35)",
};
const CHART_LEGEND_STYLE: CSSProperties = {
  color: "#bfd3ee",
  fontSize: 11,
  fontWeight: 700,
};
const BAR_RADIUS: [number, number, number, number] = [0, 8, 8, 0];
const COLUMN_BAR_RADIUS: [number, number, number, number] = [6, 6, 0, 0];

// ─── Pending Aging Waterfall helpers ────────────────────────────────────────

type WaterfallEntry = {
  name: string;
  /** invisible base that floats the visible bar */
  base: number;
  /** visible segment height */
  segment: number;
  /** true colour of this segment */
  color: string;
  /** the running total AFTER this bar (used for connector lines) */
  runningTotal: number;
  /** display label */
  ticketCount: number;
  isTotal: boolean;
};

function buildWaterfallData(
  buckets: { name: string; value: number }[],
): WaterfallEntry[] {
  const BUCKET_COLORS: Record<string, string> = {
    "0\u201324h": "#38bdf8", // sky-blue
    "1\u20133d": "#60a5fa", // blue
    "3\u20137d": "#f59e0b", // amber
    "7d+": "#ef4444",    // red
  };

  const total = buckets.reduce((s, b) => s + b.value, 0);
  const entries: WaterfallEntry[] = [];

  // First bar: full total (base=0, segment=total)
  entries.push({
    name: "Total",
    base: 0,
    segment: total,
    color: "#22d3ee", // cyan accent
    runningTotal: total,
    ticketCount: total,
    isTotal: true,
  });

  // Each bucket is a descending floating segment
  let running = total;
  for (const bucket of buckets) {
    running -= bucket.value;
    entries.push({
      name: bucket.name,
      base: running,
      segment: bucket.value,
      color: BUCKET_COLORS[bucket.name] ?? "#94a3b8",
      runningTotal: running,
      ticketCount: bucket.value,
      isTotal: false,
    });
  }

  return entries;
}

function PendingAgingWaterfall({
  buckets,
}: {
  buckets: { name: string; value: number }[];
}) {
  const data = buildWaterfallData(buckets);
  const maxY = data[0]?.segment ?? 0;

  // Dashed connector lines between bars: drawn as ReferenceLine segments
  // from the top of each bar to the base of the next.
  const connectors = data.slice(0, -1).map((entry, i) => ({
    x1Label: entry.name,
    x2Label: data[i + 1].name,
    y: entry.runningTotal,
  }));

  return (
    <ResponsiveContainer width="100%" height={260}>
      <ComposedChart
        data={data}
        margin={{ left: 4, right: 20, top: 28, bottom: 4 }}
        barCategoryGap="25%"
      >
        <CartesianGrid stroke={CHART_GRID_STROKE} vertical={false} />
        <XAxis
          dataKey="name"
          stroke={CHART_AXIS_STROKE}
          tickLine={false}
          axisLine={false}
          tick={{ fontSize: 11, fill: CHART_AXIS_STROKE }}
        />
        <YAxis
          stroke={CHART_AXIS_STROKE}
          allowDecimals={false}
          tickLine={false}
          axisLine={false}
          domain={[0, Math.ceil(maxY * 1.15) || 10]}
          tick={{ fontSize: 11, fill: CHART_AXIS_STROKE }}
        />
        <Tooltip
          contentStyle={CHART_TOOLTIP_STYLE}
          formatter={(_value: unknown, _name: string, props: { payload?: WaterfallEntry }) => [
            props.payload?.ticketCount ?? 0,
            "Tickets",
          ]}
          labelFormatter={(label: string) => label}
        />

        {/* Invisible base bar to float the visible segment */}
        <Bar dataKey="base" stackId="wf" fill="transparent" isAnimationActive={false} />

        {/* Visible coloured segment */}
        <Bar
          dataKey="segment"
          stackId="wf"
          radius={[6, 6, 0, 0]}
          isAnimationActive={true}
          label={false}
        >
          {data.map((entry) => (
            <Cell key={entry.name} fill={entry.color} />
          ))}
          <LabelList
            dataKey="ticketCount"
            position="top"
            fill={CHART_LABEL_FILL}
            fontSize={12}
            fontWeight={700}
          />
        </Bar>

        {/* Dashed connector lines between bars */}
        {connectors.map((c) => (
          <ReferenceLine
            key={`conn-${c.x1Label}`}
            y={c.y}
            stroke="rgba(148,163,184,0.45)"
            strokeDasharray="4 4"
            strokeWidth={1.5}
            ifOverflow="extendDomain"
          />
        ))}
      </ComposedChart>
    </ResponsiveContainer>
  );
}
// ─── End Waterfall helpers ───────────────────────────────────────────────────

const EMPTY_FILTERS: Filters = {
  search: "",
  status: [],
  severity: [],
  region: [],
  impact: [],
  site: [],
  openingMonth: [],
  rcaFamily: [],
};

function countBy<T>(
  items: T[],
  keyFn: (item: T) => string,
): { name: string; value: number }[] {
  const map = new Map<string, number>();
  items.forEach((item) => {
    const key = keyFn(item) || "Blank";
    map.set(key, (map.get(key) ?? 0) + 1);
  });
  return Array.from(map.entries())
    .map(([name, value]) => ({ name, value }))
    .sort((a, b) => b.value - a.value || a.name.localeCompare(b.name));
}
function pct(value: number, total: number): string {
  if (!total) return "";
  return `${((value / total) * 100).toFixed(1)}%`;
}
function metricValue(
  items: { name: string; value: number }[],
  expected: string,
): number {
  return (
    items.find((item) => item.name.toLowerCase() === expected.toLowerCase())
      ?.value ?? 0
  );
}
function renderPieLabel(props: {
  name?: string;
  value?: number;
  percent?: number;
}) {
  const value = props.value ?? 0;
  if (!value) return "";
  return `${value}`;
}

const DISTINCT_REPORT_HEADERS = [
  "#",
  "Site ID",
  "Site Name",
  "Managed Resource",
  "Severity",
  "Issues",
  "Observation Date",
  "Observation Time",
  "Recovery Date",
  "Recovery Time",
  "Escalated for L3 Support Date",
  "Escalated for L3 Support Time",
  "Total Duration Days/Hours",
  "TT",
  "Status",
  "Escalated to",
  "RCA",
  "Action",
];
const PERF_REPORT_HEADERS = [
  "S No",
  "Site ID",
  "Site Name",
  "Site Availability, Hrs",
  "Site Availability, days",
  "Channel Busy Count",
  "MW link Performance, Hrs",
  "DMR Reliability",
  "Sites Down, hrs",
];
const PERF_TEMPLATE_PDF_HEADERS = [
  "S No",
  "Site No",
  "Site ID",
  "Site Availability, Hrs",
  "Site Availability, days",
  "Channel Busy Count",
  "MW link Performance, Hrs",
  "DMR Reliability",
  "Sites Down, hrs",
];
const TICKET_TEMPLATE_PDF_HEAD = [
  [
    "No",
    "Equipment/ site",
    "Site Name",
    "Effected Managed Resource",
    "Severity",
    "Alarm Type",
    "Escalation",
    "",
    "Recovery",
    "",
    "Escalated for L3 Support",
    "",
    "Outage Duration",
    "TT Number",
    "TT Status",
    "TT Owner",
    "RCA",
    "Comments",
  ],
  [
    "",
    "",
    "",
    "",
    "",
    "",
    "Date",
    "Time",
    "Date",
    "Time",
    "Date",
    "Time",
    "",
    "",
    "",
    "",
    "",
    "",
  ],
];

const TEMPLATE_PDF_COLORS = {
  title: [192, 0, 0] as [number, number, number],
  band: [192, 0, 0] as [number, number, number],
  bandText: [255, 255, 255] as [number, number, number],
  header: [217, 217, 217] as [number, number, number],
  subHeader: [242, 242, 242] as [number, number, number],
  text: [0, 0, 0] as [number, number, number],
  muted: [89, 89, 89] as [number, number, number],
  border: [128, 128, 128] as [number, number, number],
  ok: [0, 128, 0] as [number, number, number],
  warn: [192, 0, 0] as [number, number, number],
};

type ReportLogos = { nasco: HTMLImageElement };
let reportLogoPromise: Promise<ReportLogos> | null = null;
function loadImageElement(src: string): Promise<HTMLImageElement> {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => resolve(img);
    img.onerror = reject;
    img.src = src;
  });
}
function loadReportLogos(): Promise<ReportLogos> {
  reportLogoPromise ??= loadImageElement(nascoLogoSrc).then((nasco) => ({
    nasco,
  }));
  return reportLogoPromise;
}
function drawPdfReportHeader(
  doc: jsPDF,
  pageW: number,
  title: string,
  subtitle: string,
  logos: ReportLogos,
  titleColor: [number, number, number] = TEMPLATE_PDF_COLORS.title,
) {
  doc.addImage(logos.nasco, "PNG", 12, 5, 18, 18);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(15);
  doc.setTextColor(...titleColor);
  doc.text(title, pageW / 2, 12.2, { align: "center" });
  doc.setFont("helvetica", "normal");
  doc.setFontSize(8);
  doc.setTextColor(...TEMPLATE_PDF_COLORS.muted);
  doc.text(subtitle, pageW / 2, 17.2, { align: "center" });
}

function normalizePerformanceRegionLabel(value: string): string {
  const label = clean(value);
  const region = label.toUpperCase();
  if (region === "EOA" || region === "NEOA") return "EOA";
  if (region === "SOA") return "SOA";
  if (region === "COA") return "COA";
  if (region === "WOA") return "WOA";
  if (/\bSOA\b/.test(region)) return "SOA";
  if (/\bCOA\b/.test(region)) return "COA";
  if (/\bWOA\b/.test(region)) return "WOA";
  if (/\b(?:EOA|NEOA)\b/.test(region)) return "EOA";
  return label;
}

function perfSourceLabel(row: TicketRecord): string {
  const region = normalizePerformanceRegionLabel(row.region);
  if (region) return region;
  return clean(row.sourceFile) || "Workbook";
}
function perfEntryKey(sourceLabel: string, siteId: string): string {
  return `${sourceLabel.toLowerCase()}||${normalizeSiteId(clean(siteId)).toLowerCase()}`;
}

type PerformanceSiteOrderEntry = {
  siteId: string;
  siteName: string;
  sourceLabel?: string;
};

function computePerfRows(
  allRows: TicketRecord[],
  monthKey: string,
  siteOrder: PerformanceSiteOrderEntry[] = [],
): PerfRow[] {
  const range = monthKey !== "all" ? selectedMonthRange(monthKey) : null;
  const allMonthKeys = Array.from(
    new Set(
      allRows
        .map(
          (row) => row.openingMonthKey || openingMonthKey(row.observationDate),
        )
        .filter((key) => key && key !== "Unknown"),
    ),
  );
  const monthHours =
    monthKey !== "all"
      ? totalHoursInMonth(monthKey)
      : allMonthKeys.length
        ? allMonthKeys.reduce((sum, key) => sum + totalHoursInMonth(key), 0)
        : 24 * 30;

  const siteNameMap = new Map<string, string>();
  const siteIdMap = new Map<string, string>();
  const sourceMap = new Map<string, string>();
  const siteTicketCount = new Map<string, number>();
  const siteDownHours = new Map<string, number>();
  const countedSiteTickets = new Set<string>();
  const countedSiteDowntime = new Set<string>();

  siteOrder.forEach((site) => {
    const siteId = normalizeSiteId(clean(site.siteId));
    if (!siteId) return;
    const sourceLabel = site.sourceLabel || "Workbook";
    const key = perfEntryKey(sourceLabel, siteId);
    if (!siteIdMap.has(key)) siteIdMap.set(key, siteId);
    if (!sourceMap.has(key)) sourceMap.set(key, sourceLabel);
    if (!siteNameMap.has(key) && site.siteName)
      siteNameMap.set(key, site.siteName);
  });

  allRows.forEach((row) => {
    if (!row.siteId) return;
    const sourceLabel = perfSourceLabel(row);
    const normalizedSiteId = normalizeSiteId(clean(row.siteId));
    const key = perfEntryKey(sourceLabel, normalizedSiteId);
    if (!siteIdMap.has(key)) siteIdMap.set(key, normalizedSiteId);
    if (!sourceMap.has(key)) sourceMap.set(key, sourceLabel);
    if (!siteNameMap.has(key) && row.siteName)
      siteNameMap.set(key, row.siteName);
    const ticketKey = `${key}||${clean(row.tt) || row.rowNo}`;
    if (!countedSiteTickets.has(ticketKey)) {
      countedSiteTickets.add(ticketKey);
      siteTicketCount.set(key, (siteTicketCount.get(key) ?? 0) + 1);
    }
  });

  function combineDatetime(dateStr: string, timeStr: string): Date | null {
    const d = parseDateValue(dateStr);
    if (!d) return null;
    if (timeStr) {
      const tm = timeStr.match(/^(\d{1,2}):(\d{2})/);
      if (tm) {
        d.setHours(Number(tm[1]), Number(tm[2]), 0, 0);
      }
    }
    return d;
  }

  allRows.forEach((row) => {
    if (!row.siteId) return;
    if (clean(row.impact).toLowerCase() !== "service impact") return;
    const mr = clean(row.managedResource).toLowerCase();
    if (mr !== "complete site" && mr !== "link down") return;
    const outageStart = combineDatetime(
      row.observationDate,
      row.observationTime,
    );
    if (!outageStart) return;
    let outageEnd: Date | null = combineDatetime(
      row.recoveryDate,
      row.recoveryTime,
    );
    const sourceLabel = perfSourceLabel(row);
    const normalizedSiteId = normalizeSiteId(clean(row.siteId));
    const key = perfEntryKey(sourceLabel, normalizedSiteId);
    const ticketKey = `${key}||${clean(row.tt) || row.rowNo}`;
    if (countedSiteDowntime.has(ticketKey)) return;
    countedSiteDowntime.add(ticketKey);

    if (monthKey === "all") {
      const hours = ticketDurationHours(row);
      siteDownHours.set(
        key,
        Math.round(((siteDownHours.get(key) ?? 0) + hours) * 10) / 10,
      );
      return;
    }
    const monthRange = selectedMonthRange(monthKey);
    if (!monthRange) return;
    const { start: monthStart, end: monthEnd } = monthRange;
    if (!outageEnd) {
      outageEnd = monthEnd;
    }
    if (outageEnd <= monthStart || outageStart > monthEnd) return;
    const effectiveStart = outageStart < monthStart ? monthStart : outageStart;
    const effectiveEnd = outageEnd > monthEnd ? monthEnd : outageEnd;
    const overlapMs = effectiveEnd.getTime() - effectiveStart.getTime();
    if (overlapMs <= 0) return;
    const hours = Math.round((overlapMs / (1000 * 60 * 60)) * 10) / 10;
    siteDownHours.set(
      key,
      Math.round(((siteDownHours.get(key) ?? 0) + hours) * 10) / 10,
    );
  });

  const allSiteKeys = Array.from(siteIdMap.keys()).filter(Boolean);
  allSiteKeys.sort(
    (a, b) => (siteTicketCount.get(b) ?? 0) - (siteTicketCount.get(a) ?? 0),
  );
  const siteEntries = allSiteKeys.map((key) => ({
    siteId: siteIdMap.get(key) ?? "",
    siteName: siteNameMap.get(key) ?? "",
    sourceLabel: sourceMap.get(key) ?? "Workbook",
    perfKey: key,
  }));

  return siteEntries
    .sort((a, b) => comparePerformanceSiteRows(a, b))
    .map(({ siteId, siteName, sourceLabel, perfKey }) => {
      const downHours = Math.min(siteDownHours.get(perfKey) ?? 0, monthHours);
      const availHours = Math.max(0, monthHours - downHours);
      const totalHours = availHours + downHours;
      const reliability = totalHours > 0 ? availHours / totalHours : 1;
      const totalMins = Math.round(availHours * 60);
      const dDays = Math.floor(totalMins / (60 * 24));
      const dHrs = Math.floor((totalMins % (60 * 24)) / 60);
      const dMins = Math.round(totalMins % 60);
      return {
        siteId,
        siteName,
        displayName: formatPerformanceChartLabel(siteId, siteName, sourceLabel),
        sourceLabel,
        perfKey,
        sitesDownHours: downHours,
        availHours: Math.round(availHours * 10) / 10,
        availDay: `${dDays} d, ${dHrs} h, ${dMins} m`,
        reliability: `${(reliability * 100).toFixed(2)}%`,
        channelBusy: 0,
        mwLinkPerf: 0,
        ticketCount: siteTicketCount.get(perfKey) ?? 0,
      };
    });
}

function performanceRegionSortValue(sourceLabel: string): number {
  const region = normalizePerformanceRegionLabel(sourceLabel).toUpperCase();
  const order = ["EOA", "SOA", "COA", "WOA"];
  const index = order.indexOf(region);
  return index >= 0 ? index : order.length;
}

function performanceRegionDisplayLabel(sourceLabel: string): string {
  const region = normalizePerformanceRegionLabel(sourceLabel).toUpperCase();
  if (region === "EOA") return "EOA&NEOA";
  return clean(sourceLabel);
}

function comparePerformanceSiteRows(
  a: { siteId: string; siteName?: string; sourceLabel?: string },
  b: { siteId: string; siteName?: string; sourceLabel?: string },
): number {
  const byRfNumber = rfSiteSortValue(a.siteId) - rfSiteSortValue(b.siteId);
  if (byRfNumber !== 0) return byRfNumber;
  const byRegion =
    performanceRegionSortValue(a.sourceLabel ?? "") -
    performanceRegionSortValue(b.sourceLabel ?? "");
  if (byRegion !== 0) return byRegion;
  const bySiteId = clean(a.siteId).localeCompare(clean(b.siteId), undefined, {
    numeric: true,
    sensitivity: "base",
  });
  if (bySiteId !== 0) return bySiteId;
  return clean(a.siteName).localeCompare(clean(b.siteName), undefined, {
    numeric: true,
    sensitivity: "base",
  });
}

function formatPerformanceChartLabel(
  siteId: string,
  siteName: string,
  sourceLabel: string,
): string {
  const name = clean(siteName);
  const region = performanceRegionDisplayLabel(sourceLabel);
  const regionSuffix = region && region !== "Workbook" ? ` (${region})` : "";
  return `${normalizeSiteId(clean(siteId))}${name ? ` - ${name}` : ""}${regionSuffix}`;
}
function perfReportRows(rows: PerfRow[]): string[][] {
  return rows.map((r, i) => [
    String(i + 1),
    r.siteId,
    r.siteName,
    String(r.availHours),
    r.availDay,
    "",
    "",
    r.reliability,
    String(r.sitesDownHours),
  ]);
}

function rfSiteSortValue(siteId: string): number {
  const match = clean(siteId).match(/rf\s*site\s*(\d+)/i);
  return match ? Number(match[1]) : Number.MAX_SAFE_INTEGER;
}

function aggregatePerfRowsForReport(rows: PerfRow[]): PerfRow[] {
  const grouped = new Map<string, PerfRow>();
  const periodHoursBySite = new Map<string, number>();

  rows
    .filter((row) => isRfSiteId(row.siteId))
    .forEach((row) => {
      const key =
        normalizeSiteId(clean(row.siteId)).toLowerCase() ||
        clean(row.siteId).toLowerCase();
      const periodHours = row.availHours + row.sitesDownHours;
      periodHoursBySite.set(
        key,
        Math.max(periodHoursBySite.get(key) ?? 0, periodHours),
      );
      const existing = grouped.get(key);
      if (!existing) {
        grouped.set(key, { ...row });
        return;
      }
      existing.sitesDownHours =
        Math.round((existing.sitesDownHours + row.sitesDownHours) * 10) / 10;
      existing.channelBusy =
        Math.round(
          (safePptNumber(existing.channelBusy) +
            safePptNumber(row.channelBusy)) *
            10,
        ) / 10;
      existing.mwLinkPerf =
        Math.round(
          (safePptNumber(existing.mwLinkPerf) + safePptNumber(row.mwLinkPerf)) *
            10,
        ) / 10;
      existing.ticketCount =
        (existing.ticketCount ?? 0) + (row.ticketCount ?? 0);
      const names = new Set(
        [existing.siteName, row.siteName].map((v) => clean(v)).filter(Boolean),
      );
      existing.siteName = Array.from(names).join(" / ");
    });

  return Array.from(grouped.values())
    .map((row) => {
      const key =
        normalizeSiteId(clean(row.siteId)).toLowerCase() ||
        clean(row.siteId).toLowerCase();
      const periodHours =
        periodHoursBySite.get(key) ?? row.availHours + row.sitesDownHours;
      const availHours = Math.max(0, periodHours - row.sitesDownHours);
      const totalHours = availHours + row.sitesDownHours;
      const reliability = totalHours > 0 ? availHours / totalHours : 1;
      const totalMins = Math.round(availHours * 60);
      const dDays = Math.floor(totalMins / (60 * 24));
      const dHrs = Math.floor((totalMins % (60 * 24)) / 60);
      const dMins = Math.round(totalMins % 60);
      return {
        ...row,
        availHours: Math.round(availHours * 10) / 10,
        availDay: `${dDays} d, ${dHrs} h, ${dMins} m`,
        reliability: `${(reliability * 100).toFixed(2)}%`,
      };
    })
    .sort((a, b) => {
      const byRfNumber = rfSiteSortValue(a.siteId) - rfSiteSortValue(b.siteId);
      if (byRfNumber !== 0) return byRfNumber;
      return clean(a.siteId).localeCompare(clean(b.siteId), undefined, {
        numeric: true,
        sensitivity: "base",
      });
    });
}

function computePerfKPIs(rows: PerfRow[]) {
  const totalAvail = rows.reduce((s, r) => s + r.availHours, 0);
  const totalDown = rows.reduce((s, r) => s + r.sitesDownHours, 0);
  const totalSiteIds = new Set(
    rows.map((r) => clean(r.siteId)).filter(isRfSiteId),
  );
  const affectedSiteIds = new Set(
    rows
      .filter((r) => r.sitesDownHours > 0)
      .map((r) => clean(r.siteId))
      .filter(isRfSiteId),
  );
  const sitesWithDown = affectedSiteIds.size;
  const nonAffectedSites = Math.max(0, totalSiteIds.size - sitesWithDown);
  const totalHrs = totalAvail + totalDown;
  const pctAvailability =
    totalHrs > 0 ? `${((totalAvail / totalHrs) * 100).toFixed(2)}%` : "";
  const mttr =
    sitesWithDown > 0 ? `${(totalDown / sitesWithDown).toFixed(2)} hrs` : "";
  const mtbf = totalDown > 0 ? `${(totalHrs / totalDown).toFixed(2)} hrs` : "";
  const mtbfNum = totalDown > 0 ? totalHrs / totalDown : null;
  const mttrNum = sitesWithDown > 0 ? totalDown / sitesWithDown : null;
  const mttf =
    mtbfNum !== null && mttrNum !== null
      ? `${(mtbfNum + mttrNum).toFixed(2)} hrs`
      : "";
  const totalDownRounded = Math.round(totalDown * 10) / 10;
  const tdMins = Math.round(totalDown * 60);
  const tdDays = Math.floor(tdMins / (60 * 24));
  const tdHrs = Math.floor((tdMins % (60 * 24)) / 60);
  const tdMin = Math.round(tdMins % 60);
  return {
    pctAvailability,
    mttr,
    mtbf,
    mttf,
    totalDown: totalDownRounded,
    totalAvail: Math.round(totalAvail * 10) / 10,
    affectedSites: sitesWithDown,
    nonAffectedSites,
    totalDownHrs: `${tdDays}d ${tdHrs}h ${tdMin}m`,
  };
}

async function exportPerfTemplate(
  rows: PerfRow[],
  monthKey: string,
  regions: string[],
) {
  const DATA_START = 7;
  const LAST_DATA = 40;
  const PROTECTED = 41;
  const TEMPLATE_N = LAST_DATA - DATA_START + 1;
  const OLD_SHEET = "EOA March 2026";
  const COLS = ["A", "B", "C", "D", "E", "F", "G", "H", "I"];

  const xmlEsc = (s: string) =>
    s
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");

  const setCell = (
    xml: string,
    ref: string,
    value: string | number,
  ): string => {
    const markerIdx = xml.indexOf(` r="${ref}"`);
    if (markerIdx === -1) return xml;
    const cStart = xml.lastIndexOf("<c", markerIdx);
    if (cStart === -1) return xml;
    const tagClose = xml.indexOf(">", cStart);
    if (tagClose === -1) return xml;
    const isSelfClosing = xml[tagClose - 1] === "/";
    let cEnd: number;
    if (isSelfClosing) {
      cEnd = tagClose + 1;
    } else {
      const fcIdx = xml.indexOf("</c>", tagClose);
      if (fcIdx === -1) return xml;
      cEnd = fcIdx + 4;
    }
    let attrs = xml
      .slice(cStart + 2, tagClose)
      .replace(/\s*\/$/, "")
      .replace(/\s+t="[^"]*"/g, "");
    const v = String(value ?? "");
    let cell: string;
    if (!v) cell = `<c${attrs}/>`;
    else if (typeof value === "number") cell = `<c${attrs}><v>${v}</v></c>`;
    else cell = `<c${attrs} t="inlineStr"><is><t>${xmlEsc(v)}</t></is></c>`;
    return xml.slice(0, cStart) + cell + xml.slice(cEnd);
  };

  const normalizeTableRows = (
    sheetXml: string,
    firstRow: number,
    lastRow: number,
  ) =>
    sheetXml.replace(/<row\b[^>]*>/g, (rowTag) => {
      const rowNumber = Number(rowTag.match(/\br="(\d+)"/)?.[1] ?? 0);
      if (rowNumber < firstRow || rowNumber > lastRow) return rowTag;
      return rowTag
        .replace(/\s+s="[^"]*"/g, "")
        .replace(/\s+customFormat="[^"]*"/g, "")
        .replace(/\s+ht="[^"]*"/g, "")
        .replace(/\s+customHeight="[^"]*"/g, "");
    });

  const setWorksheetColumns = (sheetXml: string): string => {
    const colsXml =
      "<cols>" +
      '<col min="1" max="1" width="8" customWidth="1"/>' +
      '<col min="2" max="2" width="14" customWidth="1"/>' +
      '<col min="3" max="3" width="34" customWidth="1"/>' +
      '<col min="4" max="4" width="20" customWidth="1"/>' +
      '<col min="5" max="5" width="28" customWidth="1"/>' +
      '<col min="6" max="6" width="20" customWidth="1"/>' +
      '<col min="7" max="7" width="24" customWidth="1"/>' +
      '<col min="8" max="8" width="18" customWidth="1"/>' +
      '<col min="9" max="9" width="16" customWidth="1"/>' +
      "</cols>";
    return sheetXml.includes("<cols>")
      ? sheetXml.replace(/<cols>[\s\S]*?<\/cols>/, colsXml)
      : sheetXml.replace("<sheetData>", `${colsXml}<sheetData>`);
  };

  const rewriteWorksheetRow = (
    sheetXml: string,
    rowNumber: number,
    rowXml: string,
  ) => {
    const rowPattern = new RegExp(
      `<row\\b[^>]*\\br="${rowNumber}"[^>]*>[\\s\\S]*?<\\/row>`,
    );
    if (rowPattern.test(sheetXml)) return sheetXml.replace(rowPattern, rowXml);
    const nextRowPattern = new RegExp(
      `<row\\b[^>]*\\br="${rowNumber + 1}"[^>]*>`,
    );
    const nextMatch = sheetXml.match(nextRowPattern);
    if (nextMatch?.index !== undefined)
      return `${sheetXml.slice(0, nextMatch.index)}${rowXml}${sheetXml.slice(nextMatch.index)}`;
    return sheetXml.replace("</sheetData>", `${rowXml}</sheetData>`);
  };

  const buildStyledCellXml = (
    col: string,
    rowNumber: number,
    value: string | number,
    styleIndex: number,
    numeric = false,
  ) => {
    if (numeric) {
      const num = Number(value);
      return `<c r="${col}${rowNumber}" s="${styleIndex}"><v>${Number.isFinite(num) ? num : 0}</v></c>`;
    }
    return `<c r="${col}${rowNumber}" s="${styleIndex}" t="inlineStr"><is><t>${xmlEsc(String(value ?? ""))}</t></is></c>`;
  };

  const stripTemplateTableFormatting = (sheetXml: string) =>
    sheetXml
      .replace(/<conditionalFormatting[\s\S]*?<\/conditionalFormatting>/g, "")
      .replace(/<extLst>[\s\S]*?<\/extLst>/g, "");

  try {
    const res = await fetch(
      publicWorkbookUrl("Network_Performance_Report.xlsx"),
      { cache: "no-store" },
    );
    if (!res.ok)
      throw new Error(
        `HTTP ${res.status} - place Network_Performance_Report.xlsx in your public/ folder`,
      );
    const rawBuf = await res.arrayBuffer();
    const { unzipSync, zipSync, strFromU8, strToU8 } = await import("fflate");
    const files = unzipSync(new Uint8Array(rawBuf));
    const exportRows = aggregatePerfRowsForReport(rows);

    const full = monthKey !== "all" ? formatMonthMMMMYYYY(monthKey) : "All";
    const parts = full.split(" ");
    const mmm =
      parts.length === 2 ? `${parts[0].slice(0, 3)}-${parts[1]}` : full;
    const regionPart = regions.length > 0 ? regions.join(", ") + " " : "";
    const sheetLabel = `${regionPart}${mmm}`
      .replace(/[\/*?[\]:]/g, "-")
      .slice(0, 31);

    const sheetKey = Object.keys(files).find((k) =>
      /^xl\/worksheets\/sheet\d+\.xml$/.test(k),
    )!;
    let xml = strFromU8(files[sheetKey]);

    const kpi = computePerfKPIs(exportRows);
    xml = setCell(xml, "C3", mmm);
    xml = setCell(xml, "D4", kpi.totalDownHrs);
    xml = setCell(xml, "E4", String(kpi.affectedSites));
    xml = setCell(xml, "F4", kpi.pctAvailability);
    xml = setCell(xml, "G4", kpi.mttr);
    xml = setCell(xml, "H4", kpi.mtbf);
    xml = setCell(xml, "I4", kpi.mttf);

    const needed = exportRows.length;
    if (needed > TEMPLATE_N) {
      const extra = needed - TEMPLATE_N;
      const tmplMarker = ` r="${DATA_START}" `;
      const tmplMarkerIdx = xml.indexOf(tmplMarker);
      const tmplRowStart =
        tmplMarkerIdx !== -1 ? xml.lastIndexOf("<row", tmplMarkerIdx) : -1;
      const tmplRowEnd =
        tmplRowStart !== -1 ? xml.indexOf("</row>", tmplRowStart) + 6 : -1;
      const tmplRow =
        tmplRowStart !== -1 && tmplRowEnd > 0
          ? xml.slice(tmplRowStart, tmplRowEnd)
          : "";

      xml = xml
        .replace(/(<row[^>]* r=")(\d+)(")/g, (_m, a, n, b) =>
          +n >= PROTECTED ? `${a}${+n + extra}${b}` : _m,
        )
        .replace(/(<c[^>]* r=")([A-Z]+)(\d+)(")/g, (_m, a, col, n, b) =>
          +n >= PROTECTED ? `${a}${col}${+n + extra}${b}` : _m,
        )
        .replace(
          /(<dimension ref="[A-Z]+\d+:[A-Z]+)(\d+)(")/,
          (_m, pre, n, post) =>
            +n >= PROTECTED ? `${pre}${+n + extra}${post}` : _m,
        );

      const sumOld = `SUM(D${DATA_START}:D${LAST_DATA})`;
      const sumNew = `SUM(D${DATA_START}:D${LAST_DATA + extra})`;
      xml = xml.split(sumOld).join(sumNew);

      if (tmplRow) {
        const newRows = Array.from({ length: extra }, (_, i) => {
          const rn = LAST_DATA + 1 + i;
          return tmplRow
            .replace(/ r="(\d+)"/, ` r="${rn}"`)
            .replace(
              /(<c[^>]* r=")([A-Z]+)\d+(")/g,
              (_m2, a, col, b) => `${a}${col}${rn}${b}`,
            )
            .replace(/<v>[^<]*<\/v>/g, "")
            .replace(/<is>[\s\S]*?<\/is>/g, "")
            .replace(/<f[^>]*\/>/g, "")
            .replace(/<f[^>]*>[\s\S]*?<\/f>/g, "")
            .replace(/\s+t="[^"]*"/g, "");
        }).join("");
        const shiftedTag = ` r="${PROTECTED + extra}" `;
        const shiftedIdx = xml.indexOf(shiftedTag);
        if (shiftedIdx !== -1) {
          const insertAt = xml.lastIndexOf("<row", shiftedIdx);
          xml = xml.slice(0, insertAt) + newRows + xml.slice(insertAt);
        } else {
          xml = xml.replace("</sheetData>", newRows + "</sheetData>");
        }
      }
    } else if (needed < TEMPLATE_N) {
      for (let r = DATA_START + needed; r <= LAST_DATA; r++) {
        COLS.forEach((col) => {
          xml = setCell(xml, `${col}${r}`, "");
        });
      }
    }

    xml = xml.split(' t="shared"').join("");
    {
      let tmp = "",
        src = xml;
      while (true) {
        const idx = src.indexOf(' si="');
        if (idx === -1) {
          tmp += src;
          break;
        }
        const eq = src.indexOf('"', idx + 5);
        if (eq === -1) {
          tmp += src;
          break;
        }
        tmp += src.slice(0, idx);
        src = src.slice(eq + 1);
      }
      xml = tmp;
    }
    {
      let result = "",
        src = xml;
      while (true) {
        const fi = src.indexOf("<f ");
        if (fi === -1) {
          result += src;
          break;
        }
        const fe = src.indexOf(">", fi);
        if (fe === -1) {
          result += src;
          break;
        }
        const fTag = src.slice(fi, fe + 1);
        const ri = fTag.indexOf(' ref="');
        if (ri !== -1) {
          const re = fTag.indexOf('"', ri + 6) + 1;
          result += src.slice(0, fi) + fTag.slice(0, ri) + fTag.slice(re);
        } else {
          result += src.slice(0, fe + 1);
        }
        src = src.slice(fe + 1);
      }
      xml = result;
    }
    if (needed > TEMPLATE_N) {
      const extra2 = needed - TEMPLATE_N;
      xml = xml.replace(
        /(<f[^>]*>)([\s\S]*?)(<\/f>)/g,
        (_m, open, body, close) => {
          const updatedBody = body.replace(
            /(\$?)([A-J])(\$?)(\d+)/g,
            (
              ref: string,
              absCol: string,
              col: string,
              absRow: string,
              rowText: string,
            ) => {
              const rowNum = Number(rowText);
              return rowNum >= PROTECTED
                ? `${absCol}${col}${absRow}${rowNum + extra2}`
                : ref;
            },
          );
          return `${open}${updatedBody}${close}`;
        },
      );
    }

    exportRows.forEach((row, i) => {
      const r = DATA_START + i;
      const availDays =
        row.availDay || String(Math.round((row.availHours / 24) * 10) / 10);
      xml = setCell(xml, `A${r}`, i + 1);
      xml = setCell(xml, `B${r}`, row.siteId);
      xml = setCell(xml, `C${r}`, row.siteName);
      xml = setCell(xml, `D${r}`, row.availHours);
      xml = setCell(xml, `E${r}`, availDays);
      xml = setCell(xml, `F${r}`, row.channelBusy);
      xml = setCell(xml, `G${r}`, row.mwLinkPerf);
      xml = setCell(xml, `H${r}`, row.reliability);
      xml = setCell(xml, `I${r}`, row.sitesDownHours);
    });

    const headerStyle =
      getExcelCellStyle(xml, `${COLS[0]}${DATA_START - 1}`) || 1;
    const serialNoStyle = getExcelCellStyle(xml, `A${DATA_START}`) || 2;
    const bodyTextStyle = getExcelCellStyle(xml, `B${DATA_START}`) || 23;
    const bodyNumberStyle = getExcelCellStyle(xml, `D${DATA_START}`) || 22;

    const headerRowXml = `<row r="${DATA_START - 1}">${PERF_TEMPLATE_PDF_HEADERS.map(
      (header, index) =>
        buildStyledCellXml(COLS[index], DATA_START - 1, header, headerStyle),
    ).join("")}</row>`;
    xml = rewriteWorksheetRow(xml, DATA_START - 1, headerRowXml);

    exportRows.forEach((row, index) => {
      const r = DATA_START + index;
      const availDays =
        row.availDay || String(Math.round((row.availHours / 24) * 10) / 10);
      const rowXml =
        `<row r="${r}">` +
        buildStyledCellXml("A", r, index + 1, serialNoStyle, true) +
        buildStyledCellXml("B", r, row.siteId, bodyTextStyle) +
        buildStyledCellXml("C", r, row.siteName, bodyTextStyle) +
        buildStyledCellXml("D", r, row.availHours, bodyNumberStyle, true) +
        buildStyledCellXml("E", r, availDays, bodyTextStyle) +
        buildStyledCellXml("F", r, row.channelBusy, bodyNumberStyle, true) +
        buildStyledCellXml("G", r, row.mwLinkPerf, bodyNumberStyle, true) +
        buildStyledCellXml("H", r, row.reliability, bodyTextStyle) +
        buildStyledCellXml("I", r, row.sitesDownHours, bodyNumberStyle, true) +
        "</row>";
      xml = rewriteWorksheetRow(xml, r, rowXml);
    });

    xml = normalizeTableRows(
      xml,
      DATA_START - 1,
      DATA_START + Math.max(exportRows.length, 1) - 1,
    );
    xml = setWorksheetColumns(xml);
    xml = stripTemplateTableFormatting(xml);
    xml = xml.replace(/<f\b[^>]*\/>|<f\b[^>]*>[\s\S]*?<\/f>/g, "");
    files[sheetKey] = strToU8(xml);

    const wbKey = "xl/workbook.xml";
    if (files[wbKey]) {
      let wbXml = strFromU8(files[wbKey]);
      wbXml = wbXml
        .split(`name="${OLD_SHEET}"`)
        .join(`name="${xmlEsc(sheetLabel)}"`);
      wbXml = wbXml.split(`'${OLD_SHEET}'!`).join(`'${sheetLabel}'!`);
      files[wbKey] = strToU8(wbXml);
    }

    const lastRow = DATA_START + Math.max(exportRows.length, 1) - 1;
    const chartKeys = Object.keys(files).filter(
      (k) => k.startsWith("xl/charts/chart") && k.endsWith(".xml"),
    );

    const flattenRange = (cxml: string, col: string): string => {
      const openTag = "<c:f>(",
        closeTag = ")</c:f>";
      let result = cxml,
        searchFrom = 0;
      while (true) {
        const start = result.indexOf(openTag, searchFrom);
        if (start === -1) break;
        const end = result.indexOf(closeTag, start);
        if (end === -1) break;
        const inner = result.slice(start, end + closeTag.length);
        if (inner.includes(`$${col}$`)) {
          const simple = `<c:f>'${sheetLabel}'!$${col}$${DATA_START}:'${sheetLabel}'!$${col}$${lastRow}</c:f>`;
          result =
            result.slice(0, start) +
            simple +
            result.slice(end + closeTag.length);
        }
        searchFrom = start + 1;
      }
      return result;
    };

    chartKeys.forEach((ck) => {
      let cxml = strFromU8(files[ck]);
      cxml = cxml.split(`'${OLD_SHEET}'!`).join(`'${sheetLabel}'!`);
      ["B", "D", "F", "G", "I"].forEach((col) => {
        cxml = cxml.split(`$${col}$${LAST_DATA}`).join(`$${col}$${lastRow}`);
      });
      cxml = flattenRange(cxml, "B");
      cxml = flattenRange(cxml, "G");
      files[ck] = strToU8(cxml);
    });

    const drawingKeys = Object.keys(files).filter(
      (k) => k.startsWith("xl/drawings/drawing") && k.endsWith(".xml"),
    );
    const chartStartRow =
      DATA_START + Math.max(exportRows.length, TEMPLATE_N) + 4;
    const chartLayout: Record<
      string,
      { fromCol: number; fromRow: number; toCol: number; toRow: number }
    > = {
      rId6: {
        fromCol: 0,
        fromRow: chartStartRow,
        toCol: 9,
        toRow: chartStartRow + 20,
      },
      rId3: {
        fromCol: 0,
        fromRow: chartStartRow + 22,
        toCol: 9,
        toRow: chartStartRow + 38,
      },
      rId4: {
        fromCol: 0,
        fromRow: chartStartRow + 40,
        toCol: 4,
        toRow: chartStartRow + 54,
      },
      rId5: {
        fromCol: 5,
        fromRow: chartStartRow + 40,
        toCol: 9,
        toRow: chartStartRow + 54,
      },
    };
    drawingKeys.forEach((dk) => {
      let dxml = strFromU8(files[dk]);
      dxml = dxml.replace(
        /<xdr:twoCellAnchor[\s\S]*?<\/xdr:twoCellAnchor>/g,
        (anchor) => {
          const relId = anchor.match(/r:id="(rId\d+)"/)?.[1];
          const layout = relId ? chartLayout[relId] : null;
          if (!layout) return anchor;
          const replaceMarker = (
            part: string,
            marker: "from" | "to",
            col: number,
            row: number,
          ) =>
            part.replace(
              new RegExp(
                `(<xdr:${marker}>[\\s\\S]*?<xdr:col>)\\d+(<\\/xdr:col>[\\s\\S]*?<xdr:row>)\\d+(<\\/xdr:row>)`,
              ),
              `$1${col}$2${row}$3`,
            );
          let updated = replaceMarker(
            anchor,
            "from",
            layout.fromCol,
            layout.fromRow,
          );
          updated = replaceMarker(updated, "to", layout.toCol, layout.toRow);
          return updated;
        },
      );
      files[dk] = strToU8(dxml);
    });

    delete files["xl/calcChain.xml"];
    if (files["[Content_Types].xml"]) {
      let ct = strFromU8(files["[Content_Types].xml"]);
      ct = ct
        .split(
          `<Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>`,
        )
        .join("");
      files["[Content_Types].xml"] = strToU8(ct);
    }

    const output = zipSync(files, { level: 0 });
    const blob = new Blob([output], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `Network_Performance_${sheetLabel}.xlsx`;
    a.click();
    URL.revokeObjectURL(url);
  } catch (err: any) {
    console.error("Perf template export failed:", err);
    alert(
      "Performance template export failed.\n\nMake sure Network_Performance_Report.xlsx is in your public/ folder.\n\nError: " +
        (err?.message ?? String(err)),
    );
  }
}

function exportPerfCsv(rows: PerfRow[], monthKey: string) {
  const monthLabel = monthKey !== "all" ? formatMonthMMMMYYYY(monthKey) : "All";
  const kpi = computePerfKPIs(rows);
  const kpiRows: string[][] = [
    [],
    ["KPI Summary"],
    ["% Availability", kpi.pctAvailability],
    ["MTTR", kpi.mttr],
    ["MTBF", kpi.mtbf],
    ["MTTF", kpi.mttf],
    ["No. of Affected Sites", String(kpi.affectedSites)],
    ["Total Down Duration", kpi.totalDownHrs],
  ];
  const availChartRows: string[][] = [
    [],
    ["Site Availability Chart Data"],
    ["Site Name", "Availability (Hrs)"],
    ...rows.map((r) => [r.siteName, String(r.availHours)]),
  ];
  const downtimeChartRows: string[][] = [
    [],
    ["Site Downtime Chart Data"],
    ["Site Name", "Down (Hrs)", "Status"],
    ...rows.map((r) => [
      r.siteName,
      String(r.sitesDownHours),
      r.sitesDownHours === 0
        ? "No Downtime"
        : r.sitesDownHours > 24
          ? "Critical (>24h)"
          : r.sitesDownHours > 8
            ? "High (8-24h)"
            : "Moderate (<8h)",
    ]),
  ];
  const csv = [
    PERF_REPORT_HEADERS,
    ...perfReportRows(rows),
    ...kpiRows,
    ...availChartRows,
    ...downtimeChartRows,
  ]
    .map((line) =>
      line
        .map((cell) => `"${String(cell ?? "").replace(/"/g, '""')}"`)
        .join(","),
    )
    .join("\n");
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `DMR-Monthly-Performance-${monthLabel}.csv`;
  link.click();
  URL.revokeObjectURL(url);
}

function exportPerfExcel(rows: PerfRow[], monthKey: string) {
  const monthLabel = monthKey !== "all" ? formatMonthMMMMYYYY(monthKey) : "All";
  const kpi = computePerfKPIs(rows);
  const kpiRows: string[][] = [
    [],
    ["KPI Summary"],
    ["% Availability", kpi.pctAvailability],
    ["MTTR", kpi.mttr],
    ["MTBF", kpi.mtbf],
    ["MTTF", kpi.mttf],
    ["No. of Affected Sites", String(kpi.affectedSites)],
    ["Total Down Duration", kpi.totalDownHrs],
  ];
  const worksheet = XLSX.utils.aoa_to_sheet([
    PERF_REPORT_HEADERS,
    ...perfReportRows(rows),
    ...kpiRows,
  ]);
  PERF_REPORT_HEADERS.forEach((_, colIndex) => {
    const horizontal = perfExportLeftIndexes.has(colIndex) ? "left" : "center";
    for (let rowIndex = 0; rowIndex <= rows.length; rowIndex++) {
      applySheetCellAlignment(
        worksheet,
        XLSX.utils.encode_cell({ r: rowIndex, c: colIndex }),
        horizontal,
      );
    }
  });
  worksheet["!cols"] = [
    { wch: 6 },
    { wch: 14 },
    { wch: 30 },
    { wch: 24 },
    { wch: 26 },
    { wch: 20 },
    { wch: 26 },
    { wch: 18 },
    { wch: 18 },
  ];
  const availSheet = XLSX.utils.aoa_to_sheet([
    ["Site Name", "Availability (Hrs)", "Down (Hrs)", "Reliability (%)"],
    ...rows.map((r) => [
      r.siteName,
      r.availHours,
      r.sitesDownHours,
      parseFloat(r.reliability) || 100,
    ]),
  ]);
  availSheet["!cols"] = [{ wch: 32 }, { wch: 20 }, { wch: 16 }, { wch: 18 }];
  const downtimeSorted = [...rows].sort(
    (a, b) => b.sitesDownHours - a.sitesDownHours,
  );
  const downtimeSheet = XLSX.utils.aoa_to_sheet([
    [
      "Site Name",
      "Down (Hrs)",
      "Availability (Hrs)",
      "Reliability (%)",
      "Status",
    ],
    ...downtimeSorted.map((r) => [
      r.siteName,
      r.sitesDownHours,
      r.availHours,
      parseFloat(r.reliability) || 100,
      r.sitesDownHours === 0
        ? "No Downtime"
        : r.sitesDownHours > 24
          ? "Critical (>24h)"
          : r.sitesDownHours > 8
            ? "High (8-24h)"
            : "Moderate (<8h)",
    ]),
  ]);
  downtimeSheet["!cols"] = [
    { wch: 32 },
    { wch: 14 },
    { wch: 20 },
    { wch: 18 },
    { wch: 20 },
  ];
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Monthly Performance");
  XLSX.utils.book_append_sheet(workbook, availSheet, "Availability Chart Data");
  XLSX.utils.book_append_sheet(workbook, downtimeSheet, "Downtime Chart Data");
  XLSX.writeFile(workbook, `DMR-Monthly-Performance-${monthLabel}.xlsx`);
}

async function exportPerfPdf(rows: PerfRow[], monthKey: string) {
  const doc = new jsPDF({ orientation: "landscape", unit: "mm", format: "a3" });
  const monthLabel =
    monthKey !== "all" ? formatMonthMMMMYYYY(monthKey) : "All Months";
  const pageW = doc.internal.pageSize.getWidth();
  const pageH = doc.internal.pageSize.getHeight();
  {
    const logos = await loadReportLogos();
    const templateDoc = new jsPDF({
      orientation: "landscape",
      unit: "mm",
      format: "a4",
    });
    const templateMonthLabel =
      monthKey !== "all" ? formatMonthMMMMYYYY(monthKey) : "All";
    const templatePageW = templateDoc.internal.pageSize.getWidth();
    const templatePageH = templateDoc.internal.pageSize.getHeight();
    const C = TEMPLATE_PDF_COLORS;
    const kpi = computePerfKPIs(rows);
    const perfTableMargin = Math.max(10, (templatePageW - 251) / 2);

    const drawTemplateHeader = (data?: { pageNumber?: number }) => {
      const pageNumber = data?.pageNumber ?? 1;
      templateDoc.setFillColor(255, 255, 255);
      templateDoc.rect(0, 0, templatePageW, templatePageH, "F");
      drawPdfReportHeader(
        templateDoc,
        templatePageW,
        "Network Performance",
        `DMR Hytera | ${templateMonthLabel}`,
        logos,
        C.title,
      );
      templateDoc.setFillColor(...C.band);
      templateDoc.rect(10, 23, templatePageW - 20, 8, "F");
      templateDoc.setTextColor(...C.bandText);
      templateDoc.setFontSize(9);
      templateDoc.text(`Network Performance: ${templateMonthLabel}`, 13, 28.5);
      if (pageNumber !== 1) return;
      const labels = [
        "Total Downtime ,hrs",
        "Total Sites Affected",
        "% Availability",
        "MTTR",
        "MTBF",
        "MTTF",
      ];
      const values = [
        kpi.totalDownHrs,
        String(kpi.affectedSites),
        kpi.pctAvailability,
        kpi.mttr,
        kpi.mtbf,
        kpi.mttf,
      ];
      const cellW = (templatePageW - 20) / labels.length;
      labels.forEach((label, i) => {
        const x = 10 + i * cellW;
        templateDoc.setDrawColor(...C.border);
        templateDoc.setLineWidth(0.15);
        templateDoc.setFillColor(...C.header);
        templateDoc.rect(x, 35, cellW, 6, "FD");
        templateDoc.setFillColor(255, 255, 255);
        templateDoc.rect(x, 41, cellW, 7, "FD");
        templateDoc.setFontSize(7);
        templateDoc.setTextColor(...C.text);
        templateDoc.setFont("helvetica", "bold");
        templateDoc.text(label, x + cellW / 2, 39, { align: "center" });
        templateDoc.setFont("helvetica", "normal");
        templateDoc.text(values[i], x + cellW / 2, 45.5, { align: "center" });
      });
    };

    autoTable(templateDoc, {
      startY: 54,
      head: [PERF_TEMPLATE_PDF_HEADERS],
      body: perfReportRows(rows),
      theme: "grid",
      styles: {
        fontSize: 7,
        cellPadding: 1.5,
        overflow: "linebreak",
        halign: "center",
        valign: "middle",
        textColor: C.text,
        fillColor: [255, 255, 255],
        lineColor: C.border,
        lineWidth: 0.15,
      },
      headStyles: {
        fillColor: C.header,
        textColor: C.text,
        fontStyle: "bold",
        fontSize: 7.2,
        cellPadding: 1.6,
        halign: "center",
      },
      alternateRowStyles: { fillColor: [248, 248, 248] },
      columnStyles: {
        0: { cellWidth: 10, halign: "center" },
        1: { cellWidth: 24, halign: "left" },
        2: { cellWidth: 44, halign: "left" },
        3: { cellWidth: 30, halign: "center" },
        4: { cellWidth: 32, halign: "center" },
        5: { cellWidth: 25, halign: "center" },
        6: { cellWidth: 33, halign: "center" },
        7: { cellWidth: 28, halign: "center", fontStyle: "bold" },
        8: { cellWidth: 25, halign: "center", fontStyle: "bold" },
      },
      margin: {
        left: perfTableMargin,
        right: perfTableMargin,
        top: 54,
        bottom: 10,
      },
      willDrawPage: drawTemplateHeader,
      didParseCell: (d) => {
        d.cell.styles.halign =
          d.column.index === 1 || d.column.index === 2 ? "left" : "center";
        d.cell.styles.valign = "middle";
        if (d.section === "body" && d.column.index === 7) {
          const v = parseFloat(String(d.cell.raw ?? "100"));
          d.cell.styles.textColor = v < 95 ? C.warn : C.ok;
        }
        if (d.section === "body" && d.column.index === 8) {
          const v = parseFloat(String(d.cell.raw ?? "0"));
          d.cell.styles.textColor = v > 0 ? C.warn : C.ok;
        }
      },
      didDrawPage: (d) => {
        templateDoc.setFont("helvetica", "normal");
        templateDoc.setFontSize(7);
        templateDoc.setTextColor(...C.muted);
        templateDoc.text(
          `Network Performance - ${templateMonthLabel} | Page ${d.pageNumber}`,
          templatePageW / 2,
          templatePageH - 5,
          { align: "center" },
        );
      },
    });

    const drawPerfChartPage = (
      title: string,
      chartRows: PerfRow[],
      valueKey: "availHours" | "sitesDownHours",
      color: [number, number, number],
    ) => {
      templateDoc.addPage();
      templateDoc.setFillColor(255, 255, 255);
      templateDoc.rect(0, 0, templatePageW, templatePageH, "F");
      drawPdfReportHeader(
        templateDoc,
        templatePageW,
        title,
        `Network Performance | ${templateMonthLabel}`,
        logos,
        C.title,
      );
      const chartW = 245,
        chartH = 115,
        chartX = (templatePageW - chartW) / 2,
        chartY = 42;
      const items = chartRows.slice(0, 34);
      const maxValue = Math.max(
        ...items.map((row) => Number(row[valueKey]) || 0),
        1,
      );
      templateDoc.setDrawColor(...C.border);
      templateDoc.setLineWidth(0.2);
      templateDoc.rect(chartX, chartY, chartW, chartH);
      for (let i = 1; i <= 4; i++) {
        const y = chartY + (chartH / 5) * i;
        templateDoc.setDrawColor(220, 220, 220);
        templateDoc.line(chartX, y, chartX + chartW, y);
      }
      const gap = chartW / Math.max(items.length, 1);
      const barW = Math.min(8, gap * 0.58);
      items.forEach((row, index) => {
        const value = Number(row[valueKey]) || 0;
        const barH = (value / maxValue) * (chartH - 16);
        const x = chartX + index * gap + (gap - barW) / 2;
        const y = chartY + chartH - barH - 8;
        templateDoc.setFillColor(...color);
        if (barH > 0) templateDoc.rect(x, y, barW, barH, "F");
        templateDoc.setFontSize(5);
        templateDoc.setTextColor(...C.muted);
        const label = row.siteName || row.siteId;
        templateDoc.text(
          label.length > 10 ? `${label.slice(0, 9)}...` : label,
          x + barW / 2,
          chartY + chartH + 6,
          { align: "center", angle: 45 },
        );
      });
      templateDoc.setFont("helvetica", "normal");
      templateDoc.setFontSize(7);
      templateDoc.setTextColor(...C.muted);
      templateDoc.text(
        `Network Performance - ${templateMonthLabel} | Page ${templateDoc.getNumberOfPages()}`,
        templatePageW / 2,
        templatePageH - 5,
        { align: "center" },
      );
    };

    drawPerfChartPage("Site Availability Chart", rows, "availHours", C.ok);
    drawPerfChartPage(
      "Site Downtime Chart",
      [...rows].sort((a, b) => b.sitesDownHours - a.sitesDownHours),
      "sitesDownHours",
      C.warn,
    );
    templateDoc.save(
      `DMR-Monthly-Performance-${templateMonthLabel.replace(/ /g, "-")}.pdf`,
    );
    return;
  }
}

// â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”
// END OF PART 1 â€” continue in Part 2
// â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”
// â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”
// PART 2 â€” paste directly after the last line of Part 1
// â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”

function uniqueTicketValues(
  ticket: TicketAggregate,
  field: keyof TicketRecord,
): string {
  return Array.from(
    new Set(ticket.rows.map((row) => clean(row[field])).filter(Boolean)),
  ).join(", ");
}

function distinctReportRow(ticket: TicketAggregate, index: number): string[] {
  const row = ticket.primary;
  const rca = row.rca || uniqueTicketValues(ticket, "rca") || "";
  const actionTaken =
    row.actionTaken || uniqueTicketValues(ticket, "actionTaken") || "";
  return [
    String(index + 1),
    Array.from(ticket.siteIds).join(", "),
    Array.from(ticket.siteNames).join(", "),
    uniqueTicketValues(ticket, "managedResource") || row.managedResource || "",
    row.severity || "",
    row.issue || "",
    row.observationDate || "",
    row.observationTime || "",
    row.recoveryDate || "",
    row.recoveryTime || "",
    row.escalatedForL3SupportDate || "",
    row.escalatedForL3SupportTime || "",
    row.duration || "",
    ticket.tt || "",
    row.status || "",
    row.escalatedTo || "",
    rca,
    actionTaken,
  ];
}

function distinctReportRows(rows: TicketAggregate[]): string[][] {
  return rows.map((ticket, index) => distinctReportRow(ticket, index));
}

function exportCsv(rows: TicketAggregate[]) {
  const csv = [DISTINCT_REPORT_HEADERS, ...distinctReportRows(rows)]
    .map((line) =>
      line
        .map((cell) => `"${String(cell ?? "").replace(/"/g, '""')}"`)
        .join(","),
    )
    .join("\n");
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = "follow-up-distinct-tt-report.csv";
  link.click();
  URL.revokeObjectURL(url);
}

function exportExcel(rows: TicketAggregate[]) {
  const worksheet = XLSX.utils.aoa_to_sheet([
    DISTINCT_REPORT_HEADERS,
    ...distinctReportRows(rows),
  ]);
  DISTINCT_REPORT_HEADERS.forEach((_, colIndex) => {
    const horizontal = ticketExportCenteredIndexes.has(colIndex)
      ? "center"
      : "left";
    for (let rowIndex = 0; rowIndex <= rows.length; rowIndex++) {
      applySheetCellAlignment(
        worksheet,
        XLSX.utils.encode_cell({ r: rowIndex, c: colIndex }),
        horizontal,
      );
    }
  });
  worksheet["!cols"] = [
    { wch: 6 },
    { wch: 24 },
    { wch: 28 },
    { wch: 30 },
    { wch: 12 },
    { wch: 32 },
    { wch: 16 },
    { wch: 16 },
    { wch: 16 },
    { wch: 16 },
    { wch: 26 },
    { wch: 26 },
    { wch: 24 },
    { wch: 18 },
    { wch: 14 },
    { wch: 20 },
    { wch: 52 },
    { wch: 34 },
  ];
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Distinct TT Report");
  XLSX.writeFile(workbook, "follow-up-distinct-tt-report.xlsx");
}

function exportAnalyticsTableExcel(
  headers: string[],
  rows: Array<Array<string | number>>,
  sheetName: string,
  fileName: string,
) {
  const worksheet = XLSX.utils.aoa_to_sheet([headers, ...rows]);
  const columnWidths = headers.map((header, colIndex) => {
    const maxCellLength = Math.max(
      header.length,
      ...rows.map((row) => String(row[colIndex] ?? "").length),
    );
    return { wch: Math.min(Math.max(maxCellLength + 2, 12), 42) };
  });
  worksheet["!cols"] = columnWidths;
  const range = XLSX.utils.decode_range(worksheet["!ref"] ?? "A1:A1");
  for (let r = range.s.r; r <= range.e.r; r++) {
    for (let c = range.s.c; c <= range.e.c; c++) {
      applySheetCellAlignment(
        worksheet,
        XLSX.utils.encode_cell({ r, c }),
        "center",
      );
    }
  }
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
  XLSX.writeFile(workbook, fileName);
}

function exportAnalyticsWorkbookExcel(
  sheets: Array<{
    name: string;
    headers: string[];
    rows: Array<Array<string | number>>;
  }>,
  fileName: string,
) {
  const workbook = XLSX.utils.book_new();
  sheets.forEach((sheet) => {
    const worksheet = XLSX.utils.aoa_to_sheet([sheet.headers, ...sheet.rows]);
    worksheet["!cols"] = sheet.headers.map((header, colIndex) => {
      const maxCellLength = Math.max(
        header.length,
        ...sheet.rows.map((row) => String(row[colIndex] ?? "").length),
      );
      return { wch: Math.min(Math.max(maxCellLength + 2, 12), 44) };
    });
    const range = XLSX.utils.decode_range(worksheet["!ref"] ?? "A1:A1");
    for (let r = range.s.r; r <= range.e.r; r++) {
      for (let c = range.s.c; c <= range.e.c; c++) {
        applySheetCellAlignment(
          worksheet,
          XLSX.utils.encode_cell({ r, c }),
          "center",
        );
      }
    }
    XLSX.utils.book_append_sheet(workbook, worksheet, sheet.name);
  });
  XLSX.writeFile(workbook, fileName);
}

// â”€â”€â”€ MODIFIED exportTicketTemplate â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
async function exportTicketTemplate(
  tickets: TicketAggregate[],
  monthKey: string,
) {
  const normalizeTicketRegionLabel = (region: string) => {
    const value = clean(region).toUpperCase();
    if (value === "EOA" || value === "NEOA") return "EOA";
    if (value === "SOA") return "SOA";
    if (value === "COA") return "COA";
    if (value === "WOA") return "WOA";
    return value || "EOA";
  };

  const ticketRegionSet = new Set(
    tickets
      .map((ticket) => normalizeTicketRegionLabel(ticket.primary.region))
      .filter(Boolean),
  );
  const templateRegionOrder = ["EOA", "SOA", "COA", "WOA"] as const;
  if (ticketRegionSet.size > 1) {
    for (const region of templateRegionOrder) {
      const regionTickets = tickets.filter(
        (ticket) =>
          normalizeTicketRegionLabel(ticket.primary.region) === region,
      );
      if (regionTickets.length) {
        await exportTicketTemplate(regionTickets, monthKey);
      }
    }
    return;
  }
  const templateKind =
    templateRegionOrder.find((region) => ticketRegionSet.has(region)) ?? "EOA";
  const templateConfig = {
    EOA: {
      url: "/EOA_DMR_Monthly_Report.xlsx",
      dataStart: 39,
      protectedRow: 48,
    },
    SOA: {
      url: "/SOA_DMR_Monthly_Report.xlsx",
      dataStart: 41,
      protectedRow: 58,
    },
    COA: {
      url: "/COA_DMR_Monthly_Report.xlsx",
      dataStart: 32,
      protectedRow: 48,
    },
    WOA: {
      url: "/WOA_DMR_Monthly_Report.xlsx",
      dataStart: 32,
      protectedRow: 35,
    },
  }[templateKind];

  const DATA_START = templateConfig.dataStart;
  const PROTECTED = templateConfig.protectedRow;
  const AVAIL = PROTECTED - DATA_START;
  const COLS = [
    "A",
    "B",
    "C",
    "D",
    "E",
    "F",
    "G",
    "H",
    "I",
    "J",
    "K",
    "L",
    "M",
    "N",
    "O",
    "P",
    "Q",
  ];

  // â”€â”€ XML helpers â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
  const xmlEsc = (s: string) =>
    s
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");

  const setCell = (
    xml: string,
    ref: string,
    value: string | number,
  ): string => {
    const markerIdx = xml.indexOf(` r="${ref}"`);
    if (markerIdx === -1) return xml;
    const cStart = xml.lastIndexOf("<c", markerIdx);
    if (cStart === -1) return xml;
    const scIdx = xml.indexOf("/>", cStart);
    const fcIdx = xml.indexOf("</c>", cStart);
    const cEnd =
      scIdx !== -1 && (fcIdx === -1 || scIdx < fcIdx) ? scIdx + 2 : fcIdx + 4;
    const tagClose = xml.indexOf(">", cStart);
    let attrs = xml
      .slice(cStart + 2, tagClose)
      .replace(/\s*\/$/, "")
      .replace(/\s+t="[^"]*"/g, "");
    const v = String(value ?? "");
    let newCell: string;
    if (!v) newCell = `<c${attrs}/>`;
    else if (typeof value === "number") newCell = `<c${attrs}><v>${v}</v></c>`;
    else newCell = `<c${attrs} t="inlineStr"><is><t>${xmlEsc(v)}</t></is></c>`;
    return xml.slice(0, cStart) + newCell + xml.slice(cEnd);
  };

  const setWorksheetColumns = (
    sheetXml: string,
    valuesByCol: Record<string, string[]>,
  ): string => {
    const headerByCol: Record<string, string> = {
      A: "No",
      B: "Equipment/ site",
      C: "Site Name",
      D: "Effected Managed Resource",
      E: "Severity",
      F: "Escalation Date",
      G: "Escalation Time",
      H: "Recovery Date",
      I: "Recovery Time",
      J: "Escalated for L3 Support Date",
      K: "Escalated for L3 Support Time",
      L: "Outage Duration",
      M: "TT Number",
      N: "TT Status",
      O: "TT Owner",
      P: "RCA",
      Q: "Action Taken",
    };
    const widthFor = (col: string) => {
      const values = [headerByCol[col] ?? "", ...(valuesByCol[col] ?? [])];
      const longest = values.reduce((max, value) => {
        const lineMax = String(value ?? "")
          .split(/\r?\n/)
          .reduce((lm, line) => Math.max(lm, line.length), 0);
        return Math.max(max, lineMax);
      }, 0);
      return Math.min(Math.max(longest + 3, 8), 80);
    };
    const colsXml =
      "<cols>" +
      COLS.map((col, index) => {
        const columnNumber = index + 1;
        return `<col min="${columnNumber}" max="${columnNumber}" width="${widthFor(col)}" customWidth="1" bestFit="1"/>`;
      }).join("") +
      "</cols>";
    return sheetXml.includes("<cols>")
      ? sheetXml.replace(/<cols>[\s\S]*?<\/cols>/, colsXml)
      : sheetXml.replace("<sheetData>", `${colsXml}<sheetData>`);
  };

  const rewriteWorksheetRow = (
    sheetXml: string,
    rowNumber: number,
    rowXml: string,
  ) => {
    const rowPattern = new RegExp(
      `<row\\b[^>]*\\br="${rowNumber}"[^>]*>[\\s\\S]*?<\\/row>`,
    );
    if (rowPattern.test(sheetXml)) return sheetXml.replace(rowPattern, rowXml);
    const nextRowPattern = new RegExp(
      `<row\\b[^>]*\\br="${rowNumber + 1}"[^>]*>`,
    );
    const nextMatch = sheetXml.match(nextRowPattern);
    if (nextMatch?.index !== undefined)
      return `${sheetXml.slice(0, nextMatch.index)}${rowXml}${sheetXml.slice(nextMatch.index)}`;
    return sheetXml.replace("</sheetData>", `${rowXml}</sheetData>`);
  };

  // â”€â”€ MODIFIED: plain cell writer â€” inlineStr for all non-numeric values â”€â”€â”€â”€â”€
  // This prevents Excel from inheriting the number format / fill color from the
  // row above (the template header row).  Every data cell is written as a plain
  // inline string so the source text is rendered exactly as-is.
  const buildStyledCellXml = (
    col: string,
    rowNumber: number,
    value: string | number,
    styleIndex: number,
    numeric = false,
  ) => {
    if (numeric) {
      const num = Number(value);
      return `<c r="${col}${rowNumber}" s="${styleIndex}"><v>${Number.isFinite(num) ? num : 0}</v></c>`;
    }
    const str = String(value ?? "").trim();
    if (!str) {
      // Truly empty â€” self-closing, no <v>, no type
      return `<c r="${col}${rowNumber}" s="${styleIndex}"/>`;
    }
    // inlineStr: Excel renders the literal string; no date-serial or number coercion
    const escaped = xmlEsc(str).replace(/\n/g, "&#10;");
    return (
      `<c r="${col}${rowNumber}" s="${styleIndex}" t="inlineStr">` +
      `<is><t xml:space="preserve">${escaped}</t></is></c>`
    );
  };

  const normalizeTicketTableRows = (
    sheetXml: string,
    firstRow: number,
    lastRow: number,
  ) =>
    sheetXml.replace(/<row\b[^>]*>/g, (rowTag) => {
      const rowNumber = Number(rowTag.match(/\br="(\d+)"/)?.[1] ?? 0);
      if (rowNumber < firstRow || rowNumber > lastRow) return rowTag;
      return rowTag
        .replace(/\s+s="[^"]*"/g, "")
        .replace(/\s+customFormat="[^"]*"/g, "");
    });

  const removeTicketSiteNameMerges = (
    sheetXml: string,
    firstRow: number,
    lastRow: number,
  ) => {
    let removed = 0;
    const nextXml = sheetXml.replace(
      /<mergeCell ref="C(\d+):D\1"\/>/g,
      (match, rowNumber) => {
        const row = Number(rowNumber);
        if (row < firstRow || row > lastRow) return match;
        removed += 1;
        return "";
      },
    );
    if (!removed) return nextXml;
    return nextXml.replace(
      /<mergeCells\b([^>]*)count="(\d+)"([^>]*)>/,
      (_match, before, count, after) =>
        `<mergeCells${before}count="${Math.max(0, Number(count) - removed)}"${after}>`,
    );
  };

  const ticketSiteLinesForTemplate = (ticket: TicketAggregate) => {
    const bySite = new Map<string, { siteId: string; siteName: string }>();
    ticket.rows.forEach((row) => {
      const siteId = normalizeSiteId(clean(row.siteId));
      if (!siteId) return;
      const key = siteId.toUpperCase();
      if (!bySite.has(key))
        bySite.set(key, { siteId, siteName: clean(row.siteName) });
    });
    const rows = Array.from(bySite.values());
    const rfRows = rows.filter((row) => isRfSiteId(row.siteId));
    const rowsToUse = rfRows.length ? rfRows : rows;
    if (!rowsToUse.length) {
      return {
        siteIds: Array.from(ticket.siteIds).join("\n"),
        siteNames: Array.from(ticket.siteNames).join("\n"),
      };
    }
    return {
      siteIds: rowsToUse.map((row) => row.siteId).join("\n"),
      siteNames: rowsToUse.map((row) => row.siteName).join("\n"),
    };
  };

  try {
    // â”€â”€ 1. Fetch template â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    const res = await fetch(publicWorkbookUrl(templateConfig.url), {
      cache: "no-store",
    });
    if (!res.ok)
      throw new Error(
        `HTTP ${res.status} - make sure ${templateConfig.url.slice(1)} is inside your project public folder`,
      );
    const buf = await res.arrayBuffer();

    // â”€â”€ 2. Unzip â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    const { unzipSync, zipSync, strFromU8, strToU8 } = await import("fflate");
    const files = unzipSync(new Uint8Array(buf));

    // â”€â”€ 3. Locate worksheet â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    const sheetKey = Object.keys(files).find((k) =>
      /^xl\/worksheets\/sheet\d+\.xml$/.test(k),
    );
    if (!sheetKey)
      throw new Error("Could not find worksheet XML inside template");
    let xml = strFromU8(files[sheetKey]);

    // â”€â”€ 4. Month label â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    const full = monthKey !== "all" ? formatMonthMMMMYYYY(monthKey) : "All";
    const parts = full.split(" ");
    const label =
      parts.length === 2 ? `${parts[0].slice(0, 3)}-${parts[1]}` : full;
    const safeName = label.replace(/[\/*?[\]:]/g, "-").slice(0, 31);
    const regionLabel =
      Array.from(
        new Set(
          tickets
            .map((ticket) => normalizeTicketRegionLabel(ticket.primary.region))
            .filter(Boolean),
        ),
      ).join(" / ") || "EOA";

    // â”€â”€ 5. Set metadata cells â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    xml = setCell(xml, "C5", regionLabel);
    xml = setCell(xml, "Q5", label);

    // â”€â”€ 6. Rename sheet â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    const wbKey = "xl/workbook.xml";
    if (files[wbKey]) {
      let wbXml = strFromU8(files[wbKey]);
      wbXml = wbXml.replace(
        /(<sheet\b[^>]*\bname=")[^"]*(")/,
        `$1${xmlEsc(safeName)}$2`,
      );
      wbXml = wbXml.replace(
        /'Month-Year'!/g,
        `'${safeName.replace(/'/g, "''")}'!`,
      );
      files[wbKey] = strToU8(wbXml);
    }

    // â”€â”€ 7. Row insertion when more tickets than template rows â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    const needed = tickets.length;
    if (needed > AVAIL) {
      const extra = needed - AVAIL;
      const tmplMarker = ` r="${DATA_START}" `;
      const tmplMarkerIdx = xml.indexOf(tmplMarker);
      const tmplRowStart =
        tmplMarkerIdx !== -1 ? xml.lastIndexOf("<row", tmplMarkerIdx) : -1;
      const tmplRowEnd =
        tmplRowStart !== -1 ? xml.indexOf("</row>", tmplRowStart) + 6 : -1;
      const tmplRow =
        tmplRowStart !== -1 && tmplRowEnd > 0
          ? xml.slice(tmplRowStart, tmplRowEnd)
          : "";

      xml = xml
        .replace(/(<row[^>]* r=")(\d+)(")/g, (_m, a, n, b) =>
          +n >= PROTECTED ? `${a}${+n + extra}${b}` : _m,
        )
        .replace(/(<c[^>]* r=")([A-Z]+)(\d+)(")/g, (_m, a, col, n, b) =>
          +n >= PROTECTED ? `${a}${col}${+n + extra}${b}` : _m,
        )
        .replace(
          /(<mergeCell ref=")([A-Z]+)(\d+)(:)([A-Z]+)(\d+)(")/g,
          (_m, a, c1, r1, sep, c2, r2, b) =>
            `${a}${c1}${+r1 >= PROTECTED ? +r1 + extra : +r1}${sep}${c2}${+r2 >= PROTECTED ? +r2 + extra : +r2}${b}`,
        )
        .replace(
          /(<dimension ref="[A-Z]+\d+:[A-Z]+)(\d+)(")/,
          (_m, pre, n, post) =>
            +n >= PROTECTED ? `${pre}${+n + extra}${post}` : _m,
        );

      if (tmplRow) {
        const newRows = Array.from({ length: extra }, (_, i) => {
          const rn = DATA_START + AVAIL + i;
          return tmplRow
            .replace(/ r="(\d+)"/, ` r="${rn}"`)
            .replace(
              /(<c[^>]* r=")([A-Z]+)\d+(")/g,
              (_m2, a, col, b) => `${a}${col}${rn}${b}`,
            )
            .replace(/<v>[^<]*<\/v>/g, "")
            .replace(/<is>[\s\S]*?<\/is>/g, "")
            .replace(/<f[^>]*\/>/g, "")
            .replace(/<f[^>]*>[\s\S]*?<\/f>/g, "")
            .replace(/\s+t="[^"]*"/g, "");
        }).join("");
        const shiftedTag = ` r="${PROTECTED + extra}" `;
        const shiftedIdx = xml.indexOf(shiftedTag);
        if (shiftedIdx !== -1) {
          const insertAt = xml.lastIndexOf("<row", shiftedIdx);
          xml = xml.slice(0, insertAt) + newRows + xml.slice(insertAt);
        } else {
          xml = xml.replace("</sheetData>", newRows + "</sheetData>");
        }
      }
    }

    // â”€â”€ 8. Build plain body style (no inherited fill/format from template) â”€
    const styleCodec = { strFromU8, strToU8 };
    const plainBodyStyle = ensureExcelPlainTicketBodyStyle(files, styleCodec);

    const ticketColumnValues: Record<string, string[]> = Object.fromEntries(
      COLS.map((col) => [col, []]),
    );

    // â”€â”€ 9. Write data rows â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    tickets.forEach((ticket, i) => {
      const row = DATA_START + i;
      const p = ticket.primary;
      const siteLines = ticketSiteLinesForTemplate(ticket);
      const actionTaken =
        p.actionTaken || uniqueTicketValues(ticket, "actionTaken") || "";
      const rca = p.rca || uniqueTicketValues(ticket, "rca") || "";

      const values: Record<string, string | number> = {
        A: i + 1, // serial â€” numeric
        B: siteLines.siteIds,
        C: siteLines.siteNames,
        D: p.managedResource || "",
        E: p.severity || "",
        F: p.observationDate || "", // kept as source string, no date serial
        G: p.observationTime || "",
        H: p.recoveryDate || "",
        I: p.recoveryTime || "",
        J: p.escalatedForL3SupportDate || "",
        K: p.escalatedForL3SupportTime || "",
        L: p.duration || "",
        M: ticket.tt || "",
        N: p.status || "",
        O: p.escalatedTo || "",
        P: rca,
        Q: actionTaken,
      };

      COLS.forEach((col) => {
        ticketColumnValues[col].push(String(values[col] ?? ""));
      });

      const rowXml =
        `<row r="${row}">` +
        COLS.map((col) =>
          buildStyledCellXml(
            col,
            row,
            values[col] ?? "",
            plainBodyStyle,
            col === "A",
          ),
        ).join("") +
        `</row>`;

      xml = rewriteWorksheetRow(xml, row, rowXml);
    });

    // â”€â”€ 10. Clear unused template rows â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    for (let r = DATA_START + needed; r < PROTECTED; r++) {
      const emptyRowXml =
        `<row r="${r}">` +
        COLS.map((col) => `<c r="${col}${r}" s="${plainBodyStyle}"/>`).join(
          "",
        ) +
        `</row>`;
      xml = rewriteWorksheetRow(xml, r, emptyRowXml);
    }

    xml = normalizeTicketTableRows(
      xml,
      DATA_START,
      DATA_START + Math.max(needed, 1) - 1,
    );
    xml = removeTicketSiteNameMerges(
      xml,
      DATA_START,
      DATA_START + Math.max(needed, AVAIL) - 1,
    );
    xml = setWorksheetColumns(xml, ticketColumnValues);
    files[sheetKey] = strToU8(xml);

    const output = zipSync(files, { level: 0 });
    const blob = new Blob([output], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    const regionForFile =
      regionLabel.replace(/\s*\/\s*/g, "_").replace(/[^A-Za-z0-9_-]/g, "") ||
      templateKind;
    const exportStamp = new Date()
      .toISOString()
      .replace(/[-:]/g, "")
      .replace(/\..+$/, "");
    a.download = `DMR_Monthly_Report_${regionForFile}_${label}_${exportStamp}.xlsx`;
    a.click();
    URL.revokeObjectURL(url);
  } catch (err: any) {
    console.error("Template export failed:", err);
    alert(
      "Template export failed.\n\nCheck:\n" +
        `  â€¢ ${templateConfig.url.slice(1)} must be inside the  public/  folder\n` +
        "  â€¢ Regional templates are: EOA_DMR_Monthly_Report.xlsx, SOA_DMR_Monthly_Report.xlsx, COA_DMR_Monthly_Report.xlsx, WOA_DMR_Monthly_Report.xlsx\n\n" +
        "Error: " +
        (err?.message ?? String(err)),
    );
  }
}

async function exportPdf(rows: TicketAggregate[], monthKey: string) {
  const doc = new jsPDF({ orientation: "landscape", unit: "mm", format: "a3" });
  const monthLabel =
    monthKey !== "all" ? formatMonthMMMMYYYY(monthKey) : "All Months";
  const normalizePdfRegion = (region: string) => {
    const value = clean(region).toUpperCase();
    if (value === "EOA" || value === "NEOA") return "EOA";
    if (value === "SOA") return "SOA";
    if (value === "COA") return "COA";
    if (value === "WOA") return "WOA";
    return value;
  };
  const regionLabel =
    Array.from(
      new Set(
        rows
          .flatMap((ticket) =>
            ticket.rows.map((row) => normalizePdfRegion(row.region)),
          )
          .filter(Boolean),
      ),
    ).join(" / ") || "All";
  const pageW = doc.internal.pageSize.getWidth();
  const pageH = doc.internal.pageSize.getHeight();
  {
    const logos = await loadReportLogos();
    const templateDoc = new jsPDF({
      orientation: "landscape",
      unit: "mm",
      format: "a3",
    });
    const templatePageW = templateDoc.internal.pageSize.getWidth();
    const templatePageH = templateDoc.internal.pageSize.getHeight();
    const C = TEMPLATE_PDF_COLORS;
    const ticketTableMargin = Math.max(12, (templatePageW - 390) / 2);

    const drawTemplateHeader = () => {
      templateDoc.setFillColor(255, 255, 255);
      templateDoc.rect(0, 0, templatePageW, templatePageH, "F");
      drawPdfReportHeader(
        templateDoc,
        templatePageW,
        "MONTHLY REPORT",
        "DMR SYSTEM | DMR Hytera",
        logos,
        C.title,
      );
      templateDoc.setFontSize(10);
      templateDoc.setTextColor(...C.text);
      templateDoc.setFont("helvetica", "bold");
      templateDoc.text("DMR SYSTEM", 12, 26);
      const info = [
        ["Region", regionLabel],
        ["Network", "DMR Hytera"],
        ["Month", monthLabel],
      ];
      templateDoc.setFontSize(8);
      info.forEach((pair, i) => {
        const x = 12 + i * 60;
        templateDoc.setDrawColor(...C.border);
        templateDoc.setFillColor(...C.header);
        templateDoc.rect(x, 32, 22, 7, "FD");
        templateDoc.setFillColor(255, 255, 255);
        templateDoc.rect(x + 22, 32, 36, 7, "FD");
        templateDoc.setTextColor(...C.text);
        templateDoc.setFont("helvetica", "bold");
        templateDoc.text(pair[0], x + 3, 36.8);
        templateDoc.setFont("helvetica", "normal");
        templateDoc.text(pair[1], x + 25, 36.8);
      });
      templateDoc.setFillColor(...C.band);
      templateDoc.rect(12, 44, templatePageW - 24, 8, "F");
      templateDoc.setTextColor(...C.bandText);
      templateDoc.setFont("helvetica", "bold");
      templateDoc.setFontSize(9);
      templateDoc.text("Outages in this Month", 15, 49.5);
    };

    autoTable(templateDoc, {
      startY: 56,
      head: TICKET_TEMPLATE_PDF_HEAD,
      body: distinctReportRows(rows),
      theme: "grid",
      styles: {
        fontSize: 6.6,
        cellPadding: 1.35,
        overflow: "linebreak",
        halign: "center",
        valign: "middle",
        textColor: C.text,
        fillColor: [255, 255, 255],
        lineColor: C.border,
        lineWidth: 0.15,
      },
      headStyles: {
        fillColor: C.header,
        textColor: C.text,
        fontStyle: "bold",
        fontSize: 6.8,
        halign: "center",
        valign: "middle",
      },
      alternateRowStyles: { fillColor: [248, 248, 248] },
      columnStyles: {
        0: { cellWidth: 8, halign: "center" },
        1: { cellWidth: 24, halign: "center" },
        2: { cellWidth: 28, halign: "center" },
        3: { cellWidth: 34, halign: "center" },
        4: { cellWidth: 16, halign: "center", fontStyle: "bold" },
        5: { cellWidth: 34, halign: "center" },
        6: { cellWidth: 18, halign: "center" },
        7: { cellWidth: 15, halign: "center" },
        8: { cellWidth: 18, halign: "center" },
        9: { cellWidth: 15, halign: "center" },
        10: { cellWidth: 24, halign: "center" },
        11: { cellWidth: 20, halign: "center" },
        12: { cellWidth: 24, halign: "center" },
        13: { cellWidth: 18, halign: "center", fontStyle: "bold" },
        14: { cellWidth: 16, halign: "center", fontStyle: "bold" },
        15: { cellWidth: 22, halign: "center" },
        16: { cellWidth: 46, halign: "center" },
        17: { cellWidth: 31, halign: "center" },
      },
      margin: {
        left: ticketTableMargin,
        right: ticketTableMargin,
        top: 56,
        bottom: 10,
      },
      willDrawPage: drawTemplateHeader,
      didParseCell: (d) => {
        if (d.section === "body" && d.column.index === 4) {
          const v = String(d.cell.raw ?? "").toLowerCase();
          d.cell.styles.textColor =
            v.includes("critical") || v.includes("p1") ? C.warn : C.text;
        }
        if (d.section === "body" && d.column.index === 14) {
          const v = String(d.cell.raw ?? "").toLowerCase();
          d.cell.styles.textColor =
            v.includes("closed") || v.includes("resolved")
              ? C.ok
              : v
                ? C.warn
                : C.text;
        }
      },
      didDrawPage: (d) => {
        templateDoc.setFont("helvetica", "normal");
        templateDoc.setFontSize(7);
        templateDoc.setTextColor(...C.muted);
        templateDoc.text(
          `DMR Monthly Report - ${monthLabel} | Page ${d.pageNumber}`,
          templatePageW / 2,
          templatePageH - 5,
          { align: "center" },
        );
      },
    });
    const suffix =
      monthKey !== "all" ? `-${monthLabel.replace(/ /g, "-")}` : "";
    templateDoc.save(`DMR-Monthly-Tickets${suffix}.pdf`);
    return;
  }
}

// â”€â”€â”€ All remaining component code (StatCard, PartnerLogoStrip, SelectFilter,
//     MultiSelectFilter, and the default export Home) is IDENTICAL to your
//     original file from this point forward. Paste it here unchanged.
// â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

function clampGaugePercent(value: number): number {
  if (!Number.isFinite(value)) return 0;
  return Math.max(0, Math.min(100, value));
}

function extractFirstNumericValue(text: string): number | null {
  const match = clean(text).match(/-?\d+(?:\.\d+)?/);
  return match ? Number(match[0]) : null;
}

type GaugeStatus = "excellent" | "good" | "warning" | "critical";
type GaugeThresholds = { excellent: number; good: number; warning: number };
type PerformanceGaugeContext = { totalSites: number; totalHours: number };
type PerformanceGaugeConfig = {
  id: string;
  label: string;
  color: string;
  icon: LucideIcon;
  direction: "higher" | "lower";
  thresholds: GaugeThresholds;
  getValue: (
    kpi: ReturnType<typeof computePerfKPIs>,
    ctx: PerformanceGaugeContext,
  ) => number;
  getScale: (
    kpi: ReturnType<typeof computePerfKPIs>,
    ctx: PerformanceGaugeContext,
  ) => { min: number; max: number };
  formatValue: (kpi: ReturnType<typeof computePerfKPIs>) => string;
  caption: (
    kpi: ReturnType<typeof computePerfKPIs>,
    ctx: PerformanceGaugeContext,
  ) => string;
  helper: (
    kpi: ReturnType<typeof computePerfKPIs>,
    ctx: PerformanceGaugeContext,
  ) => string;
  sparkline: (
    rows: PerfRow[],
    kpi: ReturnType<typeof computePerfKPIs>,
    ctx: PerformanceGaugeContext,
  ) => number[];
};
type PerformanceGaugeCardModel = {
  id: string;
  label: string;
  value: string;
  color: string;
  progress: number;
  status: GaugeStatus;
  icon: LucideIcon;
  caption: string;
  helper: string;
  sparkline: number[];
};

const PERFORMANCE_GAUGE_CONFIG: PerformanceGaugeConfig[] = [
  {
    id: "availability",
    label: "% Availability",
    color: "#22d3ee",
    icon: Activity,
    direction: "higher",
    thresholds: { excellent: 99.5, good: 98, warning: 95 },
    getValue: (kpi) => extractFirstNumericValue(kpi.pctAvailability) ?? 0,
    getScale: () => ({ min: 0, max: 100 }),
    formatValue: (kpi) => kpi.pctAvailability,
    caption: (kpi) =>
      `${(extractFirstNumericValue(kpi.pctAvailability) ?? 0).toFixed(2)}% healthy window`,
    helper: (kpi) =>
      `Available ${kpi.totalAvail.toFixed(1)} hrs across selected sites`,
    sparkline: (rows) =>
      rows.map((r) => {
        const total = r.availHours + r.sitesDownHours;
        return total > 0 ? (r.availHours / total) * 100 : 100;
      }),
  },
  {
    id: "mttr",
    label: "MTTR",
    color: "#f59e0b",
    icon: RefreshCw,
    direction: "lower",
    thresholds: { excellent: 4, good: 8, warning: 24 },
    getValue: (kpi) => extractFirstNumericValue(kpi.mttr) ?? 0,
    getScale: () => ({ min: 0, max: 24 }),
    formatValue: (kpi) => kpi.mttr,
    caption: () => "Target <= 24 hrs mean repair time",
    helper: (kpi) =>
      `${(extractFirstNumericValue(kpi.mttr) ?? 0).toFixed(2)} hrs average repair`,
    sparkline: (rows) =>
      rows.filter((r) => r.sitesDownHours > 0).map((r) => r.sitesDownHours),
  },
  {
    id: "mtbf",
    label: "MTBF",
    color: "#3b82f6",
    icon: Network,
    direction: "higher",
    thresholds: { excellent: 168, good: 72, warning: 24 },
    getValue: (kpi) => extractFirstNumericValue(kpi.mtbf) ?? 0,
    getScale: () => ({ min: 0, max: 168 }),
    formatValue: (kpi) => kpi.mtbf,
    caption: () => "Target >= 168 hrs between failures",
    helper: (kpi) =>
      `${(extractFirstNumericValue(kpi.mtbf) ?? 0).toFixed(2)} hrs between failures`,
    sparkline: (rows) => rows.map((r) => r.availHours),
  },
  {
    id: "mttf",
    label: "MTTF",
    color: "#a78bfa",
    icon: ShieldAlert,
    direction: "higher",
    thresholds: { excellent: 192, good: 96, warning: 36 },
    getValue: (kpi) => extractFirstNumericValue(kpi.mttf) ?? 0,
    getScale: () => ({ min: 0, max: 192 }),
    formatValue: (kpi) => kpi.mttf,
    caption: () => "Target >= 192 hrs expected failure-free time",
    helper: (kpi) =>
      `${(extractFirstNumericValue(kpi.mttf) ?? 0).toFixed(2)} hrs uptime horizon`,
    sparkline: (rows) =>
      rows.map((r) => r.availHours + Math.max(0, r.sitesDownHours * 0.35)),
  },
  {
    id: "affectedSites",
    label: "Affected Sites",
    color: "#f43f5e",
    icon: AlertTriangle,
    direction: "lower",
    thresholds: { excellent: 0, good: 5, warning: 15 },
    getValue: (kpi) => kpi.affectedSites,
    getScale: (_, ctx) => ({ min: 0, max: Math.max(1, ctx.totalSites) }),
    formatValue: (kpi) => String(kpi.affectedSites),
    caption: (kpi, ctx) =>
      `${((kpi.affectedSites / Math.max(1, ctx.totalSites)) * 100).toFixed(1)}% of monitored sites`,
    helper: (kpi, ctx) =>
      `${kpi.affectedSites} impacted / ${ctx.totalSites} total sites`,
    sparkline: (rows) => rows.map((r) => (r.sitesDownHours > 0 ? 1 : 0)),
  },
  {
    id: "nonAffectedSites",
    label: "Non-Affected Sites",
    color: "#10b981",
    icon: CheckCircle2,
    direction: "higher",
    thresholds: { excellent: 90, good: 75, warning: 50 },
    getValue: (kpi, ctx) =>
      (kpi.nonAffectedSites / Math.max(1, ctx.totalSites)) * 100,
    getScale: () => ({ min: 0, max: 100 }),
    formatValue: (kpi) => String(kpi.nonAffectedSites),
    caption: (kpi, ctx) =>
      `${((kpi.nonAffectedSites / Math.max(1, ctx.totalSites)) * 100).toFixed(1)}% healthy sites`,
    helper: (kpi, ctx) =>
      `${kpi.nonAffectedSites} stable / ${ctx.totalSites} total sites`,
    sparkline: (rows) => rows.map((r) => (r.sitesDownHours <= 0 ? 1 : 0)),
  },
  {
    id: "totalDown",
    label: "Total Down",
    color: "#fb923c",
    icon: CloudOff,
    direction: "lower",
    thresholds: { excellent: 12, good: 48, warning: 120 },
    getValue: (kpi) => kpi.totalDown,
    getScale: (_, ctx) => ({ min: 0, max: Math.max(1, ctx.totalHours || 1) }),
    formatValue: (kpi) => kpi.totalDownHrs,
    caption: (kpi) => `${kpi.totalDownHrs} lost during selected window`,
    helper: (kpi, ctx) =>
      `${((kpi.totalDown / Math.max(1, ctx.totalHours)) * 100).toFixed(2)}% downtime share`,
    sparkline: (rows) => rows.map((r) => r.sitesDownHours),
  },
];

function gaugeStatusFromValue(
  value: number,
  direction: PerformanceGaugeConfig["direction"],
  thresholds: GaugeThresholds,
): GaugeStatus {
  if (direction === "higher") {
    if (value >= thresholds.excellent) return "excellent";
    if (value >= thresholds.good) return "good";
    if (value >= thresholds.warning) return "warning";
    return "critical";
  }
  if (value <= thresholds.excellent) return "excellent";
  if (value <= thresholds.good) return "good";
  if (value <= thresholds.warning) return "warning";
  return "critical";
}

function gaugeProgressFromScale(
  value: number,
  direction: PerformanceGaugeConfig["direction"],
  min: number,
  max: number,
): number {
  const range = Math.max(1, max - min);
  const normalized = ((value - min) / range) * 100;
  return clampGaugePercent(
    direction === "higher" ? normalized : 100 - normalized,
  );
}

function compactSparkline(values: number[], maxPoints = 12): number[] {
  const cleanValues = values.filter((value) => Number.isFinite(value));
  if (cleanValues.length === 0) return [0, 0, 0, 0, 0];
  if (cleanValues.length <= maxPoints) return cleanValues;
  const bucketSize = cleanValues.length / maxPoints;
  return Array.from({ length: maxPoints }, (_, index) => {
    const start = Math.floor(index * bucketSize);
    const end = Math.max(start + 1, Math.floor((index + 1) * bucketSize));
    const slice = cleanValues.slice(start, end);
    return (
      slice.reduce((sum, value) => sum + value, 0) / Math.max(1, slice.length)
    );
  });
}

function sparklinePath(values: number[]): string {
  const points = compactSparkline(values, 14);
  const max = Math.max(...points, 1);
  const min = Math.min(...points, 0);
  const range = Math.max(1, max - min);
  return points
    .map((value, index) => {
      const x = points.length === 1 ? 50 : (index / (points.length - 1)) * 100;
      const y = 100 - ((value - min) / range) * 100;
      return `${x},${y}`;
    })
    .join(" ");
}

function gaugeNeedleAngle(progress: number): number {
  return -130 + clampGaugePercent(progress) * 2.6;
}

function buildPerformanceGaugeCards(
  kpi: ReturnType<typeof computePerfKPIs>,
  rows: PerfRow[],
): PerformanceGaugeCardModel[] {
  const ctx: PerformanceGaugeContext = {
    totalSites: Math.max(1, kpi.affectedSites + kpi.nonAffectedSites),
    totalHours: Math.max(1, kpi.totalAvail + kpi.totalDown),
  };
  return PERFORMANCE_GAUGE_CONFIG.map((config) => {
    const value = config.getValue(kpi, ctx);
    const scale = config.getScale(kpi, ctx);
    return {
      id: config.id,
      label: config.label,
      value: config.formatValue(kpi),
      color: config.color,
      progress: gaugeProgressFromScale(
        value,
        config.direction,
        scale.min,
        scale.max,
      ),
      status: gaugeStatusFromValue(value, config.direction, config.thresholds),
      icon: config.icon,
      caption: config.caption(kpi, ctx),
      helper: config.helper(kpi, ctx),
      sparkline: compactSparkline(config.sparkline(rows, kpi, ctx)),
    };
  });
}

function statusLabel(status: GaugeStatus): string {
  return status.charAt(0).toUpperCase() + status.slice(1);
}

function PerformanceGaugeCard({
  label,
  value,
  color,
  progress,
  status,
  icon: Icon,
  caption,
  helper,
  index,
}: PerformanceGaugeCardModel & { index: number }) {
  const safeProgress = clampGaugePercent(progress);
  const gaugeId = label.replace(/[^a-z0-9]+/gi, "-").toLowerCase();
  return (
    <article
      className="perf-gauge-card"
      style={{
        ["--gauge-color" as string]: color,
        ["--gauge-progress" as string]: `${safeProgress}%`,
        ["--card-index" as string]: index,
      }}
    >
      <div className="perf-gauge-card__topline">
        <span className="perf-gauge-card__label">{label}</span>
        <span
          className={`perf-gauge-card__status perf-gauge-card__status--${status}`}
        >
          {statusLabel(status)}
        </span>
      </div>
      <div className="perf-gauge-card__dial-wrap">
        <svg
          className="perf-gauge-card__dial"
          viewBox="0 0 240 165"
          role="img"
          aria-label={`${label} performance gauge`}
        >
          <defs>
            <linearGradient
              id={`gauge-arc-${gaugeId}`}
              x1="0%"
              x2="100%"
              y1="0%"
              y2="0%"
            >
              <stop offset="0%" stopColor={color} stopOpacity="0.32" />
              <stop offset="52%" stopColor={color} stopOpacity="0.98" />
              <stop offset="100%" stopColor={color} stopOpacity="0.46" />
            </linearGradient>
          </defs>
          <path
            className="perf-gauge-card__dial-arc-track"
            d="M 14 140 A 106 106 0 0 1 226 140"
            pathLength={100}
          />
          <path
            className="perf-gauge-card__dial-arc-glow"
            d="M 14 140 A 106 106 0 0 1 226 140"
            pathLength={100}
            stroke={`url(#gauge-arc-${gaugeId})`}
            strokeDasharray={100}
            strokeDashoffset={100 - safeProgress}
          />
          {Array.from({ length: 11 }, (_, tick) => {
            const angle = Math.PI + tick * (Math.PI / 10);
            const x1 = 120 + Math.cos(angle) * 86,
              y1 = 140 + Math.sin(angle) * 86;
            const x2 = 120 + Math.cos(angle) * 102,
              y2 = 140 + Math.sin(angle) * 102;
            return (
              <line
                key={tick}
                className="perf-gauge-card__tick"
                x1={x1}
                y1={y1}
                x2={x2}
                y2={y2}
              />
            );
          })}
        </svg>
        <div className="perf-gauge-card__center">
          <div className="perf-gauge-card__icon-wrap">
            <Icon size={20} />
          </div>
          <strong className="perf-gauge-card__value">{value || "--"}</strong>
          <span className="perf-gauge-card__caption">{caption}</span>
        </div>
      </div>
      <div className="perf-gauge-card__footer">
        <span className="perf-gauge-card__helper">{helper}</span>
      </div>
    </article>
  );
}

function StatCard({
  label,
  value,
  note,
  icon: Icon,
  tone,
  onClick,
  className = "",
  style,
}: {
  label: ReactNode;
  value: ReactNode;
  note?: ReactNode;
  icon: LucideIcon;
  tone: string;
  onClick?: () => void;
  className?: string;
  style?: CSSProperties;
}) {
  return (
    <div
      className={`stat-card ${className}`.trim()}
      style={{
        ["--tone" as string]: tone,
        cursor: onClick ? "pointer" : undefined,
        ...style,
      }}
      onClick={onClick}
      role={onClick ? "button" : undefined}
      tabIndex={onClick ? 0 : undefined}
      onKeyDown={
        onClick
          ? (e) => {
              if (e.key === "Enter" || e.key === " ") onClick();
            }
          : undefined
      }
    >
      <div className="stat-icon">
        <Icon size={28} />
      </div>
      <span>{label}</span>
      <strong>{value}</strong>
      <small>{note}</small>
    </div>
  );
}

function BannerDecorativeSvg() {
  return (
    <svg
      className="corporate-banner-svg"
      viewBox="0 0 1440 260"
      preserveAspectRatio="none"
      aria-hidden="true"
      focusable="false"
    >
      <defs>
        <pattern id="bannerDotGrid" width="28" height="28" patternUnits="userSpaceOnUse">
          <circle cx="3" cy="3" r="1.35" fill="rgba(125,211,252,0.34)" />
        </pattern>
        <linearGradient id="leftBlueWave" x1="0" y1="0" x2="1" y2="0">
          <stop offset="0" stopColor="#00d4ff" stopOpacity="0.9" />
          <stop offset="1" stopColor="#1d4ed8" stopOpacity="0.1" />
        </linearGradient>
        <linearGradient id="rightRedWave" x1="1" y1="0" x2="0" y2="0">
          <stop offset="0" stopColor="#ef4444" stopOpacity="0.95" />
          <stop offset="1" stopColor="#fb7185" stopOpacity="0.08" />
        </linearGradient>
        <filter id="bannerGlow" x="-30%" y="-30%" width="160%" height="160%">
          <feGaussianBlur stdDeviation="8" result="blur" />
          <feMerge>
            <feMergeNode in="blur" />
            <feMergeNode in="SourceGraphic" />
          </feMerge>
        </filter>
      </defs>

      <rect width="1440" height="260" fill="url(#bannerDotGrid)" opacity="0.72" />

      <ellipse cx="170" cy="118" rx="185" ry="88" fill="#0ea5e9" opacity="0.18" filter="url(#bannerGlow)" />
      <ellipse cx="1274" cy="116" rx="182" ry="86" fill="#ef4444" opacity="0.2" filter="url(#bannerGlow)" />
      <ellipse cx="720" cy="22" rx="120" ry="34" fill="#38bdf8" opacity="0.12" filter="url(#bannerGlow)" />

      <path d="M-20 70 C120 20 198 35 310 92 C395 135 486 128 610 74" fill="none" stroke="url(#leftBlueWave)" strokeWidth="4" opacity="0.82" />
      <path d="M-30 124 C112 70 210 82 330 142 C412 183 500 168 628 118" fill="none" stroke="url(#leftBlueWave)" strokeWidth="2.8" opacity="0.58" />
      <path d="M-38 174 C96 132 214 142 340 198 C435 240 530 220 660 172" fill="none" stroke="url(#leftBlueWave)" strokeWidth="2" opacity="0.45" />

      <path d="M1468 78 C1325 24 1230 38 1125 96 C1038 144 944 134 826 82" fill="none" stroke="url(#rightRedWave)" strokeWidth="4" opacity="0.85" />
      <path d="M1478 130 C1342 78 1238 86 1120 150 C1034 195 936 176 806 122" fill="none" stroke="url(#rightRedWave)" strokeWidth="2.8" opacity="0.62" />
      <path d="M1488 188 C1348 136 1246 146 1118 204 C1022 248 916 222 792 176" fill="none" stroke="url(#rightRedWave)" strokeWidth="2" opacity="0.48" />

      <line x1="350" y1="92" x2="1090" y2="92" stroke="rgba(56,189,248,0.52)" strokeWidth="1.4" filter="url(#bannerGlow)" />
      <line x1="390" y1="174" x2="1050" y2="174" stroke="rgba(20,184,166,0.48)" strokeWidth="1.2" filter="url(#bannerGlow)" />
      <line x1="438" y1="102" x2="1002" y2="102" stroke="rgba(255,255,255,0.16)" strokeWidth="0.8" />
      <line x1="438" y1="164" x2="1002" y2="164" stroke="rgba(255,255,255,0.14)" strokeWidth="0.8" />

      <g className="banner-mini-chart" transform="translate(470 54)" opacity="0.84">
        <rect x="0" y="30" width="7" height="22" rx="2" fill="#22d3ee" />
        <rect x="14" y="18" width="7" height="34" rx="2" fill="#60a5fa" />
        <rect x="28" y="8" width="7" height="44" rx="2" fill="#38bdf8" />
        <rect x="42" y="24" width="7" height="28" rx="2" fill="#14b8a6" />
        <path d="M-5 55H58" stroke="rgba(255,255,255,0.28)" strokeWidth="1.2" />
      </g>

      <g className="banner-mini-donut" transform="translate(923 54)" opacity="0.86">
        <circle cx="28" cy="28" r="22" fill="none" stroke="rgba(239,68,68,0.42)" strokeWidth="8" />
        <path d="M28 6a22 22 0 0 1 19 33" fill="none" stroke="#fb7185" strokeWidth="8" strokeLinecap="round" />
        <path d="M47 39A22 22 0 0 1 12 44" fill="none" stroke="#22d3ee" strokeWidth="8" strokeLinecap="round" />
        <circle cx="28" cy="28" r="10" fill="rgba(2,15,46,0.72)" />
      </g>

      <g className="banner-circuit" stroke="rgba(125,211,252,0.52)" strokeWidth="1.3" fill="none" opacity="0.8">
        <path d="M74 34h54v34h42" />
        <path d="M76 220h72v-32h54" />
        <path d="M1266 36h-58v34h-42" />
        <path d="M1264 218h-70v-32h-54" />
      </g>
      <g fill="rgba(34,211,238,0.82)" filter="url(#bannerGlow)">
        <circle cx="74" cy="34" r="4" /><circle cx="128" cy="68" r="3.5" /><circle cx="170" cy="68" r="4" />
        <circle cx="76" cy="220" r="4" /><circle cx="148" cy="188" r="3.5" /><circle cx="202" cy="188" r="4" />
        <circle cx="1266" cy="36" r="4" /><circle cx="1208" cy="70" r="3.5" /><circle cx="1166" cy="70" r="4" />
        <circle cx="1264" cy="218" r="4" /><circle cx="1194" cy="186" r="3.5" /><circle cx="1140" cy="186" r="4" />
      </g>

      <g className="banner-ticket-energy" transform="translate(684 14)" filter="url(#bannerGlow)">
        <path d="M10 8h56c6 0 10 4 10 10v9c-5 1-8 5-8 10s3 9 8 10v9c0 6-4 10-10 10H10C4 66 0 62 0 56v-9c5-1 8-5 8-10s-3-9-8-10v-9C0 12 4 8 10 8Z" fill="rgba(2,15,46,0.62)" stroke="rgba(125,211,252,0.76)" strokeWidth="2" />
        <path d="M42 17 27 38h13l-6 19 18-25H39l3-15Z" fill="rgba(34,211,238,0.88)" stroke="rgba(255,255,255,0.56)" strokeWidth="1.2" strokeLinejoin="round" />
        <path d="M18 22h9M18 52h9M57 22h9M57 52h9" stroke="rgba(20,184,166,0.78)" strokeWidth="2" strokeLinecap="round" />
        <circle cx="18" cy="37" r="2.4" fill="rgba(34,211,238,0.9)" />
        <circle cx="58" cy="37" r="2.4" fill="rgba(248,113,113,0.88)" />
      </g>
    </svg>
  );
}

function CorporateBannerHeader({
  title = "DMR Ticketing Dashboard",
  subtitle = "Ticketing Performance Command Center",
}: {
  title?: string;
  subtitle?: string;
}) {
  return (
    <header className="corporate-banner-header" aria-label="Page header">
      <BannerDecorativeSvg />
      <div className="banner-logo-glow banner-logo-glow--left" />
      <div className="banner-logo-glow banner-logo-glow--right" />
      <div className="banner-logo-card banner-logo-card--left" aria-label="Saudi Energy logo">
        <img src="/assets/se-logo.png" alt="Saudi Energy" />
      </div>
      <div className="banner-title-lockup">
        <h1>{title}</h1>
        <p>{subtitle}</p>
      </div>
      <div className="banner-logo-card banner-logo-card--right" aria-label="NASCO logo">
        <img src="/assets/nasco-logo.png" alt="NASCO" />
      </div>
    </header>
  );
}
function ReportFileIcon({ kind }: { kind: "xlsx" | "pdf" | "ppt" | "png" }) {
  if (kind === "png") {
    return (
      <svg
        className="file-export-svg file-export-svg-png"
        viewBox="0 0 64 64"
        aria-hidden="true"
      >
        <path className="file-page" d="M14 5h25l11 11v43H14Z" />
        <path className="file-fold" d="M39 5v12h11" />
        <circle className="file-mark" cx="25" cy="24" r="5" />
        <path className="file-mark" d="m18 49 11-13 7 8 5-6 8 11Z" />
      </svg>
    );
  }

  const meta = {
    xlsx: { color: "#21a366", label: "XLS" },
    pdf: { color: "#ef4444", label: "PDF" },
    ppt: { color: "#f97316", label: "PPT" },
  }[kind];

  return (
    <svg
      className={`file-export-svg file-export-svg-${kind}`}
      viewBox="0 0 64 64"
      aria-hidden="true"
      style={{ "--file-color": meta.color } as CSSProperties}
    >
      <path className="file-page" d="M14 5h25l11 11v43H14Z" />
      <path className="file-fold" d="M39 5v12h11" />
      <path className="file-lines" d="M22 24h18M22 31h18M22 38h12" />
      {kind === "xlsx" && (
        <path
          className="file-grid"
          d="M22 23h20M22 30h20M22 37h20M28 20v22M36 20v22"
        />
      )}
      {kind === "ppt" && (
        <path className="file-chart" d="M25 41V29h5v12m4 0V24h5v17" />
      )}
      <rect
        className="file-ribbon"
        x="6"
        y="34"
        width="52"
        height="20"
        rx="3"
      />
      <text className="file-label" x="32" y="49" textAnchor="middle">
        {meta.label}
      </text>
    </svg>
  );
}

function SelectFilter({
  label,
  value,
  options,
  optionLabels,
  onChange,
}: {
  label: string;
  value: string;
  options: string[];
  optionLabels?: Record<string, string>;
  onChange: (value: string) => void;
}) {
  const [open, setOpen] = useState(false);
  const [dropdownStyle, setDropdownStyle] = useState<CSSProperties>({});
  const wrapRef = useRef<HTMLDivElement | null>(null);
  const triggerRef = useRef<HTMLButtonElement | null>(null);
  const dropdownRef = useRef<HTMLDivElement | null>(null);

  useEffect(() => {
    function handleClickOutside(e: MouseEvent) {
      const target = e.target as Node;
      if (
        !wrapRef.current?.contains(target) &&
        !dropdownRef.current?.contains(target)
      )
        setOpen(false);
    }
    document.addEventListener("mousedown", handleClickOutside);
    return () => document.removeEventListener("mousedown", handleClickOutside);
  }, []);

  function computePosition() {
    if (!triggerRef.current) return;
    const rect = triggerRef.current.getBoundingClientRect();
    const dropW = Math.max(rect.width, 200);
    const left = Math.min(rect.left, window.innerWidth - dropW - 8);
    const spaceBelow = window.innerHeight - rect.bottom;
    if (spaceBelow >= 160) {
      setDropdownStyle({
        position: "fixed",
        top: rect.bottom + 4,
        left,
        width: dropW,
        zIndex: 99999,
      });
    } else {
      setDropdownStyle({
        position: "fixed",
        bottom: window.innerHeight - rect.top + 4,
        left,
        width: dropW,
        zIndex: 99999,
      });
    }
  }

  function handleOpen() {
    if (!open) computePosition();
    setOpen((o) => !o);
  }

  useEffect(() => {
    if (!open) return;
    const update = () => computePosition();
    window.addEventListener("scroll", update, true);
    window.addEventListener("resize", update);
    return () => {
      window.removeEventListener("scroll", update, true);
      window.removeEventListener("resize", update);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [open]);

  const displayLabel =
    value === "all" || value === "ALL"
      ? "All"
      : (optionLabels?.[value] ?? value);
  const dropdown = open
    ? createPortal(
        <div
          ref={dropdownRef}
          className="multi-select-dropdown"
          style={dropdownStyle}
        >
          <label
            className="multi-select-option"
            onClick={() => {
              onChange("all");
              setOpen(false);
            }}
          >
            <input
              type="radio"
              readOnly
              checked={value === "all" || value === "ALL"}
            />
            <span>All</span>
          </label>
          {options.map((opt) => (
            <label
              key={opt}
              className="multi-select-option"
              onClick={() => {
                onChange(opt);
                setOpen(false);
              }}
            >
              <input type="radio" readOnly checked={value === opt} />
              <span>{optionLabels?.[opt] ?? opt}</span>
            </label>
          ))}
        </div>,
        document.body,
      )
    : null;

  return (
    <div className="filter-field multi-select-filter" ref={wrapRef}>
      <span>{label}</span>
      <button
        type="button"
        className="multi-select-trigger"
        ref={triggerRef}
        onClick={handleOpen}
      >
        <span className="multi-select-value">{displayLabel}</span>
        <ChevronDown
          size={14}
          style={{
            transform: open ? "rotate(180deg)" : undefined,
            transition: "transform .15s",
          }}
        />
      </button>
      {dropdown}
    </div>
  );
}

function MultiSelectFilter({
  label,
  value,
  options,
  optionLabels,
  onChange,
  showAllOption,
}: {
  label: string;
  value: string[];
  options: string[];
  optionLabels?: Record<string, string>;
  onChange: (value: string[]) => void;
  showAllOption?: boolean;
}) {
  const [open, setOpen] = useState(false);
  const [dropdownStyle, setDropdownStyle] = useState<CSSProperties>({});
  const wrapRef = useRef<HTMLDivElement | null>(null);
  const triggerRef = useRef<HTMLButtonElement | null>(null);
  const dropdownRef = useRef<HTMLDivElement | null>(null);

  useEffect(() => {
    function handleClickOutside(e: MouseEvent) {
      const target = e.target as Node;
      if (
        !wrapRef.current?.contains(target) &&
        !dropdownRef.current?.contains(target)
      )
        setOpen(false);
    }
    document.addEventListener("mousedown", handleClickOutside);
    return () => document.removeEventListener("mousedown", handleClickOutside);
  }, []);

  function computePosition() {
    if (!triggerRef.current) return;
    const rect = triggerRef.current.getBoundingClientRect();
    const dropW = Math.max(rect.width, 220);
    const preferredDropH = Math.min(280, options.length * 36 + 48);
    const spaceBelow = window.innerHeight - rect.bottom;
    const spaceAbove = rect.top;
    const left = Math.min(rect.left, window.innerWidth - dropW - 8);
    const openDown = spaceBelow >= preferredDropH || spaceBelow >= spaceAbove;
    const availableHeight = Math.max(
      120,
      (openDown ? spaceBelow : spaceAbove) - 12,
    );
    const maxHeight = Math.min(preferredDropH, availableHeight);
    if (openDown) {
      setDropdownStyle({
        position: "fixed",
        top: rect.bottom + 4,
        left,
        width: dropW,
        maxHeight,
        zIndex: 99999,
      });
    } else {
      setDropdownStyle({
        position: "fixed",
        bottom: window.innerHeight - rect.top + 4,
        left,
        width: dropW,
        maxHeight,
        zIndex: 99999,
      });
    }
  }

  function handleOpen() {
    if (!open) computePosition();
    setOpen((o) => !o);
  }

  useEffect(() => {
    if (!open) return;
    const update = () => computePosition();
    window.addEventListener("scroll", update, true);
    window.addEventListener("resize", update);
    return () => {
      window.removeEventListener("scroll", update, true);
      window.removeEventListener("resize", update);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [open]);

  const allSelected = value.length === 0;
  const toggle = (opt: string) => {
    if (allSelected) {
      onChange([opt]);
      return;
    }
    if (value.includes(opt)) {
      onChange(value.filter((v) => v !== opt));
    } else {
      onChange([...value, opt]);
    }
  };
  const displayLabel = allSelected
    ? "All"
    : value.length === 1
      ? (optionLabels?.[value[0]] ?? value[0])
      : `${value.length} selected`;

  const dropdown = open
    ? createPortal(
        <div
          className="multi-select-dropdown"
          style={dropdownStyle}
          ref={dropdownRef}
        >
          {showAllOption && (
            <label
              className="multi-select-option multi-select-option-all"
              onClick={() => onChange([])}
            >
              <input type="checkbox" readOnly checked={allSelected} />
              <span
                style={{
                  fontWeight: allSelected ? 700 : undefined,
                  color: allSelected ? "#22d3ee" : undefined,
                }}
              >
                All
              </span>
            </label>
          )}
          {!showAllOption && value.length > 0 && (
            <button
              type="button"
              className="multi-select-clear"
              onClick={() => onChange([])}
            >
              Clear
            </button>
          )}
          {options.map((opt) => (
            <label key={opt} className="multi-select-option">
              <input
                type="checkbox"
                checked={value.includes(opt)}
                onChange={() => toggle(opt)}
              />
              <span>{optionLabels?.[opt] ?? opt}</span>
            </label>
          ))}
        </div>,
        document.body,
      )
    : null;

  return (
    <div className="filter-field multi-select-filter" ref={wrapRef}>
      <span>{label}</span>
      <button
        type="button"
        className="multi-select-trigger"
        ref={triggerRef}
        onClick={handleOpen}
      >
        <span className="multi-select-value">{displayLabel}</span>
        <ChevronDown
          size={14}
          style={{
            transform: open ? "rotate(180deg)" : undefined,
            transition: "transform .15s",
          }}
        />
      </button>
      {dropdown}
    </div>
  );
}

// â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
// The default export Home() component is IDENTICAL to your original file.
// Paste the entire "export default function Home() { â€¦ }" block here unchanged.
// â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
export default function Home() {
  const inputRef = useRef<HTMLInputElement | null>(null);
  const addRegionRef = useRef<HTMLInputElement | null>(null);
  const [data, setData] = useState<DashboardData | null>(null);
  const [regions, setRegions] = useState<DashboardData[]>([]);
  const [uploadedWorkbookSources, setUploadedWorkbookSources] = useState<
    UploadedWorkbookSource[]
  >([]);
  const [error, setError] = useState("");
  const [isDragging, setIsDragging] = useState(false);
  const statsRef = useRef<HTMLDivElement | null>(null);
  const performanceKpiCardsRef = useRef<HTMLDivElement | null>(null);
  const tableRef = useRef<HTMLDivElement | null>(null);
  const filtersRef = useRef<HTMLElement | null>(null);
  const [exportMonths, setExportMonths] = useState<string[]>([]);
  const [exportRegions, setExportRegions] = useState<string[]>([]);
  const [filters, setFilters] = useState<Filters>(EMPTY_FILTERS);
  const [tablePage, setTablePage] = useState(1);
  const [perfPage, setPerfPage] = useState(1);
  const TABLE_PAGE_SIZE = 20;
  const [perfMonths, setPerfMonths] = useState<string[]>([]);
  const [perfRegions, setPerfRegions] = useState<string[]>([]);
  const [managedReportType, setManagedReportType] =
    useState<ManagedReportType>("tickets");
  const [managedReportFormat, setManagedReportFormat] =
    useState<ManagedReportFormat>("xlsx");
  const [reportSearch, setReportSearch] = useState("");
  const [generatedReports, setGeneratedReports] = useState<GeneratedReportItem[]>([]);
  const { theme: dashboardTheme, toggleTheme } = useTheme();
  const isDark = dashboardTheme === "dark";
  const [isAuthenticated, setIsAuthenticated] = useState(getInitialLoginState);
  const [loginUsername, setLoginUsername] = useState("");
  const [loginPassword, setLoginPassword] = useState("");
  const [loginError, setLoginError] = useState("");
  const [collapsedSections, setCollapsedSections] = useState<
    Record<DashboardSectionId, boolean>
  >(INITIAL_COLLAPSED_SECTIONS);
  const [activeDashboardTab, setActiveDashboardTab] =
    useState<DashboardSectionId>("performanceKpis");
  const [savedSnapshotAvailable, setSavedSnapshotAvailable] = useState(false);
  const [googleRegionLinks, setGoogleRegionLinks] = useState<GoogleRegionLinks>(
    EMPTY_GOOGLE_REGION_LINKS,
  );
  const [googleRegionSaveLinks, setGoogleRegionSaveLinks] =
    useState<GoogleRegionSaveLinks>(EMPTY_GOOGLE_REGION_SAVE_LINKS);
  const [googleRegionSelection, setGoogleRegionSelection] =
    useState<GoogleRegionSelection>(EMPTY_GOOGLE_REGION_SELECTION);
  const [microsoftGraphConfig, setMicrosoftGraphConfig] =
    useState<MicrosoftGraphConfig>(EMPTY_MICROSOFT_GRAPH_CONFIG);
  const [googleSheetLoading, setGoogleSheetLoading] = useState(false);
  const [onlineSourceMode, setOnlineSourceMode] = useState<
    "add" | "replace" | null
  >(null);
  const [manualDrafts, setManualDrafts] = useState<ManualTicketDraft[]>([]);
  const [manualSaveLoading, setManualSaveLoading] = useState(false);
  const [manualSaveStatus, setManualSaveStatus] = useState("");
  const [existingTtSearch, setExistingTtSearch] = useState("");
  const [existingTtLoading, setExistingTtLoading] = useState(false);
  const activeThemeImage = THEME_IMAGES[dashboardTheme];
  const heroThemeOverlay =
    "linear-gradient(90deg, rgba(0,0,0,0), rgba(0,0,0,0))";
  const ribbonThemeOverlay =
    "linear-gradient(90deg, rgba(0,0,0,0), rgba(0,0,0,0))";
  function handleLoginSubmit(event: React.FormEvent<HTMLFormElement>) {
    event.preventDefault();
    const usernameMatches =
      loginUsername.trim().toLowerCase() === DASHBOARD_LOGIN_USERNAME;
    const passwordMatches = loginPassword === DASHBOARD_LOGIN_PASSWORD;

    if (!usernameMatches || !passwordMatches) {
      setLoginError("Invalid username or password.");
      return;
    }

    if (typeof window !== "undefined") {
      window.sessionStorage.setItem(LOGIN_SESSION_KEY, "authenticated");
    }
    setLoginError("");
    setLoginPassword("");
    setIsAuthenticated(true);
  }

  function handleLogout() {
    if (typeof window !== "undefined") {
      window.sessionStorage.removeItem(LOGIN_SESSION_KEY);
    }
    setIsAuthenticated(false);
    setLoginPassword("");
  }

  const renderThemeToggle = (modifierClass = "") => (
    <div
      className={`theme-toggle no-print ${modifierClass}`}
      aria-label="Dashboard theme selector"
    >
      <button
        type="button"
        className="active"
        onClick={() => toggleTheme?.()}
        aria-pressed={!isDark}
        aria-label={`Switch to ${isDark ? "light" : "dark"} theme`}
        title={`Switch to ${isDark ? "Light" : "Dark"} Theme`}
      >
        {isDark ? (
          <>
            <Moon size={18} strokeWidth={2.4} />
            <span>Dark Theme</span>
          </>
        ) : (
          <>
            <Sun size={18} strokeWidth={2.4} />
            <span>Light Theme</span>
          </>
        )}
      </button>
    </div>
  );

  // â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
  // TABLE COLUMN WIDTHS â€” edit the px values below to size each column.
  //
  //   â€¢ Each header has a clear `width =` entry.
  //   â€¢ Resizable columns: user can drag the right edge to adjust at runtime.
  //   â€¢ Wrap columns:      listed in *_WRAP_COLUMNS â€” no drag handle, text wraps.
  // â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•

  // â”€â”€ Tickets table â€” one entry per header in DISTINCT_REPORT_HEADERS â”€â”€â”€â”€â”€â”€â”€
  const TICKETS_COLUMN_WIDTHS: Record<string, number> = {
    "#": 56, // width =
    "Site ID": 110, // width =
    "Site Name": 220, // width =
    "Managed Resource": 150, // width =
    Severity: 110, // width =
    Issues: 430, // width =
    "Observation Date": 130, // width =
    "Observation Time": 100, // width =
    "Recovery Date": 130, // width =
    "Recovery Time": 100, // width =
    "Escalated for L3 Support Date": 160, // width =
    "Escalated for L3 Support Time": 130, // width =
    "Total Duration Days/Hours": 250, // width =
    TT: 90, // width =
    Status: 110, // width =
    "Escalated to": 130, // width =
    RCA: 360, // width =
    Action: 400, // width =
  };

  // â”€â”€ Performance table â€” one entry per header in PERF_REPORT_HEADERS â”€â”€â”€â”€â”€â”€â”€
  const PERF_COLUMN_WIDTHS: Record<string, number> = {
    "S No": 20, // width =
    "Site ID": 80, // width =
    "Site Name": 200, // width =
    "Site Availability, Hrs": 120, // width =
    "Site Availability, days": 120, // width =
    "Channel Busy Count": 100, // width =
    "MW link Performance, Hrs": 100, // width =
    "DMR Reliability": 120, // width =
    "Sites Down, hrs": 120, // width =
  };

  // â”€â”€ Runtime override maps â€” populated by drag handles, fall through to config above
  const [ticketColumnWidths, setTicketColumnWidths] = useState<
    Record<string, number>
  >({});
  const [perfColumnWidths, setPerfColumnWidths] = useState<
    Record<string, number>
  >({});

  const getTicketColumnWidth = (h: string): number =>
    ticketColumnWidths[h] ?? TICKETS_COLUMN_WIDTHS[h] ?? 130;
  const getPerfColumnWidth = (h: string): number =>
    perfColumnWidths[h] ?? PERF_COLUMN_WIDTHS[h] ?? 130;

  // Shared resize handler factory â€” wires a header name to the right state setter.
  function createResizeHandler(
    setWidths: React.Dispatch<React.SetStateAction<Record<string, number>>>,
    getCurrent: (h: string) => number,
  ) {
    return (header: string) => (e: React.MouseEvent) => {
      e.preventDefault();
      e.stopPropagation();
      const startX = e.clientX;
      const startW = getCurrent(header);
      const onMove = (ev: MouseEvent) => {
        const next = Math.max(40, startW + (ev.clientX - startX));
        setWidths((prev) => ({ ...prev, [header]: next }));
      };
      const onUp = () => {
        window.removeEventListener("mousemove", onMove);
        window.removeEventListener("mouseup", onUp);
        document.body.style.cursor = "";
        document.body.style.userSelect = "";
      };
      window.addEventListener("mousemove", onMove);
      window.addEventListener("mouseup", onUp);
      document.body.style.cursor = "col-resize";
      document.body.style.userSelect = "none";
    };
  }
  const startTicketColumnResize = createResizeHandler(
    setTicketColumnWidths,
    getTicketColumnWidth,
  );
  const startPerfColumnResize = createResizeHandler(
    setPerfColumnWidths,
    getPerfColumnWidth,
  );

  function fileExtension(fileName: string): string {
    const match = fileName.toLowerCase().match(/\.[^.]+$/);
    return match?.[0] ?? "";
  }

  function validateSelectedFiles(files?: FileList | File[] | null): File[] {
    const selectedFiles = Array.from(files ?? []);
    if (!selectedFiles.length) return [];

    const invalidFiles = selectedFiles.filter(
      (file) => !ALLOWED_UPLOAD_EXTENSIONS.includes(fileExtension(file.name) as (typeof ALLOWED_UPLOAD_EXTENSIONS)[number]),
    );
    if (invalidFiles.length) {
      throw new Error("Invalid file type. Please upload .xlsx, .xls, or .csv only.");
    }

    const emptyFiles = selectedFiles.filter((file) => file.size <= 0);
    if (emptyFiles.length) {
      throw new Error("One or more selected files are empty. Please upload a valid workbook.");
    }

    const oversizedFiles = selectedFiles.filter(
      (file) => file.size > MAX_UPLOAD_FILE_SIZE_BYTES,
    );
    if (oversizedFiles.length) {
      throw new Error("One or more selected files are larger than 50 MB. Please reduce the file size and try again.");
    }

    return selectedFiles;
  }

  function safeDashboardError(error: unknown): string {
    if (import.meta.env.DEV) {
      console.error("Dashboard processing error:", error);
    }
    if (error instanceof Error) {
      const message = error.message.trim();
      if (
        message.startsWith("Invalid file type") ||
        message.includes("larger than 50 MB") ||
        message.includes("empty") ||
        message.includes("Missing required columns") ||
        message.includes("Required workbook sheet") ||
        message.includes("No ticket rows")
      ) {
        return message;
      }
    }
    return SAFE_UPLOAD_ERROR_MESSAGE;
  }

  function parseWorkbookObject(
    workbook: XLSX.WorkBook,
    fileName: string,
  ): DashboardData {
    const parsed = parseRows(workbook, fileName);

    if (!parsed.rows.length || !parsed.uniqueTickets.length) {
      throw new Error(
        `Required workbook sheet or columns are missing. Required columns: ${REQUIRED_UPLOAD_COLUMNS.join(", ")}.`,
      );
    }

    return parsed;
  }

  async function parseWorkbookBuffer(
    buffer: ArrayBuffer,
    fileName: string,
  ): Promise<DashboardData> {
    try {
      const workbook = XLSX.read(buffer, { type: "array", cellDates: false });
      return parseWorkbookObject(workbook, fileName);
    } catch (error) {
      if (import.meta.env.DEV) {
        console.error(`Workbook parse failed for ${fileName}:`, error);
      }
      throw new Error(SAFE_UPLOAD_ERROR_MESSAGE);
    }
  }

  async function parseWorkbookFile(file: File): Promise<DashboardData> {
    return parseWorkbookBuffer(await file.arrayBuffer(), file.name);
  }

  async function parseUploadedWorkbookFile(
    file: File,
  ): Promise<UploadedWorkbookSource> {
    const workbook = XLSX.read(await file.arrayBuffer(), {
      type: "array",
      cellDates: false,
    });
    const parsed = parseWorkbookObject(workbook, file.name);
    return { fileName: file.name, workbook, data: parsed };
  }

  function extractGoogleSpreadsheetId(value: string): string {
    const text = value.trim();
    const match = text.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    if (match?.[1]) return match[1];
    if (/^[a-zA-Z0-9-_]{20,}$/.test(text)) return text;
    return "";
  }

  function extractGoogleSheetGid(value: string): string {
    const text = value.trim();
    const match = text.match(/[?#&]gid=(\d+)/);
    return match?.[1] ?? "";
  }

  function isMicrosoftOnlineSheetLink(value: string): boolean {
    const text = value.trim().toLowerCase();
    return (
      text.includes("sharepoint.com") ||
      text.includes("onedrive.live.com") ||
      text.includes("1drv.ms") ||
      text.includes("office.com")
    );
  }

  function microsoftShareId(value: string): string {
    const bytes = new TextEncoder().encode(value.trim());
    let binary = "";
    bytes.forEach((byte) => {
      binary += String.fromCharCode(byte);
    });
    return `u!${btoa(binary)
      .replace(/\+/g, "-")
      .replace(/\//g, "_")
      .replace(/=+$/, "")}`;
  }

  function withMicrosoftDownloadParam(value: string): string {
    const url = new URL(value.trim());
    url.searchParams.set("download", "1");
    return url.toString();
  }

  async function fetchWorkbookArrayBuffer(
    urls: string[],
  ): Promise<ArrayBuffer> {
    const errors: string[] = [];
    for (const url of urls) {
      try {
        const response = await fetch(url, { redirect: "follow" });
        if (!response.ok) {
          errors.push(`${response.status} ${response.statusText}`);
          continue;
        }

        const contentType = response.headers.get("content-type") ?? "";
        const buffer = await response.arrayBuffer();
        const firstBytes = new Uint8Array(buffer.slice(0, 4));
        const isZipWorkbook =
          firstBytes[0] === 0x50 &&
          firstBytes[1] === 0x4b &&
          firstBytes[2] === 0x03 &&
          firstBytes[3] === 0x04;
        if (
          isZipWorkbook &&
          !contentType.includes("text/html") &&
          !contentType.includes("application/json")
        ) {
          return buffer;
        }
        errors.push(`not workbook (${contentType || "unknown content"})`);
      } catch (error) {
        errors.push(error instanceof Error ? error.message : String(error));
      }
    }

    throw new Error(errors.filter(Boolean).slice(-2).join(" | "));
  }

  function googleRegionLabelToKey(label: string): GoogleRegionKey | null {
    const normalized = normalizeHeader(label);
    if (normalized === "eoaneoa") return "eoaNeoa";
    if (normalized === "soa") return "soa";
    if (normalized === "coawoa") return "coaWoa";
    return null;
  }

  function parseGoogleSheetsConfig(text: string): {
    links: Partial<GoogleRegionLinks>;
    saveLinks: Partial<GoogleRegionSaveLinks>;
  } {
    const links: Partial<GoogleRegionLinks> = {};
    const saveLinks: Partial<GoogleRegionSaveLinks> = {};
    text.split(/\r?\n/).forEach((line) => {
      const trimmed = line.trim();
      if (!trimmed || trimmed.startsWith("#")) return;
      const separator = trimmed.indexOf("=");
      if (separator === -1) return;
      const rawKey = trimmed.slice(0, separator).trim();
      const isSaveKey = /_SAVE$/i.test(rawKey);
      const regionLabel = rawKey.replace(/_SAVE$/i, "");
      const key = googleRegionLabelToKey(regionLabel);
      if (!key) return;
      const value = trimmed.slice(separator + 1).trim();
      if (isSaveKey) saveLinks[key] = value;
      else links[key] = value;
    });
    return { links, saveLinks };
  }

  function parseMicrosoftGraphConfig(text: string): MicrosoftGraphConfig {
    const next = { ...EMPTY_MICROSOFT_GRAPH_CONFIG };
    text.split(/\r?\n/).forEach((line) => {
      const trimmed = line.trim();
      if (!trimmed || trimmed.startsWith("#")) return;
      const separator = trimmed.indexOf("=");
      if (separator === -1) return;
      const key = normalizeHeader(trimmed.slice(0, separator));
      const value = trimmed.slice(separator + 1).trim();
      if (key === "clientid") next.clientId = value;
      if (key === "tenantid") next.tenantId = value || "common";
      if (key === "scopes") next.scopes = value || next.scopes;
    });
    return next;
  }

  function base64UrlEncode(input: ArrayBuffer | Uint8Array | string): string {
    const bytes =
      typeof input === "string"
        ? new TextEncoder().encode(input)
        : input instanceof Uint8Array
          ? input
          : new Uint8Array(input);
    let binary = "";
    bytes.forEach((byte) => {
      binary += String.fromCharCode(byte);
    });
    return btoa(binary)
      .replace(/\+/g, "-")
      .replace(/\//g, "_")
      .replace(/=+$/, "");
  }

  function randomPkceValue(byteLength = 32): string {
    const bytes = new Uint8Array(byteLength);
    crypto.getRandomValues(bytes);
    return base64UrlEncode(bytes);
  }

  async function pkceChallenge(verifier: string): Promise<string> {
    const digest = await crypto.subtle.digest(
      "SHA-256",
      new TextEncoder().encode(verifier),
    );
    return base64UrlEncode(digest);
  }

  async function microsoftGraphAccessToken(): Promise<string> {
    const { clientId, tenantId, scopes } = microsoftGraphConfig;
    if (!clientId.trim()) {
      throw new Error(
        "Microsoft Graph login is not configured. Add clientId and tenantId to public/microsoft-graph-config.txt.",
      );
    }

    const verifier = randomPkceValue(64);
    const challenge = await pkceChallenge(verifier);
    const state = randomPkceValue(16);
    const redirectUri = `${window.location.origin}/`;
    const tenant = tenantId.trim() || "common";
    const authUrl = new URL(
      `https://login.microsoftonline.com/${encodeURIComponent(tenant)}/oauth2/v2.0/authorize`,
    );
    authUrl.searchParams.set("client_id", clientId.trim());
    authUrl.searchParams.set("response_type", "code");
    authUrl.searchParams.set("redirect_uri", redirectUri);
    authUrl.searchParams.set("response_mode", "query");
    authUrl.searchParams.set("scope", scopes.trim());
    authUrl.searchParams.set("state", state);
    authUrl.searchParams.set("code_challenge", challenge);
    authUrl.searchParams.set("code_challenge_method", "S256");
    authUrl.searchParams.set("prompt", "select_account");

    const popup = window.open(
      authUrl.toString(),
      "microsoft-graph-login",
      "popup,width=560,height=720",
    );
    if (!popup) {
      throw new Error("Microsoft sign-in popup was blocked by the browser.");
    }

    const code = await new Promise<string>((resolve, reject) => {
      const started = Date.now();
      const timer = window.setInterval(() => {
        try {
          if (popup.closed) {
            window.clearInterval(timer);
            reject(
              new Error("Microsoft sign-in was closed before it finished."),
            );
            return;
          }
          if (Date.now() - started > 120000) {
            window.clearInterval(timer);
            popup.close();
            reject(new Error("Microsoft sign-in timed out."));
            return;
          }
          if (popup.location.origin !== window.location.origin) return;
          const params = new URLSearchParams(popup.location.search);
          const returnedState = params.get("state");
          const authCode = params.get("code");
          const authError =
            params.get("error_description") ?? params.get("error");
          if (authError) {
            window.clearInterval(timer);
            popup.close();
            reject(new Error(authError));
            return;
          }
          if (authCode) {
            window.clearInterval(timer);
            popup.close();
            if (returnedState !== state) {
              reject(new Error("Microsoft sign-in state check failed."));
              return;
            }
            resolve(authCode);
          }
        } catch {
          // Cross-origin access throws until Microsoft redirects back here.
        }
      }, 500);
    });

    const tokenResponse = await fetch(
      `https://login.microsoftonline.com/${encodeURIComponent(tenant)}/oauth2/v2.0/token`,
      {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: new URLSearchParams({
          client_id: clientId.trim(),
          grant_type: "authorization_code",
          code,
          redirect_uri: redirectUri,
          code_verifier: verifier,
          scope: scopes.trim(),
        }),
      },
    );
    const payload = await tokenResponse.json();
    if (!tokenResponse.ok || !payload?.access_token) {
      throw new Error(
        payload?.error_description ||
          payload?.error ||
          "Microsoft Graph token request failed.",
      );
    }
    return payload.access_token;
  }

  async function fetchMicrosoftWorkbookWithGraph(
    url: string,
  ): Promise<ArrayBuffer> {
    const accessToken = await microsoftGraphAccessToken();
    const response = await fetch(
      `https://graph.microsoft.com/v1.0/shares/${microsoftShareId(url)}/driveItem/content`,
      {
        headers: { Authorization: `Bearer ${accessToken}` },
        redirect: "follow",
      },
    );
    if (!response.ok) {
      throw new Error(`${response.status} ${response.statusText}`.trim());
    }
    return response.arrayBuffer();
  }

  function setGoogleRegionLink(key: GoogleRegionKey, value: string) {
    setGoogleRegionLinks((prev) => ({ ...prev, [key]: value }));
  }

  function setGoogleRegionChecked(key: GoogleRegionKey, checked: boolean) {
    setGoogleRegionSelection((prev) => ({ ...prev, [key]: checked }));
  }

  function loadGoogleSheetTable(
    spreadsheetId: string,
    options: { gid?: string; sheetName?: string },
  ): Promise<any> {
    return new Promise((resolve, reject) => {
      const callbackName = `__googleSheetCallback_${Date.now()}_${Math.random()
        .toString(36)
        .slice(2)}`;
      const script = document.createElement("script");
      const cleanup = () => {
        script.remove();
        delete (window as any)[callbackName];
      };

      (window as any)[callbackName] = (payload: any) => {
        cleanup();
        resolve(payload);
      };

      script.onerror = () => {
        cleanup();
        reject(
          new Error(
            `Could not read ${options.sheetName || "TT-History"} from Google Sheets. Share the file as 'Anyone with the link can view' and try again.`,
          ),
        );
      };

      const tabSelector = options.gid
        ? `gid=${encodeURIComponent(options.gid)}`
        : `sheet=${encodeURIComponent(options.sheetName || "TT-History")}`;
      script.src =
        `https://docs.google.com/spreadsheets/d/${spreadsheetId}/gviz/tq?` +
        `tqx=responseHandler:${callbackName}&${tabSelector}&headers=0`;
      document.body.appendChild(script);
    });
  }

  function googleCellValue(cell: any): unknown {
    const value = cell?.f ?? cell?.v ?? "";
    if (typeof value !== "string") return value;
    const match = value.match(
      /^Date\((\d+),(\d+),(\d+)(?:,(\d+),(\d+),(\d+))?\)$/,
    );
    if (!match) return value;

    const year = Number(match[1]);
    const month = Number(match[2]);
    const day = Number(match[3]);
    const hour = Number(match[4] ?? 0);
    const minute = Number(match[5] ?? 0);
    const second = Number(match[6] ?? 0);
    return new Date(year, month, day, hour, minute, second);
  }

  function googlePayloadToAoa(payload: any): unknown[][] {
    const table = payload?.table;
    const cols = table?.cols ?? [];
    const rows = table?.rows ?? [];
    if (!cols.length && !rows.length) return [];
    const values = rows.map((row: any) =>
      cols.map((_col: any, index: number) => googleCellValue(row?.c?.[index])),
    );
    const looksLikeTicketHeader = (row: unknown[]) => {
      const normalized = row.map((value) => normalizeHeader(String(value)));
      return (
        normalized.includes("tt") ||
        (normalized.includes("siteid") && normalized.includes("sitename")) ||
        (normalized.includes("severity") && normalized.includes("region"))
      );
    };
    const headerRowIndex = values.findIndex(looksLikeTicketHeader);
    let headers = cols.map((col: any) => clean(col?.label || col?.id || ""));
    let dataRows = values;
    if (headerRowIndex >= 0) {
      headers = values[headerRowIndex].map(
        (value: unknown, index: number) =>
          clean(value) || GOOGLE_TT_HISTORY_HEADERS[index] || "",
      );
      dataRows = values.slice(headerRowIndex + 1);
    }
    if (!headers.some((header: unknown) => clean(header))) {
      headers = GOOGLE_TT_HISTORY_HEADERS.slice(0, cols.length);
    }
    return [
      headers,
      ...dataRows.map((row: unknown[]) =>
        headers.map((_header: unknown, index: number) => row[index] ?? ""),
      ),
    ];
  }

  type GoogleSheetTicketRow = {
    rowNumber: number;
    headers: string[];
    values: unknown[];
  };

  function googlePayloadToTicketRows(payload: any): GoogleSheetTicketRow[] {
    const table = payload?.table;
    const cols = table?.cols ?? [];
    const rows = table?.rows ?? [];
    if (!cols.length && !rows.length) return [];
    const values = rows.map((row: any) =>
      cols.map((_col: any, index: number) => googleCellValue(row?.c?.[index])),
    );
    const looksLikeTicketHeader = (row: unknown[]) => {
      const normalized = row.map((value) => normalizeHeader(String(value)));
      return (
        normalized.includes("tt") ||
        (normalized.includes("siteid") && normalized.includes("sitename")) ||
        (normalized.includes("severity") && normalized.includes("region"))
      );
    };
    const headerRowIndex = values.findIndex(looksLikeTicketHeader);
    let headers = cols.map((col: any) => clean(col?.label || col?.id || ""));
    let dataRows = values;
    let firstDataSheetRow = 1;
    if (headerRowIndex >= 0) {
      headers = values[headerRowIndex].map(
        (value: unknown, index: number) =>
          clean(value) || GOOGLE_TT_HISTORY_HEADERS[index] || "",
      );
      dataRows = values.slice(headerRowIndex + 1);
      firstDataSheetRow = headerRowIndex + 2;
    } else if (!headers.some((header: unknown) => clean(header))) {
      headers = GOOGLE_TT_HISTORY_HEADERS.slice(0, cols.length);
      firstDataSheetRow = 2;
    }
    return dataRows.map((row: unknown[], index: number) => ({
      rowNumber: firstDataSheetRow + index,
      headers,
      values: headers.map((_header: unknown, cellIndex: number) => row[cellIndex] ?? ""),
    }));
  }

  function sheetRowValue(row: GoogleSheetTicketRow, names: string[]): string {
    const normalizedNames = names.map((name) => normalizeHeader(name));
    const index = row.headers.findIndex((header) =>
      normalizedNames.includes(normalizeHeader(header)),
    );
    return index >= 0 ? clean(row.values[index]) : "";
  }

  function dateForManualInput(value: string): string {
    const text = clean(value);
    if (!text) return "";
    const iso = text.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (iso) return text;
    const dmy = text.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2}|\d{4})$/);
    if (!dmy) return text;
    const day = dmy[1].padStart(2, "0");
    const month = dmy[2].padStart(2, "0");
    const year = dmy[3].length === 2 ? `20${dmy[3]}` : dmy[3];
    return `${year}-${month}-${day}`;
  }
  function googlePayloadDataAoa(payload: any): unknown[][] {
    const table = payload?.table;
    const cols = table?.cols ?? [];
    const rows = table?.rows ?? [];
    return rows.map((row: any) =>
      cols.map((_col: any, index: number) => googleCellValue(row?.c?.[index])),
    );
  }
  async function loadOptionalGoogleSheetAoa(
    spreadsheetId: string,
    sheetName: string,
  ): Promise<unknown[][]> {
    try {
      const payload = await loadGoogleSheetTable(spreadsheetId, { sheetName });
      if (payload?.status && payload.status !== "ok") return [];
      return googlePayloadDataAoa(payload);
    } catch {
      return [];
    }
  }
  async function parseGoogleSheetLink(value: string): Promise<DashboardData> {
    const spreadsheetId = extractGoogleSpreadsheetId(value);
    if (!spreadsheetId) {
      throw new Error("Paste a valid Google Sheets link or spreadsheet ID.");
    }

    const gid = extractGoogleSheetGid(value);
    if (!gid) {
      throw new Error(
        "Open the TT-History tab in Google Sheets, then copy/paste the full browser URL. The link must include gid= so the dashboard can load the correct tab.",
      );
    }
    const payload = await loadGoogleSheetTable(spreadsheetId, { gid });
    if (payload?.status && payload.status !== "ok") {
      throw new Error(
        "Could not read TT-History from Google Sheets. Share the file as 'Anyone with the link can view' and try again.",
      );
    }

    const ttAoa = googlePayloadToAoa(payload);
    if (!ttAoa.length || ttAoa.length < 2) {
      throw new Error(
        "TT-History is empty or could not be read from Google Sheets.",
      );
    }

    const headers = ttAoa[0].map((value) => clean(value));
    const hasTicketHeader = headers.some((header) =>
      ["tt", "ttnumber", "ttno", "ticket", "ticketnumber"].includes(
        normalizeHeader(header),
      ),
    );
    if (!hasTicketHeader && headers.length > 1) {
      headers[1] = "TT";
      ttAoa[0] = headers;
    }

    const [siteIdAoa, rcaAoa] = await Promise.all([
      loadOptionalGoogleSheetAoa(spreadsheetId, "Site ID"),
      loadOptionalGoogleSheetAoa(spreadsheetId, "RCA"),
    ]);

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.aoa_to_sheet(ttAoa),
      "TT-History",
    );
    if (siteIdAoa.length) {
      XLSX.utils.book_append_sheet(
        workbook,
        XLSX.utils.aoa_to_sheet(siteIdAoa),
        "Site ID",
      );
    }
    if (rcaAoa.length) {
      XLSX.utils.book_append_sheet(
        workbook,
        XLSX.utils.aoa_to_sheet(rcaAoa),
        "RCA",
      );
    }

    const parsed = parseRows(workbook, "Google Sheet TT-History");
    if (!parsed.rows.length || !parsed.uniqueTickets.length) {
      throw new Error(
        "TT-History was loaded from Google Sheets, but no ticket rows with TT numbers were found.",
      );
    }

    return parsed;
  }

  async function parseMicrosoftOnlineSheetLink(
    value: string,
  ): Promise<DashboardData> {
    const url = value.trim();
    if (!url) {
      throw new Error("Paste a valid Microsoft 365 or OneDrive Excel link.");
    }

    let proxyMessage = "";
    try {
      const proxyResponse = await fetch("/api/download-workbook", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ url }),
      });
      if (proxyResponse.ok) {
        const buffer = await proxyResponse.arrayBuffer();
        return parseWorkbookBuffer(buffer, "Microsoft 365 Online Workbook");
      }

      proxyMessage = `${proxyResponse.status} ${proxyResponse.statusText}`;
      try {
        const payload = await proxyResponse.json();
        proxyMessage = [payload?.error, payload?.details]
          .filter(Boolean)
          .join(" ");
      } catch {
        // Keep the HTTP status message when the proxy did not return JSON.
      }

      const buffer = await fetchWorkbookArrayBuffer([
        `https://api.onedrive.com/v1.0/shares/${microsoftShareId(url)}/root/content`,
        withMicrosoftDownloadParam(url),
      ]);
      return parseWorkbookBuffer(buffer, "Microsoft 365 Online Workbook");
    } catch (error) {
      const anonymousDetails = [
        proxyMessage,
        error instanceof Error ? error.message : String(error),
      ]
        .filter(Boolean)
        .join(" | ");
      try {
        const graphBuffer = await fetchMicrosoftWorkbookWithGraph(url);
        return parseWorkbookBuffer(graphBuffer, "Microsoft Graph Workbook");
      } catch (graphError) {
        throw new Error(
          "Could not download the Microsoft 365 workbook anonymously or with Microsoft Graph login. Check that public/microsoft-graph-config.txt has a valid clientId/tenantId, the app registration redirect URI is this dashboard URL, and Graph permissions allow file access. Details: " +
            [
              anonymousDetails,
              graphError instanceof Error
                ? graphError.message
                : String(graphError),
            ]
              .filter(Boolean)
              .join(" | "),
        );
      }
    }
  }

  async function parseOnlineSheetLink(value: string): Promise<DashboardData> {
    return isMicrosoftOnlineSheetLink(value)
      ? parseMicrosoftOnlineSheetLink(value)
      : parseGoogleSheetLink(value);
  }

  function mergeDashboardRegions(
    regionsToMerge: DashboardData[],
  ): DashboardData {
    const allRows = regionsToMerge.flatMap((region) => region.rows);
    const ttMap = new Map<string, TicketAggregate>();

    for (const row of allRows) {
      if (!row.tt) continue;
      const existing = ttMap.get(row.tt);

      if (!existing) {
        ttMap.set(row.tt, {
          tt: row.tt,
          primary: row,
          siteIds: new Set([row.siteId].filter(Boolean)),
          siteNames: new Set([row.siteName].filter(Boolean)),
          rows: [row],
        });
      } else {
        existing.rows.push(row);
        if (row.siteId) existing.siteIds.add(row.siteId);
        if (row.siteName) existing.siteNames.add(row.siteName);
      }
    }

    const mergedSiteOrder = [...(regionsToMerge[0]?.siteOrder ?? [])];
    const mergedIds = new Set(mergedSiteOrder.map((site) => site.siteId));
    const mergedRcaLookup: { action: string; rca: string }[] = [];
    const mergedRcaKeys = new Set<string>();
    regionsToMerge.forEach((region) => {
      (region.rcaLookup ?? []).forEach((item) => {
        const key = normalizeHeader(item.action);
        if (!key || mergedRcaKeys.has(key)) return;
        mergedRcaKeys.add(key);
        mergedRcaLookup.push(item);
      });
    });

    regionsToMerge.slice(1).forEach((region) => {
      region.siteOrder.forEach((site) => {
        if (!mergedIds.has(site.siteId)) {
          mergedIds.add(site.siteId);
          mergedSiteOrder.push(site);
        }
      });
    });

    return {
      fileName: regionsToMerge.map((region) => region.fileName).join(" + "),
      sheetName: regionsToMerge[0]?.sheetName ?? "Tickets_Data",
      generatedAt: new Date().toLocaleString(),
      rows: allRows,
      uniqueTickets: Array.from(ttMap.values()),
      siteOrder: mergedSiteOrder,
      rcaLookup: mergedRcaLookup,
    };
  }

  async function handleFile(files?: FileList | File[] | null) {
    const selectedFiles = validateSelectedFiles(files);
    if (!selectedFiles.length) return;

    setError("");
    try {
      const parsedFiles = await Promise.all(
        selectedFiles.map(parseWorkbookFile),
      );
      const merged = mergeDashboardRegions(parsedFiles);

      setData(merged);
      setRegions(parsedFiles);
      setExportMonths([]);
      setExportRegions([]);
      setPerfMonths([]);
      setPerfRegions([]);
      setFilters(EMPTY_FILTERS);
      setTablePage(1);
      setOnlineSourceMode(null);
      if (inputRef.current) inputRef.current.value = "";
    } catch (err) {
      setError(safeDashboardError(err));
    }
  }

  async function handleGoogleSheetLoad(mode: "replace" | "add" = "replace") {
    if (googleSheetLoading) return;
    const selectedLinks = GOOGLE_REGION_LINKS.filter(
      (region) => googleRegionSelection[region.key],
    ).map((region) => ({
      ...region,
      url: googleRegionLinks[region.key].trim(),
    }));
    if (!selectedLinks.length) {
      setError("Select at least one online sheet region to load.");
      return;
    }
    const missingLink = selectedLinks.find((region) => !region.url);
    if (missingLink) {
      setError(
        `Add the online sheet link for ${missingLink.label}, or uncheck it.`,
      );
      return;
    }

    setError("");
    setGoogleSheetLoading(true);
    try {
      const parsedSheets = await Promise.all(
        selectedLinks.map((region) => parseOnlineSheetLink(region.url)),
      );
      if (mode === "add" && data) {
        setRegions((prev) => {
          const updated = [...prev, ...parsedSheets];
          const merged = mergeDashboardRegions(updated);
          setData(merged);
          return updated;
        });
      } else {
        const merged = mergeDashboardRegions(parsedSheets);
        setData(merged);
        setRegions(parsedSheets);
        setExportMonths([]);
        setExportRegions([]);
        setPerfMonths([]);
        setPerfRegions([]);
        setFilters(EMPTY_FILTERS);
      }
      setTablePage(1);
      setPerfPage(1);
      setOnlineSourceMode(null);
    } catch (err) {
      setError(
        err instanceof Error
          ? err.message
          : "Unable to read the online sheet link.",
      );
    } finally {
      setGoogleSheetLoading(false);
    }
  }

  async function handleAddRegion(files?: FileList | File[] | null) {
    const selectedFiles = validateSelectedFiles(files);
    if (!selectedFiles.length) return;

    setError("");
    try {
      const parsedFiles = await Promise.all(
        selectedFiles.map(parseWorkbookFile),
      );

      setRegions((prev) => {
        const updated = [...prev, ...parsedFiles];
        const merged = mergeDashboardRegions(updated);
        setData(merged);
        return updated;
      });

      setTablePage(1);
      setOnlineSourceMode(null);
      if (addRegionRef.current) addRegionRef.current.value = "";
    } catch (err) {
      setError(safeDashboardError(err));
    }
  }

  const workbookRows = data?.rows ?? [];
  const manualTicketRows = useMemo(
    () =>
      manualDrafts
        .map((draft, index) => manualDraftToTicketRecord(draft, index))
        .filter((row): row is TicketRecord => Boolean(row)),
    [manualDrafts],
  );
  const allDataRows = useMemo(
    () => [...workbookRows, ...manualTicketRows],
    [workbookRows, manualTicketRows],
  );
  const uniqueRows = useMemo(() => groupTickets(allDataRows), [allDataRows]);
  const filterOptions = useMemo(() => {
    const primaryRows = uniqueRows.map((ticket) => ticket.primary);
    const uniq = (field: keyof TicketRecord) =>
      Array.from(
        new Set(primaryRows.map((row) => clean(row[field])).filter(Boolean)),
      ).sort();
    const openingMonths = Array.from(
      new Set(
        primaryRows
          .map(
            (row) =>
              row.openingMonthKey || openingMonthKey(row.observationDate),
          )
          .filter(Boolean),
      ),
    ).sort((a, b) => {
      if (a === "Unknown") return 1;
      if (b === "Unknown") return -1;
      return a.localeCompare(b);
    });
    const exportMonths = Array.from(
      new Set(
        uniqueRows
          .flatMap((ticket) =>
            ticket.rows.flatMap((row) => coveredMonthKeys(row)),
          )
          .filter((key) => key && key !== "Unknown"),
      ),
    ).sort((a, b) => a.localeCompare(b));
    const rcaFamilyOptions = Array.from(
      new Set(
        primaryRows
          .map((row) => row.rcaFamily || getRcaFamily(row.rca))
          .filter(Boolean),
      ),
    ).sort() as string[];
    return {
      status: uniq("status"),
      severity: uniq("severity"),
      region: uniq("region"),
      impact: uniq("impact"),
      site: uniq("siteId"),
      openingMonth: openingMonths,
      openingMonthLabels: Object.fromEntries(
        openingMonths.map((key) => [key, openingMonthLabel(key)]),
      ),
      exportMonth: exportMonths,
      exportMonthLabels: Object.fromEntries(
        exportMonths.map((key) => [key, openingMonthLabel(key)]),
      ),
      rcaFamily: rcaFamilyOptions,
    };
  }, [uniqueRows]);

  // Reset page when filters change
  useEffect(() => {
    setTablePage(1);
  }, [filters]);
  useEffect(() => {
    setPerfPage(1);
  }, [perfMonths, perfRegions, filters.openingMonth]);

  const filteredTickets = useMemo(() => {
    const q = filters.search.toLowerCase().trim();
    return uniqueRows.filter((ticket) => {
      const row = ticket.primary;
      const allSites = Array.from(ticket.siteIds).join(" ");
      const allSiteNames = Array.from(ticket.siteNames).join(" ");
      const haystack = [
        ticket.tt,
        row.siteId,
        row.siteName,
        allSites,
        allSiteNames,
        row.managedResource,
        row.issue,
        row.status,
        row.severity,
        row.region,
        row.impact,
        row.escalationLevel,
        row.escalatedTo,
        row.rca,
        row.rcaFamily || getRcaFamily(row.rca),
        row.responsibleTeam ||
          getResponsibleTeam(row.rcaFamily || getRcaFamily(row.rca)),
        row.escalatedForL3SupportDate,
        row.escalatedForL3SupportTime,
      ]
        .join(" ")
        .toLowerCase();
      return (
        (!q || haystack.includes(q)) &&
        (!filters.status.length || filters.status.includes(row.status)) &&
        (!filters.severity.length || filters.severity.includes(row.severity)) &&
        (!filters.region.length || filters.region.includes(row.region)) &&
        (!filters.impact.length || filters.impact.includes(row.impact)) &&
        (!filters.openingMonth.length ||
          filters.openingMonth.includes(
            row.openingMonthKey || openingMonthKey(row.observationDate),
          )) &&
        (!filters.site.length ||
          filters.site.some((s) => ticket.siteIds.has(s))) &&
        (!filters.rcaFamily.length ||
          filters.rcaFamily.includes(row.rcaFamily || getRcaFamily(row.rca)))
      );
    });
  }, [filters, uniqueRows]);

  const monthlyExportTickets = useMemo(() => {
    const sourceRows = allDataRows.filter((row) => {
      const matchesReportRegion =
        !exportRegions.length || exportRegions.includes(row.region);
      const matchesReportMonth =
        !exportMonths.length ||
        exportMonths.some((m) =>
          ticketMatchesMonthlyExport(
            {
              tt: row.tt,
              primary: row,
              siteIds: new Set(row.siteId ? [row.siteId] : []),
              siteNames: new Set(row.siteName ? [row.siteName] : []),
              rows: [row],
            },
            m,
          ),
        );
      return matchesReportRegion && matchesReportMonth;
    });
    return groupTickets(sourceRows);
  }, [allDataRows, exportMonths, exportRegions]);

  const monthlyExcelExportGroups = useMemo(() => {
    const selectedRegions = exportRegions.length
      ? new Set(exportRegions)
      : null;
    const sourceRows = regions.length
      ? regions.flatMap((region) => region.rows)
      : allDataRows;
    const regionOrder = ["EOA", "SOA", "COA", "WOA"] as const;
    const normalizeTemplateRegion = (region: string) => {
      const value = clean(region).toUpperCase();
      if (value === "EOA" || value === "NEOA") return "EOA";
      if (value === "SOA") return "SOA";
      if (value === "COA") return "COA";
      if (value === "WOA") return "WOA";
      return value || "EOA";
    };
    const rowsByTemplateRegion = new Map<string, TicketRecord[]>();

    sourceRows.forEach((row) => {
      const matchesReportRegion =
        !selectedRegions || selectedRegions.has(row.region);
      const matchesReportMonth =
        !exportMonths.length ||
        exportMonths.some((m) =>
          ticketMatchesMonthlyExport(
            {
              tt: row.tt,
              primary: row,
              siteIds: new Set(row.siteId ? [row.siteId] : []),
              siteNames: new Set(row.siteName ? [row.siteName] : []),
              rows: [row],
            },
            m,
          ),
        );
      if (!matchesReportRegion || !matchesReportMonth) return;

      const key = normalizeTemplateRegion(row.region);
      rowsByTemplateRegion.set(key, [
        ...(rowsByTemplateRegion.get(key) ?? []),
        row,
      ]);
    });

    return regionOrder
      .map((region) => ({
        region,
        tickets: groupTickets(rowsByTemplateRegion.get(region) ?? []),
      }))
      .filter((group) => group.tickets.length);
  }, [allDataRows, exportMonths, exportRegions, regions]);

  const selectedExportMonthLabel =
    exportMonths.length === 0
      ? "All export-eligible TT"
      : exportMonths.length === 1
        ? openingMonthLabel(exportMonths[0])
        : `${exportMonths.length} months`;

  // Monthly Performance rows -- computed from raw rows (not aggregated tickets)
  const siteOrder = data?.siteOrder ?? [];
  const buildPerformanceSiteOrder = (
    selectedRegions: string[] = [],
    selectedSites: string[] = [],
  ): PerformanceSiteOrderEntry[] => {
    const selectedRegionSet = new Set(
      selectedRegions.map((region) => normalizePerformanceRegionLabel(region)),
    );
    const selectedSiteSet = new Set(
      selectedSites.map((site) => normalizeSiteId(clean(site)).toUpperCase()),
    );
    const sourceDatasets = regions.length ? regions : data ? [data] : [];
    const entries: PerformanceSiteOrderEntry[] = [];
    const seen = new Set<string>();

    sourceDatasets.forEach((dataset) => {
      const sourceLabel =
        dataset.rows
          .map((row) => normalizePerformanceRegionLabel(row.region))
          .find(Boolean) ||
        normalizePerformanceRegionLabel(dataset.fileName) ||
        "Workbook";
      const rowRegions = new Set(
        dataset.rows
          .map((row) => normalizePerformanceRegionLabel(row.region))
          .filter(Boolean),
      );
      const matchesRegion =
        selectedRegionSet.size === 0 ||
        selectedRegionSet.has(sourceLabel) ||
        Array.from(rowRegions).some((region) => selectedRegionSet.has(region));
      if (!matchesRegion) return;

      dataset.siteOrder.forEach((site) => {
        const siteId = normalizeSiteId(clean(site.siteId));
        if (!siteId) return;
        if (selectedSiteSet.size && !selectedSiteSet.has(siteId.toUpperCase()))
          return;
        const key = perfEntryKey(sourceLabel, siteId);
        if (seen.has(key)) return;
        seen.add(key);
        entries.push({ siteId, siteName: site.siteName, sourceLabel });
      });
    });

    return entries.sort((a, b) => comparePerformanceSiteRows(a, b));
  };
  const performanceReportSiteOrder = buildPerformanceSiteOrder(perfRegions);

  const perfRows = useMemo(() => {
    // Filter by region
    const sourceRows =
      perfRegions.length === 0
        ? allDataRows
        : allDataRows.filter((r) => perfRegions.includes(r.region));
    // For multi-month: pass "all" if none selected, or the first month if one selected,
    // or compute combined rows for multiple months
    if (perfMonths.length === 0) {
      return computePerfRows(sourceRows, "all", performanceReportSiteOrder);
    } else if (perfMonths.length === 1) {
      return computePerfRows(
        sourceRows,
        perfMonths[0],
        performanceReportSiteOrder,
      );
    } else {
      // Sum down hours across all selected months per site
      const combined = new Map<string, PerfRow>();
      perfMonths.forEach((mk) => {
        const rows = computePerfRows(
          sourceRows,
          mk,
          performanceReportSiteOrder,
        );
        rows.forEach((r) => {
          const existing = combined.get(r.perfKey);
          if (!existing) {
            combined.set(r.perfKey, { ...r });
          } else {
            existing.sitesDownHours =
              Math.round((existing.sitesDownHours + r.sitesDownHours) * 10) /
              10;
            // availHours: recalculate based on total month hours across selected months
          }
        });
      });
      // Recalculate availHours for combined rows
      const totalMonthHours = perfMonths.reduce(
        (s, mk) => s + totalHoursInMonth(mk),
        0,
      );
      return Array.from(combined.values())
        .map((r) => {
          const availHours = Math.max(0, totalMonthHours - r.sitesDownHours);
          const totalHrs = availHours + r.sitesDownHours;
          const reliability = totalHrs > 0 ? availHours / totalHrs : 1;
          const totalMins = Math.round(availHours * 60);
          const dDays = Math.floor(totalMins / (60 * 24));
          const dHrs = Math.floor((totalMins % (60 * 24)) / 60);
          const dMins = Math.round(totalMins % 60);
          return {
            ...r,
            displayName: formatPerformanceChartLabel(
              r.siteId,
              r.siteName,
              r.sourceLabel,
            ),
            availHours: Math.round(availHours * 10) / 10,
            availDay: `${dDays} d, ${dHrs} h, ${dMins} m`,
            reliability: `${(reliability * 100).toFixed(2)}%`,
          };
        })
        .sort((a, b) => comparePerformanceSiteRows(a, b));
    }
  }, [allDataRows, perfMonths, perfRegions, performanceReportSiteOrder]);

  const performanceDashboardSiteOrder = buildPerformanceSiteOrder(
    filters.region,
    filters.site,
  );

  const performanceKpiRows = useMemo(() => {
    const q = filters.search.toLowerCase().trim();
    const sourceRows = allDataRows.filter((row) => {
      const haystack = [
        row.tt,
        row.siteId,
        row.siteName,
        row.managedResource,
        row.issue,
        row.status,
        row.severity,
        row.region,
        row.impact,
        row.escalationLevel,
        row.escalatedTo,
        row.rca,
        row.rcaFamily || getRcaFamily(row.rca),
        row.responsibleTeam ||
          getResponsibleTeam(row.rcaFamily || getRcaFamily(row.rca)),
        row.escalatedForL3SupportDate,
        row.escalatedForL3SupportTime,
      ]
        .join(" ")
        .toLowerCase();
      return (
        (!q || haystack.includes(q)) &&
        (!filters.status.length || filters.status.includes(row.status)) &&
        (!filters.severity.length || filters.severity.includes(row.severity)) &&
        (!filters.region.length || filters.region.includes(row.region)) &&
        (!filters.impact.length || filters.impact.includes(row.impact)) &&
        (!filters.openingMonth.length ||
          filters.openingMonth.includes(
            row.openingMonthKey || openingMonthKey(row.observationDate),
          )) &&
        (!filters.site.length || filters.site.includes(row.siteId)) &&
        (!filters.rcaFamily.length ||
          filters.rcaFamily.includes(row.rcaFamily || getRcaFamily(row.rca)))
      );
    });

    if (filters.openingMonth.length === 0) {
      return computePerfRows(sourceRows, "all", performanceDashboardSiteOrder);
    }
    if (filters.openingMonth.length === 1) {
      return computePerfRows(
        sourceRows,
        filters.openingMonth[0],
        performanceDashboardSiteOrder,
      );
    }

    const combined = new Map<string, PerfRow>();
    filters.openingMonth.forEach((monthKey) => {
      computePerfRows(
        sourceRows,
        monthKey,
        performanceDashboardSiteOrder,
      ).forEach((row) => {
        const existing = combined.get(row.perfKey);
        if (!existing) {
          combined.set(row.perfKey, { ...row });
        } else {
          existing.sitesDownHours =
            Math.round((existing.sitesDownHours + row.sitesDownHours) * 10) /
            10;
        }
      });
    });

    const totalMonthHours = filters.openingMonth.reduce(
      (sum, monthKey) => sum + totalHoursInMonth(monthKey),
      0,
    );
    return Array.from(combined.values())
      .map((row) => {
        const availHours = Math.max(0, totalMonthHours - row.sitesDownHours);
        const totalHrs = availHours + row.sitesDownHours;
        const reliability = totalHrs > 0 ? availHours / totalHrs : 1;
        const totalMins = Math.round(availHours * 60);
        const dDays = Math.floor(totalMins / (60 * 24));
        const dHrs = Math.floor((totalMins % (60 * 24)) / 60);
        const dMins = Math.round(totalMins % 60);
        return {
          ...row,
          displayName: formatPerformanceChartLabel(
            row.siteId,
            row.siteName,
            row.sourceLabel,
          ),
          availHours: Math.round(availHours * 10) / 10,
          availDay: `${dDays} d, ${dHrs} h, ${dMins} m`,
          reliability: `${(reliability * 100).toFixed(2)}%`,
        };
      })
      .sort((a, b) => comparePerformanceSiteRows(a, b));
  }, [allDataRows, filters, performanceDashboardSiteOrder]);

  const analytics = useMemo(() => {
    const primaryRows = filteredTickets.map((ticket) => ticket.primary);
    const totalUnique = filteredTickets.length;
    const status = countBy(primaryRows, (row) => row.status);
    const severity = countBy(primaryRows, (row) => row.severity);
    const region = countBy(primaryRows, (row) => row.region);
    const impact = countBy(primaryRows, (row) => row.impact);
    const escalation = countBy(primaryRows, (row) => row.escalationLevel);
    const trendGrain = filters.openingMonth.length === 0 ? "month" : "week";
    const trendMap = new Map<
      string,
      { key: string; name: string; opened: number; resolved: number }
    >();
    const ensureTrendBucket = (key: string) => {
      const fallbackKey = key || "Unknown";
      if (!trendMap.has(fallbackKey)) {
        trendMap.set(fallbackKey, {
          key: fallbackKey,
          name:
            trendGrain === "month"
              ? openingMonthLabel(fallbackKey)
              : weekLabel(fallbackKey),
          opened: 0,
          resolved: 0,
        });
      }
      return trendMap.get(fallbackKey)!;
    };
    primaryRows.forEach((row) => {
      const openedKey =
        trendGrain === "month"
          ? row.openingMonthKey || openingMonthKey(row.observationDate)
          : weekKey(row.observationDate);
      ensureTrendBucket(openedKey).opened += 1;

      if (clean(row.status).toLowerCase().includes("resolved")) {
        const resolvedKey =
          trendGrain === "month"
            ? recordDateMonthKey(row.recoveryDate || row.observationDate)
            : weekKey(row.recoveryDate || row.observationDate);
        ensureTrendBucket(resolvedKey).resolved += 1;
      }
    });
    const monthly = Array.from(trendMap.entries())
      .sort(([a], [b]) => {
        if (a === "Unknown") return 1;
        if (b === "Unknown") return -1;
        return a.localeCompare(b);
      })
      .map(([, value]) => value);
    const replyTimeTrend = monthly.map((bucket) => {
      const rowsInBucket = primaryRows.filter((row) => {
        const key =
          trendGrain === "month"
            ? row.openingMonthKey || openingMonthKey(row.observationDate)
            : weekKey(row.observationDate);
        return key === bucket.key;
      });
      return {
        name: bucket.name,
        frt:
          Math.round(average(rowsInBucket.map((row) => row.frtHours)) * 10) /
          10,
        response:
          Math.round(
            average(rowsInBucket.map((row) => row.responseHours)) * 10,
          ) / 10,
        resolution:
          Math.round(
            average(rowsInBucket.map((row) => row.resolutionHours)) * 10,
          ) / 10,
      };
    });

    const monthlyResolutionMap = new Map<
      string,
      { sum: number; count: number; label: string }
    >();
    primaryRows.forEach((row) => {
      const s = String(row.status ?? "")
        .toLowerCase()
        .trim();
      if (s !== "closed" && s !== "resolved") return;
      const hours = ticketDurationHours(row);
      if (!Number.isFinite(hours) || hours <= 0) return;
      const key = row.openingMonthKey || openingMonthKey(row.observationDate);
      if (!key || key === "Unknown") return;
      const label = openingMonthLabel(key);
      const bucket = monthlyResolutionMap.get(key) ?? {
        sum: 0,
        count: 0,
        label,
      };
      bucket.sum += hours;
      bucket.count += 1;
      monthlyResolutionMap.set(key, bucket);
    });

    // Pending-age aggregate â€” all Pending tickets land in current calendar month.
    const nowDate = new Date();
    const currentMonthKey = `${nowDate.getFullYear()}-${String(nowDate.getMonth() + 1).padStart(2, "0")}`;
    const pendingSamples = primaryRows
      .filter((row) => isPendingStatus(row.status))
      .map((row) => {
        const obs = combineDateTime(row.observationDate, row.observationTime);
        if (!obs) return null;
        const hours = (nowDate.getTime() - obs.getTime()) / (1000 * 60 * 60);
        return Number.isFinite(hours) && hours > 0 ? hours : null;
      })
      .filter((v): v is number => v !== null);

    const pendingTotal = pendingSamples.reduce((s, v) => s + v, 0);
    const pendingAvg = pendingSamples.length
      ? pendingTotal / pendingSamples.length
      : 0;

    // Merge both series into a single chart-ready array (one row per month).
    const monthlyResolutionTime = (() => {
      const arr = Array.from(monthlyResolutionMap.entries())
        .sort(([a], [b]) => a.localeCompare(b))
        .map(([key, b]) => ({
          key,
          name: b.label,
          avgHours: Math.round((b.sum / b.count) * 100) / 100,
          totalHours: Math.round(b.sum * 100) / 100,
          count: b.count,
          pendingAvg: 0,
          pendingTotal: 0,
          pendingCount: 0,
        }));
      if (pendingSamples.length) {
        let idx = arr.findIndex((m) => m.key === currentMonthKey);
        if (idx === -1) {
          arr.push({
            key: currentMonthKey,
            name: openingMonthLabel(currentMonthKey),
            avgHours: 0,
            totalHours: 0,
            count: 0,
            pendingAvg: 0,
            pendingTotal: 0,
            pendingCount: 0,
          });
          arr.sort((a, b) => a.key.localeCompare(b.key));
          idx = arr.findIndex((m) => m.key === currentMonthKey);
        }
        arr[idx].pendingAvg = Math.round(pendingAvg * 100) / 100;
        arr[idx].pendingTotal = Math.round(pendingTotal * 100) / 100;
        arr[idx].pendingCount = pendingSamples.length;
      }
      return arr;
    })();

    // Grand mean retained for any consumer that still wants a single number.
    const resolutionSamples = primaryRows
      .filter((row) => {
        const s = String(row.status ?? "")
          .toLowerCase()
          .trim();
        return s === "closed" || s === "resolved";
      })
      .map((row) => ticketDurationHours(row))
      .filter((value) => Number.isFinite(value) && value > 0);

    const avgReplyTime = {
      frt: average(primaryRows.map((row) => row.frtHours)),
      response: average(primaryRows.map((row) => row.responseHours)),
      resolution: resolutionSamples.length
        ? resolutionSamples.reduce((sum, v) => sum + v, 0) /
          resolutionSamples.length
        : 0,
    };

    const avgHoursSource = primaryRows
      .map((row) => ticketDurationHours(row))
      .filter((value) => Number.isFinite(value));
    const avgHours = avgHoursSource.length
      ? avgHoursSource.reduce((sum, value) => sum + value, 0) /
        avgHoursSource.length
      : 0;
    const uniqueSites = new Set(
      performanceKpiRows
        .map((row) => normalizeSiteId(clean(row.siteId)).toUpperCase())
        .filter(isRfSiteId),
    ).size;
    const sourceSiteSets = new Map<string, Set<string>>();
    filteredTickets.forEach((ticket) => {
      ticket.rows.forEach((row) => {
        const siteId = normalizeSiteId(clean(row.siteId)).toUpperCase();
        if (!isRfSiteId(siteId)) return;
        const source = row.region || row.sourceFile || "Workbook";
        if (!sourceSiteSets.has(source)) sourceSiteSets.set(source, new Set());
        sourceSiteSets.get(source)!.add(siteId);
      });
    });
    const regionSiteTotal = Array.from(sourceSiteSets.values()).reduce(
      (sum, siteIds) => sum + siteIds.size,
      0,
    );
    const rootCauseUpdated = primaryRows.filter(
      (row) => row.actionTaken || !rcaNotProvided(row.rca),
    ).length;
    const totalSiteAffected = filteredTickets.reduce(
      (sum, ticket) =>
        sum + Math.max(ticket.siteIds.size, ticket.primary.siteId ? 1 : 0),
      0,
    );
    const rcaByCount = countBy(primaryRows, (row) =>
      rcaNotProvided(row.rca) ? "RCA not Provided" : row.rca,
    );
    const topRcaByCount = rcaByCount[0] ?? { name: "", value: 0 };
    const selectedDurationMonthKeys = filters.openingMonth.length
      ? filters.openingMonth
      : Array.from(
          new Set(
            primaryRows
              .map(
                (row) =>
                  row.openingMonthKey || openingMonthKey(row.observationDate),
              )
              .filter((key) => key && key !== "Unknown"),
          ),
        );
    const selectedDurationPeriodHours = selectedDurationMonthKeys.length
      ? selectedDurationMonthKeys.reduce(
          (sum, key) => sum + totalHoursInMonth(key),
          0,
        )
      : 24 * 30;
    const cappedTicketDuration = (row: TicketRecord) =>
      Math.min(ticketDurationHours(row), selectedDurationPeriodHours);
    const downtimeByRcaMap = new Map<
      string,
      { name: string; value: number; count: number }
    >();
    const countedRcaTickets = new Set<string>();
    primaryRows.forEach((row) => {
      const name = rcaNotProvided(row.rca) ? "RCA not Provided" : row.rca;
      const key = `${clean(row.tt) || row.rowNo}||${name}`;
      if (countedRcaTickets.has(key)) return;
      countedRcaTickets.add(key);
      const hours = cappedTicketDuration(row);
      const current = downtimeByRcaMap.get(name) ?? {
        name,
        value: 0,
        count: 0,
      };
      current.value += hours;
      if (hours) current.count += 1;
      downtimeByRcaMap.set(name, current);
    });
    const downtimeByRca = Array.from(downtimeByRcaMap.values()).sort(
      (a, b) => b.value - a.value || a.name.localeCompare(b.name),
    );
    const mttrByRca = Array.from(downtimeByRcaMap.values())
      .map((item) => ({
        name: item.name,
        value: item.count ? item.value / item.count : 0,
        count: item.count,
      }))
      .sort((a, b) => b.value - a.value || a.name.localeCompare(b.name));
    const repeatedSiteRcaMap = new Map<string, number>();
    primaryRows.forEach((row) => {
      const site = row.siteId || "Blank";
      const rca = rcaNotProvided(row.rca) ? "RCA not Provided" : row.rca;
      const key = `${site}||${rca}`;
      repeatedSiteRcaMap.set(key, (repeatedSiteRcaMap.get(key) ?? 0) + 1);
    });
    const repeatedRcaSites = Array.from(repeatedSiteRcaMap.values()).filter(
      (value) => value > 1,
    ).length;
    const rcaNotProvidedCount = primaryRows.filter((row) =>
      rcaNotProvided(row.rca),
    ).length;
    const preventableCount = primaryRows.filter(
      (row) =>
        (row.preventability || getPreventability(row.rca)) === "Preventable",
    ).length;
    const rcaFamily = countBy(
      primaryRows,
      (row) => row.rcaFamily || getRcaFamily(row.rca),
    );
    const siteNameById = new Map<string, string>();
    filteredTickets.forEach((ticket) => {
      ticket.rows.forEach((row) => {
        if (row.siteId && row.siteName && !siteNameById.has(row.siteId))
          siteNameById.set(row.siteId, row.siteName);
      });
    });
    const siteLabel = (siteId: string) => {
      const siteName = siteNameById.get(siteId);
      if (!siteId || siteId === "Blank") return "Blank";
      return siteName ? `${siteId} -- ${siteName}` : siteId;
    };
    const siteMap = new Map<
      string,
      { name: string; value: number; exposure?: number }
    >();
    filteredTickets.forEach((ticket) => {
      const sites = ticket.siteIds.size
        ? Array.from(ticket.siteIds)
        : [ticket.primary.siteId || "Blank"];
      sites.forEach((site) => {
        const current = siteMap.get(site) ?? {
          name: siteLabel(site),
          value: 0,
        };
        current.value += 1;
        siteMap.set(site, current);
      });
    });
    const topSites = Array.from(siteMap.values())
      .sort((a, b) => b.value - a.value || a.name.localeCompare(b.name))
      .slice(0, 12);
    const preventableBreakdown = [
      { name: "Preventable", value: preventableCount },
      { name: "Non-preventable", value: primaryRows.length - preventableCount },
    ].filter((item) => item.value > 0);

    const RCA_FAMILIES = [
      "Power & Environment",
      "Fiber & Physical",
      "Transmission & Link",
      "Hardware & Device",
      "Configuration / Software",
      "Human / Process / Planned",
      "Other / Review",
    ];
    const monthlyRcaFamilyMap = new Map<
      string,
      Record<string, string | number>
    >();
    primaryRows.forEach((row) => {
      const key = row.openingMonthKey || openingMonthKey(row.observationDate);
      if (!key || key === "Unknown") return;
      const family = row.rcaFamily || getRcaFamily(row.rca);
      if (!monthlyRcaFamilyMap.has(key)) {
        const entry: Record<string, string | number> = {
          monthKey: key,
          name: openingMonthLabel(key),
        };
        RCA_FAMILIES.forEach((f) => {
          entry[f] = 0;
        });
        monthlyRcaFamilyMap.set(key, entry);
      }
      const entry = monthlyRcaFamilyMap.get(key)!;
      entry[family] = ((entry[family] as number) || 0) + 1;
    });
    const monthlyRcaFamily = Array.from(monthlyRcaFamilyMap.entries())
      .sort(([a], [b]) => a.localeCompare(b))
      .map(([, value]) => value);

    const managedResourceByCount = countBy(
      primaryRows,
      (row) => row.managedResource || "Unknown",
    );
    const topManagedResources = managedResourceByCount.slice(0, 12);

    return {
      totalUnique,
      status,
      severity,
      region,
      impact,
      escalation,
      monthly,
      trendGrain,
      replyTimeTrend,
      avgReplyTime,
      monthlyResolutionTime,
      avgHours,
      uniqueSites,
      regionSiteTotal,
      rootCauseUpdated,
      totalSiteAffected,
      topSites,
      rcaByCount,
      rcaFamily,
      topRcaByCount,
      topRcaByDowntime: downtimeByRca[0] ?? { name: "", value: 0, count: 0 },
      highestMttrRca: mttrByRca[0] ?? { name: "", value: 0, count: 0 },
      repeatedRcaSites,
      rcaNotProvidedCount,
      preventableCount,
      rcaByDowntime: downtimeByRca.map((item) => ({
        name: item.name,
        value: Math.round(item.value * 10) / 10,
      })),
      rcaByMttr: mttrByRca.map((item) => ({
        name: item.name,
        value: Math.round(item.value * 10) / 10,
      })),
      preventableBreakdown,
      monthlyRcaFamily,
      rcaFamilyKeys: RCA_FAMILIES,
      topManagedResources,
    };
  }, [filteredTickets, filters.openingMonth, performanceKpiRows]);

  const closed = metricValue(analytics.status, "Closed");
  const pending = metricValue(analytics.status, "Pending");
  const resolved = metricValue(analytics.status, "Resolved");
  const critical = metricValue(analytics.severity, "Critical");
  const major = metricValue(analytics.severity, "Major");
  const minor = metricValue(analytics.severity, "Minor");
  const serviceImpact = metricValue(analytics.impact, "Service Impact");
  const nonServiceImpact = metricValue(analytics.impact, "Non-Service Impact");
  const rcaNotProvidedPct = pct(
    analytics.rcaNotProvidedCount,
    analytics.totalUnique,
  );
  const preventableRcaPct = pct(
    analytics.preventableCount,
    analytics.totalUnique,
  );

  const executiveInsights = useMemo(
    () =>
      calculateExecutiveInsights({
        tickets: filteredTickets,
        performanceRows: performanceKpiRows,
      }),
    [filteredTickets, performanceKpiRows],
  );

  const deepDiveAnalytics = useMemo(
    () =>
      calculateDeepDiveAnalytics({
        tickets: filteredTickets,
        performanceRows: performanceKpiRows,
      }),
    [filteredTickets, performanceKpiRows],
  );
  function scrollToTopCards() {
    // Scroll to the filters card. Fall back to the stats grid if filters
    // aren't mounted (e.g. before a workbook is loaded).
    const target = filtersRef.current ?? statsRef.current;
    if (target) {
      target.scrollIntoView({ behavior: "smooth", block: "start" });
    } else {
      window.scrollTo({ top: 0, behavior: "smooth" });
    }
  }

  const ticketExportMetrics = useMemo(() => {
    const closedOrResolved = monthlyExportTickets.filter((ticket) => {
      const status = clean(ticket.primary.status).toLowerCase();
      return status === "closed" || status === "resolved";
    }).length;
    const pendingCount = monthlyExportTickets.filter(
      (ticket) => clean(ticket.primary.status).toLowerCase() === "pending",
    ).length;
    const criticalCount = monthlyExportTickets.filter(
      (ticket) => clean(ticket.primary.severity).toLowerCase() === "critical",
    ).length;
    return { closedOrResolved, pendingCount, criticalCount };
  }, [monthlyExportTickets]);

  const performanceExportMetrics = useMemo(() => {
    const scopedSourceRows = allDataRows.filter((row) => {
      const regionMatch =
        perfRegions.length === 0 || perfRegions.includes(row.region);
      const monthMatch =
        perfMonths.length === 0 ||
        coveredMonthKeys(row).some((monthKey) => perfMonths.includes(monthKey));
      return regionMatch && monthMatch;
    });
    const scopedSiteIds = new Set(
      scopedSourceRows.map((row) => clean(row.siteId)).filter(isRfSiteId),
    );
    const rfPerfRows = scopedSiteIds.size
      ? perfRows.filter((row) => scopedSiteIds.has(clean(row.siteId)))
      : perfRows.filter((row) => isRfSiteId(row.siteId));
    const rfSiteIds = new Set(
      rfPerfRows.map((row) => clean(row.siteId)).filter(isRfSiteId),
    );
    const affectedIds = new Set(
      rfPerfRows
        .filter((row) => row.sitesDownHours > 0)
        .map((row) => clean(row.siteId))
        .filter(isRfSiteId),
    );
    return {
      totalSites: rfSiteIds.size,
      affectedSites: affectedIds.size,
      nonAffectedSites: Math.max(0, rfSiteIds.size - affectedIds.size),
      reportRows: rfPerfRows.length,
    };
  }, [allDataRows, perfMonths, perfRegions, perfRows]);


  const reportPerformanceExportMetrics = useMemo(() => {
    const scopedSourceRows = allDataRows.filter((row) => {
      const regionMatch =
        exportRegions.length === 0 || exportRegions.includes(row.region);
      const monthMatch =
        exportMonths.length === 0 ||
        coveredMonthKeys(row).some((monthKey) => exportMonths.includes(monthKey));
      return regionMatch && monthMatch;
    });
    const scopedSiteIds = new Set(
      scopedSourceRows.map((row) => clean(row.siteId)).filter(isRfSiteId),
    );
    const rfPerfRows = scopedSiteIds.size
      ? perfRows.filter((row) => scopedSiteIds.has(clean(row.siteId)))
      : perfRows.filter((row) => isRfSiteId(row.siteId));
    const rfSiteIds = new Set(
      rfPerfRows.map((row) => clean(row.siteId)).filter(isRfSiteId),
    );
    const affectedIds = new Set(
      rfPerfRows
        .filter((row) => row.sitesDownHours > 0)
        .map((row) => clean(row.siteId))
        .filter(isRfSiteId),
    );
    return {
      totalSites: rfSiteIds.size,
      affectedSites: affectedIds.size,
      nonAffectedSites: Math.max(0, rfSiteIds.size - affectedIds.size),
      reportRows: rfPerfRows.length,
    };
  }, [allDataRows, exportMonths, exportRegions, perfRows]);

  const managedReportDefinitions = useMemo(
    () => [
      {
        id: "tickets" as ManagedReportType,
        title: "Tickets Monthly Report",
        description: "Filtered ticket register, status, severity, RCA and monthly follow-up export.",
        formats: ["xlsx", "pdf", "ppt"] as ManagedReportFormat[],
        records: monthlyExportTickets.length,
      },
      {
        id: "performance" as ManagedReportType,
        title: "Performance KPI Report",
        description: "Site availability, downtime, reliability and monthly performance export.",
        formats: ["xlsx", "pdf", "ppt"] as ManagedReportFormat[],
        records: reportPerformanceExportMetrics.reportRows,
      },
      {
        id: "executive" as ManagedReportType,
        title: "Executive Insights Workbook",
        description: "Network health score, insight cards and high-risk site ranking.",
        formats: ["xlsx"] as ManagedReportFormat[],
        records: executiveInsights.highRiskSites.length + executiveInsights.cards.length,
      },
      {
        id: "quality" as ManagedReportType,
        title: "RCA / SLA Deep-Dive Workbook",
        description: "Preventability, RCA families, pending aging and repeated offender sites.",
        formats: ["xlsx"] as ManagedReportFormat[],
        records:
          deepDiveAnalytics.repeatedOffenderSites.length +
          deepDiveAnalytics.rcaFamilyDeepDive.length,
      },
      {
        id: "kpiCards" as ManagedReportType,
        title: "KPI Cards Export",
        description: "Export the full KPI card grid as a clean management PNG or PDF.",
        formats: ["png", "pdf"] as ManagedReportFormat[],
        records: analytics.totalUnique,
      },
      {
        id: "performanceCards" as ManagedReportType,
        title: "Performance KPI Cards Export",
        description: "Export the full Performance KPI gauge card grid as a clean management PNG or PDF.",
        formats: ["png", "pdf"] as ManagedReportFormat[],
        records: performanceKpiRows.length,
      },
    ],
    [
      deepDiveAnalytics.rcaFamilyDeepDive.length,
      deepDiveAnalytics.repeatedOffenderSites.length,
      executiveInsights.cards.length,
      executiveInsights.highRiskSites.length,
      monthlyExportTickets.length,
      reportPerformanceExportMetrics.reportRows,
      analytics.totalUnique,
      performanceKpiRows.length,
    ],
  );

  const selectedManagedReport =
    managedReportDefinitions.find((report) => report.id === managedReportType) ??
    managedReportDefinitions[0];

  const managedReportFormatOptions = selectedManagedReport.formats;

  useEffect(() => {
    if (!managedReportFormatOptions.includes(managedReportFormat)) {
      setManagedReportFormat(managedReportFormatOptions[0] ?? "xlsx");
    }
  }, [managedReportFormat, managedReportFormatOptions]);

  const visibleGeneratedReports = useMemo(() => {
    const query = clean(reportSearch).toLowerCase();
    if (!query) return generatedReports;
    return generatedReports.filter((report) =>
      [report.fileName, report.reportType, report.format, report.generatedAt]
        .join(" ")
        .toLowerCase()
        .includes(query),
    );
  }, [generatedReports, reportSearch]);

  function addGeneratedReportHistory(item: Omit<GeneratedReportItem, "id" | "generatedAt">) {
    const generatedAt = new Date().toLocaleString("en", {
      year: "numeric",
      month: "short",
      day: "2-digit",
      hour: "2-digit",
      minute: "2-digit",
    });
    setGeneratedReports((prev) => [
      {
        ...item,
        id: `${Date.now()}-${Math.random().toString(16).slice(2)}`,
        generatedAt,
      },
      ...prev,
    ].slice(0, 12));
  }

  function clearGeneratedReportsHistory() {
    setGeneratedReports([]);
  }

  async function handleGenerateManagedReport() {
    const report = selectedManagedReport;
    const format = report.formats.includes(managedReportFormat)
      ? managedReportFormat
      : report.formats[0];
    const monthLabel =
      exportMonths.length === 1
        ? (filterOptions.exportMonthLabels?.[exportMonths[0]] ?? exportMonths[0])
        : exportMonths.length > 1
          ? `${exportMonths.length} months`
          : "All";
    const perfMonthLabel =
      exportMonths.length === 1
        ? (filterOptions.exportMonthLabels?.[exportMonths[0]] ?? exportMonths[0])
        : exportMonths.length > 1
          ? `${exportMonths.length} months`
          : "All";

    try {
      if (report.id === "tickets") {
        if (format === "xlsx") {
          const groupsToExport = monthlyExcelExportGroups.length
            ? monthlyExcelExportGroups
            : [{ tickets: monthlyExportTickets }];
          for (const group of groupsToExport) {
            await exportTicketTemplate(group.tickets, exportMonths[0] ?? "all");
          }
        } else if (format === "pdf") {
          exportPdf(monthlyExportTickets, exportMonths[0] ?? "all");
        } else {
          exportTicketsPpt(
            monthlyExportTickets,
            monthLabel,
            executiveInsights,
            deepDiveAnalytics,
          );
        }
      } else if (report.id === "performance") {
        if (format === "xlsx") {
          await exportPerfTemplate(perfRows, exportMonths[0] ?? "all", exportRegions);
        } else if (format === "pdf") {
          exportPerfPdf(perfRows, exportMonths[0] ?? "all");
        } else {
          exportPerfPpt(perfRows, perfMonthLabel, executiveInsights);
        }
      } else if (report.id === "kpiCards") {
        setActiveDashboardTab("kpis");
        await new Promise((resolve) =>
          requestAnimationFrame(() => requestAnimationFrame(resolve)),
        );
        if (format === "pdf") {
          await exportCardsCanvasPdf(statsRef.current, "KPI-Cards", ".stat-card");
        } else {
          await exportCardsCanvasPng(statsRef.current, "KPI-Cards", ".stat-card");
        }
        setActiveDashboardTab("reports");
      } else if (report.id === "performanceCards") {
        setActiveDashboardTab("performanceKpis");
        await new Promise((resolve) =>
          requestAnimationFrame(() => requestAnimationFrame(resolve)),
        );
        if (format === "pdf") {
          await exportCardsCanvasPdf(
            performanceKpiCardsRef.current,
            "Performance-KPI-Gauge-Cards",
            ".perf-gauge-card",
          );
        } else {
          await exportCardsCanvasPng(
            performanceKpiCardsRef.current,
            "Performance-KPI-Gauge-Cards",
            ".perf-gauge-card",
          );
        }
        setActiveDashboardTab("reports");
      } else if (report.id === "executive") {
        await exportDashboardSectionExcel("executive");
      } else {
        await exportDashboardSectionExcel("deepDive");
      }

      addGeneratedReportHistory({
        fileName: `${report.title.replace(/\s+/g, "_")}_${format.toUpperCase()}`,
        reportType: report.title,
        format,
        records: report.records,
      });
    } catch (err) {
      if (import.meta.env.DEV) console.error(err);
      setError("The report could not be generated. Please check the selected filters and try again.");
    }
  }

  const activeDashboardSection =
    DASHBOARD_SECTIONS.find((section) => section.id === activeDashboardTab) ??
    DASHBOARD_SECTIONS[0];

  const executiveRcaSlaRows = useMemo(
    () => deepDiveAnalytics.repeatedOffenderSites,
    [deepDiveAnalytics.repeatedOffenderSites],
  );

  function makeFileSafeName(value: string) {
    return (
      clean(value)
        .replace(/[^a-z0-9]+/gi, "_")
        .replace(/^_+|_+$/g, "") || "dashboard_export"
    );
  }

  function serializeDashboardData(snapshot: DashboardData) {
    return {
      fileName: snapshot.fileName,
      sheetName: snapshot.sheetName,
      generatedAt: snapshot.generatedAt,
      rows: snapshot.rows,
      siteOrder: snapshot.siteOrder,
      rcaLookup: snapshot.rcaLookup ?? [],
    };
  }

  function hydrateDashboardData(
    snapshot: ReturnType<typeof serializeDashboardData>,
  ): DashboardData {
    return {
      ...snapshot,
      uniqueTickets: groupTickets(snapshot.rows),
      rcaLookup: snapshot.rcaLookup ?? [],
    };
  }

  function loadSavedDashboardSnapshot() {
    try {
      const raw = localStorage.getItem(SAVED_DASHBOARD_KEY);
      if (!raw) return;
      const parsed = JSON.parse(raw) as {
        data?: ReturnType<typeof serializeDashboardData>;
        regions?: Array<ReturnType<typeof serializeDashboardData>>;
        googleRegionLinks?: Partial<GoogleRegionLinks>;
        googleRegionSaveLinks?: Partial<GoogleRegionSaveLinks>;
        googleRegionSelection?: Partial<GoogleRegionSelection>;
      };
      if (!parsed.data?.rows?.length) return;
      const restoredData = hydrateDashboardData(parsed.data);
      const restoredRegions = (parsed.regions ?? []).map(hydrateDashboardData);
      setData(restoredData);
      setRegions(restoredRegions.length ? restoredRegions : [restoredData]);
      setGoogleRegionLinks((prev) => ({
        ...prev,
        ...(parsed.googleRegionLinks ?? {}),
      }));
      setGoogleRegionSaveLinks((prev) => ({
        ...prev,
        ...(parsed.googleRegionSaveLinks ?? {}),
      }));
      if (parsed.googleRegionSelection) {
        setGoogleRegionSelection((prev) => ({
          ...prev,
          ...parsed.googleRegionSelection,
        }));
      }
      setError("");
      setFilters(EMPTY_FILTERS);
      setTablePage(1);
      setPerfPage(1);
    } catch (err) {
      console.error("Failed to restore previous dashboard snapshot:", err);
      setError(
        "Could not restore the previous workbook session. Please upload the workbook again.",
      );
    }
  }

  function returnToWelcomeUploadScreen() {
    setData(null);
    setRegions([]);
    setError("");
    setFilters(EMPTY_FILTERS);
    setExportMonths([]);
    setExportRegions([]);
    setPerfMonths([]);
    setPerfRegions([]);
    setTablePage(1);
    setPerfPage(1);
    setCollapsedSections(INITIAL_COLLAPSED_SECTIONS);
    setActiveDashboardTab("performanceKpis");
    setSavedSnapshotAvailable(
      Boolean(localStorage.getItem(SAVED_DASHBOARD_KEY)),
    );
    window.scrollTo({ top: 0, behavior: "smooth" });
  }

  async function clearSavedDashboardData() {
    try {
      window.localStorage.removeItem(SAVED_DASHBOARD_KEY);
      window.localStorage.removeItem(SAVED_MANUAL_TICKETS_KEY);
      window.sessionStorage.removeItem("followup-dashboard:last-upload-error");

      const dashboardIndexedDB = indexedDB as IDBFactory & {
        databases?: () => Promise<Array<{ name?: string }>>;
      };
      if ("indexedDB" in window && typeof dashboardIndexedDB.databases === "function") {
        const databases = await dashboardIndexedDB.databases();
        await Promise.all(
          databases
            .map((database) => database.name)
            .filter(
              (name): name is string =>
                typeof name === "string" && /followup|dashboard|ticket|workbook|risk/i.test(name),
            )
            .map(
              (name) =>
                new Promise<void>((resolve) => {
                  const request = dashboardIndexedDB.deleteDatabase(name);
                  request.onsuccess = () => resolve();
                  request.onerror = () => resolve();
                  request.onblocked = () => resolve();
                }),
            ),
        );
      }

      setSavedSnapshotAvailable(false);
      window.alert("Saved dashboard data cleared successfully.");
    } catch (error) {
      if (import.meta.env.DEV) {
        console.error("Failed to clear saved dashboard data:", error);
      }
      window.alert("Saved dashboard data could not be cleared. Please try again.");
    }
  }

  function scrollToDashboardTop() {
    window.scrollTo({ top: 0, behavior: "smooth" });
  }

  function openDashboardTab(sectionId: DashboardSectionId) {
    setActiveDashboardTab(sectionId);
    window.setTimeout(() => {
      window.scrollTo({ top: 0, behavior: "smooth" });
    }, 30);
  }

  function toggleDashboardSection(sectionId: DashboardSectionId) {
    setActiveDashboardTab(sectionId);
  }

  function setAllDashboardSections(_collapsed: boolean) {
    setActiveDashboardTab("performanceKpis");
  }

  async function exportSvgFallbackToPng(
    element: HTMLElement,
    fileName: string,
  ) {
    const svg =
      element instanceof SVGSVGElement
        ? element
        : element.querySelector<SVGSVGElement>(
            "svg.recharts-surface, .recharts-wrapper svg",
          );
    if (!svg || svg.classList.contains("file-export-svg")) return false;

    const clonedSvg = svg.cloneNode(true) as SVGSVGElement;
    const sourceBox = svg.getBoundingClientRect();
    const width = Math.max(
      Math.ceil(sourceBox.width || svg.clientWidth || 900),
      1,
    );
    const height = Math.max(
      Math.ceil(sourceBox.height || svg.clientHeight || 420),
      1,
    );
    clonedSvg.setAttribute("xmlns", "http://www.w3.org/2000/svg");
    clonedSvg.setAttribute("width", String(width));
    clonedSvg.setAttribute("height", String(height));
    clonedSvg.setAttribute(
      "viewBox",
      clonedSvg.getAttribute("viewBox") || `0 0 ${width} ${height}`,
    );

    const svgText = new XMLSerializer().serializeToString(clonedSvg);
    const svgBlob = new Blob([svgText], {
      type: "image/svg+xml;charset=utf-8",
    });
    const svgUrl = URL.createObjectURL(svgBlob);

    try {
      const img = new Image();
      img.decoding = "async";
      img.src = svgUrl;
      await new Promise<void>((resolve, reject) => {
        img.onload = () => resolve();
        img.onerror = () =>
          reject(new Error("SVG image export fallback failed."));
      });

      const canvas = document.createElement("canvas");
      canvas.width = width * 2;
      canvas.height = height * 2;
      const ctx = canvas.getContext("2d");
      if (!ctx) return false;
      ctx.fillStyle = getComputedStyle(element).backgroundColor || "#071c2a";
      ctx.fillRect(0, 0, canvas.width, canvas.height);
      ctx.scale(2, 2);
      ctx.drawImage(img, 0, 0, width, height);

      const link = document.createElement("a");
      link.href = canvas.toDataURL("image/png");
      link.download = `${makeFileSafeName(fileName)}.png`;
      document.body.appendChild(link);
      link.click();
      link.remove();
      return true;
    } finally {
      URL.revokeObjectURL(svgUrl);
    }
  }
  async function exportElementToPng(
    element: HTMLElement | null,
    fileName: string,
  ) {
    if (!element) {
      setError(
        "PNG export target was not found. Please open the section and try again.",
      );
      return;
    }
    const { default: html2canvas } = await import("html2canvas");
    try {
      element.classList.add("png-export-active");
      const canvas = await html2canvas(element, {
        backgroundColor: null,
        scale: Math.min(2, window.devicePixelRatio || 1.5),
        useCORS: true,
        logging: false,
        ignoreElements: (node) =>
          node !== element &&
          node instanceof HTMLElement &&
          (node.classList.contains("section-control-panel") ||
            node.classList.contains("chart-export-png-button") ||
            node.classList.contains("no-print")),
      });
      const link = document.createElement("a");
      link.href = canvas.toDataURL("image/png");
      link.download = `${makeFileSafeName(fileName)}.png`;
      document.body.appendChild(link);
      link.click();
      link.remove();
    } catch (err) {
      console.error("PNG export failed:", err);
      try {
        const fallbackWorked = await exportSvgFallbackToPng(element, fileName);
        if (!fallbackWorked) {
          setError(
            "PNG export failed. Please try again after the section finishes rendering.",
          );
        }
      } catch (fallbackErr) {
        console.error("PNG SVG fallback failed:", fallbackErr);
        setError(
          "PNG export failed. Please try again after the section finishes rendering.",
        );
      }
    } finally {
      element.classList.remove("png-export-active");
    }
  }


  function drawRoundedRect(
    ctx: CanvasRenderingContext2D,
    x: number,
    y: number,
    width: number,
    height: number,
    radius: number,
  ) {
    const r = Math.min(radius, width / 2, height / 2);
    ctx.beginPath();
    ctx.moveTo(x + r, y);
    ctx.lineTo(x + width - r, y);
    ctx.quadraticCurveTo(x + width, y, x + width, y + r);
    ctx.lineTo(x + width, y + height - r);
    ctx.quadraticCurveTo(x + width, y + height, x + width - r, y + height);
    ctx.lineTo(x + r, y + height);
    ctx.quadraticCurveTo(x, y + height, x, y + height - r);
    ctx.lineTo(x, y + r);
    ctx.quadraticCurveTo(x, y, x + r, y);
    ctx.closePath();
  }

  function wrapCanvasText(
    ctx: CanvasRenderingContext2D,
    text: string,
    x: number,
    y: number,
    maxWidth: number,
    lineHeight: number,
    maxLines = 2,
  ) {
    const words = text.split(/\s+/).filter(Boolean);
    const lines: string[] = [];
    let current = "";
    words.forEach((word) => {
      const next = current ? `${current} ${word}` : word;
      if (ctx.measureText(next).width <= maxWidth || !current) {
        current = next;
      } else {
        lines.push(current);
        current = word;
      }
    });
    if (current) lines.push(current);
    lines.slice(0, maxLines).forEach((line, index) => {
      const finalLine =
        index === maxLines - 1 && lines.length > maxLines ? `${line}…` : line;
      ctx.fillText(finalLine, x, y + index * lineHeight);
    });
    return Math.min(lines.length, maxLines) * lineHeight;
  }

  function getCardText(card: Element, selector: string, fallback = "") {
    return clean(card.querySelector(selector)?.textContent ?? fallback);
  }

  async function renderCardsCanvas(
    element: HTMLElement | null,
    fileName: string,
    cardSelector: string,
  ): Promise<HTMLCanvasElement | null> {
    if (!element) {
      setError("PNG/PDF export target was not found. Please open the tab and try again.");
      return null;
    }

    await new Promise((resolve) => requestAnimationFrame(resolve));

    const cards = Array.from(element.querySelectorAll<HTMLElement>(cardSelector)).filter(
      (card) => card.offsetWidth > 0 && card.offsetHeight > 0,
    );

    if (!cards.length) {
      setError("No cards were found to export. Please open the tab and try again.");
      return null;
    }

    const cardBoxes = cards.map((card) => card.getBoundingClientRect());
    const minX = Math.min(...cardBoxes.map((box) => box.left));
    const minY = Math.min(...cardBoxes.map((box) => box.top));
    const maxX = Math.max(...cardBoxes.map((box) => box.right));
    const maxY = Math.max(...cardBoxes.map((box) => box.bottom));
    const padding = 34;
    const titleHeight = 54;
    const width = Math.max(900, Math.ceil(maxX - minX + padding * 2));
    const height = Math.max(360, Math.ceil(maxY - minY + padding * 2 + titleHeight));
    const scale = Math.min(2, window.devicePixelRatio || 1.5);

    const canvas = document.createElement("canvas");
    canvas.width = width * scale;
    canvas.height = height * scale;
    const ctx = canvas.getContext("2d");
    if (!ctx) return null;
    ctx.scale(scale, scale);

    const isLight = document
      .querySelector(".dashboard-shell")
      ?.getAttribute("data-dashboard-theme") === "light";
    const background = isLight ? "#eef6ff" : "#07111f";
    const panel = isLight ? "#ffffff" : "#101a2b";
    const border = isLight ? "#c6d9ea" : "#24445f";
    const primaryText = isLight ? "#0f2740" : "#f8fafc";
    const mutedText = isLight ? "#55708a" : "#9fb4ca";

    ctx.fillStyle = background;
    ctx.fillRect(0, 0, width, height);

    ctx.fillStyle = primaryText;
    ctx.font = "800 24px Arial, sans-serif";
    ctx.fillText(fileName.replace(/[-_]/g, " "), padding, 34);

    cards.forEach((card, index) => {
      const box = cardBoxes[index];
      const x = Math.round(box.left - minX + padding);
      const y = Math.round(box.top - minY + padding + titleHeight);
      const w = Math.max(170, Math.round(box.width));
      const h = Math.max(110, Math.round(box.height));
      const color =
        card.style.getPropertyValue("--tone") ||
        card.style.getPropertyValue("--gauge-color") ||
        getComputedStyle(card).getPropertyValue("--tone") ||
        getComputedStyle(card).getPropertyValue("--gauge-color") ||
        "#22d3ee";

      ctx.save();
      drawRoundedRect(ctx, x, y, w, h, 22);
      ctx.fillStyle = panel;
      ctx.fill();
      ctx.strokeStyle = border;
      ctx.lineWidth = 1.5;
      ctx.stroke();

      ctx.strokeStyle = color.trim() || "#22d3ee";
      ctx.lineWidth = 4;
      ctx.beginPath();
      ctx.moveTo(x + 20, y + 10);
      ctx.lineTo(x + w - 20, y + 10);
      ctx.stroke();

      const isGauge = card.classList.contains("perf-gauge-card");
      if (isGauge) {
        ctx.strokeStyle = color.trim() || "#22d3ee";
        ctx.lineWidth = 10;
        ctx.beginPath();
        ctx.arc(x + w / 2, y + 132, Math.min(w * 0.42, 104), Math.PI, 0);
        ctx.stroke();
        ctx.fillStyle = `${color.trim() || "#22d3ee"}22`;
        ctx.beginPath();
        ctx.arc(x + w / 2, y + 92, 22, 0, Math.PI * 2);
        ctx.fill();
      } else {
        ctx.fillStyle = `${color.trim() || "#22d3ee"}22`;
        ctx.beginPath();
        ctx.arc(x + 42, y + 44, 24, 0, Math.PI * 2);
        ctx.fill();
      }

      const label = isGauge
        ? getCardText(card, ".perf-gauge-card__label")
        : getCardText(card, "span");
      const value = isGauge
        ? getCardText(card, ".perf-gauge-card__value")
        : getCardText(card, "strong");
      const caption = isGauge
        ? getCardText(card, ".perf-gauge-card__caption")
        : getCardText(card, "small");
      const helper = isGauge ? getCardText(card, ".perf-gauge-card__helper") : "";

      ctx.textAlign = "center";
      ctx.fillStyle = mutedText;
      ctx.font = "800 12px Arial, sans-serif";
      wrapCanvasText(ctx, label.toUpperCase(), x + w / 2, y + 34, w - 26, 15, 2);

      ctx.fillStyle = color.trim() || "#22d3ee";
      ctx.font = "900 24px Arial, sans-serif";
      wrapCanvasText(ctx, value || "--", x + w / 2, y + (isGauge ? 126 : 78), w - 24, 26, 2);

      ctx.fillStyle = primaryText;
      ctx.font = "700 12px Arial, sans-serif";
      wrapCanvasText(ctx, caption || helper, x + w / 2, y + h - 38, w - 26, 15, 2);
      ctx.restore();
    });

    return canvas;
  }

  async function exportCardsCanvasPng(
    element: HTMLElement | null,
    fileName: string,
    cardSelector: string,
  ) {
    const canvas = await renderCardsCanvas(element, fileName, cardSelector);
    if (!canvas) return;
    const link = document.createElement("a");
    link.href = canvas.toDataURL("image/png");
    link.download = `${makeFileSafeName(fileName)}.png`;
    document.body.appendChild(link);
    link.click();
    link.remove();
  }

  async function exportCardsCanvasPdf(
    element: HTMLElement | null,
    fileName: string,
    cardSelector: string,
  ) {
    const canvas = await renderCardsCanvas(element, fileName, cardSelector);
    if (!canvas) return;
    const imgData = canvas.toDataURL("image/png");
    const pageWidth = Math.ceil(canvas.width / Math.min(2, window.devicePixelRatio || 1.5)) + 40;
    const pageHeight = Math.ceil(canvas.height / Math.min(2, window.devicePixelRatio || 1.5)) + 40;
    const doc = new jsPDF({
      orientation: pageWidth >= pageHeight ? "landscape" : "portrait",
      unit: "px",
      format: [pageWidth, pageHeight],
    });
    doc.addImage(imgData, "PNG", 20, 20, pageWidth - 40, pageHeight - 40);
    doc.save(`${makeFileSafeName(fileName)}.pdf`);
  }

  function hasDashboardSectionExcel(sectionId: DashboardSectionId) {
    switch (sectionId) {
      case "performanceKpis":
        return performanceKpiRows.length > 0;
      case "ticketsTable":
        return filteredTickets.length > 0;
      case "executive":
      case "deepDive":
        return false;
      default:
        return false;
    }
  }

  function exportDashboardSectionExcel(sectionId: DashboardSectionId) {
    if (sectionId === "performanceKpis") {
      exportAnalyticsTableExcel(
        PERF_REPORT_HEADERS,
        perfReportRows(performanceKpiRows),
        "Performance KPIs",
        "Performance-KPI-Data.xlsx",
      );
      return;
    }

    if (sectionId === "ticketsTable") {
      exportExcel(filteredTickets);
      return;
    }

    if (sectionId === "executive") {
      exportAnalyticsWorkbookExcel(
        [
          {
            name: "Executive Summary",
            headers: ["Metric", "Value", "Note"],
            rows: [
              [
                "Network Health Score",
                executiveInsights.healthScore.score,
                executiveInsights.healthScore.status,
              ],
              ["Health Reason", executiveInsights.healthScore.mainReason, ""],
              ["Executive Summary", executiveInsights.summaryText, ""],
              ...executiveInsights.cards.map((card) => [
                card.label,
                card.value,
                card.note,
              ]),
            ],
          },
          {
            name: "High Risk Sites",
            headers: [
              "Region",
              "Rank",
              "Site ID",
              "Site Name",
              "Tickets",
              "Downtime",
              "Reliability",
              "Top RCA",
              "Risk Level",
              "Risk Score",
            ],
            rows: executiveInsights.highRiskSites.map((site) => [
              site.region,
              site.rank,
              site.siteId,
              site.siteName || "-",
              site.ticketCount,
              `${site.downtimeHours} hrs`,
              `${site.reliability.toFixed(2)}%`,
              site.topRca,
              site.riskLevel,
              site.riskScore,
            ]),
          },
        ],
        "Executive-Insights-Risk-Summary.xlsx",
      );
      return;
    }

    if (sectionId === "deepDive") {
      exportAnalyticsWorkbookExcel(
        [
          {
            name: "KPI Summary",
            headers: ["Metric", "Value", "Note"],
            rows: [
              [
                "Avg FRT",
                `${deepDiveAnalytics.slaSummary.avgFrtHours.toFixed(1)}h`,
                `${deepDiveAnalytics.slaSummary.frtBreaches} above 1h target`,
              ],
              [
                "Avg Response",
                `${deepDiveAnalytics.slaSummary.avgResponseHours.toFixed(1)}h`,
                `${deepDiveAnalytics.slaSummary.responseBreaches} above 4h target`,
              ],
              [
                "Avg Resolution",
                `${deepDiveAnalytics.slaSummary.avgResolutionHours.toFixed(1)}h`,
                `${deepDiveAnalytics.slaSummary.resolutionBreaches} above 24h target`,
              ],
              [
                "Repeated Sites",
                deepDiveAnalytics.repeatedOffenderSites.length,
                "Top sites requiring technical follow-up",
              ],
            ],
          },
          {
            name: "Chart Preventability",
            headers: ["Category", "Tickets", "Percentage"],
            rows: deepDiveAnalytics.preventabilityByCount.map((row) => [
              row.name,
              row.value,
              `${row.percentage ?? 0}%`,
            ]),
          },
          {
            name: "Chart Pending Aging",
            headers: ["Bucket", "Tickets", "Label"],
            rows: deepDiveAnalytics.slaSummary.pendingAgingBuckets.map(
              (row) => [row.name, row.value, row.label],
            ),
          },
          {
            name: "Chart RCA Downtime",
            headers: [
              "RCA Family",
              "Tickets",
              "Downtime",
              "Missing RCA",
              "Preventable",
              "Service Impact",
              "Responsible Team",
              "Recommended Action",
            ],
            rows: deepDiveAnalytics.rcaFamilyDeepDive.map((row) => [
              row.family,
              row.tickets,
              `${row.downtimeHours} hrs`,
              row.missingRca,
              row.preventableTickets,
              row.serviceImpactTickets,
              row.responsibleTeam,
              row.recommendedAction,
            ]),
          },
          {
            name: "Repeated Offenders",
            headers: [
              "Region",
              "Site ID",
              "Site Name",
              "TT's",
              "Availability",
              "Downtime",
              "Reliability",
              "Top RCA",
            ],
            rows: deepDiveAnalytics.repeatedOffenderSites.map((site) => [
              site.region,
              site.siteId,
              site.siteName || "-",
              site.tickets,
              site.performanceAvailabilityHours === null
                ? "-"
                : `${site.performanceAvailabilityHours} hrs`,
              site.performanceDowntimeHours === null
                ? "-"
                : `${site.performanceDowntimeHours} hrs`,
              site.reliability === null ? "-" : `${site.reliability}%`,
              site.topRca,
            ]),
          },
          {
            name: "Actions",
            headers: ["No", "Recommended Management Action"],
            rows: deepDiveAnalytics.recommendations.map((item, index) => [
              index + 1,
              item,
            ]),
          },
        ],
        "Operational-Quality-Follow-Up-Priorities.xlsx",
      );
    }
  }
  useEffect(() => {
    try {
      const raw = localStorage.getItem(SAVED_MANUAL_TICKETS_KEY);
      if (!raw) return;
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed)) setManualDrafts(parsed);
    } catch (error) {
      console.warn("Could not restore manual ticket input rows", error);
    }
  }, []);

  useEffect(() => {
    try {
      localStorage.setItem(
        SAVED_MANUAL_TICKETS_KEY,
        JSON.stringify(manualDrafts),
      );
    } catch (error) {
      console.warn("Could not save manual ticket input rows", error);
    }
  }, [manualDrafts]);
  useEffect(() => {
    const raw = localStorage.getItem(SAVED_DASHBOARD_KEY);
    setSavedSnapshotAvailable(Boolean(raw));
    if (!raw) return;
    try {
      const parsed = JSON.parse(raw) as {
        googleRegionLinks?: Partial<GoogleRegionLinks>;
        googleRegionSaveLinks?: Partial<GoogleRegionSaveLinks>;
        googleRegionSelection?: Partial<GoogleRegionSelection>;
      };
      if (parsed.googleRegionLinks) {
        setGoogleRegionLinks((prev) => ({
          ...prev,
          ...parsed.googleRegionLinks,
        }));
      }
      if (parsed.googleRegionSelection) {
        setGoogleRegionSelection((prev) => ({
          ...prev,
          ...parsed.googleRegionSelection,
        }));
      }
    } catch {
      // Ignore older saved payloads that do not include Google links.
    }
  }, []);

  useEffect(() => {
    fetch(`/google-sheets-config.txt?v=${Date.now()}`, { cache: "no-store" })
      .then((response) => (response.ok ? response.text() : ""))
      .then((text) => {
        if (!text) return;
        const config = parseGoogleSheetsConfig(text);
        setGoogleRegionLinks((prev) => {
          const next = { ...prev };
          GOOGLE_REGION_LINKS.forEach((region) => {
            if (!next[region.key] && config.links[region.key]) {
              next[region.key] = config.links[region.key] ?? "";
            }
          });
          return next;
        });
        setGoogleRegionSaveLinks((prev) => ({
          ...prev,
          ...config.saveLinks,
        }));
      })
      .catch(() => {
        // The config file is optional; users can still paste links in the UI.
      });
  }, []);

  useEffect(() => {
    fetch("/microsoft-graph-config.txt")
      .then((response) => (response.ok ? response.text() : ""))
      .then((text) => {
        if (!text) return;
        setMicrosoftGraphConfig(parseMicrosoftGraphConfig(text));
      })
      .catch(() => {
        // The Graph config is optional; anonymous/proxy loading can still work.
      });
  }, []);

  useEffect(() => {
    if (!data) return;
    try {
      localStorage.setItem(
        SAVED_DASHBOARD_KEY,
        JSON.stringify({
          data: serializeDashboardData(data),
          regions: regions.map(serializeDashboardData),
          googleRegionLinks,
          googleRegionSaveLinks,
          googleRegionSelection,
        }),
      );
      setSavedSnapshotAvailable(true);
    } catch (err) {
      console.warn("Could not save dashboard snapshot:", err);
    }
  }, [
    data,
    regions,
    googleRegionLinks,
    googleRegionSaveLinks,
    googleRegionSelection,
  ]);

  useEffect(() => {
    if (!data) return;
    document
      .querySelectorAll(".section-control-panel")
      .forEach((panel) => panel.remove());
    const cleanup: Array<() => void> = [];

    DASHBOARD_SECTIONS.forEach((section) => {
      const target = document.querySelector<HTMLElement>(section.selector);
      if (!target) return;
      const isActive = section.id === activeDashboardTab;
      const isVisible =
        isActive ||
        (activeDashboardTab === "executive" && section.id === "deepDive");
      target.classList.add("dashboard-section-content-block");
      target.classList.remove("dashboard-section-collapsed");
      target.classList.toggle("dashboard-section-tab-hidden", !isVisible);

      if (!isActive) return;

      const panel = document.createElement("div");
      panel.id = `section-control-${section.id}`;
      panel.className =
        "section-control-panel section-control-panel-tab-active no-print";
      const iconSvg = {
        top: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 19V5"/><path d="M5 12l7-7 7 7"/></svg>',
      };
      panel.innerHTML = `
        <div class="section-control-title section-control-title-static${section.id === "reports" ? " section-control-title--with-subtitle" : ""}">
          <span class="section-control-tab-dot" aria-hidden="true"></span>
          <span class="section-control-title-copy">
            <span class="section-control-title-text">${section.title}</span>
            ${section.id === "reports" ? '<span class="section-control-subtitle">Generate filtered ticket and performance reports, track export actions, and keep management outputs in one place.</span>' : ""}
          </span>
        </div>
        <div class="section-control-actions">
          ${section.id === "reports" ? `
            <label class="section-header-search reports-search-box" data-role="section-report-search">
              <svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="11" cy="11" r="7"></circle><path d="m20 20-3.5-3.5"></path></svg>
              <input type="search" data-action="report-search" value="${reportSearch.replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;").replace(/>/g, "&gt;")}" placeholder="Search reports" />
            </label>
          ` : ""}
          <button type="button" class="section-tool-button" data-action="top">${iconSvg.top}<span>Top Nav</span></button>
        </div>
      `;

      const topButton = panel.querySelector<HTMLButtonElement>(
        '[data-action="top"]',
      );
      const reportSearchInput = panel.querySelector<HTMLInputElement>(
        '[data-action="report-search"]',
      );
      const topHandler = () => scrollToDashboardTop();
      const reportSearchHandler = (event: Event) => {
        setReportSearch((event.currentTarget as HTMLInputElement).value);
      };
      topButton?.addEventListener("click", topHandler);
      reportSearchInput?.addEventListener("input", reportSearchHandler);
      target.parentElement?.insertBefore(panel, target);

      cleanup.push(() => {
        topButton?.removeEventListener("click", topHandler);
        reportSearchInput?.removeEventListener("input", reportSearchHandler);
        panel.remove();
      });
    });

    const performanceGaugeTarget = document.querySelector<HTMLElement>(
      "#section-performance-kpis",
    );
    if (performanceGaugeTarget) {
      performanceGaugeTarget.classList.add("dashboard-section-content-block");
      performanceGaugeTarget.classList.remove("dashboard-section-collapsed");
      performanceGaugeTarget.classList.toggle(
        "dashboard-section-tab-hidden",
        activeDashboardTab !== "performanceKpis",
      );
    }

    return () => cleanup.forEach((remove) => remove());
  }, [data, activeDashboardTab, reportSearch]);

  useEffect(() => {
    if (!data) return;
    document
      .querySelectorAll(".chart-export-png-button")
      .forEach((button) => button.remove());
    const chartHosts = new Set<HTMLElement>();

    const explicitChartSelector = [
      ".dashboard-chart-grid > .glass-card",
      ".chart-mosaic > .glass-card",
      ".chart-2col > .glass-card",
      ".dashboard-chart-grid .glass-card",
      ".dashboard-chart-grid article",
      ".chart-mosaic article",
      ".deep-dive-chart-grid .deep-dive-panel",
      ".client-delivery-section .deep-dive-panel",
      ".glass-card:has(.recharts-wrapper)",
      ".glass-card:has(.recharts-surface)",
      ".deep-dive-panel:has(.recharts-wrapper)",
      ".deep-dive-panel:has(.recharts-surface)",
      "article:has(.recharts-wrapper)",
      "article:has(.recharts-surface)",
    ].join(",");

    document
      .querySelectorAll<HTMLElement>(explicitChartSelector)
      .forEach((card) => {
        chartHosts.add(card);
      });

    document
      .querySelectorAll<HTMLElement>(".recharts-wrapper, .recharts-surface")
      .forEach((chart) => {
        const host = chart.closest<HTMLElement>(
          ".glass-card, .deep-dive-panel, .chart-card, article",
        );
        if (host) chartHosts.add(host);
      });

    const cleanup: Array<() => void> = [];
    Array.from(chartHosts).forEach((card, index) => {
      if (card.querySelector(".chart-export-png-button")) return;
      card.classList.add("chart-export-target");
      const title =
        card
          .querySelector("h3, h4, strong, .card-heading")
          ?.textContent?.trim() ?? `Chart ${index + 1}`;
      const button = document.createElement("button");
      button.type = "button";
      button.className = "chart-export-png-button no-print";
      button.title = `Export ${title} as PNG`;
      button.innerHTML =
        '<svg class="file-export-svg file-export-svg-png" viewBox="0 0 64 64" aria-hidden="true"><path class="file-page" d="M14 5h25l11 11v43H14Z"/><path class="file-fold" d="M39 5v12h11"/><circle class="file-mark" cx="25" cy="24" r="5"/><path class="file-mark" d="m18 49 11-13 7 8 5-6 8 11Z"/></svg><span>PNG</span>';
      const handler = (event: MouseEvent) => {
        event.preventDefault();
        event.stopPropagation();
        exportElementToPng(card, title);
      };
      button.addEventListener("click", handler);
      card.appendChild(button);
      cleanup.push(() => {
        button.removeEventListener("click", handler);
        button.remove();
        card.classList.remove("chart-export-target");
      });
    });
    return () => cleanup.forEach((remove) => remove());
  }, [
    data,
    analytics,
    executiveInsights,
    deepDiveAnalytics,
    activeDashboardTab,
  ]);

  const selectedManualRegionKey = (region: string): GoogleRegionKey | null => {
    const normalized = normalizeHeader(region);
    if (normalized === "eoa" || normalized === "neoa") return "eoaNeoa";
    if (normalized === "soa") return "soa";
    if (normalized === "coa" || normalized === "woa") return "coaWoa";
    return null;
  };

  const inferSiteRegionKey = (
    siteName: string,
    datasetRegions: Set<string>,
  ): GoogleRegionKey | null => {
    const normalizedName = normalizeHeader(siteName);
    if (normalizedName.includes("neoa") || normalizedName.includes("eoa"))
      return "eoaNeoa";
    if (normalizedName.includes("soa")) return "soa";
    if (normalizedName.includes("coa") || normalizedName.includes("woa"))
      return "coaWoa";
    const regionKeys = Array.from(datasetRegions)
      .map(selectedManualRegionKey)
      .filter((key): key is GoogleRegionKey => Boolean(key));
    const uniqueKeys = Array.from(new Set(regionKeys));
    return uniqueKeys.length === 1 ? uniqueKeys[0] : null;
  };

  const siteLookupOptions = useMemo(() => {
    const datasets = regions.length ? regions : data ? [data] : [];
    const options: Array<{
      siteId: string;
      siteName: string;
      regionKey: GoogleRegionKey | null;
    }> = [];
    const seen = new Set<string>();
    datasets.forEach((dataset) => {
      const datasetRegions = new Set(
        dataset.rows.map((row) => clean(row.region)),
      );
      dataset.siteOrder.forEach((site) => {
        const siteId = normalizeSiteId(clean(site.siteId));
        if (!siteId) return;
        const siteName = clean(site.siteName);
        const regionKey = inferSiteRegionKey(siteName, datasetRegions);
        const key = `${regionKey ?? "all"}||${siteId}`;
        if (seen.has(key)) return;
        seen.add(key);
        options.push({ siteId, siteName, regionKey });
      });
    });
    return options.sort((a, b) => {
      const byRegion = clean(a.regionKey ?? "").localeCompare(
        clean(b.regionKey ?? ""),
      );
      if (byRegion !== 0) return byRegion;
      return clean(a.siteId).localeCompare(clean(b.siteId), undefined, {
        numeric: true,
        sensitivity: "base",
      });
    });
  }, [data, regions]);

  const siteLookupOptionsForRegion = (region: string) => {
    const regionKey = selectedManualRegionKey(region);
    if (!regionKey) return [];
    return siteLookupOptions.filter(
      (site) => !site.regionKey || site.regionKey === regionKey,
    );
  };
  const actionRcaLookup = useMemo(
    () => data?.rcaLookup ?? [],
    [data?.rcaLookup],
  );
  const managedResourceOptions = useMemo(
    () =>
      Array.from(
        new Set(
          allDataRows.map((row) => clean(row.managedResource)).filter(Boolean),
        ),
      ).sort(),
    [allDataRows],
  );
  const escalatedToOptions = useMemo(
    () =>
      Array.from(
        new Set(
          allDataRows.map((row) => clean(row.escalatedTo)).filter(Boolean),
        ),
      ).sort(),
    [allDataRows],
  );

  const findSiteLookup = (value: string, region = "") => {
    const siteId = normalizeSiteId(clean(value));
    const scopedOptions = siteLookupOptionsForRegion(region);
    return scopedOptions.find(
      (item) => normalizeSiteId(clean(item.siteId)) === siteId,
    );
  };

  const findActionLookup = (value: string) => {
    const action = clean(value);
    return actionRcaLookup.find(
      (item) => normalizeHeader(item.action) === normalizeHeader(action),
    );
  };

  const updateManualDraft = (id: string, patch: Partial<ManualTicketDraft>) => {
    setManualSaveStatus("");
    setManualDrafts((prev) =>
      prev.map((draft) => (draft.id === id ? { ...draft, ...patch } : draft)),
    );
  };

  const addManualDraft = () => {
    setManualDrafts((prev) => [...prev, createManualTicketDraft()]);
    setManualSaveStatus("");
    setActiveDashboardTab("input");
  };

  const duplicateManualDraft = (id: string) => {
    setManualDrafts((prev) => {
      const source = prev.find((draft) => draft.id === id);
      if (!source) return prev;
      return [
        ...prev,
        {
          ...source,
          id: createManualTicketDraft().id,
          tt: "",
        },
      ];
    });
    setManualSaveStatus("");
  };

  const deleteManualDraft = (id: string) => {
    setManualDrafts((prev) => prev.filter((draft) => draft.id !== id));
    setManualSaveStatus("");
  };

  const handleManualSiteIdChange = (id: string, value: string) => {
    const siteId = normalizeSiteId(clean(value));
    setManualDrafts((prev) =>
      prev.map((draft) => {
        if (draft.id !== id) return draft;
        const site = findSiteLookup(siteId, draft.region);
        const customSite = draft.customSite || !site;
        return {
          ...draft,
          siteId,
          customSite,
          siteName: customSite ? draft.siteName : site?.siteName || "",
        };
      }),
    );
    setManualSaveStatus("");
  };

  const setManualSiteOverride = (id: string, customSite: boolean) => {
    setManualDrafts((prev) =>
      prev.map((draft) => {
        if (draft.id !== id) return draft;
        const site = findSiteLookup(draft.siteId, draft.region);
        return {
          ...draft,
          customSite,
          siteName: customSite ? draft.siteName : site?.siteName || "",
        };
      }),
    );
    setManualSaveStatus("");
  };

  const handleManualActionChange = (id: string, value: string) => {
    const action = clean(value);
    const match = findActionLookup(action);
    setManualDrafts((prev) =>
      prev.map((draft) => {
        if (draft.id !== id) return draft;
        const customRca = draft.customRca || !match;
        return {
          ...draft,
          action,
          customRca,
          rca: customRca ? draft.rca : match?.rca || "",
        };
      }),
    );
    setManualSaveStatus("");
  };

  const setManualRcaOverride = (id: string, customRca: boolean) => {
    setManualDrafts((prev) =>
      prev.map((draft) => {
        if (draft.id !== id) return draft;
        const match = findActionLookup(draft.action);
        return {
          ...draft,
          customRca,
          rca: customRca ? draft.rca : match?.rca || "",
        };
      }),
    );
    setManualSaveStatus("");
  };

  function googleRowToManualDraft(
    row: GoogleSheetTicketRow,
    fallbackKey: GoogleRegionKey,
  ): ManualTicketDraft {
    const siteId = normalizeSiteId(
      sheetRowValue(row, ["Site ID", "Site No", "Equipment/site"]),
    );
    const siteName = sheetRowValue(row, ["Site Name"]);
    const fallbackRegion =
      fallbackKey === "soa" ? "SOA" : fallbackKey === "coaWoa" ? "COA" : "EOA";
    const sourceRegion = sheetRowValue(row, ["Region"]) || fallbackRegion;
    const site = findSiteLookup(siteId, sourceRegion);
    const action = sheetRowValue(row, ["Action"]);
    const rca = sheetRowValue(row, ["RCA"]);
    const actionMatch = findActionLookup(action);
    return {
      id: `${Date.now()}-${row.rowNumber}-${Math.random().toString(36).slice(2)}`,
      tt: sheetRowValue(row, ["TT", "TT's", "TT Number", "Ticket Number"]),
      siteId,
      siteName: siteName || site?.siteName || "",
      customSite: Boolean(
        !site || (siteName && clean(site.siteName) !== clean(siteName)),
      ),
      managedResource: sheetRowValue(row, [
        "Managed Resource",
        "Managed Resource ",
        "Effected Managed Resource",
      ]),
      issue: sheetRowValue(row, ["Issues", "Alarm Type"]),
      severity: sheetRowValue(row, ["Severity"]),
      region: sourceRegion,
      observationDate: dateForManualInput(
        sheetRowValue(row, ["Observation Date", "Escalation Date"]),
      ),
      observationTime: normalizeManualTime(
        sheetRowValue(row, ["Observation Time", "Escalation Time"]),
      ),
      recoveryDate: dateForManualInput(sheetRowValue(row, ["Recovery Date"])),
      recoveryTime: normalizeManualTime(sheetRowValue(row, ["Recovery Time"])),
      escalatedForL3SupportDate: dateForManualInput(
        sheetRowValue(row, [
          "Escalated for L3 Support Date",
          "Escalated for L3 Support Date Time",
        ]),
      ),
      escalatedForL3SupportTime: normalizeManualTime(
        sheetRowValue(row, ["Escalated for L3 Support Time"]),
      ),
      duration: sheetRowValue(row, [
        "Total Duration/Days/Hours",
        "Total Durration Days/Hours",
        "Outage Duration",
      ]),
      durationHours: sheetRowValue(row, [
        "Duration (hrs)",
        "Duration",
        "Duration hrs",
      ]),
      neDetail: sheetRowValue(row, ["NE Detail/Impacted Object"]),
      impact:
        sheetRowValue(row, ["Service Impaction Status"]) || "Service Impact",
      correlatedAlarms: sheetRowValue(row, [
        "No. of correlated Alarms",
        "No. of corelated Alarms",
      ]),
      escalatedTo: sheetRowValue(row, ["Escalated to", "Escalated to "]),
      escalationLevel: sheetRowValue(row, ["Escalation Level"]) || "L1",
      status: sheetRowValue(row, ["Status", "TT Status"]) || "Pending",
      commentsDate: dateForManualInput(sheetRowValue(row, ["Comments Date"])),
      commentsFeedback: sheetRowValue(row, ["Comments-Feedback"]),
      maintenanceTeam: sheetRowValue(row, [
        "Maintenance person/Team",
        "Maintenance person",
        "Maintenance Team",
      ]),
      maintenanceContact: sheetRowValue(row, ["Maintenance Contact Details"]),
      actionTaken: sheetRowValue(row, ["Action Taken/RCA", "Action Taken"]),
      action,
      customRca: Boolean(
        !actionMatch || (rca && clean(actionMatch.rca) !== clean(rca)),
      ),
      rca: rca || actionMatch?.rca || "",
    };
  }

  const loadExistingTicketForEditing = async () => {
    const targetTt = clean(existingTtSearch);
    if (!targetTt || existingTtLoading) return;
    setExistingTtLoading(true);
    setManualSaveStatus(`Searching for TT ${targetTt} in online sheets...`);
    try {
      const matches: ManualTicketDraft[] = [];
      for (const region of GOOGLE_REGION_LINKS) {
        const link = googleRegionLinks[region.key]?.trim();
        const spreadsheetId = extractGoogleSpreadsheetId(link || "");
        const gid = extractGoogleSheetGid(link || "");
        if (!spreadsheetId || !gid) continue;
        const payload = await loadGoogleSheetTable(spreadsheetId, { gid });
        if (payload?.status && payload.status !== "ok") continue;
        googlePayloadToTicketRows(payload)
          .filter((row) => {
            const tt = sheetRowValue(row, [
              "TT",
              "TT's",
              "TT Number",
              "Ticket Number",
            ]);
            return clean(tt).toLowerCase() === targetTt.toLowerCase();
          })
          .forEach((row) =>
            matches.push(googleRowToManualDraft(row, region.key)),
          );
      }

      if (!matches.length) {
        setManualSaveStatus(`No rows found for TT ${targetTt}.`);
        return;
      }
      setManualDrafts(matches);
      setActiveDashboardTab("input");
      setManualSaveStatus(
        `Loaded ${matches.length} source row${matches.length === 1 ? "" : "s"} for TT ${targetTt}. Edit the fields, then Save to Source to update the same row${matches.length === 1 ? "" : "s"}.`,
      );
    } catch (error) {
      setManualSaveStatus(
        `Could not load TT ${targetTt}: ${error instanceof Error ? error.message : String(error)}`,
      );
    } finally {
      setExistingTtLoading(false);
    }
  };
  function workbookTicketSheetName(workbook: XLSX.WorkBook): string {
    return (
      workbook.SheetNames.find((name) => {
        const normalized = normalizeHeader(name);
        return normalized.includes("ticketsdata") || normalized === "tthistory";
      }) ??
      workbook.SheetNames[0] ??
      "TT-History"
    );
  }

  function regionToGoogleKey(region: string): GoogleRegionKey | null {
    const normalized = normalizeHeader(region);
    if (normalized === "eoa" || normalized === "neoa") return "eoaNeoa";
    if (normalized === "soa") return "soa";
    if (normalized === "coa" || normalized === "woa") return "coaWoa";
    return null;
  }

  function sourceWorkbookKeys(
    source: UploadedWorkbookSource,
  ): Set<GoogleRegionKey> {
    const keys = new Set<GoogleRegionKey>();
    source.data.rows.forEach((row) => {
      const key = regionToGoogleKey(row.region);
      if (key) keys.add(key);
    });
    return keys;
  }

  function sourceMatchesManualDraft(
    source: UploadedWorkbookSource,
    draft: ManualTicketDraft,
  ): boolean {
    const sourceKeys = sourceWorkbookKeys(source);
    if (!sourceKeys.size) return true;
    const draftKey = manualDraftRegionKey(draft);
    return Boolean(draftKey && sourceKeys.has(draftKey));
  }

  function nextManualWorkbookSerial(sheet: XLSX.WorkSheet): number {
    const rows = XLSX.utils.sheet_to_json<unknown[]>(sheet, {
      header: 1,
      defval: "",
    }) as unknown[][];
    for (let index = rows.length - 1; index >= 0; index -= 1) {
      const match = String(rows[index]?.[0] ?? "").match(/\d+/);
      if (match) return Number(match[0]) + 1;
    }
    return 1;
  }

  function cloneWorkbook(workbook: XLSX.WorkBook): XLSX.WorkBook {
    const buffer = XLSX.write(workbook, { type: "array", bookType: "xlsx" });
    return XLSX.read(buffer, { type: "array", cellDates: false });
  }

  function appendManualRowsToWorkbook(
    workbook: XLSX.WorkBook,
    drafts: ManualTicketDraft[],
  ): number {
    if (!drafts.length) return 0;
    const sheetName = workbookTicketSheetName(workbook);
    let sheet = workbook.Sheets[sheetName];
    if (!sheet) {
      sheet = XLSX.utils.aoa_to_sheet([[...MANUAL_TICKET_EXPORT_HEADERS]]);
      XLSX.utils.book_append_sheet(workbook, sheet, sheetName || "TT-History");
    }
    let nextSerial = nextManualWorkbookSerial(sheet);
    const rows = drafts.map((draft, index) => {
      const row = manualDraftToSheetRow(draft, index) as Record<
        string,
        unknown
      >;
      row["SN."] = nextSerial;
      nextSerial += 1;
      return MANUAL_TICKET_EXPORT_HEADERS.map(
        (header: string) => (row as Record<string, unknown>)[header] ?? "",
      );
    });
    XLSX.utils.sheet_add_aoa(sheet, rows, { origin: -1 });
    sheet["!cols"] =
      sheet["!cols"] ?? MANUAL_TICKET_EXPORT_HEADERS.map(() => ({ wch: 16 }));
    return rows.length;
  }

  const downloadUpdatedUploadedWorkbooks = () => {
    const validDrafts = manualDrafts.filter((draft) =>
      manualDraftToTicketRecord(draft, 0),
    );
    if (!validDrafts.length) {
      setManualSaveStatus(
        "Add at least one valid manual ticket row before downloading an updated workbook.",
      );
      return;
    }
    if (!uploadedWorkbookSources.length) {
      setManualSaveStatus(
        "Load an Excel workbook first, then use this button to download an updated workbook copy.",
      );
      return;
    }

    let exportedCount = 0;
    uploadedWorkbookSources.forEach((source) => {
      const sourceDrafts = validDrafts.filter((draft) =>
        sourceMatchesManualDraft(source, draft),
      );
      if (!sourceDrafts.length) return;
      const workbook = cloneWorkbook(source.workbook);
      const appended = appendManualRowsToWorkbook(workbook, sourceDrafts);
      if (!appended) return;
      const baseName = makeFileSafeName(
        source.fileName.replace(/\.[^.]+$/, ""),
      );
      XLSX.writeFile(workbook, `Updated_${baseName}.xlsx`);
      exportedCount += 1;
    });

    setManualSaveStatus(
      exportedCount
        ? `Downloaded ${exportedCount} updated workbook${exportedCount === 1 ? "" : "s"} with the manual rows appended.`
        : "No manual rows matched the loaded workbook regions.",
    );
  };
  async function postToGoogleAppsScript(
    url: string,
    payload: unknown,
  ): Promise<void> {
    const trimmedUrl = url.trim();
    if (!trimmedUrl) {
      throw new Error("Google Apps Script save URL is empty.");
    }
    await fetch(trimmedUrl, {
      method: "POST",
      mode: "no-cors",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify(payload),
    });
  }
  const saveManualDraftsToOnlineSheets = async () => {
    if (manualSaveLoading) return;
    const rows = manualDrafts
      .map((draft, index) => ({ draft, index }))
      .filter(({ draft }) => manualDraftToTicketRecord(draft, 0));
    if (!rows.length) {
      setManualSaveStatus(
        "Add at least one valid manual ticket row before saving.",
      );
      return;
    }

    const grouped = new Map<GoogleRegionKey, typeof rows>();
    rows.forEach((row) => {
      const key = manualDraftRegionKey(row.draft);
      if (!key) return;
      const current = grouped.get(key) ?? [];
      current.push(row);
      grouped.set(key, current);
    });

    const missing = Array.from(grouped.keys()).filter(
      (key) => !googleRegionSaveLinks[key]?.trim(),
    );
    if (missing.length) {
      const labels = missing
        .map(
          (key) =>
            GOOGLE_REGION_LINKS.find((region) => region.key === key)?.label,
        )
        .filter(Boolean)
        .join(", ");
      setManualSaveStatus(
        `Missing save URL in public/google-sheets-config.txt for: ${labels}. Add lines like SOA_SAVE=https://script.google.com/macros/s/.../exec`,
      );
      return;
    }

    setManualSaveLoading(true);
    setManualSaveStatus("Saving manual rows to online sheet...");
    try {
      await Promise.all(
        Array.from(grouped.entries()).map(([key, regionRows]) =>
          postToGoogleAppsScript(googleRegionSaveLinks[key].trim(), {
            region: key,
            spreadsheetId: extractGoogleSpreadsheetId(googleRegionLinks[key]),
            sheetGid: extractGoogleSheetGid(googleRegionLinks[key]),
            sheetName: "TT-History",
            rows: regionRows.map(({ draft, index }) =>
              manualDraftToSheetRow(draft, index),
            ),
          }),
        ),
      );
      setManualSaveStatus(
        `Manual rows submitted to: ${Array.from(grouped.keys())
          .map(
            (key) =>
              GOOGLE_REGION_LINKS.find((region) => region.key === key)?.label,
          )
          .filter(Boolean)
          .join(
            ", ",
          )}. Reload the Google Sheet to confirm the appended rows, then clear local rows when ready.`,
      );
    } catch (error) {
      setManualSaveStatus(
        `Could not submit manual rows: ${error instanceof Error ? error.message : String(error)}`,
      );
    } finally {
      setManualSaveLoading(false);
    }
  };
  const getDashboardSectionIcon = (sectionId: DashboardSectionId) => {
    const Icon =
      sectionId === "input"
        ? FileSpreadsheet
        : sectionId === "performanceKpis"
          ? Activity
          : sectionId === "ticketsTable"
            ? FileSpreadsheet
            : sectionId === "overviewCharts"
              ? BarChart3
              : sectionId === "executive"
                ? ShieldAlert
                : sectionId === "trendCharts"
                  ? BarChart3
                  : Presentation;

    return <Icon size={15} strokeWidth={2.5} aria-hidden="true" />;
  };
  const filtersPanel =
    data &&
    activeDashboardTab !== "reports" &&
    activeDashboardTab !== "input" ? (
      <section
        ref={filtersRef}
        className="filters-panel no-print dashboard-filters-panel"
        style={{ width: "100%" }}
      >
        <label className="search-box" style={{ width: "100%", margin: 0 }}>
          <span>Search</span>
          <Search size={16} />
          <input
            value={filters.search}
            onChange={(event) =>
              setFilters((prev) => ({ ...prev, search: event.target.value }))
            }
            placeholder="Search..."
            style={{ width: "100%" }}
          />
        </label>
        <div style={{ width: "100%", overflow: "hidden" }}>
          <MultiSelectFilter
            label="Status"
            value={filters.status}
            options={filterOptions.status}
            onChange={(value) =>
              setFilters((prev) => ({ ...prev, status: value }))
            }
          />
        </div>
        <div style={{ width: "100%", overflow: "hidden" }}>
          <MultiSelectFilter
            label="Severity"
            value={filters.severity}
            options={filterOptions.severity}
            onChange={(value) =>
              setFilters((prev) => ({ ...prev, severity: value }))
            }
          />
        </div>
        <div style={{ width: "100%", overflow: "hidden" }}>
          <MultiSelectFilter
            label="Region"
            value={filters.region}
            options={filterOptions.region}
            onChange={(value) =>
              setFilters((prev) => ({ ...prev, region: value }))
            }
          />
        </div>
        <div style={{ width: "100%", overflow: "hidden" }}>
          <MultiSelectFilter
            label="Impact"
            value={filters.impact}
            options={filterOptions.impact}
            onChange={(value) =>
              setFilters((prev) => ({ ...prev, impact: value }))
            }
          />
        </div>
        <div style={{ width: "100%", overflow: "hidden" }}>
          <MultiSelectFilter
            label="Opening Month"
            value={filters.openingMonth}
            options={filterOptions.openingMonth}
            optionLabels={filterOptions.openingMonthLabels}
            onChange={(value) =>
              setFilters((prev) => ({ ...prev, openingMonth: value }))
            }
          />
        </div>
        <div style={{ width: "100%", overflow: "hidden" }}>
          <MultiSelectFilter
            label="Site"
            value={filters.site}
            options={filterOptions.site}
            onChange={(value) =>
              setFilters((prev) => ({ ...prev, site: value }))
            }
          />
        </div>
        <div style={{ width: "100%", overflow: "hidden" }}>
          <MultiSelectFilter
            label="RCA Family"
            value={filters.rcaFamily}
            options={filterOptions.rcaFamily}
            onChange={(value) =>
              setFilters((prev) => ({ ...prev, rcaFamily: value }))
            }
          />
        </div>
        <button
          className="ghost-button"
          style={{ margin: 0, whiteSpace: "nowrap" }}
          onClick={() => {
            setFilters(EMPTY_FILTERS);
            setTablePage(1);
          }}
        >
          <Filter size={16} /> Clear
        </button>
      </section>
    ) : null;

  if (!isAuthenticated) {
    return (
      <main
        className="dashboard-shell dashboard-login-shell"
        data-dashboard-theme={dashboardTheme}
      >
        <section className="hero-panel hero-panel--corporate-banner hero-panel--login-banner">
          <CorporateBannerHeader />
        </section>
        <section
          className="dashboard-login-panel"
          style={{
            backgroundImage: `${heroThemeOverlay}, url(${activeThemeImage})`,
          }}
        >
          <div className="dashboard-login-brand">
            <img src={nascoLogoSrc} alt="NASCO" />
            <div>
              <span>DMR Ticketing Dashboard</span>
              <strong>Secure access</strong>
            </div>
          </div>

          <form className="dashboard-login-card" onSubmit={handleLoginSubmit}>
            <div className="dashboard-login-heading">
              <ShieldAlert size={34} />
              <div>
                <span>Welcome Back</span>
                <h1>Sign in</h1>
              </div>
            </div>

            <label className="dashboard-login-field">
              <span>Username</span>
              <input
                type="text"
                value={loginUsername}
                onChange={(event) => {
                  setLoginUsername(event.target.value);
                  setLoginError("");
                }}
                autoComplete="username"
                autoFocus
              />
            </label>

            <label className="dashboard-login-field">
              <span>Password</span>
              <input
                type="password"
                value={loginPassword}
                onChange={(event) => {
                  setLoginPassword(event.target.value);
                  setLoginError("");
                }}
                autoComplete="current-password"
              />
            </label>

            {loginError && (
              <p className="dashboard-login-error">{loginError}</p>
            )}

            <div className="dashboard-login-actions">
              <button type="submit" className="primary-button">
                Login
              </button>
              {renderThemeToggle("theme-toggle--action")}
            </div>
          </form>
        </section>
      </main>
    );
  }
  return (
    <main
      className={`dashboard-shell ${data ? "dashboard-shell--loaded" : "dashboard-shell--welcome"}`}
      data-dashboard-theme={dashboardTheme}
      data-active-tab={activeDashboardTab}
    >
      <section className="hero-panel hero-panel--corporate-banner">
        <CorporateBannerHeader />

        {data && (
          <div className="dashboard-action-bar no-print" aria-label="Dashboard actions">
            <button
              type="button"
              className="ghost-button"
              onClick={returnToWelcomeUploadScreen}
              title="Return to the welcome upload screen"
            >
              <HomeIcon size={30} /> Home
            </button>
            <button
              className="ghost-button"
              onClick={() =>
                setOnlineSourceMode((prev) => (prev === "add" ? null : "add"))
              }
            >
              <UploadCloud size={30} /> Add regions
            </button>
            <button
              className="ghost-button"
              onClick={() =>
                setOnlineSourceMode((prev) =>
                  prev === "replace" ? null : "replace",
                )
              }
            >
              <RefreshCw size={30} /> New workbook(s)
            </button>
            <button
              type="button"
              className="ghost-button"
              onClick={clearSavedDashboardData}
              title="Clear saved workbook/session data stored by this dashboard"
            >
              <ShieldAlert size={30} /> Clear Saved Data
            </button>
            <button className="primary-button" onClick={() => window.print()}>
              <Printer size={30} /> Dashboard PDF
            </button>
            {renderThemeToggle("theme-toggle--action")}
            <button
              type="button"
              className="ghost-button"
              onClick={handleLogout}
              title="Logout from the dashboard"
            >
              <LogOut size={30} /> Logout
            </button>
          </div>
        )}
        {data && (
          <div
            id="dashboard-section-nav"
            className="dashboard-section-nav dashboard-tabs-nav no-print"
            aria-label="Dashboard tabs navigation"
          >
            <div
              className="section-nav-buttons section-tabs-buttons"
              role="tablist"
              aria-label="Dashboard sections"
            >
              {DASHBOARD_SECTIONS.filter(
                (section) => section.id !== "deepDive",
              ).map((section) => (
                <button
                  key={section.id}
                  type="button"
                  role="tab"
                  aria-selected={activeDashboardTab === section.id}
                  className={`section-nav-button section-tab-button${activeDashboardTab === section.id ? " section-tab-button-active" : ""}`}
                  onClick={() => openDashboardTab(section.id)}
                >
                  {section.label}
                </button>
              ))}
            </div>
            <div className="section-tabs-current" aria-live="polite">
              {activeDashboardSection.title}
            </div>
          </div>
        )}
        {data && onlineSourceMode && (
          <div className="online-source-panel no-print">
            <div className="online-source-panel__head">
              <div>
                <span>
                  {onlineSourceMode === "add"
                    ? "Add regional sources"
                    : "Load new workbook sources"}
                </span>
                <strong>
                  Choose local Excel files or load the selected Google/Microsoft
                  365 links.
                </strong>
              </div>
              <button
                type="button"
                className="ghost-button"
                onClick={() => setOnlineSourceMode(null)}
              >
                Close
              </button>
            </div>
            <div className="online-source-actions">
              <button
                type="button"
                className="primary-button"
                onClick={() =>
                  onlineSourceMode === "add"
                    ? addRegionRef.current?.click()
                    : inputRef.current?.click()
                }
              >
                <FileSpreadsheet size={18} />
                {onlineSourceMode === "add"
                  ? "Select Excel region file(s)"
                  : "Select Excel workbook(s)"}
              </button>
              <button
                type="button"
                className="ghost-button"
                onClick={() => handleGoogleSheetLoad(onlineSourceMode)}
                disabled={
                  googleSheetLoading ||
                  !GOOGLE_REGION_LINKS.some(
                    (region) =>
                      googleRegionSelection[region.key] &&
                      googleRegionLinks[region.key].trim(),
                  )
                }
              >
                <LinkIcon size={18} />
                {googleSheetLoading
                  ? "Loading..."
                  : onlineSourceMode === "add"
                    ? "Add selected online sheets"
                    : "Load selected online sheets"}
              </button>
            </div>
            <div className="google-sheet-loader online-source-links">
              <label>Online sheet region links</label>
              <div className="google-region-grid">
                {GOOGLE_REGION_LINKS.map((region) => (
                  <label key={region.key} className="google-region-row">
                    <input
                      type="checkbox"
                      checked={googleRegionSelection[region.key]}
                      onChange={(event) =>
                        setGoogleRegionChecked(region.key, event.target.checked)
                      }
                    />
                    <span>{region.label}</span>
                    <input
                      type="url"
                      value={googleRegionLinks[region.key]}
                      onChange={(event) =>
                        setGoogleRegionLink(region.key, event.target.value)
                      }
                      placeholder="Google TT-History URL or Microsoft 365 Excel link"
                    />
                  </label>
                ))}
              </div>
            </div>
          </div>
        )}
        {filtersPanel}

        {data && (
          <>
            <section
              id="section-manual-input"
              className="table-card manual-input-section dashboard-section-content-block"
            >
              <div className="table-heading manual-input-heading">
                <div>
                  <h2>Manual Ticket Input</h2>
                  <p style={{ color: "#94a3b8", fontSize: 13, marginTop: 4 }}>
                    Add tickets manually; saved rows are included in dashboard
                    tables, charts, KPIs, and exports.
                  </p>
                </div>
                <div className="manual-input-actions no-print">
                  <button
                    className="primary-button"
                    type="button"
                    onClick={addManualDraft}
                  >
                    Add Row
                  </button>
                  <button
                    className="primary-button"
                    type="button"
                    onClick={saveManualDraftsToOnlineSheets}
                    disabled={!manualDrafts.length || manualSaveLoading}
                  >
                    <Save size={16} />
                    {manualSaveLoading ? "Saving..." : "Save to Source"}
                  </button>
                  <button
                    className="ghost-button"
                    type="button"
                    onClick={downloadUpdatedUploadedWorkbooks}
                    disabled={
                      !manualDrafts.length || !uploadedWorkbookSources.length
                    }
                  >
                    <Download size={16} />
                    Download Updated Workbook
                  </button>
                  <button
                    className="ghost-button"
                    type="button"
                    onClick={() => setManualDrafts([])}
                    disabled={!manualDrafts.length}
                  >
                    Clear Manual Rows
                  </button>
                </div>
              </div>

              <div className="manual-input-tt-lookup no-print">
                <label>
                  <span>Open old TT</span>
                  <input
                    value={existingTtSearch}
                    placeholder="Enter TT number"
                    onChange={(event) =>
                      setExistingTtSearch(event.target.value)
                    }
                    onKeyDown={(event) => {
                      if (event.key === "Enter") loadExistingTicketForEditing();
                    }}
                  />
                </label>
                <button
                  className="primary-button"
                  type="button"
                  onClick={loadExistingTicketForEditing}
                  disabled={!clean(existingTtSearch) || existingTtLoading}
                >
                  <Search size={16} />
                  {existingTtLoading ? "Searching..." : "Load TT"}
                </button>
              </div>

              <datalist id="manual-managed-resource-options">
                {managedResourceOptions.map((value) => (
                  <option key={value} value={value} />
                ))}
              </datalist>
              <datalist id="manual-escalated-to-options">
                {escalatedToOptions.map((value) => (
                  <option key={value} value={value} />
                ))}
              </datalist>
              <datalist id="manual-action-options">
                {actionRcaLookup.map((item) => (
                  <option
                    key={item.action}
                    value={item.action}
                    label={item.rca}
                  />
                ))}
              </datalist>

              {manualSaveStatus && (
                <div className="manual-input-status">{manualSaveStatus}</div>
              )}

              {manualDrafts.length === 0 ? (
                <div className="manual-input-empty">
                  No manual ticket rows yet. Use Add Row to start entering
                  tickets.
                </div>
              ) : (
                <div className="manual-input-table-wrap">
                  <table className="manual-input-table">
                    <thead>
                      <tr>
                        <th>SN.</th>
                        <th>Region</th>
                        <th>TT's</th>
                        <th>Site ID</th>
                        <th>Site Name</th>
                        <th>Managed Resource</th>
                        <th>Issues</th>
                        <th>Severity</th>
                        <th>Observation Date</th>
                        <th>Observation Time</th>
                        <th>Recovery Date</th>
                        <th>Recovery Time</th>
                        <th>L3 Date</th>
                        <th>L3 Time</th>
                        <th>Total Duration</th>
                        <th>Duration (hrs)</th>
                        <th>NE Detail</th>
                        <th>Service Impact</th>
                        <th>Correlated Alarms</th>
                        <th>Escalated to</th>
                        <th>Escalation Level</th>
                        <th>Status</th>
                        <th>Comments Date</th>
                        <th>Comments Feedback</th>
                        <th>Maintenance Team</th>
                        <th>Maintenance Contact</th>
                        <th>Action Taken/RCA</th>
                        <th>Action</th>
                        <th>RCA</th>
                        <th>Tools</th>
                      </tr>
                    </thead>
                    <tbody>
                      {manualDrafts.map((draft, index) => {
                        const rowEnabled = Boolean(clean(draft.region));
                        return (
                          <tr
                            key={draft.id}
                            className={
                              !rowEnabled
                                ? "manual-input-row-locked"
                                : undefined
                            }
                          >
                            <td>{index + 1}</td>
                            <td>
                              <select
                                value={draft.region}
                                onChange={(event) => {
                                  const nextRegion = event.target.value;
                                  const site = draft.siteId
                                    ? findSiteLookup(draft.siteId, nextRegion)
                                    : null;
                                  updateManualDraft(draft.id, {
                                    region: nextRegion,
                                    customSite:
                                      draft.customSite ||
                                      !draft.siteId ||
                                      !site,
                                    siteName: draft.customSite
                                      ? draft.siteName
                                      : site?.siteName || "",
                                  });
                                }}
                              >
                                <option value="">Select region</option>
                                {REGION_OPTIONS.map((option) => (
                                  <option key={option} value={option}>
                                    {option}
                                  </option>
                                ))}
                              </select>
                            </td>
                            <td>
                              <input
                                value={draft.tt}
                                disabled={!rowEnabled}
                                onChange={(event) =>
                                  updateManualDraft(draft.id, {
                                    tt: event.target.value,
                                  })
                                }
                              />
                            </td>
                            <td>
                              <input
                                list={`manual-site-id-options-${draft.id}`}
                                value={draft.siteId}
                                disabled={!rowEnabled}
                                onChange={(event) =>
                                  handleManualSiteIdChange(
                                    draft.id,
                                    event.target.value,
                                  )
                                }
                              />
                              <datalist
                                id={`manual-site-id-options-${draft.id}`}
                              >
                                {siteLookupOptionsForRegion(draft.region).map(
                                  (site) => (
                                    <option
                                      key={`${site.regionKey ?? "all"}-${site.siteId}`}
                                      value={normalizeSiteId(
                                        clean(site.siteId),
                                      )}
                                      label={site.siteName}
                                    />
                                  ),
                                )}
                              </datalist>
                            </td>
                            <td>
                              <div className="manual-input-lookup-cell">
                                <input
                                  value={draft.siteName}
                                  disabled={!rowEnabled}
                                  readOnly={!draft.customSite}
                                  title={
                                    draft.customSite
                                      ? "Manual site name"
                                      : "Auto-filled from Site ID sheet"
                                  }
                                  onChange={(event) =>
                                    updateManualDraft(draft.id, {
                                      siteName: event.target.value,
                                    })
                                  }
                                />
                                <label className="manual-input-override">
                                  <input
                                    type="checkbox"
                                    disabled={!rowEnabled}
                                    checked={Boolean(draft.customSite)}
                                    onChange={(event) =>
                                      setManualSiteOverride(
                                        draft.id,
                                        event.target.checked,
                                      )
                                    }
                                  />
                                  Manual site
                                </label>
                              </div>
                            </td>
                            <td>
                              <input
                                list="manual-managed-resource-options"
                                value={draft.managedResource}
                                disabled={!rowEnabled}
                                onChange={(event) =>
                                  updateManualDraft(draft.id, {
                                    managedResource: event.target.value,
                                  })
                                }
                              />
                            </td>
                            <td>
                              <input
                                value={draft.issue}
                                disabled={!rowEnabled}
                                onChange={(event) =>
                                  updateManualDraft(draft.id, {
                                    issue: event.target.value,
                                  })
                                }
                              />
                            </td>
                            <td>
                              <select
                                value={draft.severity}
                                disabled={!rowEnabled}
                                onChange={(event) =>
                                  updateManualDraft(draft.id, {
                                    severity: event.target.value,
                                  })
                                }
                              >
                                {SEVERITY_OPTIONS.map((option) => (
                                  <option key={option} value={option}>
                                    {option}
                                  </option>
                                ))}
                              </select>
                            </td>
                            <td>
                              <input
                                type="date"
                                value={draft.observationDate}
                                onChange={(event) =>
                                  updateManualDraft(draft.id, {
                                    observationDate: event.target.value,
                                  })
                                }
                              />
                            </td>
                            <td>
                              <input
                                inputMode="numeric"
                                maxLength={5}
                                pattern="^([01]\\d|2[0-3]):[0-5]\\d$"
                                placeholder="HH:mm"
                                value={draft.observationTime}
                                onChange={(event) =>
                                  updateManualDraft(draft.id, {
                                    observationTime: normalizeManualTime(
                                      event.target.value,
                                    ),
                                  })
                                }
                              />
                            </td>
                            <td>
                              <input
                                type="date"
                                value={draft.recoveryDate}
                                onChange={(event) =>
                                  updateManualDraft(draft.id, {
                                    recoveryDate: event.target.value,
                                  })
                                }
                              />
                            </td>
                            <td>
                              <input
                                inputMode="numeric"
                                maxLength={5}
                                pattern="^([01]\\d|2[0-3]):[0-5]\\d$"
                                placeholder="HH:mm"
                                value={draft.recoveryTime}
                                onChange={(event) =>
                                  updateManualDraft(draft.id, {
                                    recoveryTime: normalizeManualTime(
                                      event.target.value,
                                    ),
                                  })
                                }
                              />
                            </td>
                            <td>
                              <input
                                type="date"
                                value={draft.escalatedForL3SupportDate}
                                onChange={(event) =>
                                  updateManualDraft(draft.id, {
                                    escalatedForL3SupportDate:
                                      event.target.value,
                                  })
                                }
                              />
                            </td>
                            <td>
                              <input
                                inputMode="numeric"
                                maxLength={5}
                                pattern="^([01]\\d|2[0-3]):[0-5]\\d$"
                                placeholder="HH:mm"
                                value={draft.escalatedForL3SupportTime}
                                onChange={(event) =>
                                  updateManualDraft(draft.id, {
                                    escalatedForL3SupportTime:
                                      normalizeManualTime(event.target.value),
                                  })
                                }
                              />
                            </td>
                            <td>
                              <input
                                className="manual-input-calculated"
                                value={formatManualTotalDuration(draft)}
                                readOnly
                                title="Calculated from observation/recovery date and HH:mm time"
                              />
                            </td>
                            <td>
                              <input
                                className="manual-input-calculated"
                                value={formatManualDurationHours(draft)}
                                readOnly
                                title="Calculated from observation/recovery date and HH:mm time"
                              />
                            </td>
                            <td>
                              <input
                                value={draft.neDetail}
                                onChange={(event) =>
                                  updateManualDraft(draft.id, {
                                    neDetail: event.target.value,
                                  })
                                }
                              />
                            </td>
                            <td>
                              <select
                                value={draft.impact}
                                onChange={(event) =>
                                  updateManualDraft(draft.id, {
                                    impact: event.target.value,
                                  })
                                }
                              >
                                {IMPACT_OPTIONS.map((option) => (
                                  <option key={option} value={option}>
                                    {option}
                                  </option>
                                ))}
                              </select>
                            </td>
                            <td>
                              <input
                                value={draft.correlatedAlarms}
                                onChange={(event) =>
                                  updateManualDraft(draft.id, {
                                    correlatedAlarms: event.target.value,
                                  })
                                }
                              />
                            </td>
                            <td>
                              <input
                                list="manual-escalated-to-options"
                                value={draft.escalatedTo}
                                onChange={(event) =>
                                  updateManualDraft(draft.id, {
                                    escalatedTo: event.target.value,
                                  })
                                }
                              />
                            </td>
                            <td>
                              <select
                                value={draft.escalationLevel}
                                onChange={(event) =>
                                  updateManualDraft(draft.id, {
                                    escalationLevel: event.target.value,
                                  })
                                }
                              >
                                {ESCALATION_LEVEL_OPTIONS.map((option) => (
                                  <option key={option} value={option}>
                                    {option}
                                  </option>
                                ))}
                              </select>
                            </td>
                            <td>
                              <select
                                value={draft.status}
                                onChange={(event) =>
                                  updateManualDraft(draft.id, {
                                    status: event.target.value,
                                  })
                                }
                              >
                                {STATUS_OPTIONS.map((option) => (
                                  <option key={option} value={option}>
                                    {option}
                                  </option>
                                ))}
                              </select>
                            </td>
                            <td>
                              <input
                                type="date"
                                value={draft.commentsDate}
                                onChange={(event) =>
                                  updateManualDraft(draft.id, {
                                    commentsDate: event.target.value,
                                  })
                                }
                              />
                            </td>
                            <td>
                              <input
                                value={draft.commentsFeedback}
                                onChange={(event) =>
                                  updateManualDraft(draft.id, {
                                    commentsFeedback: event.target.value,
                                  })
                                }
                              />
                            </td>
                            <td>
                              <input
                                value={draft.maintenanceTeam}
                                onChange={(event) =>
                                  updateManualDraft(draft.id, {
                                    maintenanceTeam: event.target.value,
                                  })
                                }
                              />
                            </td>
                            <td>
                              <input
                                value={draft.maintenanceContact}
                                onChange={(event) =>
                                  updateManualDraft(draft.id, {
                                    maintenanceContact: event.target.value,
                                  })
                                }
                              />
                            </td>
                            <td>
                              <input
                                value={draft.actionTaken}
                                onChange={(event) =>
                                  updateManualDraft(draft.id, {
                                    actionTaken: event.target.value,
                                  })
                                }
                              />
                            </td>
                            <td>
                              <input
                                list="manual-action-options"
                                value={draft.action}
                                onChange={(event) =>
                                  handleManualActionChange(
                                    draft.id,
                                    event.target.value,
                                  )
                                }
                              />
                            </td>
                            <td>
                              <div className="manual-input-lookup-cell">
                                <input
                                  value={draft.rca}
                                  readOnly={!draft.customRca}
                                  title={
                                    draft.customRca
                                      ? "Manual RCA"
                                      : "Auto-filled from RCA sheet by Action"
                                  }
                                  onChange={(event) =>
                                    updateManualDraft(draft.id, {
                                      rca: event.target.value,
                                    })
                                  }
                                />
                                <label className="manual-input-override">
                                  <input
                                    type="checkbox"
                                    checked={Boolean(draft.customRca)}
                                    onChange={(event) =>
                                      setManualRcaOverride(
                                        draft.id,
                                        event.target.checked,
                                      )
                                    }
                                  />
                                  Manual RCA
                                </label>
                              </div>
                            </td>
                            <td>
                              <div className="manual-input-row-actions">
                                <button
                                  type="button"
                                  className="ghost-button"
                                  onClick={() => duplicateManualDraft(draft.id)}
                                >
                                  Copy
                                </button>
                                <button
                                  type="button"
                                  className="ghost-button"
                                  onClick={() => deleteManualDraft(draft.id)}
                                >
                                  Delete
                                </button>
                              </div>
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              )}
            </section>

            <section
              id="section-reports"
              className="reports-management-center dashboard-section-content-block no-print"
            >
              <div className="reports-management-grid">
                <aside className="reports-generator-panel">
                  <button
                    type="button"
                    className="reports-generate-hero"
                    onClick={handleGenerateManagedReport}
                  >
                    <span className="reports-generate-icon">
                      <ReportFileIcon kind={managedReportFormat} />
                    </span>
                    <span>
                      <strong>Generate New Report</strong>
                      <small>{selectedManagedReport.title}</small>
                    </span>
                  </button>

                  <div className="reports-form-card">
                    <h3>Report Generation Form</h3>
                    <label className="reports-field">
                      <span>Select Report Type</span>
                      <select
                        value={managedReportType}
                        onChange={(event) =>
                          setManagedReportType(event.target.value as ManagedReportType)
                        }
                      >
                        {managedReportDefinitions.map((report) => (
                          <option key={report.id} value={report.id}>
                            {report.title}
                          </option>
                        ))}
                      </select>
                    </label>

                    <div className="reports-form-row reports-form-row--two">
                      <label className="reports-field">
                        <span>Filtered Records</span>
                        <input readOnly value={`${selectedManagedReport.records.toLocaleString()} records`} />
                      </label>

                      <label className="reports-field">
                        <span>Output Format</span>
                        <select
                          value={managedReportFormat}
                          onChange={(event) =>
                            setManagedReportFormat(event.target.value as ManagedReportFormat)
                          }
                        >
                          {managedReportFormatOptions.map((format) => (
                            <option key={format} value={format}>
                              {format.toUpperCase()}
                            </option>
                          ))}
                        </select>
                      </label>
                    </div>

                    <div className="reports-selected-note">
                      <span>Report Scope</span>
                      <p>{selectedManagedReport.description}</p>
                    </div>
                  </div>

                  <div className="reports-filter-card reports-filter-card--shared">
                    <h3>Shared Report Filters</h3>
                    <div className="reports-filter-row reports-filter-row--two">
                      <MultiSelectFilter
                        label="Performance Month"
                        value={exportMonths}
                        options={filterOptions.exportMonth}
                        optionLabels={filterOptions.exportMonthLabels}
                        onChange={setExportMonths}
                        showAllOption
                      />
                      <MultiSelectFilter
                        label="Performance Region"
                        value={exportRegions}
                        options={filterOptions.region}
                        onChange={setExportRegions}
                        showAllOption
                      />
                    </div>
                    <p className="reports-filter-hint">
                      These filters are shared by all report types and stay unchanged when switching reports.
                    </p>
                  </div>
                </aside>

                <section className="reports-history-panel">
                  <div className="reports-history-title-row">
                    <div>
                      <h3>Generated Reports History</h3>
                      <p className="reports-download-note">
                        Generated files are saved by your browser to the Downloads folder.
                      </p>
                    </div>

                    <div className="reports-history-actions">
                      <span>{visibleGeneratedReports.length.toLocaleString()} recent</span>
                      <button
                        type="button"
                        className="reports-clear-history-button"
                        onClick={clearGeneratedReportsHistory}
                        disabled={generatedReports.length === 0}
                      >
                        Clear History
                      </button>
                    </div>
                  </div>

                  <div className="reports-quick-actions">
                    <button
                      type="button"
                      className="reports-action-card"
                      onClick={() => {
                        setManagedReportType("tickets");
                        setManagedReportFormat("xlsx");
                      }}
                    >
                      <ReportFileIcon kind="xlsx" />
                      <span>
                        <strong>Tickets XLSX</strong>
                        <small>{monthlyExportTickets.length.toLocaleString()} filtered tickets</small>
                      </span>
                    </button>
                    <button
                      type="button"
                      className="reports-action-card"
                      onClick={() => {
                        setManagedReportType("tickets");
                        setManagedReportFormat("ppt");
                      }}
                    >
                      <ReportFileIcon kind="ppt" />
                      <span>
                        <strong>Tickets PPT</strong>
                        <small>Executive ticket briefing</small>
                      </span>
                    </button>
                    <button
                      type="button"
                      className="reports-action-card"
                      onClick={() => {
                        setManagedReportType("performance");
                        setManagedReportFormat("xlsx");
                      }}
                    >
                      <ReportFileIcon kind="xlsx" />
                      <span>
                        <strong>Performance XLSX</strong>
                        <small>{reportPerformanceExportMetrics.reportRows.toLocaleString()} site rows</small>
                      </span>
                    </button>
                    <button
                      type="button"
                      className="reports-action-card"
                      onClick={() => {
                        setManagedReportType("executive");
                        setManagedReportFormat("xlsx");
                      }}
                    >
                      <ReportFileIcon kind="xlsx" />
                      <span>
                        <strong>Executive Workbook</strong>
                        <small>Health score and high-risk sites</small>
                      </span>
                    </button>
                    <button
                      type="button"
                      className="reports-action-card"
                      onClick={() => {
                        setManagedReportType("quality");
                        setManagedReportFormat("xlsx");
                      }}
                    >
                      <ReportFileIcon kind="xlsx" />
                      <span>
                        <strong>RCA / SLA Workbook</strong>
                        <small>Preventability and repeated offenders</small>
                      </span>
                    </button>
                    <button
                      type="button"
                      className="reports-action-card"
                      onClick={() => {
                        setManagedReportType("kpiCards");
                        setManagedReportFormat("png");
                      }}
                    >
                      <ReportFileIcon kind="png" />
                      <span>
                        <strong>KPI Cards PNG</strong>
                        <small>Full KPI card grid image</small>
                      </span>
                    </button>
                    <button
                      type="button"
                      className="reports-action-card"
                      onClick={() => {
                        setManagedReportType("kpiCards");
                        setManagedReportFormat("pdf");
                      }}
                    >
                      <ReportFileIcon kind="pdf" />
                      <span>
                        <strong>KPI Cards PDF</strong>
                        <small>Full KPI card grid PDF</small>
                      </span>
                    </button>
                    <button
                      type="button"
                      className="reports-action-card"
                      onClick={() => {
                        setManagedReportType("performanceCards");
                        setManagedReportFormat("png");
                      }}
                    >
                      <ReportFileIcon kind="png" />
                      <span>
                        <strong>Performance Cards PNG</strong>
                        <small>Full gauge card grid image</small>
                      </span>
                    </button>
                    <button
                      type="button"
                      className="reports-action-card"
                      onClick={() => {
                        setManagedReportType("performanceCards");
                        setManagedReportFormat("pdf");
                      }}
                    >
                      <ReportFileIcon kind="pdf" />
                      <span>
                        <strong>Performance Cards PDF</strong>
                        <small>Full gauge card grid PDF</small>
                      </span>
                    </button>
                  </div>

                  <div className="reports-summary-strip">
                    <span><b>{monthlyExportTickets.length.toLocaleString()}</b>Tickets</span>
                    <span><b>{ticketExportMetrics.pendingCount.toLocaleString()}</b>Pending</span>
                    <span><b>{ticketExportMetrics.criticalCount.toLocaleString()}</b>Critical</span>
                    <span><b>{reportPerformanceExportMetrics.totalSites.toLocaleString()}</b>RF Sites</span>
                    <span><b>{reportPerformanceExportMetrics.affectedSites.toLocaleString()}</b>Affected</span>
                    <span><b>{executiveInsights.healthScore.score}</b>Health</span>
                  </div>

                  <div className="reports-history-table-wrap">
                    <table className="reports-history-table">
                      <thead>
                        <tr>
                          <th>File Name</th>
                          <th>Report Type</th>
                          <th>Generated Date</th>
                          <th>Records</th>
                          <th>Format</th>
                        </tr>
                      </thead>
                      <tbody>
                        {visibleGeneratedReports.length === 0 ? (
                          <tr>
                            <td colSpan={5}>No reports generated in this session yet.</td>
                          </tr>
                        ) : (
                          visibleGeneratedReports.map((report) => (
                            <tr key={report.id}>
                              <td>{report.fileName}</td>
                              <td>{report.reportType}</td>
                              <td>{report.generatedAt}</td>
                              <td>{report.records.toLocaleString()}</td>
                              <td>
                                <span className={`reports-format-pill reports-format-pill--${report.format}`}>
                                  {report.format.toUpperCase()}
                                </span>
                              </td>
                            </tr>
                          ))
                        )}
                      </tbody>
                    </table>
                  </div>
                </section>
              </div>
            </section>

          </>
        )}
      </section>

      <input
        ref={inputRef}
        className="sr-only"
        type="file"
        accept=".xlsx,.xls,.csv"
        multiple
        onChange={(event) => handleFile(event.target.files)}
      />
      <input
        ref={addRegionRef}
        className="sr-only"
        type="file"
        accept=".xlsx,.xls,.csv"
        multiple
        onChange={(event) => handleAddRegion(event.target.files)}
      />

      {!data ? (
        <section className="upload-stage no-print">
          <div
            className={`upload-card ${isDragging ? "dragging" : ""}`}
            onDragOver={(event) => {
              event.preventDefault();
              setIsDragging(true);
            }}
            onDragLeave={() => setIsDragging(false)}
            onDrop={(event) => {
              event.preventDefault();
              setIsDragging(false);
              handleFile(event.dataTransfer.files);
            }}
          >
            <div className="upload-copy">
              <div className="upload-kicker-row">
                <span className="section-kicker">
                  <UploadCloud size={14} /> Workbooks Upload
                </span>
              </div>
              <h2>Load the tickets workbook(s)</h2>
              <p>
                Start with the main Follow-Up Sheets workbook, then add regional
                workbooks from the dashboard header after the first file is
                loaded.
              </p>
              {error && (
                <div className="error-banner">
                  <AlertTriangle size={16} /> {error}
                </div>
              )}
              <div className="upload-actions">
                <button
                  className="primary-button large"
                  onClick={() => inputRef.current?.click()}
                >
                  <FileSpreadsheet size={20} /> Select Excel workbook
                </button>
                {savedSnapshotAvailable && (
                  <button
                    type="button"
                    className="ghost-button large"
                    onClick={loadSavedDashboardSnapshot}
                  >
                    <RefreshCw size={20} /> Continue previous workbook
                  </button>
                )}
                {savedSnapshotAvailable && (
                  <button
                    type="button"
                    className="ghost-button large"
                    onClick={clearSavedDashboardData}
                  >
                    <ShieldAlert size={20} /> Clear Saved Dashboard Data
                  </button>
                )}
                {renderThemeToggle("theme-toggle--action theme-toggle--upload-action")}
                <span>or drop the workbook here</span>
              </div>
              <p className="upload-privacy-note">
                Uploaded files are processed in your browser. No files are sent to a server by this dashboard.
              </p>
              <div className="google-sheet-loader">
                <label>Online sheet region links</label>
                <div className="google-region-grid">
                  {GOOGLE_REGION_LINKS.map((region) => (
                    <label key={region.key} className="google-region-row">
                      <input
                        type="checkbox"
                        checked={googleRegionSelection[region.key]}
                        onChange={(event) =>
                          setGoogleRegionChecked(
                            region.key,
                            event.target.checked,
                          )
                        }
                      />
                      <span>{region.label}</span>
                      <input
                        type="url"
                        value={googleRegionLinks[region.key]}
                        onChange={(event) =>
                          setGoogleRegionLink(region.key, event.target.value)
                        }
                        placeholder="Google TT-History URL or Microsoft 365 Excel link"
                      />
                    </label>
                  ))}
                </div>
                <button
                  type="button"
                  className="ghost-button large google-load-selected"
                  onClick={() => void handleGoogleSheetLoad("replace")}
                  disabled={
                    googleSheetLoading ||
                    !GOOGLE_REGION_LINKS.some(
                      (region) =>
                        googleRegionSelection[region.key] &&
                        googleRegionLinks[region.key].trim(),
                    )
                  }
                >
                  <LinkIcon size={20} />
                  {googleSheetLoading ? "Loading..." : "Load selected"}
                </button>
              </div>
            </div>
            <div className="upload-visual">
              <img src="/h.png" alt="Dashboard workbook preview" />
            </div>
          </div>
        </section>
      ) : (
        <>
          <section
            ref={statsRef}
            id="section-kpis"
            className="stats-grid workbook-cards dashboard-section-content-block"
          >
            {/* Standard StatCards */}
            <StatCard
              label="Total TT's"
              value={analytics.totalUnique.toLocaleString()}
              // note={<span>TT's Opened</span>}
              icon={Layers3}
              tone="#fff200"
              onClick={() => {
                setFilters(EMPTY_FILTERS);
                setTablePage(1);
              }}
            />

            <StatCard
              label="Closed TT's"
              value={closed.toLocaleString()}
              // note={<span>Closed TT's</span>}
              icon={CheckCircle2}
              tone="#34d399"
              onClick={() => {
                setFilters({ ...EMPTY_FILTERS, status: ["Closed"] });
                setTablePage(1);
              }}
            />

            <StatCard
              label="Pending TT's"
              value={pending.toLocaleString()}
              // note={<span>Pending TT's</span>}
              icon={ShieldAlert}
              tone="#f59e0b"
              onClick={() => {
                setFilters({ ...EMPTY_FILTERS, status: ["Pending"] });
                setTablePage(1);
              }}
            />

            <StatCard
              label="Resolved TT's"
              value={resolved.toLocaleString()}
              // note={<span>Resolved TT's</span>}
              icon={CheckCircle2}
              tone="#60a5fa"
              onClick={() => {
                setFilters({ ...EMPTY_FILTERS, status: ["Resolved"] });
                setTablePage(1);
              }}
            />

            <StatCard
              label="Critical TT's"
              value={critical.toLocaleString()}
              // note={<span>Critical TT's</span>}
              icon={AlertTriangle}
              tone="#ef4444"
              onClick={() => {
                setFilters({ ...EMPTY_FILTERS, severity: ["Critical"] });
                setTablePage(1);
              }}
            />

            <StatCard
              label="Major TT's"
              value={major.toLocaleString()}
              // note={<span>Major TT's</span>}
              icon={Activity}
              tone="#f59e0b"
              onClick={() => {
                setFilters({ ...EMPTY_FILTERS, severity: ["Major"] });
                setTablePage(1);
              }}
            />

            <StatCard
              label="Minor TT's"
              value={
                minor !== undefined && minor !== null
                  ? minor.toLocaleString()
                  : "0"
              }
              // note={<span>Minor TT's</span>}
              icon={CircleDot}
              tone="#22d3ee"
              onClick={() => {
                setFilters({ ...EMPTY_FILTERS, severity: ["Minor"] });
                setTablePage(1);
              }}
            />

            <StatCard
              label="Region Sites"
              value={analytics.regionSiteTotal.toLocaleString()}
              // note={<span>{`${analytics.uniqueSites.toLocaleString()} Unique RF Sites`}</span>}
              icon={BarChart3}
              tone="#60a5fa"
            />

            <StatCard
              label="Regions"
              value={analytics.region.length.toLocaleString()}
              // note={<span>Total Regions</span>}
              icon={CircleDot}
              tone="#60a5fa"
            />

            <StatCard
              label="Non-Service Impact"
              value={nonServiceImpact.toLocaleString()}
              // note={<span>No Service Impact</span>}
              icon={CloudOff}
              tone="#94a3b8"
              onClick={() => {
                setFilters({
                  ...EMPTY_FILTERS,
                  impact: ["Non-Service Impact"],
                });
                setTablePage(1);
              }}
            />

            <StatCard
              label="Service Impact"
              value={serviceImpact.toLocaleString()}
              // note={<span>Exact Service Impact</span>}
              icon={Network}
              tone="#ef4444"
              onClick={() => {
                setFilters({ ...EMPTY_FILTERS, impact: ["Service Impact"] });
                setTablePage(1);
              }}
            />

            {/* Top RCA by Tickets Count â€” direct grid child so it stretches to the same height as siblings */}
            <StatCard
              label="Top RCA / TT's"
              value={analytics.topRcaByCount.name || "N/A"}
              note={
                <span>
                  {analytics.topRcaByCount.value
                    ? `${analytics.topRcaByCount.value.toLocaleString()} Tickets`
                    : ""}
                </span>
              }
              icon={BarChart3}
              tone="#22d3ee"
              className="rca-inline-fix"
            />

            {/* Top RCA by Downtime */}
            <StatCard
              label="Top RCA / Downtime"
              value={analytics.topRcaByDowntime.name || "N/A"}
              note={
                <span>{formatHours(analytics.topRcaByDowntime.value)}</span>
              }
              icon={Activity}
              tone="#f59e0b"
            />

            {/* Highest MTTR RCA */}
            <StatCard
              label="Highest MTTR RCA"
              value={analytics.highestMttrRca.name || "N/A"}
              note={<span>{formatHours(analytics.highestMttrRca.value)}</span>}
              icon={AlertTriangle}
              tone="#ef4444"
            />

            <StatCard
              label="Repeated RCA/Sites"
              value={analytics.repeatedRcaSites.toLocaleString()}
              note={<span />}
              icon={Network}
              tone="#a78bfa"
            />

            <StatCard
              label="RCA not Provided %"
              value={analytics.rcaNotProvidedCount.toLocaleString()}
              note={<span />}
              icon={ShieldAlert}
              tone="#ff0000"
            />
          </section>

          {/* Performance gauges -- live with dashboard filters */}
          {performanceKpiRows.length > 0 &&
            (() => {
              const kpi = computePerfKPIs(performanceKpiRows);
              return (
                <div
                  ref={performanceKpiCardsRef}
                  id="section-performance-kpis"
                  className="hero-export-row no-print dashboard-section-content-block"
                  style={{
                    paddingBottom: 0,
                    paddingTop: 0,
                    marginTop: "4px",
                  }}
                >
                  <div className="perf-kpi-row perf-kpi-row--gauges">
                    {(() => {
                      const gaugeCards = buildPerformanceGaugeCards(
                        kpi,
                        performanceKpiRows,
                      );

                      return gaugeCards.map((card, index) => (
                        <PerformanceGaugeCard
                          key={card.id}
                          {...card}
                          index={index}
                        />
                      ));
                    })()}
                  </div>
                </div>
              );
            })()}

          <section
            id="section-executive"
            className="glass-card executive-insights-section dashboard-section-content-block"
          >
            <div className="card-heading executive-insights-header">
              <div className="executive-insights-copy">
                <span className="section-kicker">
                  <Activity size={14} /> Executive Insights
                </span>
                <h3>Network Health & Risk Summary</h3>
                <p>{executiveInsights.summaryText}</p>
              </div>

              <div
                className={`executive-health-score executive-health-score--${executiveInsights.healthScore.status.toLowerCase()}`}
              >
                <span>Network Health Score</span>
                <strong>{executiveInsights.healthScore.score}</strong>
                <em>{executiveInsights.healthScore.status}</em>
                <small>{executiveInsights.healthScore.mainReason}</small>
              </div>
            </div>

            <div className="executive-insight-grid">
              {executiveInsights.cards.map((card) => (
                <div
                  key={card.label}
                  className="executive-insight-card"
                  style={{ ["--insight-tone" as string]: card.tone }}
                >
                  <span>{card.label}</span>
                  <strong>{card.value}</strong>
                  <small>{card.note}</small>
                </div>
              ))}
            </div>

            <section
              id="section-deep-dive"
              className="glass-card deep-dive-section client-delivery-section dashboard-section-content-block"
            >
              <div className="card-heading deep-dive-header">
                <div>
                  <span className="section-kicker">
                    <ShieldAlert size={14} /> RCA / Preventability / SLA
                    Deep-Dive
                  </span>
                  <h3>Operational Quality & Follow-Up Priorities</h3>
                  <p>
                    This layer converts RCA quality, preventability, SLA
                    response, and repeated-site patterns into management
                    follow-up actions.
                  </p>
                </div>
              </div>

              <div className="deep-dive-kpi-grid">
                <div className="deep-dive-kpi-card">
                  <span>Avg FRT</span>
                  <strong>
                    {deepDiveAnalytics.slaSummary.avgFrtHours.toFixed(1)}h
                  </strong>
                  <small>
                    {deepDiveAnalytics.slaSummary.frtBreaches.toLocaleString()}{" "}
                    above 1h target
                  </small>
                </div>
                <div className="deep-dive-kpi-card">
                  <span>Avg Response</span>
                  <strong>
                    {deepDiveAnalytics.slaSummary.avgResponseHours.toFixed(1)}h
                  </strong>
                  <small>
                    {deepDiveAnalytics.slaSummary.responseBreaches.toLocaleString()}{" "}
                    above 4h target
                  </small>
                </div>
                <div className="deep-dive-kpi-card">
                  <span>Avg Resolution</span>
                  <strong>
                    {deepDiveAnalytics.slaSummary.avgResolutionHours.toFixed(1)}
                    h
                  </strong>
                  <small>
                    {deepDiveAnalytics.slaSummary.resolutionBreaches.toLocaleString()}{" "}
                    above 24h target
                  </small>
                </div>
                <div className="deep-dive-kpi-card">
                  <span>Repeated Sites</span>
                  <strong>
                    {deepDiveAnalytics.repeatedOffenderSites.length.toLocaleString()}
                  </strong>
                  <small>Top sites requiring technical follow-up</small>
                </div>
              </div>

              <div className="deep-dive-chart-grid">
                <article className="deep-dive-panel">
                  <div className="deep-dive-panel-heading">
                    <strong>Preventability by Tickets</strong>
                  </div>
                  <ResponsiveContainer width="100%" height={230}>
                    <PieChart>
                      <Pie
                        data={deepDiveAnalytics.preventabilityByCount}
                        dataKey="value"
                        nameKey="name"
                        innerRadius={52}
                        outerRadius={82}
                        paddingAngle={4}
                        label={({ percentage }) => `${percentage ?? 0}%`}
                      >
                        {deepDiveAnalytics.preventabilityByCount.map(
                          (entry, index) => (
                            <Cell
                              key={entry.name}
                              fill={COLORS[index % COLORS.length]}
                            />
                          ),
                        )}
                      </Pie>
                      <Tooltip contentStyle={CHART_TOOLTIP_STYLE} />
                      <Legend wrapperStyle={CHART_LEGEND_STYLE} />
                    </PieChart>
                  </ResponsiveContainer>
                </article>

                <article className="deep-dive-panel">
                  <div className="deep-dive-panel-heading">
                    <strong>Pending Aging Buckets</strong>
                    <small style={{ color: "#8ea4c2", fontSize: 10, marginLeft: 8 }}>
                      Total → breakdown by age
                    </small>
                  </div>
                  <PendingAgingWaterfall
                    buckets={deepDiveAnalytics.slaSummary.pendingAgingBuckets}
                  />
                  {/* Colour legend */}
                  <div style={{
                    display: "flex",
                    gap: 12,
                    flexWrap: "wrap",
                    marginTop: 6,
                    paddingLeft: 4,
                    fontSize: 10,
                    color: "#8ea4c2",
                  }}>
                    <span style={{ display: "flex", alignItems: "center", gap: 4 }}>
                      <span style={{ width: 10, height: 10, borderRadius: 2, background: "#22d3ee", display: "inline-block" }} />
                      Total Pending
                    </span>
                    <span style={{ display: "flex", alignItems: "center", gap: 4 }}>
                      <span style={{ width: 10, height: 10, borderRadius: 2, background: "#38bdf8", display: "inline-block" }} />
                      0–24h
                    </span>
                    <span style={{ display: "flex", alignItems: "center", gap: 4 }}>
                      <span style={{ width: 10, height: 10, borderRadius: 2, background: "#60a5fa", display: "inline-block" }} />
                      1–3d
                    </span>
                    <span style={{ display: "flex", alignItems: "center", gap: 4 }}>
                      <span style={{ width: 10, height: 10, borderRadius: 2, background: "#f59e0b", display: "inline-block" }} />
                      3–7d
                    </span>
                    <span style={{ display: "flex", alignItems: "center", gap: 4 }}>
                      <span style={{ width: 10, height: 10, borderRadius: 2, background: "#ef4444", display: "inline-block" }} />
                      7d+ (Escalation)
                    </span>
                  </div>
                </article>

                <article className="deep-dive-panel deep-dive-panel--wide">
                  <div className="deep-dive-panel-heading">
                    <strong>Top RCA Families by Downtime</strong>
                  </div>
                  <ResponsiveContainer width="100%" height={280}>
                    <BarChart
                      data={deepDiveAnalytics.rcaFamilyDeepDive.slice(0, 8)}
                      layout="vertical"
                      margin={{ left: 18, right: 48, top: 8, bottom: 8 }}
                    >
                      <CartesianGrid
                        stroke={CHART_GRID_STROKE}
                        horizontal={false}
                      />
                      <XAxis
                        type="number"
                        stroke={CHART_AXIS_STROKE}
                        allowDecimals={false}
                        tickLine={false}
                        axisLine={false}
                      />
                      <YAxis
                        dataKey="family"
                        type="category"
                        stroke={CHART_AXIS_STROKE}
                        width={190}
                        tickLine={false}
                        axisLine={false}
                        tick={{ fontSize: 11 }}
                      />
                      <Tooltip contentStyle={CHART_TOOLTIP_STYLE} />
                      <Bar
                        dataKey="downtimeHours"
                        name="Downtime hrs"
                        radius={BAR_RADIUS}
                        fill="#22d3ee"
                      >
                        <LabelList
                          dataKey="downtimeHours"
                          position="right"
                          fill={CHART_LABEL_FILL}
                          fontSize={11}
                        />
                      </Bar>
                    </BarChart>
                  </ResponsiveContainer>
                </article>
              </div>

              <div className="deep-dive-bottom-grid deep-dive-bottom-grid--actions-only">
                <div className="deep-dive-actions-card">
                  <div className="deep-dive-panel-heading">
                    <strong>Recommended Management Actions</strong>
                    <small>Auto-generated from current filtered scope</small>
                  </div>
                  <ol>
                    {deepDiveAnalytics.recommendations.map((item) => (
                      <li key={item}>{item}</li>
                    ))}
                  </ol>
                </div>
              </div>

              <div className="executive-risk-table-card">
                <div className="executive-risk-table-header">
                  <div>
                    <strong>Repeated Offender Sites</strong>
                    <p>
                      Sites prioritized by repeated TT count, downtime, and top
                      RCA.
                    </p>
                  </div>
                  <div
                    style={{ display: "flex", alignItems: "center", gap: 10 }}
                  >
                    <span>Top {executiveRcaSlaRows.length}</span>
                  </div>
                </div>

                <div className="table-scroll executive-risk-table-scroll">
                  <table className="data-table executive-risk-table executive-rca-sla-table">
                    <thead>
                      <tr>
                        <th>Rank</th>
                        <th>Region</th>
                        <th>Site ID</th>
                        <th>Site Name</th>
                        <th>TT's</th>
                        <th>Availability</th>
                        <th>Downtime</th>
                        <th>Reliability</th>
                        <th>Top RCA</th>
                      </tr>
                    </thead>
                    <tbody>
                      {executiveRcaSlaRows.length ? (
                        executiveRcaSlaRows.map((site, index) => (
                          <tr key={`${site.region}-${site.siteId}`}>
                            <td>{index + 1}</td>
                            <td>{site.region}</td>
                            <td>{site.siteId}</td>
                            <td>{site.siteName || "-"}</td>
                            <td>{site.tickets.toLocaleString()}</td>
                            <td>
                              {site.performanceAvailabilityHours === null
                                ? "-"
                                : `${site.performanceAvailabilityHours.toLocaleString()} hrs`}
                            </td>
                            <td>
                              {site.performanceDowntimeHours === null
                                ? "-"
                                : `${site.performanceDowntimeHours.toLocaleString()} hrs`}
                            </td>
                            <td>
                              {site.reliability === null
                                ? "-"
                                : `${site.reliability.toFixed(2)}%`}
                            </td>
                            <td>{site.topRca}</td>
                          </tr>
                        ))
                      ) : (
                        <tr>
                          <td colSpan={9} className="empty-table-cell">
                            No repeated offender sites found for the selected
                            filters.
                          </td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </div>
              </div>
            </section>
          </section>

          {/* â•â• Charts â•â• */}
          <div
            id="section-overview-charts"
            className="chart-2col dashboard-chart-grid dashboard-section-content-block"
          >
            {/* â”€â”€ Column 1 â€” five charts, one per row â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€ */}
            <article
              className="glass-card"
              style={{ gridColumn: "1 / 7", gridRow: 2 }}
            >
              <div
                className="card-heading"
                style={{
                  display: "flex",
                  justifyContent: "center",
                  alignItems: "center",
                  gap: "10px",
                }}
              >
                {" "}
                <h3>
                  {analytics.trendGrain === "month"
                    ? "Tickets per month"
                    : "Tickets per week"}
                </h3>{" "}
                <BarChart3 size={18} />
              </div>
              <ResponsiveContainer width="100%" height={290}>
                <LineChart
                  data={analytics.monthly}
                  margin={{ left: 0, right: 22, top: 16, bottom: 8 }}
                >
                  <defs>
                    <linearGradient id="openedLine" x1="0" y1="0" x2="1" y2="0">
                      <stop offset="0%" stopColor="#38bdf8" />
                      <stop offset="100%" stopColor="#2dd4bf" />
                    </linearGradient>
                  </defs>
                  <CartesianGrid stroke={CHART_GRID_STROKE} vertical={false} />
                  <XAxis
                    dataKey="name"
                    stroke={CHART_AXIS_STROKE}
                    tickLine={false}
                    axisLine={false}
                    tick={{ fontSize: 11, fill: "#aac0dc" }}
                  />
                  <YAxis
                    stroke={CHART_AXIS_STROKE}
                    tickLine={false}
                    axisLine={false}
                    allowDecimals={false}
                    tick={{ fontSize: 11, fill: "#aac0dc" }}
                  />
                  <Tooltip
                    contentStyle={CHART_TOOLTIP_STYLE}
                    cursor={{
                      stroke: "rgba(125, 211, 252, .28)",
                      strokeWidth: 1,
                    }}
                  />
                  <Legend wrapperStyle={CHART_LEGEND_STYLE} iconType="circle" />
                  <Line
                    type="monotone"
                    name="Opened"
                    dataKey="opened"
                    stroke="url(#openedLine)"
                    strokeWidth={3.5}
                    dot={{
                      r: 3,
                      fill: "#071426",
                      stroke: "#67e8f9",
                      strokeWidth: 2,
                    }}
                    activeDot={{
                      r: 6,
                      fill: "#22d3ee",
                      stroke: "#ecfeff",
                      strokeWidth: 2,
                    }}
                  >
                    <LabelList
                      dataKey="opened"
                      position="top"
                      fill={CHART_LABEL_FILL}
                      fontSize={11}
                    />
                  </Line>
                </LineChart>
              </ResponsiveContainer>
            </article>

            <article
              className="glass-card"
              style={{ gridColumn: "1 / 7", gridRow: 3 }}
            >
              <div
                className="card-heading"
                style={{
                  display: "flex",
                  justifyContent: "center",
                  alignItems: "center",
                  gap: "10px",
                }}
              >
                {" "}
                <h3>Top sites by unique Tickets</h3> <BarChart3 size={18} />
              </div>
              <ResponsiveContainer width="100%" height={290}>
                <BarChart
                  data={analytics.topSites.slice(0, 10)}
                  layout="vertical"
                  margin={{ left: 18, right: 44, top: 8, bottom: 8 }}
                >
                  <defs>
                    <linearGradient
                      id="topSitesBar"
                      x1="0"
                      y1="0"
                      x2="1"
                      y2="0"
                    >
                      <stop offset="0%" stopColor="#0ea5e9" />
                      <stop offset="100%" stopColor="#67e8f9" />
                    </linearGradient>
                  </defs>
                  <CartesianGrid
                    stroke={CHART_GRID_STROKE}
                    horizontal={false}
                  />
                  <XAxis
                    type="number"
                    stroke={CHART_AXIS_STROKE}
                    allowDecimals={false}
                    tick={{ fontSize: 11, fill: "#aac0dc" }}
                  />
                  <YAxis
                    dataKey="name"
                    type="category"
                    stroke={CHART_AXIS_STROKE}
                    width={210}
                    tickLine={false}
                    axisLine={false}
                    tick={{ fontSize: 11, fill: "#d8e6f7" }}
                    interval={0}
                  />
                  <Tooltip
                    contentStyle={CHART_TOOLTIP_STYLE}
                    cursor={{ fill: "rgba(125, 211, 252, .06)" }}
                  />
                  <Bar
                    dataKey="value"
                    radius={BAR_RADIUS}
                    fill="url(#topSitesBar)"
                    maxBarSize={18}
                  >
                    <LabelList
                      dataKey="value"
                      position="right"
                      fill={CHART_LABEL_FILL}
                      fontSize={12}
                    />
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </article>

            <article className="glass-card" style={{ display: "none" }}>
              <div
                className="card-heading"
                style={{
                  display: "flex",
                  justifyContent: "center",
                  alignItems: "center",
                  gap: "10px",
                }}
              >
                {" "}
                <h3>Top 10 RCA by unique Tickets count</h3>{" "}
                <BarChart3 size={18} />
              </div>
              <ResponsiveContainer width="100%" height={290}>
                <BarChart
                  data={analytics.rcaByCount.slice(0, 10)}
                  layout="vertical"
                  margin={{ left: 18, right: 56, top: 8, bottom: 8 }}
                >
                  <defs>
                    <linearGradient
                      id="rcaCountBar"
                      x1="0"
                      y1="0"
                      x2="1"
                      y2="0"
                    >
                      <stop offset="0%" stopColor="#2563eb" />
                      <stop offset="100%" stopColor="#22d3ee" />
                    </linearGradient>
                  </defs>
                  <CartesianGrid
                    stroke={CHART_GRID_STROKE}
                    horizontal={false}
                  />
                  <XAxis
                    type="number"
                    stroke={CHART_AXIS_STROKE}
                    allowDecimals={false}
                    tick={{ fontSize: 11, fill: "#aac0dc" }}
                  />
                  <YAxis
                    dataKey="name"
                    type="category"
                    stroke={CHART_AXIS_STROKE}
                    width={200}
                    tickLine={false}
                    axisLine={false}
                    tick={{ fontSize: 12, fill: "#d8e6f7" }}
                    interval={0}
                  />
                  <Tooltip
                    contentStyle={CHART_TOOLTIP_STYLE}
                    cursor={{ fill: "rgba(125, 211, 252, .06)" }}
                  />
                  <Bar
                    dataKey="value"
                    radius={BAR_RADIUS}
                    fill="url(#rcaCountBar)"
                    maxBarSize={18}
                  >
                    <LabelList
                      dataKey="value"
                      position="right"
                      fill={CHART_LABEL_FILL}
                      fontSize={13}
                    />
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </article>

            <article className="glass-card" style={{ display: "none" }}>
              <div
                className="card-heading"
                style={{
                  display: "flex",
                  justifyContent: "center",
                  alignItems: "center",
                  gap: "10px",
                }}
              >
                {" "}
                <h3>Operational families - Tickets distribution</h3>{" "}
                <BarChart3 size={18} />
              </div>
              <div className="rca-family-layout">
                <ResponsiveContainer width="100%" height={290}>
                  <PieChart>
                    <Pie
                      data={analytics.rcaFamily}
                      dataKey="value"
                      nameKey="name"
                      innerRadius={68}
                      outerRadius={110}
                      paddingAngle={4}
                      labelLine={false}
                      label={renderPieLabel}
                    >
                      {analytics.rcaFamily.map((entry, index) => (
                        <Cell
                          key={entry.name}
                          fill={COLORS[index % COLORS.length]}
                        />
                      ))}
                    </Pie>
                    <Tooltip
                      contentStyle={{
                        background: "#071426",
                        border: "1px solid rgba(34,211,238,.25)",
                        borderRadius: 14,
                        color: "#e2e8f0",
                      }}
                    />
                  </PieChart>
                </ResponsiveContainer>
                <div className="rca-family-legend">
                  {analytics.rcaFamily.map((entry, index) => (
                    <div key={entry.name} className="rca-legend-row">
                      <span
                        className="rca-legend-dot"
                        style={{ background: COLORS[index % COLORS.length] }}
                      />
                      <span className="rca-legend-name">{entry.name}</span>
                      <strong className="rca-legend-val">
                        {entry.value}{" "}
                        <small>
                          ({pct(entry.value, analytics.totalUnique)})
                        </small>
                      </strong>
                    </div>
                  ))}
                </div>
              </div>
            </article>

            <article className="glass-card full" style={{ display: "none" }}>
              {/* Row 5 of column 1 is split 50/50 horizontally: Top Managed Resources | Average Ticket Resolution Time. */}
              <div
                style={{
                  display: "grid",
                  gridTemplateColumns: "1fr 1fr",
                  gap: "16px",
                  height: "100%",
                }}
              >
                <article className="glass-card" style={{ margin: 0 }}>
                  <div
                    className="card-heading"
                    style={{
                      display: "flex",
                      justifyContent: "center",
                      alignItems: "center",
                      gap: "10px",
                    }}
                  >
                    {" "}
                    <h3>Top Managed Resources by Ticket Count</h3>{" "}
                    <BarChart3 size={18} />
                  </div>
                  <ResponsiveContainer width="100%" height={290}>
                    <BarChart
                      data={analytics.topManagedResources}
                      layout="vertical"
                      margin={{ left: 18, right: 72, top: 8, bottom: 8 }}
                    >
                      <CartesianGrid
                        stroke="rgba(148,163,184,.12)"
                        horizontal={false}
                      />
                      <XAxis
                        type="number"
                        stroke="#94a3b8"
                        allowDecimals={false}
                      />
                      <YAxis
                        dataKey="name"
                        type="category"
                        stroke="#cbd5e1"
                        width={220}
                        tickLine={false}
                        axisLine={false}
                        tick={{ fontSize: 12 }}
                        interval={0}
                      />
                      <Tooltip
                        contentStyle={{
                          background: "#071426",
                          border: "1px solid rgba(34,211,238,.25)",
                          borderRadius: 14,
                          color: "#e2e8f0",
                        }}
                        formatter={(v: number) => [
                          v.toLocaleString(),
                          "Tickets",
                        ]}
                      />
                      <Bar
                        dataKey="value"
                        radius={[0, 10, 10, 0]}
                        fill="#22d3ee"
                      >
                        <LabelList
                          dataKey="value"
                          position="right"
                          fill="#e2e8f0"
                          fontSize={12}
                        />
                      </Bar>
                    </BarChart>
                  </ResponsiveContainer>
                </article>

                <article
                  className="glass-card"
                  style={{
                    margin: 0,
                    display: "flex",
                    flexDirection: "column",
                    minHeight: 0,
                  }}
                >
                  <div
                    className="card-heading"
                    style={{
                      display: "flex",
                      justifyContent: "center",
                      alignItems: "center",
                      gap: "10px",
                    }}
                  >
                    {" "}
                    <h3>Average Ticket Resolution Time</h3>{" "}
                    <Activity size={18} />
                  </div>
                  <div style={{ padding: "0 8px 8px", flex: 1, minHeight: 0 }}>
                    {(analytics.monthlyResolutionTime ?? []).length === 0 ? (
                      <div
                        style={{
                          color: "#94a3b8",
                          fontSize: "12px",
                          textAlign: "center",
                          padding: "12px",
                        }}
                      >
                        No closed / resolved tickets with a duration in the
                        current selection.
                      </div>
                    ) : (
                      <ResponsiveContainer width="100%" height={280}>
                        <LineChart
                          data={analytics.monthlyResolutionTime}
                          margin={{ top: 4, right: 18, left: 0, bottom: 36 }}
                        >
                          <CartesianGrid
                            strokeDasharray="3 3"
                            stroke="#1e293b"
                            vertical={false}
                          />
                          <XAxis
                            dataKey="name"
                            tick={{ fill: "#cbd5e1", fontSize: 10 }}
                            angle={-40}
                            textAnchor="end"
                            interval={0}
                            height={40}
                          />
                          <YAxis
                            tick={{ fill: "#cbd5e1", fontSize: 10 }}
                            tickFormatter={(v: number) => `${v}h`}
                            width={42}
                          />
                          <Tooltip
                            contentStyle={{
                              background: "#0f172a",
                              border: "1px solid #334155",
                              borderRadius: 8,
                              fontSize: 12,
                            }}
                            labelStyle={{ color: "#e2e8f0", fontWeight: 700 }}
                            formatter={(
                              value: number,
                              name: string,
                              props: {
                                payload?: {
                                  count?: number;
                                  totalHours?: number;
                                  pendingCount?: number;
                                  pendingTotal?: number;
                                };
                              },
                            ) => {
                              const p = props.payload ?? {};
                              if (name === "Resolution (avg h)") {
                                return [
                                  `${value.toFixed(2)} h (n=${p.count ?? 0}, total ${p.totalHours ?? 0} h)`,
                                  name,
                                ];
                              }
                              if (name === "Pending (avg h)") {
                                return value > 0
                                  ? [
                                      `${value.toFixed(2)} h (n=${p.pendingCount ?? 0}, total ${p.pendingTotal ?? 0} h)`,
                                      name,
                                    ]
                                  : ["-", name];
                              }
                              return [value, name];
                            }}
                          />
                          <Legend
                            verticalAlign="top"
                            height={22}
                            wrapperStyle={{
                              fontSize: "10px",
                              color: "#cbd5e1",
                              paddingBottom: 2,
                            }}
                            iconType="circle"
                            iconSize={8}
                          />
                          <Line
                            type="monotone"
                            dataKey="avgHours"
                            name="Resolution (avg h)"
                            stroke="#93c5fd"
                            strokeWidth={2.5}
                            dot={{ r: 4 }}
                            activeDot={{ r: 6 }}
                            isAnimationActive={false}
                          />
                          <Line
                            type="monotone"
                            dataKey="pendingAvg"
                            name="Pending (avg h)"
                            stroke="#f59e0b"
                            strokeWidth={2.5}
                            dot={{ r: 4 }}
                            activeDot={{ r: 6 }}
                            isAnimationActive={false}
                          />
                        </LineChart>
                      </ResponsiveContainer>
                    )}
                  </div>
                </article>
              </div>
            </article>

            {/* â”€â”€ Column 2 â€” five charts stacked vertically, spans column 1 rows 1-4 â”€â”€ */}
            <div
              style={{
                display: "contents",
              }}
            >
              <article
                className="glass-card"
                style={{ gridColumn: "7 / 10", gridRow: 2 }}
              >
                <div
                  className="card-heading"
                  style={{
                    display: "flex",
                    justifyContent: "center",
                    alignItems: "center",
                    gap: "10px",
                  }}
                >
                  {" "}
                  <h3>Status</h3> <BarChart3 size={18} />
                </div>
                <ResponsiveContainer width="100%" height={290}>
                  <PieChart>
                    <Pie
                      data={analytics.status}
                      dataKey="value"
                      nameKey="name"
                      innerRadius={48}
                      outerRadius={76}
                      paddingAngle={4}
                      cornerRadius={7}
                      labelLine={false}
                      label={renderPieLabel}
                    >
                      {analytics.status.map((entry, index) => (
                        <Cell
                          key={entry.name}
                          fill={
                            STATUS_COLORS[entry.name] ??
                            COLORS[index % COLORS.length]
                          }
                        />
                      ))}
                    </Pie>
                    <Tooltip contentStyle={CHART_TOOLTIP_STYLE} />
                    <Legend
                      wrapperStyle={CHART_LEGEND_STYLE}
                      iconType="circle"
                      iconSize={8}
                    />
                  </PieChart>
                </ResponsiveContainer>
              </article>

              <article
                className="glass-card"
                style={{ gridColumn: "10 / 13", gridRow: 2 }}
              >
                <div
                  className="card-heading"
                  style={{
                    display: "flex",
                    justifyContent: "center",
                    alignItems: "center",
                    gap: "10px",
                  }}
                >
                  {" "}
                  <h3>Severity</h3> <BarChart3 size={18} />
                </div>
                <ResponsiveContainer width="100%" height={290}>
                  <PieChart>
                    <Pie
                      data={analytics.severity}
                      dataKey="value"
                      nameKey="name"
                      innerRadius={48}
                      outerRadius={76}
                      paddingAngle={4}
                      cornerRadius={7}
                      labelLine={false}
                      label={renderPieLabel}
                    >
                      {analytics.severity.map((entry, index) => (
                        <Cell
                          key={entry.name}
                          fill={
                            SEVERITY_COLORS[entry.name] ??
                            COLORS[index % COLORS.length]
                          }
                        />
                      ))}
                    </Pie>
                    <Tooltip contentStyle={CHART_TOOLTIP_STYLE} />
                    <Legend
                      wrapperStyle={CHART_LEGEND_STYLE}
                      iconType="circle"
                      iconSize={8}
                    />
                  </PieChart>
                </ResponsiveContainer>
              </article>

              <article
                className="glass-card"
                style={{ gridColumn: "7 / 10", gridRow: 3 }}
              >
                <div
                  className="card-heading"
                  style={{
                    display: "flex",
                    justifyContent: "center",
                    alignItems: "center",
                    gap: "10px",
                  }}
                >
                  {" "}
                  <h3>Region</h3> <BarChart3 size={18} />
                </div>
                <ResponsiveContainer width="100%" height={290}>
                  <PieChart>
                    <Pie
                      data={analytics.region}
                      dataKey="value"
                      nameKey="name"
                      innerRadius={48}
                      outerRadius={76}
                      paddingAngle={4}
                      cornerRadius={7}
                      labelLine={false}
                      label={renderPieLabel}
                    >
                      {analytics.region.map((entry, index) => (
                        <Cell
                          key={entry.name}
                          fill={COLORS[index % COLORS.length]}
                        />
                      ))}
                    </Pie>
                    <Tooltip contentStyle={CHART_TOOLTIP_STYLE} />
                    <Legend
                      wrapperStyle={CHART_LEGEND_STYLE}
                      iconType="circle"
                      iconSize={8}
                    />
                  </PieChart>
                </ResponsiveContainer>
              </article>

              <article
                className="glass-card"
                style={{ gridColumn: "10 / 13", gridRow: 3 }}
              >
                <div
                  className="card-heading"
                  style={{
                    display: "flex",
                    justifyContent: "center",
                    alignItems: "center",
                    gap: "10px",
                  }}
                >
                  {" "}
                  <h3>Escalation level</h3> <BarChart3 size={18} />
                </div>
                <ResponsiveContainer width="100%" height={290}>
                  <PieChart>
                    <Pie
                      data={analytics.escalation}
                      dataKey="value"
                      nameKey="name"
                      innerRadius={48}
                      outerRadius={76}
                      paddingAngle={4}
                      cornerRadius={7}
                      labelLine={false}
                      label={renderPieLabel}
                    >
                      {analytics.escalation.map((entry, index) => (
                        <Cell
                          key={entry.name}
                          fill={COLORS[index % COLORS.length]}
                        />
                      ))}
                    </Pie>
                    <Tooltip contentStyle={CHART_TOOLTIP_STYLE} />
                    <Legend
                      wrapperStyle={CHART_LEGEND_STYLE}
                      iconType="circle"
                      iconSize={8}
                    />
                  </PieChart>
                </ResponsiveContainer>
              </article>

              <article className="glass-card" style={{ display: "none" }}>
                <div
                  className="card-heading"
                  style={{
                    display: "flex",
                    justifyContent: "center",
                    alignItems: "center",
                    gap: "10px",
                  }}
                >
                  {" "}
                  <h3>Service impact</h3> <BarChart3 size={18} />
                </div>
                <ResponsiveContainer width="100%" height={290}>
                  <BarChart
                    data={analytics.impact}
                    margin={{ left: 0, right: 24, top: 18, bottom: 42 }}
                  >
                    <CartesianGrid
                      stroke="rgba(148,163,184,.12)"
                      vertical={false}
                    />
                    <XAxis
                      dataKey="name"
                      stroke="#cbd5e1"
                      tickLine={false}
                      axisLine={false}
                      interval={0}
                      angle={-18}
                      textAnchor="end"
                      height={58}
                      tick={{ fontSize: 11 }}
                    />
                    <YAxis
                      stroke="#94a3b8"
                      tickLine={false}
                      axisLine={false}
                      allowDecimals={false}
                    />
                    <Tooltip
                      contentStyle={{
                        background: "#071426",
                        border: "1px solid rgba(34,211,238,.25)",
                        borderRadius: 14,
                        color: "#e2e8f0",
                      }}
                    />
                    <Bar dataKey="value" radius={[10, 10, 0, 0]} fill="#a78bfa">
                      <LabelList
                        dataKey="value"
                        position="top"
                        fill="#e2e8f0"
                        fontSize={12}
                      />
                      {analytics.impact.map((entry, index) => (
                        <Cell
                          key={entry.name}
                          fill={COLORS[index % COLORS.length]}
                        />
                      ))}
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </article>
            </div>
          </div>
          {/* /chart-2col */}
          {/* â•â• Full-width + bottom-row charts â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â• */}
          <section
            id="section-trend-charts"
            className="chart-mosaic dashboard-section-content-block"
          >
            <article className="glass-card" style={{ gridColumn: "1 / 5" }}>
              <div
                className="card-heading"
                style={{
                  display: "flex",
                  justifyContent: "center",
                  alignItems: "center",
                  gap: "10px",
                }}
              >
                {" "}
                <h3>Top 10 RCA by unique Tickets count</h3>{" "}
                <BarChart3 size={18} />
              </div>
              <ResponsiveContainer width="100%" height={290}>
                <BarChart
                  data={analytics.rcaByCount.slice(0, 10)}
                  layout="vertical"
                  margin={{ left: 18, right: 56, top: 8, bottom: 8 }}
                >
                  <CartesianGrid
                    stroke="rgba(148,163,184,.12)"
                    horizontal={false}
                  />
                  <XAxis type="number" stroke="#94a3b8" allowDecimals={false} />
                  <YAxis
                    dataKey="name"
                    type="category"
                    stroke="#cbd5e1"
                    width={200}
                    tickLine={false}
                    axisLine={false}
                    tick={{ fontSize: 12 }}
                    interval={0}
                  />
                  <Tooltip
                    contentStyle={{
                      background: "#071426",
                      border: "1px solid rgba(34,211,238,.25)",
                      borderRadius: 14,
                      color: "#e2e8f0",
                    }}
                  />
                  <Bar dataKey="value" radius={[0, 10, 10, 0]} fill="#22d3ee">
                    <LabelList
                      dataKey="value"
                      position="right"
                      fill="#e2e8f0"
                      fontSize={13}
                    />
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </article>

            <article className="glass-card" style={{ gridColumn: "5 / 9" }}>
              <div
                className="card-heading"
                style={{
                  display: "flex",
                  justifyContent: "center",
                  alignItems: "center",
                  gap: "10px",
                }}
              >
                {" "}
                <h3>Top 10 RCA by total downtime (hrs)</h3>{" "}
                <BarChart3 size={18} />
              </div>
              <ResponsiveContainer width="100%" height={290}>
                <BarChart
                  data={analytics.rcaByDowntime.slice(0, 10)}
                  layout="vertical"
                  margin={{ left: 18, right: 72, top: 8, bottom: 8 }}
                >
                  <defs>
                    <linearGradient
                      id="rcaDowntimeBar"
                      x1="0"
                      y1="0"
                      x2="1"
                      y2="0"
                    >
                      <stop offset="0%" stopColor="#f97316" />
                      <stop offset="100%" stopColor="#facc15" />
                    </linearGradient>
                  </defs>
                  <CartesianGrid
                    stroke={CHART_GRID_STROKE}
                    horizontal={false}
                  />
                  <XAxis
                    type="number"
                    stroke={CHART_AXIS_STROKE}
                    allowDecimals={false}
                    tick={{ fontSize: 11, fill: "#aac0dc" }}
                  />
                  <YAxis
                    dataKey="name"
                    type="category"
                    stroke={CHART_AXIS_STROKE}
                    width={200}
                    tickLine={false}
                    axisLine={false}
                    tick={{ fontSize: 12, fill: "#d8e6f7" }}
                    interval={0}
                  />
                  <Tooltip
                    contentStyle={CHART_TOOLTIP_STYLE}
                    formatter={(v: number) => [
                      `${v.toLocaleString()} hrs`,
                      "Downtime",
                    ]}
                    cursor={{ fill: "rgba(250, 204, 21, .06)" }}
                  />
                  <Bar
                    dataKey="value"
                    radius={BAR_RADIUS}
                    fill="url(#rcaDowntimeBar)"
                    maxBarSize={18}
                  >
                    <LabelList
                      dataKey="value"
                      position="right"
                      fill={CHART_LABEL_FILL}
                      fontSize={12}
                      formatter={(v: number) => `${v.toLocaleString()}h`}
                    />
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </article>

            <article className="glass-card" style={{ gridColumn: "9 / 13" }}>
              <div
                className="card-heading"
                style={{
                  display: "flex",
                  justifyContent: "center",
                  alignItems: "center",
                  gap: "10px",
                }}
              >
                {" "}
                <h3>Highest MTTR by RCA</h3> <BarChart3 size={18} />
              </div>
              <ResponsiveContainer width="100%" height={290}>
                <BarChart
                  data={analytics.rcaByMttr.slice(0, 8)}
                  layout="vertical"
                  margin={{ left: 18, right: 72, top: 8, bottom: 8 }}
                >
                  <defs>
                    <linearGradient id="mttrBar" x1="0" y1="0" x2="1" y2="0">
                      <stop offset="0%" stopColor="#dc2626" />
                      <stop offset="100%" stopColor="#fb7185" />
                    </linearGradient>
                  </defs>
                  <CartesianGrid
                    stroke={CHART_GRID_STROKE}
                    horizontal={false}
                  />
                  <XAxis
                    type="number"
                    stroke={CHART_AXIS_STROKE}
                    allowDecimals={false}
                    tick={{ fontSize: 11, fill: "#aac0dc" }}
                  />
                  <YAxis
                    dataKey="name"
                    type="category"
                    stroke={CHART_AXIS_STROKE}
                    width={200}
                    tickLine={false}
                    axisLine={false}
                    tick={{ fontSize: 12, fill: "#d8e6f7" }}
                    interval={0}
                  />
                  <Tooltip
                    contentStyle={CHART_TOOLTIP_STYLE}
                    formatter={(v: number) => [
                      `${v.toLocaleString()} hrs avg`,
                      "MTTR",
                    ]}
                    cursor={{ fill: "rgba(248, 113, 113, .06)" }}
                  />
                  <Bar
                    dataKey="value"
                    radius={BAR_RADIUS}
                    fill="url(#mttrBar)"
                    maxBarSize={18}
                  >
                    <LabelList
                      dataKey="value"
                      position="right"
                      fill={CHART_LABEL_FILL}
                      fontSize={12}
                      formatter={(v: number) => `${v.toLocaleString()}h`}
                    />
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </article>

            <article className="glass-card" style={{ gridColumn: "1 / 7" }}>
              <div
                className="card-heading"
                style={{
                  display: "flex",
                  justifyContent: "center",
                  alignItems: "center",
                  gap: "10px",
                }}
              >
                {" "}
                <h3>Top Managed Resources by Ticket Count</h3>{" "}
                <BarChart3 size={18} />
              </div>
              <ResponsiveContainer width="100%" height={290}>
                <BarChart
                  data={analytics.topManagedResources}
                  layout="vertical"
                  margin={{ left: 18, right: 72, top: 8, bottom: 8 }}
                >
                  <defs>
                    <linearGradient
                      id="managedResourceBar"
                      x1="0"
                      y1="0"
                      x2="1"
                      y2="0"
                    >
                      <stop offset="0%" stopColor="#0891b2" />
                      <stop offset="100%" stopColor="#5eead4" />
                    </linearGradient>
                  </defs>
                  <CartesianGrid
                    stroke={CHART_GRID_STROKE}
                    horizontal={false}
                  />
                  <XAxis
                    type="number"
                    stroke={CHART_AXIS_STROKE}
                    allowDecimals={false}
                    tick={{ fontSize: 11, fill: "#aac0dc" }}
                  />
                  <YAxis
                    dataKey="name"
                    type="category"
                    stroke={CHART_AXIS_STROKE}
                    width={220}
                    tickLine={false}
                    axisLine={false}
                    tick={{ fontSize: 12, fill: "#d8e6f7" }}
                    interval={0}
                  />
                  <Tooltip
                    contentStyle={CHART_TOOLTIP_STYLE}
                    formatter={(v: number) => [v.toLocaleString(), "Tickets"]}
                    cursor={{ fill: "rgba(45, 212, 191, .06)" }}
                  />
                  <Bar
                    dataKey="value"
                    radius={BAR_RADIUS}
                    fill="url(#managedResourceBar)"
                    maxBarSize={18}
                  >
                    <LabelList
                      dataKey="value"
                      position="right"
                      fill={CHART_LABEL_FILL}
                      fontSize={12}
                    />
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </article>

            <article
              className="glass-card"
              style={{
                gridColumn: "7 / 13",
                display: "flex",
                flexDirection: "column",
                minHeight: 0,
              }}
            >
              <div
                className="card-heading"
                style={{
                  display: "flex",
                  justifyContent: "center",
                  alignItems: "center",
                  gap: "10px",
                }}
              >
                {" "}
                <h3>Average Ticket Resolution Time</h3> <Activity size={18} />
              </div>
              <div style={{ padding: "0 8px 8px", flex: 1, minHeight: 0 }}>
                {(analytics.monthlyResolutionTime ?? []).length === 0 ? (
                  <div
                    style={{
                      color: "#94a3b8",
                      fontSize: "12px",
                      textAlign: "center",
                      padding: "12px",
                    }}
                  >
                    No closed / resolved tickets with a duration in the current
                    selection.
                  </div>
                ) : (
                  <ResponsiveContainer width="100%" height={280}>
                    <LineChart
                      data={analytics.monthlyResolutionTime}
                      margin={{ top: 4, right: 18, left: 0, bottom: 36 }}
                    >
                      <CartesianGrid
                        strokeDasharray="3 3"
                        stroke={CHART_GRID_STROKE}
                        vertical={false}
                      />
                      <XAxis
                        dataKey="name"
                        tick={{ fill: "#aac0dc", fontSize: 10 }}
                        angle={-40}
                        textAnchor="end"
                        interval={0}
                        height={40}
                      />
                      <YAxis
                        tick={{ fill: "#aac0dc", fontSize: 10 }}
                        tickFormatter={(v: number) => `${v}h`}
                        width={42}
                      />
                      <Tooltip
                        contentStyle={CHART_TOOLTIP_STYLE}
                        labelStyle={{ color: "#e2e8f0", fontWeight: 700 }}
                        formatter={(
                          value: number,
                          name: string,
                          props: {
                            payload?: {
                              count?: number;
                              totalHours?: number;
                              pendingCount?: number;
                              pendingTotal?: number;
                            };
                          },
                        ) => {
                          const p = props.payload ?? {};
                          if (name === "Resolution (avg h)") {
                            return [
                              `${value.toFixed(2)} h (n=${p.count ?? 0}, total ${p.totalHours ?? 0} h)`,
                              name,
                            ];
                          }
                          if (name === "Pending (avg h)") {
                            return value > 0
                              ? [
                                  `${value.toFixed(2)} h (n=${p.pendingCount ?? 0}, total ${p.pendingTotal ?? 0} h)`,
                                  name,
                                ]
                              : ["-", name];
                          }
                          return [value, name];
                        }}
                      />
                      <Legend
                        verticalAlign="top"
                        height={22}
                        wrapperStyle={{
                          ...CHART_LEGEND_STYLE,
                          paddingBottom: 2,
                        }}
                        iconType="circle"
                        iconSize={8}
                      />
                      <Line
                        type="monotone"
                        dataKey="avgHours"
                        name="Resolution (avg h)"
                        stroke="#93c5fd"
                        strokeWidth={3}
                        dot={{ r: 3, fill: "#071426", strokeWidth: 2 }}
                        activeDot={{ r: 6, fill: "#93c5fd" }}
                        isAnimationActive={false}
                      />
                      <Line
                        type="monotone"
                        dataKey="pendingAvg"
                        name="Pending (avg h)"
                        stroke="#f59e0b"
                        strokeWidth={3}
                        dot={{ r: 3, fill: "#071426", strokeWidth: 2 }}
                        activeDot={{ r: 6, fill: "#f59e0b" }}
                        isAnimationActive={false}
                      />
                    </LineChart>
                  </ResponsiveContainer>
                )}
              </div>
            </article>

            <article className="glass-card" style={{ gridColumn: "1 / 13" }}>
              <div
                className="card-heading"
                style={{
                  display: "flex",
                  justifyContent: "center",
                  alignItems: "center",
                  gap: "10px",
                }}
              >
                {" "}
                <h3>Site Availability (All Sites)</h3> <BarChart3 size={18} />
              </div>

              {/* The Safety Guard: Only render if data exists and is not empty */}
              {performanceKpiRows && performanceKpiRows.length > 0 ? (
                <ResponsiveContainer width="100%" height={320}>
                  <BarChart
                    key={JSON.stringify(performanceKpiRows)}
                    data={performanceKpiRows}
                    layout="horizontal"
                    margin={{ left: 4, right: 18, top: 24, bottom: 6 }}
                    barCategoryGap="22%"
                  >
                    <defs>
                      <linearGradient
                        id="availabilityBar"
                        x1="0"
                        y1="1"
                        x2="0"
                        y2="0"
                      >
                        <stop offset="0%" stopColor="#059669" />
                        <stop offset="100%" stopColor="#86efac" />
                      </linearGradient>
                    </defs>
                    <CartesianGrid
                      stroke={CHART_GRID_STROKE}
                      vertical={false}
                    />
                    <XAxis
                      dataKey="displayName"
                      type="category"
                      stroke={CHART_AXIS_STROKE}
                      tick={{ fontSize: 8, fill: "#aac0dc" }}
                      angle={-65}
                      textAnchor="end"
                      interval={0}
                      height={88}
                    />
                    <YAxis
                      type="number"
                      stroke={CHART_AXIS_STROKE}
                      tickLine={false}
                      axisLine={false}
                      width={38}
                      tick={{ fontSize: 10, fill: "#aac0dc" }}
                    />
                    <Tooltip
                      contentStyle={CHART_TOOLTIP_STYLE}
                      cursor={{ fill: "rgba(16, 185, 129, .06)" }}
                    />
                    <Bar
                      dataKey="availHours"
                      radius={COLUMN_BAR_RADIUS}
                      fill="url(#availabilityBar)"
                      maxBarSize={12}
                    >
                      <LabelList
                        dataKey="availHours"
                        position="top"
                        fill="#d1fae5"
                        fontSize={9}
                        formatter={(value: number) => `${value}h`}
                      />
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              ) : (
                // What shows if data is missing or empty
                <div
                  style={{
                    height: "300px",
                    display: "flex",
                    alignItems: "center",
                    justifyContent: "center",
                    color: "#94a3b8",
                  }}
                >
                  No performance data available.
                </div>
              )}
            </article>

            <article className="glass-card" style={{ gridColumn: "1 / 13" }}>
              <div
                className="card-heading"
                style={{
                  display: "flex",
                  justifyContent: "center",
                  alignItems: "center",
                  gap: "5px",
                }}
              >
                {" "}
                <h3>Site Downtime (All Sites)</h3> <BarChart3 size={18} />
              </div>
              <div className="card-subheading">
                <span>Hours down per site</span>
              </div>
              {performanceKpiRows && performanceKpiRows.length > 0 ? (
                <ResponsiveContainer width="100%" height={320}>
                  <BarChart
                    key={"downtime-" + JSON.stringify(performanceKpiRows)}
                    data={performanceKpiRows}
                    layout="horizontal"
                    margin={{ left: 4, right: 18, top: 24, bottom: 6 }}
                    barCategoryGap="22%"
                  >
                    <CartesianGrid
                      stroke={CHART_GRID_STROKE}
                      vertical={false}
                    />
                    <XAxis
                      dataKey="displayName"
                      type="category"
                      stroke={CHART_AXIS_STROKE}
                      tick={{ fontSize: 8, fill: "#aac0dc" }}
                      angle={-65}
                      textAnchor="end"
                      interval={0}
                      height={88}
                    />
                    <YAxis
                      type="number"
                      stroke={CHART_AXIS_STROKE}
                      tickLine={false}
                      axisLine={false}
                      width={38}
                      tick={{ fontSize: 10, fill: "#aac0dc" }}
                    />
                    <Tooltip
                      contentStyle={CHART_TOOLTIP_STYLE}
                      formatter={(value: number) => [
                        `${value} hrs`,
                        "Down Hours",
                      ]}
                      cursor={{ fill: "rgba(248, 113, 113, .06)" }}
                    />
                    <Bar
                      dataKey="sitesDownHours"
                      radius={COLUMN_BAR_RADIUS}
                      maxBarSize={12}
                    >
                      {performanceKpiRows.map((row, index) => (
                        <Cell
                          key={`cell-${index}`}
                          fill={
                            row.sitesDownHours === 0
                              ? "rgba(148,163,184,0.25)"
                              : row.sitesDownHours > 24
                                ? "#ef4444"
                                : row.sitesDownHours > 8
                                  ? "#f59e0b"
                                  : "#fb923c"
                          }
                        />
                      ))}
                      <LabelList
                        dataKey="sitesDownHours"
                        position="top"
                        fill="#fecaca"
                        fontSize={9}
                        formatter={(value: number) =>
                          value > 0 ? `${value}h` : ""
                        }
                      />
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              ) : (
                <div
                  style={{
                    height: "300px",
                    display: "flex",
                    alignItems: "center",
                    justifyContent: "center",
                    color: "#94a3b8",
                  }}
                >
                  No performance data available.
                </div>
              )}
            </article>
          </section>

          {/* Tickets Table */}
          <section
            id="section-tickets-table"
            className="table-card dashboard-section-content-block"
          >
            <div className="table-heading">
              <div>
                <h2>
                  {filteredTickets.length.toLocaleString()} distinct Ticket
                  records
                </h2>
              </div>
              <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
                <span style={{ color: "#94a3b8", fontSize: 13 }}>
                  Page {tablePage} of{" "}
                  {Math.max(
                    1,
                    Math.ceil(filteredTickets.length / TABLE_PAGE_SIZE),
                  )}{" "}
                  &mdash; {filteredTickets.length.toLocaleString()} total
                </span>
              </div>
            </div>
            <div className="table-scroll" id="ticket-table-wrapper">
              <div ref={tableRef}>
                <div style={{ marginBottom: "12px" }}>
                  <button className="ghost-button" onClick={scrollToTopCards}>
                    <ArrowUp size={16} /> Back to Summary
                  </button>
                </div>

                <table
                  style={{
                    tableLayout: "fixed",
                    width: "max-content",
                    minWidth: "100%",
                    borderCollapse: "separate",
                  }}
                >
                  <colgroup>
                    {DISTINCT_REPORT_HEADERS.map((header) => (
                      <col
                        key={header}
                        style={{ width: `${getTicketColumnWidth(header)}px` }}
                      />
                    ))}
                  </colgroup>
                  <thead>
                    <tr>
                      {DISTINCT_REPORT_HEADERS.map((header) => {
                        return (
                          <th
                            key={header}
                            style={{
                              textAlign: "center",
                              padding: "12px",
                              verticalAlign: "middle",
                              position: "relative",
                              whiteSpace: "normal",
                              overflow: "hidden",
                              textOverflow: "ellipsis",
                            }}
                            title={header}
                          >
                            {header}
                            <span
                              onMouseDown={startTicketColumnResize(header)}
                              style={{
                                position: "absolute",
                                right: 0,
                                top: 0,
                                bottom: 0,
                                width: 6,
                                cursor: "col-resize",
                                userSelect: "none",
                                background:
                                  "linear-gradient(transparent 30%, rgba(148,163,184,0.35) 30% 70%, transparent 70%)",
                                backgroundSize: "100% 0",
                                backgroundRepeat: "no-repeat",
                                backgroundPosition: "center",
                                transition: "background-size 120ms ease",
                              }}
                              onMouseEnter={(e) => {
                                (
                                  e.currentTarget as HTMLSpanElement
                                ).style.backgroundSize = "100% 100%";
                              }}
                              onMouseLeave={(e) => {
                                (
                                  e.currentTarget as HTMLSpanElement
                                ).style.backgroundSize = "100% 0";
                              }}
                              title="Drag to resize column"
                              aria-label={`Resize ${header} column`}
                            />
                          </th>
                        );
                      })}
                    </tr>
                  </thead>
                  <tbody>
                    {filteredTickets
                      .slice(
                        (tablePage - 1) * TABLE_PAGE_SIZE,
                        tablePage * TABLE_PAGE_SIZE,
                      )
                      .map((ticket, index) => {
                        const absoluteIndex =
                          (tablePage - 1) * TABLE_PAGE_SIZE + index;
                        const row = ticket.primary;
                        const reportRow = distinctReportRow(
                          ticket,
                          absoluteIndex,
                        );
                        const siteIds = Array.from(ticket.siteIds).filter(
                          Boolean,
                        );
                        const siteNames = Array.from(ticket.siteNames).filter(
                          Boolean,
                        );

                        // Columns rendered with monospace font (codes, IDs, dates, times, durations).
                        const monoHeaders = new Set([
                          "#",
                          "Site ID",
                          "TT",
                          "Observation Date",
                          "Observation Time",
                          "Recovery Date",
                          "Recovery Time",
                          "Escalated for L3 Support Date",
                          "Escalated for L3 Support Time",
                          "Total Duration Days/Hours",
                        ]);

                        // Every cell follows the same pattern:
                        //   <td>  â†’  <div width=col-width, flex-column, overflow:hidden>  â†’  <span overflow:ellipsis>
                        // Outer <td> keeps `title` for hover-tooltip; inner div constrains the visual width.
                        const baseStyle: React.CSSProperties = {
                          verticalAlign: "top",
                          padding: "10px 12px",
                        };
                        const centeredTicketHeaders = new Set([
                          "Managed Resource",
                          "Severity",
                          "Observation Date",
                          "Observation Time",
                          "Recovery Date",
                          "Recovery Time",
                        ]);
                        const ticketCellStyle = (
                          header: string,
                        ): React.CSSProperties => ({
                          ...baseStyle,
                          textAlign: centeredTicketHeaders.has(header)
                            ? "center"
                            : undefined,
                          verticalAlign: centeredTicketHeaders.has(header)
                            ? "middle"
                            : "top",
                        });
                        const innerDivStyle = (
                          header: string,
                        ): React.CSSProperties => ({
                          display: "flex",
                          flexDirection: "column",
                          width: getTicketColumnWidth(header),
                          gap: "4px",
                          overflow: "hidden",
                          alignItems: centeredTicketHeaders.has(header)
                            ? "center"
                            : "stretch",
                          justifyContent: centeredTicketHeaders.has(header)
                            ? "center"
                            : "flex-start",
                          textAlign: centeredTicketHeaders.has(header)
                            ? "center"
                            : undefined,
                        });
                        const spanStyle: React.CSSProperties = {
                          overflow: "hidden",
                          textOverflow: "ellipsis",
                          whiteSpace: "normal",
                        };

                        return (
                          <tr key={ticket.tt}>
                            {reportRow.map((cell, cellIndex) => {
                              const header = DISTINCT_REPORT_HEADERS[cellIndex];
                              const isMono = monoHeaders.has(header);

                              // Multi-value: Site ID
                              if (header === "Site ID") {
                                const items = siteIds.length
                                  ? siteIds
                                  : [String(cell ?? "")];
                                return (
                                  <td
                                    key={header}
                                    className="mono"
                                    style={ticketCellStyle(header)}
                                    title={siteIds.join(", ")}
                                  >
                                    <div style={innerDivStyle(header)}>
                                      {items.map((id) => (
                                        <span key={id} style={spanStyle}>
                                          {id}
                                        </span>
                                      ))}
                                    </div>
                                  </td>
                                );
                              }

                              // Multi-value: Site Name
                              if (header === "Site Name") {
                                const items = siteNames.length
                                  ? siteNames
                                  : [String(cell ?? "")];
                                return (
                                  <td
                                    key={header}
                                    style={ticketCellStyle(header)}
                                    title={siteNames.join(", ")}
                                  >
                                    <div style={innerDivStyle(header)}>
                                      {items.map((name) => (
                                        <span key={name} style={spanStyle}>
                                          {name}
                                        </span>
                                      ))}
                                    </div>
                                  </td>
                                );
                              }

                              // Pill: Severity / Status â€” the pill itself replaces the span.
                              if (
                                header === "Severity" ||
                                header === "Status"
                              ) {
                                const tone =
                                  (header === "Severity"
                                    ? SEVERITY_COLORS
                                    : STATUS_COLORS)[String(cell ?? "")] ??
                                  "#64748b";
                                return (
                                  <td
                                    key={header}
                                    style={ticketCellStyle(header)}
                                    title={String(cell ?? "")}
                                  >
                                    <div style={innerDivStyle(header)}>
                                      <span
                                        className="pill"
                                        style={{ ["--pill" as string]: tone }}
                                      >
                                        {cell}
                                      </span>
                                    </div>
                                  </td>
                                );
                              }

                              // Default: single span, ellipsis truncation, hover for full value.
                              return (
                                <td
                                  key={header}
                                  className={isMono ? "mono" : undefined}
                                  style={ticketCellStyle(header)}
                                  title={String(cell ?? "")}
                                >
                                  <div style={innerDivStyle(header)}>
                                    <span style={spanStyle}>{cell}</span>
                                  </div>
                                </td>
                              );
                            })}
                          </tr>
                        );
                      })}
                  </tbody>
                </table>
              </div>
            </div>
            {Math.ceil(filteredTickets.length / TABLE_PAGE_SIZE) > 1 && (
              <div className="pagination-bar no-print">
                <button
                  className="ghost-button"
                  disabled={tablePage <= 1}
                  onClick={() => setTablePage(1)}
                >
                  <ChevronLeft size={16} />
                  <ChevronLeft size={16} />
                </button>
                <button
                  className="ghost-button"
                  disabled={tablePage <= 1}
                  onClick={() => setTablePage((p) => p - 1)}
                >
                  <ChevronLeft size={16} /> Prev
                </button>
                {Array.from(
                  {
                    length: Math.ceil(filteredTickets.length / TABLE_PAGE_SIZE),
                  },
                  (_, i) => i + 1,
                )
                  .filter(
                    (p) =>
                      p === 1 ||
                      p ===
                        Math.ceil(filteredTickets.length / TABLE_PAGE_SIZE) ||
                      Math.abs(p - tablePage) <= 2,
                  )
                  .reduce<(number | "...")[]>((acc, p, idx, arr) => {
                    if (idx > 0 && (p as number) - (arr[idx - 1] as number) > 1)
                      acc.push("...");
                    acc.push(p);
                    return acc;
                  }, [])
                  .map((p, idx) =>
                    p === "..." ? (
                      <span
                        key={`ellipsis-${idx}`}
                        style={{ color: "#94a3b8", padding: "0 4px" }}
                      >
                        ...
                      </span>
                    ) : (
                      <button
                        key={p}
                        className={`ghost-button${p === tablePage ? " active-page" : ""}`}
                        onClick={() => setTablePage(p as number)}
                      >
                        {p}
                      </button>
                    ),
                  )}
                <button
                  className="ghost-button"
                  disabled={
                    tablePage >=
                    Math.ceil(filteredTickets.length / TABLE_PAGE_SIZE)
                  }
                  onClick={() => setTablePage((p) => p + 1)}
                >
                  Next <ChevronRight size={16} />
                </button>
                <button
                  className="ghost-button"
                  disabled={
                    tablePage >=
                    Math.ceil(filteredTickets.length / TABLE_PAGE_SIZE)
                  }
                  onClick={() =>
                    setTablePage(
                      Math.ceil(filteredTickets.length / TABLE_PAGE_SIZE),
                    )
                  }
                >
                  <ChevronRight size={16} />
                  <ChevronRight size={16} />
                </button>
              </div>
            )}
          </section>

          {/* Monthly Performance Table */}
          {perfRows.length > 0 && (
            <section
              id="section-performance-table"
              className="table-card dashboard-section-content-block dashboard-section-tab-hidden"
            >
              <div className="table-heading">
                <div>
                  <h2>
                    Monthly Performance -{" "}
                    {perfMonths.length === 0
                      ? "All Months"
                      : perfMonths.length === 1
                        ? formatMonthMMMMYYYY(perfMonths[0])
                        : `${perfMonths.length} months`}
                    {perfRegions.length > 0
                      ? ` - ${perfRegions.join(", ")}`
                      : ""}
                  </h2>
                  <p style={{ color: "#94a3b8", fontSize: 13, marginTop: 4 }}>
                    {perfRows.length.toLocaleString()} sites &nbsp;|&nbsp; Total
                    hours in month:{" "}
                    {perfMonths.length === 1
                      ? totalHoursInMonth(perfMonths[0]).toLocaleString()
                      : perfMonths.length > 1
                        ? perfMonths
                            .reduce((s, m) => s + totalHoursInMonth(m), 0)
                            .toLocaleString()
                        : "N/A"}{" "}
                    hrs
                  </p>
                </div>
                <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
                  <span style={{ color: "#94a3b8", fontSize: 13 }}>
                    Page {perfPage} of{" "}
                    {Math.max(1, Math.ceil(perfRows.length / TABLE_PAGE_SIZE))}{" "}
                    &mdash; {perfRows.length.toLocaleString()} total
                  </span>
                </div>
              </div>
              <div style={{ marginBottom: "12px" }}>
                <button className="ghost-button" onClick={scrollToTopCards}>
                  <ArrowUp size={16} /> Back to Summary
                </button>
              </div>
              <div className="table-scroll">
                <table
                  style={{
                    tableLayout: "fixed",
                    width: "max-content",
                    minWidth: "100%",
                    borderCollapse: "separate",
                  }}
                >
                  <colgroup>
                    {PERF_REPORT_HEADERS.map((h) => (
                      <col
                        key={h}
                        style={{ width: `${getPerfColumnWidth(h)}px` }}
                      />
                    ))}
                  </colgroup>
                  <thead>
                    <tr>
                      {PERF_REPORT_HEADERS.map((header) => {
                        return (
                          <th
                            key={header}
                            style={{
                              textAlign:
                                header === "Site ID" || header === "Site Name"
                                  ? "left"
                                  : "center",
                              padding: "12px",
                              verticalAlign: "middle",
                              position: "relative",
                              whiteSpace: "nowrap",
                              overflow: "hidden",
                              textOverflow: "ellipsis",
                            }}
                            title={header}
                          >
                            {header}
                            <span
                              onMouseDown={startPerfColumnResize(header)}
                              style={{
                                position: "absolute",
                                right: 0,
                                top: 0,
                                bottom: 0,
                                width: 6,
                                cursor: "col-resize",
                                userSelect: "none",
                                background:
                                  "linear-gradient(transparent 30%, rgba(148,163,184,0.35) 30% 70%, transparent 70%)",
                                backgroundSize: "100% 0",
                                backgroundRepeat: "no-repeat",
                                backgroundPosition: "center",
                                transition: "background-size 120ms ease",
                              }}
                              onMouseEnter={(e) => {
                                (
                                  e.currentTarget as HTMLSpanElement
                                ).style.backgroundSize = "100% 100%";
                              }}
                              onMouseLeave={(e) => {
                                (
                                  e.currentTarget as HTMLSpanElement
                                ).style.backgroundSize = "100% 0";
                              }}
                              title="Drag to resize column"
                              aria-label={`Resize ${header} column`}
                            />
                          </th>
                        );
                      })}
                    </tr>
                  </thead>
                  <tbody>
                    {perfRows
                      .slice(
                        (perfPage - 1) * TABLE_PAGE_SIZE,
                        perfPage * TABLE_PAGE_SIZE,
                      )
                      .map((row, i) => {
                        // Original index (1-based) across the whole result set, not the current page.
                        const absoluteIndex =
                          (perfPage - 1) * TABLE_PAGE_SIZE + i + 1;
                        // Parse reliability % for color-coding.
                        const relNum = parseFloat(row.reliability);
                        const relColor = !row.sitesDownHours
                          ? undefined
                          : relNum < 95
                            ? "#ef4444" // red below 95%
                            : relNum < 99
                              ? "#f59e0b" // amber below 99%
                              : undefined; // no highlight at/above 99%

                        // Every cell follows the same pattern (matches the tickets table):
                        //   <td>  â†’  <div width=col-width, flex-column, overflow:hidden>  â†’  <span overflow:ellipsis>
                        const baseStyle: React.CSSProperties = {
                          verticalAlign: "top",
                          padding: "10px 12px",
                        };
                        const isPerfLeftAligned = (header: string) =>
                          header === "Site ID" || header === "Site Name";
                        const perfCellStyle = (
                          header: string,
                        ): React.CSSProperties => ({
                          ...baseStyle,
                          textAlign: isPerfLeftAligned(header)
                            ? "left"
                            : "center",
                          verticalAlign: isPerfLeftAligned(header)
                            ? "top"
                            : "middle",
                        });
                        const innerDivStyle = (
                          header: string,
                        ): React.CSSProperties => ({
                          display: "flex",
                          flexDirection: "column",
                          width: getPerfColumnWidth(header),
                          gap: "4px",
                          overflow: "hidden",
                          alignItems: isPerfLeftAligned(header)
                            ? "stretch"
                            : "center",
                          justifyContent: isPerfLeftAligned(header)
                            ? "flex-start"
                            : "center",
                          textAlign: isPerfLeftAligned(header)
                            ? "left"
                            : "center",
                          margin: isPerfLeftAligned(header)
                            ? undefined
                            : "0 auto",
                        });
                        const spanStyle: React.CSSProperties = {
                          overflow: "hidden",
                          textOverflow: "ellipsis",
                          whiteSpace: "nowrap",
                        };

                        const cells: Array<{
                          header: string;
                          value: React.ReactNode;
                          mono?: boolean;
                          extra?: React.CSSProperties;
                        }> = [
                          { header: "S No", value: absoluteIndex, mono: true },
                          { header: "Site ID", value: row.siteId, mono: true },
                          { header: "Site Name", value: row.siteName },
                          {
                            header: "Site Availability, Hrs",
                            value: row.availHours,
                            mono: true,
                          },
                          {
                            header: "Site Availability, days",
                            value: row.availDay,
                          },
                          {
                            header: "Channel Busy Count",
                            value: "",
                            mono: true,
                          },
                          {
                            header: "MW link Performance, Hrs",
                            value: "",
                            mono: true,
                          },
                          {
                            header: "DMR Reliability",
                            value: row.reliability,
                            mono: true,
                            extra: relColor
                              ? { color: relColor, fontWeight: 700 }
                              : undefined,
                          },
                          {
                            header: "Sites Down, hrs",
                            value: row.sitesDownHours,
                            mono: true,
                          },
                        ];

                        return (
                          <tr key={row.siteId}>
                            {cells.map(({ header, value, mono, extra }) => (
                              <td
                                key={header}
                                className={mono ? "mono" : undefined}
                                style={perfCellStyle(header)}
                                title={String(value ?? "")}
                              >
                                <div style={innerDivStyle(header)}>
                                  <span style={{ ...spanStyle, ...extra }}>
                                    {value}
                                  </span>
                                </div>
                              </td>
                            ))}
                          </tr>
                        );
                      })}
                    {/* Total summary row â€” always computed over the full dataset, shown only on the last page */}
                    {perfPage ===
                      Math.max(
                        1,
                        Math.ceil(perfRows.length / TABLE_PAGE_SIZE),
                      ) &&
                      perfRows.length > 0 &&
                      (() => {
                        const totalDown =
                          Math.round(
                            perfRows.reduce((s, r) => s + r.sitesDownHours, 0) *
                              10,
                          ) / 10;
                        const totalAvail =
                          Math.round(
                            perfRows.reduce((s, r) => s + r.availHours, 0) * 10,
                          ) / 10;
                        const totalHrs = totalAvail + totalDown;
                        const overallRel =
                          totalHrs > 0
                            ? ((totalAvail / totalHrs) * 100).toFixed(2) + "%"
                            : "";
                        const relNum = parseFloat(overallRel);
                        const relColor =
                          totalDown === 0
                            ? undefined
                            : relNum < 95
                              ? "#ef4444"
                              : relNum < 99
                                ? "#f59e0b"
                                : "#22c55e";
                        return (
                          <tr
                            style={{
                              fontWeight: 700,
                              borderTop: "2px solid rgba(148,163,184,0.3)",
                              background: "rgba(255,255,255,0.04)",
                            }}
                          >
                            <td
                              className="mono"
                              colSpan={3}
                              style={{
                                textAlign: "right",
                                paddingRight: 16,
                                color: "#94a3b8",
                              }}
                            >
                              TOTAL
                            </td>
                            <td
                              className="mono"
                              style={{ textAlign: "center" }}
                            >
                              {totalAvail}
                            </td>
                            <td style={{ textAlign: "center" }}></td>
                            <td
                              className="mono"
                              style={{ textAlign: "center" }}
                            ></td>
                            <td
                              className="mono"
                              style={{ textAlign: "center" }}
                            ></td>
                            <td
                              className="mono"
                              style={{
                                textAlign: "center",
                                ...(relColor
                                  ? { color: relColor, fontWeight: 700 }
                                  : {}),
                              }}
                            >
                              {overallRel}
                            </td>
                            <td
                              className="mono"
                              style={{ textAlign: "center" }}
                            >
                              {totalDown}
                            </td>
                          </tr>
                        );
                      })()}
                  </tbody>
                </table>
              </div>
              {Math.ceil(perfRows.length / TABLE_PAGE_SIZE) > 1 && (
                <div className="pagination-bar no-print">
                  <button
                    className="ghost-button"
                    disabled={perfPage <= 1}
                    onClick={() => setPerfPage(1)}
                  >
                    <ChevronLeft size={16} />
                    <ChevronLeft size={16} />
                  </button>
                  <button
                    className="ghost-button"
                    disabled={perfPage <= 1}
                    onClick={() => setPerfPage((p) => p - 1)}
                  >
                    <ChevronLeft size={16} /> Prev
                  </button>
                  {Array.from(
                    { length: Math.ceil(perfRows.length / TABLE_PAGE_SIZE) },
                    (_, i) => i + 1,
                  )
                    .filter(
                      (p) =>
                        p === 1 ||
                        p === Math.ceil(perfRows.length / TABLE_PAGE_SIZE) ||
                        Math.abs(p - perfPage) <= 2,
                    )
                    .reduce<(number | "...")[]>((acc, p, idx, arr) => {
                      if (
                        idx > 0 &&
                        (p as number) - (arr[idx - 1] as number) > 1
                      )
                        acc.push("...");
                      acc.push(p);
                      return acc;
                    }, [])
                    .map((p, idx) =>
                      p === "..." ? (
                        <span
                          key={`ellipsis-${idx}`}
                          style={{ color: "#94a3b8", padding: "0 4px" }}
                        >
                          ...
                        </span>
                      ) : (
                        <button
                          key={p}
                          className={`ghost-button${p === perfPage ? " active-page" : ""}`}
                          onClick={() => setPerfPage(p as number)}
                        >
                          {p}
                        </button>
                      ),
                    )}
                  <button
                    className="ghost-button"
                    disabled={
                      perfPage >= Math.ceil(perfRows.length / TABLE_PAGE_SIZE)
                    }
                    onClick={() => setPerfPage((p) => p + 1)}
                  >
                    Next <ChevronRight size={16} />
                  </button>
                  <button
                    className="ghost-button"
                    disabled={
                      perfPage >= Math.ceil(perfRows.length / TABLE_PAGE_SIZE)
                    }
                    onClick={() =>
                      setPerfPage(Math.ceil(perfRows.length / TABLE_PAGE_SIZE))
                    }
                  >
                    <ChevronRight size={16} />
                    <ChevronRight size={16} />
                  </button>
                </div>
              )}
            </section>
          )}
        </>
      )}
    </main>
  );
}
