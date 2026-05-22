/*
Design philosophy: Calm Enterprise Glass for a premium telecom network-observability cockpit.
Use deep ink backgrounds, translucent panels, cyan focus accents, tabular TT numerals, and restrained motion.
Does this choice reinforce or dilute our design philosophy?
*/
import { useEffect, useMemo, useRef, useState } from "react";
import { createPortal } from "react-dom";
import * as XLSX from "xlsx";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import {
  Activity,
  AlertTriangle,
  BarChart3,
  CheckCircle2,
  ChevronDown,
  CircleDot,
  Download,
  FileSpreadsheet,
  Filter,
  Layers3,
  Network,
  Printer,
  RefreshCw,
  Search,
  ShieldAlert,
  UploadCloud,
  XCircle,
} from "lucide-react";
import {
  Bar,
  BarChart,
  CartesianGrid,
  Cell,
  LabelList,
  Legend,
  Line,
  LineChart,
  Pie,
  PieChart,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from "recharts";

import nascoLogoSrc from "/nascologo.png";
import ngLogoSrc from "/nglogo.png";
const HERO_IMAGE = "/h.png";
const UPLOAD_IMAGE = "/h.png";
const RIBBON_IMAGE = "/h.png";

const COLORS = ["#22d3ee", "#60a5fa", "#f59e0b", "#ef4444", "#34d399", "#a78bfa", "#f472b6", "#94a3b8"];
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

type TicketRecord = {
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
  status: string;
  rca: string;
  rcaFamily: string;
  preventability: string;
  responsibleTeam: string;
  recommendedAction: string;
  actionTaken: string;
};

type TicketAggregate = {
  tt: string;
  primary: TicketRecord;
  siteIds: Set<string>;
  siteNames: Set<string>;
  rows: TicketRecord[];
};

type DashboardData = {
  fileName: string;
  sheetName: string;
  generatedAt: string;
  rows: TicketRecord[];
  uniqueTickets: TicketAggregate[];
  /** Ordered sites from the SiteID sheet (if present), used to sort the Monthly Performance table */
  siteOrder: { siteId: string; siteName: string }[];
};


type Filters = {
  search: string;
  status: string[];
  severity: string[];
  region: string[];
  impact: string[];
  site: string[];
  openingMonth: string[];
  rcaFamily: string[];
};

const EMPTY_FILTERS: Filters = { search: "", status: [], severity: [], region: [], impact: [], site: [], openingMonth: [], rcaFamily: [] };

function clean(value: unknown): string {
  if (value === null || value === undefined) return "";
  return String(value).replace(/\s+/g, " ").trim();
}

function normalizeHeader(value: string): string {
  return clean(value).toLowerCase().replace(/[^a-z0-9]/g, "");
}

function getField(row: Record<string, unknown>, aliases: string[]): string {
  const map = new Map<string, unknown>();
  Object.entries(row).forEach(([key, value]) => map.set(normalizeHeader(key), value));
  for (const alias of aliases) {
    const found = map.get(normalizeHeader(alias));
    if (found !== undefined) return clean(found);
  }
  return "";
}

// Like getField but returns the raw value without converting to string (needed for Date/time cells)
function getRawField(row: Record<string, unknown>, aliases: string[]): unknown {
  const map = new Map<string, unknown>();
  Object.entries(row).forEach(([key, value]) => map.set(normalizeHeader(key), value));
  for (const alias of aliases) {
    const found = map.get(normalizeHeader(alias));
    if (found !== undefined && found !== null && found !== "") return found;
  }
  return "";
}

function parseDateValue(value: unknown): Date | null {
  if (value === null || value === undefined || value === "") return null;
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
  const numericValue = typeof value === "number" ? value : /^\d+(\.\d+)?$/.test(clean(value)) ? Number(clean(value)) : null;
  if (numericValue !== null && numericValue > 20000 && numericValue < 90000) {
    const excelDate = XLSX.SSF.parse_date_code(numericValue);
    if (excelDate) return new Date(excelDate.y, excelDate.m - 1, excelDate.d);
  }
  const text = clean(value);
  // Handle dd/mm/yyyy and dd-mm-yyyy (produced by toDateStr helper)
  const dmyMatch = text.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/);
  if (dmyMatch) {
    const d = Number(dmyMatch[1]);
    const m = Number(dmyMatch[2]);
    const y = Number(dmyMatch[3]);
    // Detect dd/mm/yyyy vs mm/dd/yyyy: if day > 12 it must be dd/mm/yyyy
    if (d > 12) {
      const candidate = new Date(y, m - 1, d);
      if (!Number.isNaN(candidate.getTime())) return candidate;
    } else {
      // Ambiguous: assume dd/mm/yyyy (our toDateStr always produces dd/mm/yyyy)
      const candidate = new Date(y, m - 1, d);
      if (!Number.isNaN(candidate.getTime())) return candidate;
    }
  }
  const direct = new Date(text);
  if (!Number.isNaN(direct.getTime())) return direct;
  const match = text.match(/(\d{1,2})[-/ ]([A-Za-z]{3,}|\d{1,2})[-/ ](\d{2,4})/);
  if (!match) return null;
  const parsed = new Date(text.replace(/-/g, " "));
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

function normalizeMonthKey(value: unknown): string | null {
  const text = clean(value);
  if (!text) return null;
  const keyMatch = text.match(/^(\d{4})[-/](\d{1,2})$/);
  if (keyMatch) return `${keyMatch[1]}-${keyMatch[2].padStart(2, "0")}`;
  const parsed = parseDateValue(text.startsWith("1 ") ? text : `1 ${text}`);
  if (!parsed) return null;
  return `${parsed.getFullYear()}-${String(parsed.getMonth() + 1).padStart(2, "0")}`;
}

function openingMonthKey(value: unknown): string {
  const parsed = parseDateValue(value);
  if (!parsed) return "Unknown";
  const month = String(parsed.getMonth() + 1).padStart(2, "0");
  return `${parsed.getFullYear()}-${month}`;
}

function resolveOpeningMonthKey(sourceKey: unknown, sourceLabel: unknown, observationDate: unknown): string {
  return normalizeMonthKey(sourceKey) ?? normalizeMonthKey(sourceLabel) ?? openingMonthKey(observationDate);
}

function openingMonthLabel(key: string): string {
  if (!key || key === "Unknown") return "Unknown";
  const parsed = new Date(`${key}-01T00:00:00`);
  if (Number.isNaN(parsed.getTime())) return key;
  return parsed.toLocaleDateString("en", { month: "short", year: "numeric" });
}

function recordDateMonthKey(value: string): string {
  const parsed = parseDateValue(value);
  if (!parsed) return "Unknown";
  return `${parsed.getFullYear()}-${String(parsed.getMonth() + 1).padStart(2, "0")}`;
}

function selectedMonthRange(selectedMonth: string): { start: Date; end: Date } | null {
  const match = selectedMonth.match(/^(\d{4})-(\d{2})$/);
  if (!match) return null;
  const year = Number(match[1]);
  const monthIndex = Number(match[2]) - 1;
  const start = new Date(year, monthIndex, 1);
  const end = new Date(year, monthIndex + 1, 0, 23, 59, 59, 999);
  return Number.isNaN(start.getTime()) || Number.isNaN(end.getTime()) ? null : { start, end };
}

/** Total hours in a given month key (e.g. "2025-03" = 31 * 24 = 744) */
function totalHoursInMonth(monthKey: string): number {
  const range = selectedMonthRange(monthKey);
  if (!range) return 0;
  // days in month = day of last day
  const daysInMonth = range.end.getDate();
  return daysInMonth * 24;
}

function coveredMonthKeys(row: TicketRecord): string[] {
  const observation = parseDateValue(row.observationDate);
  const recovery = parseDateValue(row.recoveryDate);
  const keys = new Set<string>();
  const addEndpoint = (value: string) => {
    const key = recordDateMonthKey(value);
    if (key !== "Unknown") keys.add(key);
  };
  addEndpoint(row.observationDate);
  addEndpoint(row.recoveryDate);
  if (!observation || !recovery) return Array.from(keys);
  const startDate = observation <= recovery ? observation : recovery;
  const endDate = recovery >= observation ? recovery : observation;
  const cursor = new Date(startDate.getFullYear(), startDate.getMonth(), 1);
  const final = new Date(endDate.getFullYear(), endDate.getMonth(), 1);
  let guard = 0;
  while (cursor <= final && guard < 240) {
    keys.add(`${cursor.getFullYear()}-${String(cursor.getMonth() + 1).padStart(2, "0")}`);
    cursor.setMonth(cursor.getMonth() + 1);
    guard += 1;
  }
  return Array.from(keys);
}

function isPendingStatus(value: string): boolean {
  return clean(value).toLowerCase() === "pending";
}

function dateWithinMonth(dateValue: string, selectedMonth: string): boolean {
  const range = selectedMonthRange(selectedMonth);
  if (!range) return false;
  const parsed = parseDateValue(dateValue);
  if (!parsed) return false;
  return parsed >= range.start && parsed <= range.end;
}

function ticketMatchesMonthlyExport(ticket: TicketAggregate, selectedMonth: string): boolean {
  if (selectedMonth === "all") return true;
  const range = selectedMonthRange(selectedMonth);
  if (!range) return false;

  // Check all rows of the ticket (a TT can span multiple site rows)
  return ticket.rows.some((row) => {
    const obsDate = parseDateValue(row.observationDate);
    const recDate = parseDateValue(row.recoveryDate);

    // Criterion 1: Observation Date falls within the selected month
    const observationInMonth = obsDate !== null && obsDate >= range.start && obsDate <= range.end;

    // Criterion 2: Recovery Date falls within the selected month
    const recoveryInMonth = recDate !== null && recDate >= range.start && recDate <= range.end;

    // Criterion 3: Status is Pending AND Observation Date <= last day of the selected month
    const pendingBeforeMonthEnd = isPendingStatus(row.status) && obsDate !== null && obsDate <= range.end;

    // Criterion 4: Ticket was active throughout the entire month
    const spansEntireMonth = obsDate !== null && recDate !== null && obsDate < range.start && recDate > range.end;

    return observationInMonth || recoveryInMonth || pendingBeforeMonthEnd || spansEntireMonth;
  });
}

function dateKey(value: string): number {
  const parsed = new Date(value);
  return Number.isNaN(parsed.getTime()) ? 0 : parsed.getTime();
}

function formatDateDDMMYYYY(value: string): string {
  const parsed = parseDateValue(value);
  if (!parsed) return value || "";
  const d = String(parsed.getDate()).padStart(2, "0");
  const m = String(parsed.getMonth() + 1).padStart(2, "0");
  const y = parsed.getFullYear();
  return `${d}/${m}/${y}`;
}

function formatMonthMMMMYYYY(monthKey: string): string {
  if (!monthKey || monthKey === "all") return "";
  const parsed = new Date(`${monthKey}-01T00:00:00`);
  if (Number.isNaN(parsed.getTime())) return monthKey;
  return parsed.toLocaleDateString("en", { month: "short", year: "numeric" });
}

async function exportFormattedExcel(rows: TicketAggregate[], monthKey: string) {
  const templateUrl = "/manus-storage/DMRMonthlyReportEOA_76429c79.xlsx";
  let templateBuffer: ArrayBuffer;
  try {
    const response = await fetch(templateUrl);
    if (!response.ok) throw new Error("Template not found");
    templateBuffer = await response.arrayBuffer();
  } catch {
    exportExcel(rows);
    return;
  }
  const wb = XLSX.read(templateBuffer, { type: "array", cellStyles: true });
  const ws = wb.Sheets[wb.SheetNames[0]];

  const monthLabel = formatMonthMMMMYYYY(monthKey);
  if (monthLabel) {
    ws["R5"] = { t: "s", v: monthLabel };
  }

  const DATA_START_ROW = 39;
  const colLetter = (n: number) => {
    let s = "";
    while (n > 0) { s = String.fromCharCode(((n - 1) % 26) + 65) + s; n = Math.floor((n - 1) / 26); }
    return s;
  };
  const setCell = (row: number, col: number, value: string | number) => {
    const addr = `${colLetter(col)}${row}`;
    ws[addr] = { t: typeof value === "number" ? "n" : "s", v: value };
  };

  for (let r = DATA_START_ROW; r < 60; r++) {
    for (let c = 1; c <= 18; c++) {
      delete ws[`${colLetter(c)}${r}`];
    }
  }

  rows.forEach((ticket, idx) => {
    const row = ticket.primary;
    const r = DATA_START_ROW + idx;
    const siteIds = Array.from(ticket.siteIds).join(", ");
    const siteNames = Array.from(ticket.siteNames).join(", ");
    setCell(r, 1, idx + 1);
    setCell(r, 2, siteIds || row.siteId || "");
    setCell(r, 3, siteNames || row.siteName || "");
    setCell(r, 5, row.managedResource || "");
    setCell(r, 6, row.severity || "");
    setCell(r, 7, row.issue || "");
    setCell(r, 8, formatDateDDMMYYYY(row.observationDate));
    setCell(r, 9, row.observationTime || "");
    setCell(r, 10, formatDateDDMMYYYY(row.recoveryDate));
    setCell(r, 11, row.recoveryTime || "");
    setCell(r, 12, formatDateDDMMYYYY(row.escalatedForL3SupportDate));
    setCell(r, 13, row.escalatedForL3SupportTime || "");
    setCell(r, 14, row.duration || "");
    setCell(r, 15, ticket.tt || "");
    setCell(r, 16, row.status || "");
    setCell(r, 17, row.escalatedTo || "");
    setCell(r, 18, row.actionTaken || "");
  });

  if (!ws["!ref"]) ws["!ref"] = "A1:R62";
  else {
    const lastDataRow = Math.max(DATA_START_ROW + rows.length - 1, 59);
    ws["!ref"] = `A1:R${lastDataRow}`;
  }

  const monthSuffix = monthKey !== "all" ? `-${formatMonthMMMMYYYY(monthKey).replace(" ", "-")}` : "";
  XLSX.writeFile(wb, `DMR-Monthly-Report${monthSuffix}.xlsx`);
}

function parseDurationHours(duration: string): number | null {
  if (!duration) return null;
  // Handle text format: "3 days 21 hrs 29 mins" or "0 days 2 hrs 21 mins"
  if (/days?/i.test(duration) || /hrs?/i.test(duration)) {
    const days = Number(duration.match(/(\d+)\s*days?/i)?.[1] ?? 0);
    const hrs = Number(duration.match(/(\d+)\s*hrs?/i)?.[1] ?? 0);
    const mins = Number(duration.match(/(\d+)\s*mins?/i)?.[1] ?? 0);
    const total = days * 24 + hrs + mins / 60;
    return Number.isFinite(total) ? total : null;
  }
  // Numeric value: Duration column stores hours directly (e.g. 6.32 = 6.32 hrs)
  // Small fractions < 1 may be Excel day-fractions from some workbook versions -- convert those
  const num = Number(duration);
  if (Number.isFinite(num) && num >= 0) {
    if (num > 0 && num < 1) return Math.round(num * 24 * 10) / 10; // day-fraction edge case
    return Math.round(num * 10) / 10; // already in hours
  }
  return null;
}

const RCA_FAMILY_MAP: Record<string, string> = {
  "Power Issue": "Power & Environment",
  "High Temperature": "Power & Environment",
  "High VSWR": "Power & Environment",
  "Weather Issue": "Power & Environment",
  "DC Charger Faulty": "Power & Environment",
  "Link Down": "Transmission & Link",
  "Link Flapping": "Transmission & Link",
  "Transmission": "Transmission & Link",
  "MW Issue": "Transmission & Link",
  "MPLS Issue": "Transmission & Link",
  "SDH Hanged": "Transmission & Link",
  "Hardware Failure": "Hardware & Device",
  "Hardware Faulty": "Hardware & Device",
  "Device Hanged": "Hardware & Device",
  "Fiber Cut": "Fiber & Physical",
  "Cabling": "Fiber & Physical",
  "Port Disable": "Fiber & Physical",
  "Port Hang": "Fiber & Physical",
  "Loss of Signal": "Fiber & Physical",
  "Media Converter Faulty": "Fiber & Physical",
  "Configuration Issue": "Configuration / Software",
  "Software Issue": "Configuration / Software",
  "Application Issue": "Configuration / Software",
  "Human Mistake": "Human / Process / Planned",
  "Approved Activity": "Human / Process / Planned",
  "Un-Approved Activity": "Human / Process / Planned",
  "Planned Activity": "Human / Process / Planned",
  "Project Team": "Human / Process / Planned",
  "FMD Team": "Human / Process / Planned",
  "NG FO Team": "Human / Process / Planned",
  "RCA not Provided": "Unknown / Missing",
};

const NON_PREVENTABLE_RCAS = new Set(["Weather Issue", "Approved Activity", "Planned Activity", "RCA not Provided"]);

const RESPONSIBLE_TEAM_BY_FAMILY: Record<string, string> = {
  "Power & Environment": "Power / Facilities Team",
  "Transmission & Link": "Transmission / NOC Team",
  "Hardware & Device": "Field Maintenance / Vendor",
  "Fiber & Physical": "Fiber / Physical Maintenance Team",
  "Configuration / Software": "NOC / Configuration Team",
  "Human / Process / Planned": "Process Owner / Project Team",
  "Unknown / Missing": "RCA Owner / Follow-up Required",
  "Other / Review": "Operations Review Team",
};

const RECOMMENDED_ACTION_BY_FAMILY: Record<string, string> = {
  "Power & Environment": "Check power source, rectifier, batteries, grounding, cooling, and repeated environmental alarms.",
  "Transmission & Link": "Review link stability, transmission path, MPLS/MW/SDH health, and vendor escalation history.",
  "Hardware & Device": "Inspect device health, replace faulty hardware, verify spares, and monitor repeated failures.",
  "Fiber & Physical": "Inspect fiber/cabling route, port status, optical levels, patching quality, and civil-work exposure.",
  "Configuration / Software": "Review recent changes, configuration backup, software version, rollback records, and approval controls.",
  "Human / Process / Planned": "Validate activity approval, handover, method of procedure, and team process compliance.",
  "Unknown / Missing": "Complete RCA, assign owner, and update action taken before closure reporting.",
  "Other / Review": "Review RCA text manually and assign the correct operational owner.",
};

function getRcaFamily(rca: string): string {
  const normalized = clean(rca) || "RCA not Provided";
  return RCA_FAMILY_MAP[normalized] ?? "Other / Review";
}

function getPreventability(rca: string): string {
  const normalized = clean(rca) || "RCA not Provided";
  return NON_PREVENTABLE_RCAS.has(normalized) ? "Non-preventable" : "Preventable";
}

function getResponsibleTeam(rcaFamily: string): string {
  return RESPONSIBLE_TEAM_BY_FAMILY[rcaFamily] ?? RESPONSIBLE_TEAM_BY_FAMILY["Other / Review"];
}

function getRecommendedAction(rcaFamily: string): string {
  return RECOMMENDED_ACTION_BY_FAMILY[rcaFamily] ?? RECOMMENDED_ACTION_BY_FAMILY["Other / Review"];
}

function rcaNotProvided(rca: string): boolean {
  const normalized = clean(rca).toLowerCase();
  return !normalized || normalized === "rca not provided";
}

function formatHours(value: number): string {
  if (!value || !Number.isFinite(value)) return "";
  return `${value.toFixed(1)} hrs`;
}

function groupTickets(rows: TicketRecord[]): TicketAggregate[] {
  const grouped = new Map<string, TicketAggregate>();
  rows.forEach((row) => {
    if (!row.tt) return;
    const existing = grouped.get(row.tt);
    if (!existing) {
      grouped.set(row.tt, {
        tt: row.tt,
        primary: row,
        siteIds: new Set(row.siteId ? [row.siteId] : []),
        siteNames: new Set(row.siteName ? [row.siteName] : []),
        rows: [row],
      });
    } else {
      existing.rows.push(row);
      if (row.siteId) existing.siteIds.add(row.siteId);
      if (row.siteName) existing.siteNames.add(row.siteName);
      const currentDate = dateKey(existing.primary.observationDate);
      const nextDate = dateKey(row.observationDate);
      if (!currentDate || (nextDate && nextDate < currentDate)) existing.primary = row;
    }
  });
  return Array.from(grouped.values());
}

function parseSiteOrder(workbook: XLSX.WorkBook): { siteId: string; siteName: string }[] {
  // The SiteID table lives on the "Dashboard_Data" sheet at columns J-L (indices 9-11),
  // with the header row containing "#", "Site ID", "Site Name".
  // We scan the sheet as a 2D array to find the header row, then read data rows below it.
  const siteSheetName = workbook.SheetNames.find((name) => {
    const n = name.toLowerCase().replace(/[^a-z0-9]/g, "");
    return n === "dashboarddata" || n === "siteid" || n === "sites" || n === "sitelist" || n === "siteids" || n === "sitedata";
  });
  if (!siteSheetName) return [];
  const sheet = workbook.Sheets[siteSheetName];
  // Read as 2D array (no header inference)
  const raw2d = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: null }) as unknown[][];

  // Find the header row: look for a row that has "Site ID" somewhere
  let siteIdCol = -1;
  let siteNameCol = -1;
  let headerRowIdx = -1;
  for (let ri = 0; ri < raw2d.length; ri++) {
    const row = raw2d[ri];
    for (let ci = 0; ci < row.length; ci++) {
      const cell = String(row[ci] ?? "").trim().toLowerCase();
      if (cell === "site id" || cell === "siteid") {
        siteIdCol = ci;
        headerRowIdx = ri;
      }
      if (cell === "site name" || cell === "sitename") {
        siteNameCol = ci;
      }
    }
    if (siteIdCol >= 0 && siteNameCol >= 0) break;
  }

  if (headerRowIdx < 0 || siteIdCol < 0) return [];

  const seen = new Set<string>();
  const result: { siteId: string; siteName: string }[] = [];

  for (let ri = headerRowIdx + 1; ri < raw2d.length; ri++) {
    const row = raw2d[ri];
    const id = clean(String(row[siteIdCol] ?? ""));
    const name = siteNameCol >= 0 ? clean(String(row[siteNameCol] ?? "")) : "";
    // Stop if we hit an empty Site ID cell (end of table)
    if (!id) break;
    if (!seen.has(id)) {
      seen.add(id);
      result.push({ siteId: id, siteName: name });
    }
  }

  return result;
}

function parseRows(workbook: XLSX.WorkBook, fileName: string): DashboardData {
  const preferred = workbook.SheetNames.find((name) => name.toLowerCase().includes("tickets_data"));
  const sheetName = preferred ?? workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const raw = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, { defval: "", raw: true });
  const siteOrder = parseSiteOrder(workbook);

  function toDateStr(val: unknown): string {
    if (val === null || val === undefined || val === "") return "";
    const parsed = parseDateValue(val);
    if (!parsed) return String(val);
    const d = String(parsed.getDate()).padStart(2, "0");
    const m = String(parsed.getMonth() + 1).padStart(2, "0");
    return `${d}/${m}/${parsed.getFullYear()}`;
  }

  function toTimeStr(val: unknown): string {
    if (val === null || val === undefined || val === "") return "";
    if (val instanceof Date) {
      const hh = String(val.getUTCHours()).padStart(2, "0");
      const mm = String(val.getUTCMinutes()).padStart(2, "0");
      return `${hh}:${mm}`;
    }
    if (typeof val === "number") {
      const totalMinutes = Math.round(val * 24 * 60);
      const hh = String(Math.floor(totalMinutes / 60) % 24).padStart(2, "0");
      const mm = String(totalMinutes % 60).padStart(2, "0");
      return `${hh}:${mm}`;
    }
    const text = String(val).trim();
    const hhmmss = text.match(/^(\d{1,2}):(\d{2})(:\d{2})?$/);
    if (hhmmss) return `${hhmmss[1].padStart(2, "0")}:${hhmmss[2]}`;
    return text;
  }

  const rows: TicketRecord[] = raw
    .map((row, index) => {
      const observationDate = toDateStr(getRawField(row, ["Observation Date", "Observed Date"]) || "");
      const monthKey = resolveOpeningMonthKey(
        getField(row, ["Opening Month Key", "OpeningMonthKey"]),
        getField(row, ["Opening Month", "OpeningMonth"]),
        observationDate,
      );
      const rca = getField(row, ["RCA", "Root Cause Analysis", "Root Cause", "Action Taken/RCA"]);
      const rcaFamily = getRcaFamily(rca);
      return {
        rowNo: index + 2,
        tt: getField(row, ["TT", "Ticket", "Ticket Number"]),
        siteId: getField(row, ["Site ID", "SiteID", "Site Name", "SiteName"]),
        siteName: getField(row, ["Site Name", "SiteName"]),
        managedResource: getField(row, ["Managed Resource", "ManagedResource", "Managed Resources", "Resource", "NE Name", "Network Element"]),
        issue: getField(row, ["Issues", "Issue"]),
        severity: getField(row, ["Severity"]),
        region: getField(row, ["Region"]),
        observationDate,
        observationTime: toTimeStr(getRawField(row, ["Observation Time", "Observed Time", "ObservationTime"])),
        openingMonthKey: monthKey,
        openingMonthLabel: openingMonthLabel(monthKey),
        recoveryDate: toDateStr(getRawField(row, ["Recovery Date"]) || ""),
        recoveryTime: toTimeStr(getRawField(row, ["Recovery Time", "RecoveryTime"])),
        duration: String(getRawField(row, ["Total Duration Days/Hours", "Total Durration Days/Hours", "Duration"]) ?? ""),
        impact: getField(row, ["Service Impaction Status", "Service Impact Status"]),
        escalatedTo: getField(row, ["Escalated to", "Escalated To", "Escalated to "]),
        escalationLevel: getField(row, ["Escalation Level", "Esclation Level"]),
        escalatedForL3SupportDate: toDateStr(getField(row, ["Escalated for L3 Support Date", "Escalated For L3 Support Date", "L3 Support Date", "L3 Escalation Date", "Escalation L3 Date", "Escalated L3 Date"]) || ""),
        escalatedForL3SupportTime: toTimeStr(getRawField(row, ["Escalated for L3 Support Time", "Escalated For L3 Support Time", "L3 Support Time", "L3 Escalation Time", "Escalation L3 Time", "Escalated L3 Time"])),
        status: getField(row, ["Status"]),
        rca,
        rcaFamily,
        preventability: getPreventability(rca),
        responsibleTeam: getResponsibleTeam(rcaFamily),
        recommendedAction: getRecommendedAction(rcaFamily),
        actionTaken: getField(row, ["Action Taken/RCA", "Action Taken"]),
      };
    })
    .filter((row) => row.tt || row.siteId || row.siteName || row.issue);

  return {
    fileName,
    sheetName,
    generatedAt: new Date().toLocaleString(),
    rows,
    uniqueTickets: groupTickets(rows),
    siteOrder,
  };
}

function countBy<T>(items: T[], keyFn: (item: T) => string): { name: string; value: number }[] {
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

function metricValue(items: { name: string; value: number }[], expected: string): number {
  return items.find((item) => item.name.toLowerCase() === expected.toLowerCase())?.value ?? 0;
}

function renderPieLabel(props: { name?: string; value?: number; percent?: number }) {
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
  "Action",
];

// Monthly Performance table columns (report order)
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

type PerfRow = {
  siteId: string;      // e.g. "RF Site 1"
  siteName: string;    // e.g. "Karma S_S (EOA)"
  sitesDownHours: number;
  availHours: number;
  availDay: string;
  reliability: string;
  channelBusy: number;
  mwLinkPerf: number;
  ticketCount: number; // used for fallback sorting
};

/**
 * Compute Monthly Performance rows.
 * Sites Down (Hour) = SUMPRODUCT of duration hours for rows where:
 *   - Site ID matches
 *   - Opening Month is within the selected month
 *   - Service Impaction Status = "Service Impact"
 *   - Managed Resource = "Complete site" OR "Link Down"
 * Sorted by the SiteID sheet order (siteOrder). Falls back to ticket-count descending if no SiteID sheet.
 */
/** Normalize site ID: strip leading zeros and normalize case to match SiteID table.
 * "RF Site 01" -> "RF Site 1", "Rf Site 09" -> "RF Site 9" */
function normalizeSiteId(id: string): string {
  return id
    .replace(/(\D+)(\d+)$/, (_, prefix, num) => prefix.toUpperCase() + String(parseInt(num, 10)))
    .trim();
}

function computePerfRows(allRows: TicketRecord[], monthKey: string, siteOrder: { siteId: string; siteName: string }[] = []): PerfRow[] {
  const range = monthKey !== "all" ? selectedMonthRange(monthKey) : null;
  const monthHours = monthKey !== "all" ? totalHoursInMonth(monthKey) : 24 * 30; // fallback 30d

  // Build site map: siteId -> siteName (first seen)
  const siteNameMap = new Map<string, string>();
  // Build site -> ticket count (for sorting fallback)
  const siteTicketCount = new Map<string, number>();
  // Build site -> down hours.
  // KEY: normalizeSiteId(row.siteId).toLowerCase() from TicketsData "Site ID" column (e.g. "RF Site 1")
  // LOOKUP: normalizeSiteId(siteId).toLowerCase() from SiteID table "Site ID" column (e.g. "RF Site 1")
  // Both sides hold the RF Site ID — join is Site ID to Site ID.
  const siteDownHours = new Map<string, number>();

  allRows.forEach((row) => {
    if (!row.siteId) return;
    if (!siteNameMap.has(row.siteId) && row.siteName) {
      siteNameMap.set(row.siteId, row.siteName);
    }
    siteTicketCount.set(row.siteId, (siteTicketCount.get(row.siteId) ?? 0) + 1);
  });

  // Helper: combine a date string (dd/mm/yyyy) and time string (HH:MM) into a Date object
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

  // Compute down hours per site using overlap between ticket outage window and selected month
  allRows.forEach((row) => {
    if (!row.siteId) return;

    // Filter: Service Impact
    if (clean(row.impact).toLowerCase() !== "service impact") return;

    // Filter: Managed Resource = "Complete site" OR "Link Down"
    const mr = clean(row.managedResource).toLowerCase();
    if (mr !== "complete site" && mr !== "link down") return;

    // Determine the ticket's outage start and end datetimes
    const outageStart = combineDatetime(row.observationDate, row.observationTime);
    if (!outageStart) return; // can't calculate without a start

    // outageEnd: use recovery date/time if available, otherwise treat as still open (use month end or now)
    let outageEnd: Date | null = combineDatetime(row.recoveryDate, row.recoveryTime);

    if (monthKey === "all") {
      // For "all" months: use the stored duration value directly (no clipping needed)
      const hours = parseDurationHours(row.duration) ?? 0;
      const key = normalizeSiteId(clean(row.siteId)).toLowerCase();
      siteDownHours.set(key, Math.round(((siteDownHours.get(key) ?? 0) + hours) * 10) / 10);
      return;
    }

    // For a specific month: clip the outage window to [monthStart, monthEnd]
    const monthRange = selectedMonthRange(monthKey);
    if (!monthRange) return;
    const { start: monthStart, end: monthEnd } = monthRange;

    // If ticket has no recovery date (still open), treat end as the end of the selected month
    if (!outageEnd) {
      outageEnd = monthEnd;
    }

    // Skip if outage ended before the month started, or started after the month ended
    if (outageEnd <= monthStart || outageStart > monthEnd) return;

    // Clip to month boundaries
    const effectiveStart = outageStart < monthStart ? monthStart : outageStart;
    const effectiveEnd = outageEnd > monthEnd ? monthEnd : outageEnd;

    const overlapMs = effectiveEnd.getTime() - effectiveStart.getTime();
    if (overlapMs <= 0) return;

    const hours = Math.round((overlapMs / (1000 * 60 * 60)) * 10) / 10;
    const key = normalizeSiteId(clean(row.siteId)).toLowerCase();
    siteDownHours.set(key, Math.round(((siteDownHours.get(key) ?? 0) + hours) * 10) / 10);
  });

  // Build site list from SiteID sheet order (only sites in that sheet).
  // If no SiteID sheet, fall back to all sites sorted by ticket count descending.
  let siteEntries: { siteId: string; siteName: string }[];
  if (siteOrder.length > 0) {
    siteEntries = siteOrder;
  } else {
    const allSiteIds = Array.from(siteNameMap.keys()).filter((id) => id !== "");
    allSiteIds.sort((a, b) => (siteTicketCount.get(b) ?? 0) - (siteTicketCount.get(a) ?? 0));
    siteEntries = allSiteIds.map((id) => ({ siteId: id, siteName: siteNameMap.get(id) ?? "" }));
  }

  return siteEntries.map(({ siteId, siteName }) => {
    // Join: Site ID from SiteID table (normalized, lowercase) matches Site ID from TicketsData.
    const key = normalizeSiteId(clean(siteId)).toLowerCase();
    const downHours = siteDownHours.get(key) ?? 0;
    const availHours = Math.max(0, monthHours - downHours);
    const totalHours = availHours + downHours;
    const reliability = totalHours > 0 ? availHours / totalHours : 1;

    // Format availability as "X d, Y h, Z m"
    const totalMins = Math.round(availHours * 60);
    const dDays = Math.floor(totalMins / (60 * 24));
    const dHrs = Math.floor((totalMins % (60 * 24)) / 60);
    const dMins = Math.round(totalMins % 60);
    const availDay = `${dDays} d, ${dHrs} h, ${dMins} m`;

    return {
      siteId,
      siteName,
      sitesDownHours: downHours,
      availHours: Math.round(availHours * 10) / 10,
      availDay,
      reliability: `${(reliability * 100).toFixed(2)}%`,
      channelBusy: 0,
      mwLinkPerf: 0,
      ticketCount: siteTicketCount.get(siteId) ?? 0,
    };
  });
}


function perfReportRows(rows: PerfRow[]): string[][] {
  // Column order: S No, Site ID, Site Name, Site Availability Hrs, Site Availability days,
  //               Channel Busy Count, MW link Performance Hrs, DMR Reliability, Sites Down hrs
  return rows.map((r, i) => [
    String(i + 1),
    r.siteId,
    r.siteName,
    String(r.availHours),
    r.availDay,
    "",  // Channel Busy Count placeholder
    "",  // MW link Performance placeholder
    r.reliability,
    String(r.sitesDownHours),
  ]);
}

/**
 * Compute the 4 KPI metrics from perfRows.
 * % Availability = SUM(availHours) / (SUM(availHours) + SUM(downHours))
 * MTTR           = SUM(downHours) / COUNT(sites where downHours > 0)
 * MTBF           = SUM(availHours + downHours) / SUM(downHours)
 * MTTF           = MTBF + MTTR
 */
function computePerfKPIs(rows: PerfRow[]): {
  pctAvailability: string;
  mttr: string;
  mtbf: string;
  mttf: string;
  totalDown: number;
  totalAvail: number;
  affectedSites: number;
  totalDownHrs: string;
} {
  const totalAvail = rows.reduce((s, r) => s + r.availHours, 0);
  const totalDown  = rows.reduce((s, r) => s + r.sitesDownHours, 0);
  const sitesWithDown = rows.filter((r) => r.sitesDownHours > 0).length;
  const totalHrs = totalAvail + totalDown;

  const pctAvailability = totalHrs > 0
    ? `${((totalAvail / totalHrs) * 100).toFixed(2)}%`
    : "";
  const mttr = sitesWithDown > 0
    ? `${(totalDown / sitesWithDown).toFixed(2)} hrs`
    : "";
  const mtbf = totalDown > 0
    ? `${(totalHrs / totalDown).toFixed(2)} hrs`
    : "";
  const mtbfNum = totalDown > 0 ? totalHrs / totalDown : null;
  const mttrNum = sitesWithDown > 0 ? totalDown / sitesWithDown : null;
  const mttf = mtbfNum !== null && mttrNum !== null
    ? `${(mtbfNum + mttrNum).toFixed(2)} hrs`
    : "";

  // Format total down hours as "X d, Y h, Z m"
  const totalDownRounded = Math.round(totalDown * 10) / 10;
  const tdMins = Math.round(totalDown * 60);
  const tdDays = Math.floor(tdMins / (60 * 24));
  const tdHrs = Math.floor((tdMins % (60 * 24)) / 60);
  const tdMin = Math.round(tdMins % 60);
  const totalDownHrs = `${tdDays}d ${tdHrs}h ${tdMin}m`;

  return { pctAvailability, mttr, mtbf, mttf, totalDown: totalDownRounded, totalAvail: Math.round(totalAvail * 10) / 10, affectedSites: sitesWithDown, totalDownHrs };
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
  const csv = [PERF_REPORT_HEADERS, ...perfReportRows(rows), ...kpiRows]
    .map((line) => line.map((cell) => `"${String(cell ?? "").replace(/"/g, '""')}"`).join(","))
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
  const worksheet = XLSX.utils.aoa_to_sheet([PERF_REPORT_HEADERS, ...perfReportRows(rows), ...kpiRows]);
  worksheet["!cols"] = [
    { wch: 6 },  // S No
    { wch: 14 }, // Site ID
    { wch: 30 }, // Site Name
    { wch: 24 }, // Site Availability, Hrs
    { wch: 26 }, // Site Availability, days
    { wch: 20 }, // Channel Busy Count
    { wch: 26 }, // MW link Performance, Hrs
    { wch: 18 }, // DMR Reliability
    { wch: 18 }, // Sites Down, hrs
  ];
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Monthly Performance");
  XLSX.writeFile(workbook, `DMR-Monthly-Performance-${monthLabel}.xlsx`);
}

function exportPerfPdf(rows: PerfRow[], monthKey: string) {
  const doc = new jsPDF({ orientation: "landscape", unit: "mm", format: "a3" });
  const monthLabel = monthKey !== "all" ? formatMonthMMMMYYYY(monthKey) : "All Months";
  doc.setFontSize(14);
  doc.setTextColor(10, 30, 60);
  doc.text(`DMR Monthly Performance Report -- ${monthLabel}`, 14, 16);
  doc.setFontSize(9);
  doc.setTextColor(100);
  doc.text(`Generated: ${new Date().toLocaleString()} | Total Sites: ${rows.length}`, 14, 22);

  const kpi = computePerfKPIs(rows);
  const kpiSummaryY = (doc as jsPDF & { lastAutoTable: { finalY: number } }).lastAutoTable?.finalY;
  autoTable(doc, {
    startY: 26,
    head: [PERF_REPORT_HEADERS],
    body: perfReportRows(rows),
    styles: { fontSize: 8, cellPadding: 2, overflow: "linebreak", valign: "middle" },
    headStyles: { fillColor: [4, 60, 100], textColor: 255, fontStyle: "bold", fontSize: 8.5 },
    alternateRowStyles: { fillColor: [240, 246, 255] },
    columnStyles: {
      0: { cellWidth: 10 },  // S No
      1: { cellWidth: 22 },  // Site ID
      2: { cellWidth: 38 },  // Site Name
      3: { cellWidth: 28 },  // Site Availability, Hrs
      4: { cellWidth: 36 },  // Site Availability, days
      5: { cellWidth: 24 },  // Channel Busy Count
      6: { cellWidth: 32 },  // MW link Performance, Hrs
      7: { cellWidth: 22 },  // DMR Reliability
      8: { cellWidth: 22 },  // Sites Down, hrs
    },
    margin: { left: 10, right: 10 },
    didDrawPage: (data) => {
      const pageCount = (doc as jsPDF & { internal: { getNumberOfPages: () => number } }).internal.getNumberOfPages();
      doc.setFontSize(7);
      doc.setTextColor(150);
      doc.text(`Page ${data.pageNumber} of ${pageCount}`, doc.internal.pageSize.getWidth() - 30, doc.internal.pageSize.getHeight() - 6);
    },
  });

  const finalY = (doc as jsPDF & { lastAutoTable: { finalY: number } }).lastAutoTable?.finalY ?? 26;
  void kpiSummaryY; // used above
  const kpiStartY = finalY + 8;
  doc.setFontSize(10);
  doc.setTextColor(10, 30, 60);
  doc.text("KPI Summary", 14, kpiStartY);
  const kpiTableData = [
    ["% Availability", kpi.pctAvailability],
    ["MTTR", kpi.mttr],
    ["MTBF", kpi.mtbf],
    ["MTTF", kpi.mttf],
    ["No. of Affected Sites", String(kpi.affectedSites)],
    ["Total Down Duration", kpi.totalDownHrs],
  ];
  autoTable(doc, {
    startY: kpiStartY + 4,
    head: [["Metric", "Value"]],
    body: kpiTableData,
    styles: { fontSize: 9, cellPadding: 2 },
    headStyles: { fillColor: [4, 60, 100], textColor: 255, fontStyle: "bold" },
    columnStyles: { 0: { cellWidth: 40 }, 1: { cellWidth: 40 } },
    margin: { left: 14 },
    tableWidth: "wrap",
  });

  doc.save(`DMR-Monthly-Performance-${monthLabel.replace(" ", "-")}.pdf`);
}

function uniqueTicketValues(ticket: TicketAggregate, field: keyof TicketRecord): string {
  return Array.from(new Set(ticket.rows.map((row) => clean(row[field])).filter(Boolean))).join(", ");
}

function distinctReportRow(ticket: TicketAggregate, index: number): string[] {
  const row = ticket.primary;
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
    row.actionTaken || "",
  ];
}

function distinctReportRows(rows: TicketAggregate[]): string[][] {
  return rows.map((ticket, index) => distinctReportRow(ticket, index));
}

function exportCsv(rows: TicketAggregate[]) {
  const csv = [DISTINCT_REPORT_HEADERS, ...distinctReportRows(rows)]
    .map((line) => line.map((cell) => `"${String(cell ?? "").replace(/"/g, '""')}"`).join(","))
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
  const worksheet = XLSX.utils.aoa_to_sheet([DISTINCT_REPORT_HEADERS, ...distinctReportRows(rows)]);
  worksheet["!cols"] = [
    { wch: 6 }, { wch: 24 }, { wch: 28 }, { wch: 30 }, { wch: 12 }, { wch: 32 },
    { wch: 16 }, { wch: 16 }, { wch: 16 }, { wch: 16 }, { wch: 26 }, { wch: 26 },
    { wch: 24 }, { wch: 18 }, { wch: 14 }, { wch: 20 }, { wch: 34 },
  ];
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Distinct TT Report");
  XLSX.writeFile(workbook, "follow-up-distinct-tt-report.xlsx");
}

function exportPdf(rows: TicketAggregate[], monthKey: string) {
  const doc = new jsPDF({ orientation: "landscape", unit: "mm", format: "a3" });
  const monthLabel = monthKey !== "all" ? formatMonthMMMMYYYY(monthKey) : "All Months";
  doc.setFontSize(14);
  doc.setTextColor(10, 30, 60);
  doc.text(`DMR Monthly Tickets Report -- ${monthLabel}`, 14, 16);
  doc.setFontSize(9);
  doc.setTextColor(100);
  doc.text(`Generated: ${new Date().toLocaleString()} | Total Tickets: ${rows.length}`, 14, 22);

  autoTable(doc, {
    startY: 26,
    head: [DISTINCT_REPORT_HEADERS],
    body: distinctReportRows(rows),
    styles: { fontSize: 7, cellPadding: 1.5, overflow: "linebreak", valign: "middle" },
    headStyles: { fillColor: [4, 60, 100], textColor: 255, fontStyle: "bold", fontSize: 7.5 },
    alternateRowStyles: { fillColor: [240, 246, 255] },
    columnStyles: {
      0:  { cellWidth: 8 },
      1:  { cellWidth: 18 },
      2:  { cellWidth: 22 },
      3:  { cellWidth: 26 },
      4:  { cellWidth: 14 },
      5:  { cellWidth: 30 },
      6:  { cellWidth: 18 },
      7:  { cellWidth: 16 },
      8:  { cellWidth: 18 },
      9:  { cellWidth: 16 },
      10: { cellWidth: 22 },
      11: { cellWidth: 20 },
      12: { cellWidth: 22 },
      13: { cellWidth: 14 },
      14: { cellWidth: 14 },
      15: { cellWidth: 18 },
      16: { cellWidth: 36 },
    },
    margin: { left: 10, right: 10 },
    didDrawPage: (data) => {
      const pageCount = (doc as jsPDF & { internal: { getNumberOfPages: () => number } }).internal.getNumberOfPages();
      doc.setFontSize(7);
      doc.setTextColor(150);
      doc.text(`Page ${data.pageNumber} of ${pageCount}`, doc.internal.pageSize.getWidth() - 30, doc.internal.pageSize.getHeight() - 6);
    },
  });

  const suffix = monthKey !== "all" ? `-${monthLabel.replace(" ", "-")}` : "";
  doc.save(`DMR-Monthly-Report${suffix}.pdf`);
}

function StatCard({ label, value, note, icon: Icon, tone, onClick }: { label: string; value: string | number; note: string; icon: typeof Activity; tone: string; onClick?: () => void }) {
  return (
    <div className="stat-card" style={{ ["--tone" as string]: tone, cursor: onClick ? "pointer" : undefined }} onClick={onClick} role={onClick ? "button" : undefined} tabIndex={onClick ? 0 : undefined} onKeyDown={onClick ? (e) => { if (e.key === "Enter" || e.key === " ") onClick(); } : undefined}>
      <div className="stat-icon"><Icon size={18} /></div>
      <span>{label}</span>
      <strong>{value}</strong>
      <small>{note}</small>
    </div>
  );
}

function PartnerLogoStrip() {
  return (
    <div className="partner-logo-strip" aria-label="Project partner logos">
      <img src={ngLogoSrc} alt="National Grid SA" className="partner-logo-img ng-logo" />
      <span className="logo-divider" aria-hidden="true" />
      <img src={nascoLogoSrc} alt="NASCO" className="partner-logo-img nasco-logo" />
    </div>
  );
}

/**
 * SelectFilter using a portal-based dropdown so it is never clipped
 * by overflow:hidden containers (hero panel, export card).
 */
function SelectFilter({ label, value, options, optionLabels, onChange }: {
  label: string;
  value: string;
  options: string[];
  optionLabels?: Record<string, string>;
  onChange: (value: string) => void;
}) {
  const [open, setOpen] = useState(false);
  const [dropdownStyle, setDropdownStyle] = useState<React.CSSProperties>({});
  const wrapRef = useRef<HTMLDivElement | null>(null);
  const triggerRef = useRef<HTMLButtonElement | null>(null);
  const dropdownRef = useRef<HTMLDivElement | null>(null);

  useEffect(() => {
    function handleClickOutside(e: MouseEvent) {
      const target = e.target as Node;
      if (!wrapRef.current?.contains(target) && !dropdownRef.current?.contains(target)) setOpen(false);
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
      setDropdownStyle({ position: "fixed", top: rect.bottom + 4, left, width: dropW, zIndex: 99999 });
    } else {
      setDropdownStyle({ position: "fixed", bottom: window.innerHeight - rect.top + 4, left, width: dropW, zIndex: 99999 });
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

  const displayLabel = value === "all" || value === "ALL"
    ? "All"
    : (optionLabels?.[value] ?? value);

  const dropdown = open ? createPortal(
    <div
      ref={dropdownRef}
      className="multi-select-dropdown"
      style={dropdownStyle}
    >
      <label className="multi-select-option" onClick={() => { onChange("all"); setOpen(false); }}>
        <input type="radio" readOnly checked={value === "all" || value === "ALL"} />
        <span>All</span>
      </label>
      {options.map((opt) => (
        <label key={opt} className="multi-select-option" onClick={() => { onChange(opt); setOpen(false); }}>
          <input type="radio" readOnly checked={value === opt} />
          <span>{optionLabels?.[opt] ?? opt}</span>
        </label>
      ))}
    </div>,
    document.body
  ) : null;

  return (
    <div className="filter-field multi-select-filter" ref={wrapRef}>
      <span>{label}</span>
      <button type="button" className="multi-select-trigger" ref={triggerRef} onClick={handleOpen}>
        <span className="multi-select-value">{displayLabel}</span>
        <ChevronDown size={14} style={{ transform: open ? "rotate(180deg)" : undefined, transition: "transform .15s" }} />
      </button>
      {dropdown}
    </div>
  );
}

function MultiSelectFilter({ label, value, options, optionLabels, onChange, showAllOption }: {
  label: string;
  value: string[];
  options: string[];
  optionLabels?: Record<string, string>;
  onChange: (value: string[]) => void;
  showAllOption?: boolean;
}) {
  const [open, setOpen] = useState(false);
  const [dropdownStyle, setDropdownStyle] = useState<React.CSSProperties>({});
  const wrapRef = useRef<HTMLDivElement | null>(null);
  const triggerRef = useRef<HTMLButtonElement | null>(null);
  const dropdownRef = useRef<HTMLDivElement | null>(null);

  useEffect(() => {
    function handleClickOutside(e: MouseEvent) {
      const target = e.target as Node;
      const insideTrigger = wrapRef.current?.contains(target);
      const insideDropdown = dropdownRef.current?.contains(target);
      if (!insideTrigger && !insideDropdown) setOpen(false);
    }
    document.addEventListener("mousedown", handleClickOutside);
    return () => document.removeEventListener("mousedown", handleClickOutside);
  }, []);

  function computePosition() {
    if (!triggerRef.current) return;
    const rect = triggerRef.current.getBoundingClientRect();
    const dropW = Math.max(rect.width, 220);
    const dropH = Math.min(280, options.length * 36 + 48);
    const spaceBelow = window.innerHeight - rect.bottom;
    const left = Math.min(rect.left, window.innerWidth - dropW - 8);
    if (spaceBelow >= dropH || spaceBelow >= 160) {
      setDropdownStyle({ position: "fixed", top: rect.bottom + 4, left, width: dropW, zIndex: 99999 });
    } else {
      setDropdownStyle({ position: "fixed", bottom: window.innerHeight - rect.top + 4, left, width: dropW, zIndex: 99999 });
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
    // If All is currently active, selecting an option means: pick just that option
    if (allSelected) { onChange([opt]); return; }
    if (value.includes(opt)) {
      const next = value.filter((v) => v !== opt);
      // If deselecting last item, revert to All
      onChange(next);
    } else {
      onChange([...value, opt]);
    }
  };
  const displayLabel = allSelected ? "All" : value.length === 1 ? (optionLabels?.[value[0]] ?? value[0]) : `${value.length} selected`;

  const dropdown = open ? createPortal(
    <div className="multi-select-dropdown" style={dropdownStyle} ref={dropdownRef}>
      {showAllOption && (
        <label className="multi-select-option multi-select-option-all" onClick={() => { onChange([]); }}>
          <input type="checkbox" readOnly checked={allSelected} />
          <span style={{ fontWeight: allSelected ? 700 : undefined, color: allSelected ? "#22d3ee" : undefined }}>All</span>
        </label>
      )}
      {!showAllOption && value.length > 0 && (
        <button type="button" className="multi-select-clear" onClick={() => onChange([])}>✕ Clear</button>
      )}
      {options.map((opt) => (
        <label key={opt} className="multi-select-option">
          <input type="checkbox" checked={value.includes(opt)} onChange={() => toggle(opt)} />
          <span>{optionLabels?.[opt] ?? opt}</span>
        </label>
      ))}
    </div>,
    document.body
  ) : null;

  return (
    <div className="filter-field multi-select-filter" ref={wrapRef}>
      <span>{label}</span>
      <button type="button" className="multi-select-trigger" ref={triggerRef} onClick={handleOpen}>
        <span className="multi-select-value">{displayLabel}</span>
        <ChevronDown size={14} style={{ transform: open ? "rotate(180deg)" : undefined, transition: "transform .15s" }} />
      </button>
      {dropdown}
    </div>
  );
}

export default function Home() {
  const inputRef = useRef<HTMLInputElement | null>(null);
  const addRegionRef = useRef<HTMLInputElement | null>(null);
  const [data, setData] = useState<DashboardData | null>(null);
  const [regions, setRegions] = useState<DashboardData[]>([]);
  const [error, setError] = useState("");
  const [isDragging, setIsDragging] = useState(false);

  const [exportMonths, setExportMonths] = useState<string[]>([]);
  const [exportRegions, setExportRegions] = useState<string[]>([]);
  const [filters, setFilters] = useState<Filters>(EMPTY_FILTERS);
  const [tablePage, setTablePage] = useState(1);
  const TABLE_PAGE_SIZE = 50;

  const [perfMonths, setPerfMonths] = useState<string[]>([]);
  const [perfRegions, setPerfRegions] = useState<string[]>([]);

  async function handleFile(file?: File) {
    if (!file) return;
    setError("");
    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
      const parsed = parseRows(workbook, file.name);
      if (!parsed.rows.length || !parsed.uniqueTickets.length) {
        throw new Error("No ticket rows with TT numbers were found. Please upload the Follow-Up Sheets workbook with the Tickets_Data sheet.");
      }
      setData(parsed);
      setRegions([parsed]);
      setExportMonths([]);
      setExportRegions([]);
      setPerfMonths([]);
      setPerfRegions([]);
      setFilters(EMPTY_FILTERS);
      setTablePage(1);
      if (inputRef.current) inputRef.current.value = "";
    } catch (err) {
      setError(err instanceof Error ? err.message : "Unable to read this workbook.");
    }
  }

  async function handleAddRegion(file?: File) {
    if (!file) return;
    setError("");
    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
      const parsed = parseRows(workbook, file.name);
      if (!parsed.rows.length || !parsed.uniqueTickets.length) {
        throw new Error("No ticket rows with TT numbers were found in the additional workbook.");
      }
      setRegions((prev) => {
        const updated = [...prev, parsed];
        const allRows = updated.flatMap((r) => r.rows);
        const ttMap = new Map<string, TicketAggregate>();
        for (const row of allRows) {
          if (!row.tt) continue;
          const existing = ttMap.get(row.tt);
          if (!existing) {
            ttMap.set(row.tt, { tt: row.tt, primary: row, siteIds: new Set([row.siteId].filter(Boolean)), siteNames: new Set([row.siteName].filter(Boolean)), rows: [row] });
          } else {
            existing.rows.push(row);
            if (row.siteId) existing.siteIds.add(row.siteId);
            if (row.siteName) existing.siteNames.add(row.siteName);
          }
        }
        // Merge site orders: use first region's order, append any new entries from additional regions
        const mergedSiteOrder = [...updated[0].siteOrder];
        const mergedIds = new Set(mergedSiteOrder.map((s) => s.siteId));
        updated.slice(1).forEach((r) => {
          r.siteOrder.forEach((s) => { if (!mergedIds.has(s.siteId)) { mergedIds.add(s.siteId); mergedSiteOrder.push(s); } });
        });
        const merged: DashboardData = {
          fileName: updated.map((r) => r.fileName).join(" + "),
          sheetName: updated[0].sheetName,
          generatedAt: new Date().toLocaleString(),
          rows: allRows,
          uniqueTickets: Array.from(ttMap.values()),
          siteOrder: mergedSiteOrder,
        };
        setData(merged);
        return updated;
      });
      setTablePage(1);
      if (addRegionRef.current) addRegionRef.current.value = "";
    } catch (err) {
      setError(err instanceof Error ? err.message : "Unable to read the additional workbook.");
    }
  }

  const uniqueRows = data?.uniqueTickets ?? [];
  const allDataRows = data?.rows ?? [];

  const filterOptions = useMemo(() => {
    const primaryRows = uniqueRows.map((ticket) => ticket.primary);
    const uniq = (field: keyof TicketRecord) => Array.from(new Set(primaryRows.map((row) => clean(row[field])).filter(Boolean))).sort();
    const openingMonths = Array.from(new Set(primaryRows.map((row) => row.openingMonthKey || openingMonthKey(row.observationDate)).filter(Boolean))).sort((a, b) => {
      if (a === "Unknown") return 1;
      if (b === "Unknown") return -1;
      return a.localeCompare(b);
    });
    const exportMonths = Array.from(new Set(uniqueRows.flatMap((ticket) =>
      ticket.rows.flatMap((row) => coveredMonthKeys(row)),
    ).filter((key) => key && key !== "Unknown"))).sort((a, b) => a.localeCompare(b));
    const rcaFamilyOptions = Array.from(new Set(primaryRows.map((row) => row.rcaFamily || getRcaFamily(row.rca)).filter(Boolean))).sort() as string[];
    return {
      status: uniq("status"),
      severity: uniq("severity"),
      region: uniq("region"),
      impact: uniq("impact"),
      site: uniq("siteId"),
      openingMonth: openingMonths,
      openingMonthLabels: Object.fromEntries(openingMonths.map((key) => [key, openingMonthLabel(key)])),
      exportMonth: exportMonths,
      exportMonthLabels: Object.fromEntries(exportMonths.map((key) => [key, openingMonthLabel(key)])),
      rcaFamily: rcaFamilyOptions,
    };
  }, [uniqueRows]);

  // Reset page when filters change
  useEffect(() => { setTablePage(1); }, [filters]);

  const filteredTickets = useMemo(() => {
    const q = filters.search.toLowerCase().trim();
    return uniqueRows.filter((ticket) => {
      const row = ticket.primary;
      const allSites = Array.from(ticket.siteIds).join(" ");
      const allSiteNames = Array.from(ticket.siteNames).join(" ");
      const haystack = [
        ticket.tt, row.siteId, row.siteName, allSites, allSiteNames,
        row.managedResource, row.issue, row.status, row.severity, row.region,
        row.impact, row.escalationLevel, row.escalatedTo, row.rca,
        row.rcaFamily || getRcaFamily(row.rca),
        row.responsibleTeam || getResponsibleTeam(row.rcaFamily || getRcaFamily(row.rca)),
        row.escalatedForL3SupportDate, row.escalatedForL3SupportTime,
      ].join(" ").toLowerCase();
      return (
        (!q || haystack.includes(q)) &&
        (!filters.status.length || filters.status.includes(row.status)) &&
        (!filters.severity.length || filters.severity.includes(row.severity)) &&
        (!filters.region.length || filters.region.includes(row.region)) &&
        (!filters.impact.length || filters.impact.includes(row.impact)) &&
        (!filters.openingMonth.length || filters.openingMonth.includes(row.openingMonthKey || openingMonthKey(row.observationDate))) &&
        (!filters.site.length || filters.site.some((s) => ticket.siteIds.has(s))) &&
        (!filters.rcaFamily.length || filters.rcaFamily.includes(row.rcaFamily || getRcaFamily(row.rca)))
      );
    });
  }, [filters, uniqueRows]);

  const monthlyExportBaseTickets = useMemo(() => {
    const q = filters.search.toLowerCase().trim();
    return uniqueRows.filter((ticket) => {
      const row = ticket.primary;
      const allSites = Array.from(ticket.siteIds).join(" ");
      const allSiteNames = Array.from(ticket.siteNames).join(" ");
      const haystack = [
        ticket.tt, row.siteId, row.siteName, allSites, allSiteNames,
        row.managedResource, row.issue, row.status, row.severity, row.region,
        row.impact, row.escalationLevel, row.escalatedTo, row.rca,
        row.rcaFamily || getRcaFamily(row.rca),
        row.responsibleTeam || getResponsibleTeam(row.rcaFamily || getRcaFamily(row.rca)),
        row.escalatedForL3SupportDate, row.escalatedForL3SupportTime,
      ].join(" ").toLowerCase();
      return (
        (!q || haystack.includes(q)) &&
        (!filters.severity.length || filters.severity.includes(row.severity)) &&
        (!filters.region.length || filters.region.includes(row.region)) &&
        (!filters.impact.length || filters.impact.includes(row.impact)) &&
        (!filters.site.length || filters.site.some((s) => ticket.siteIds.has(s))) &&
        (!filters.rcaFamily.length || filters.rcaFamily.includes(row.rcaFamily || getRcaFamily(row.rca)))
      );
    });
  }, [filters.impact, filters.rcaFamily, filters.region, filters.search, filters.severity, filters.site, uniqueRows]);

  const monthlyExportTickets = useMemo(() => {
    // Apply export region filter
    const regionFiltered = exportRegions.length === 0
      ? monthlyExportBaseTickets
      : monthlyExportBaseTickets.filter((ticket) => exportRegions.includes(ticket.primary.region));
    // Apply export month filter
    if (exportMonths.length === 0) return regionFiltered;
    return regionFiltered.filter((ticket) => exportMonths.some((m) => ticketMatchesMonthlyExport(ticket, m)));
  }, [exportMonths, exportRegions, monthlyExportBaseTickets]);

  const selectedExportMonthLabel = exportMonths.length === 0 ? "All export-eligible TT" : exportMonths.length === 1 ? openingMonthLabel(exportMonths[0]) : `${exportMonths.length} months`;

  // Monthly Performance rows -- computed from raw rows (not aggregated tickets)
  const siteOrder = data?.siteOrder ?? [];
  const perfRows = useMemo(() => {
    // Filter by region
    const sourceRows = perfRegions.length === 0
      ? allDataRows
      : allDataRows.filter((r) => perfRegions.includes(r.region));
    // For multi-month: pass "all" if none selected, or the first month if one selected,
    // or compute combined rows for multiple months
    if (perfMonths.length === 0) {
      return computePerfRows(sourceRows, "all", siteOrder);
    } else if (perfMonths.length === 1) {
      return computePerfRows(sourceRows, perfMonths[0], siteOrder);
    } else {
      // Sum down hours across all selected months per site
      const combined = new Map<string, PerfRow>();
      perfMonths.forEach((mk) => {
        const rows = computePerfRows(sourceRows, mk, siteOrder);
        rows.forEach((r) => {
          const existing = combined.get(r.siteId);
          if (!existing) {
            combined.set(r.siteId, { ...r });
          } else {
            existing.sitesDownHours = Math.round((existing.sitesDownHours + r.sitesDownHours) * 10) / 10;
            // availHours: recalculate based on total month hours across selected months
          }
        });
      });
      // Recalculate availHours for combined rows
      const totalMonthHours = perfMonths.reduce((s, mk) => s + totalHoursInMonth(mk), 0);
      return Array.from(combined.values()).map((r) => {
        const availHours = Math.max(0, totalMonthHours - r.sitesDownHours);
        const totalHrs = availHours + r.sitesDownHours;
        const reliability = totalHrs > 0 ? availHours / totalHrs : 1;
        const totalMins = Math.round(availHours * 60);
        const dDays = Math.floor(totalMins / (60 * 24));
        const dHrs = Math.floor((totalMins % (60 * 24)) / 60);
        const dMins = Math.round(totalMins % 60);
        return {
          ...r,
          availHours: Math.round(availHours * 10) / 10,
          availDay: `${dDays} d, ${dHrs} h, ${dMins} m`,
          reliability: `${(reliability * 100).toFixed(2)}%`,
        };
      });
    }
  }, [allDataRows, perfMonths, perfRegions, siteOrder]);

  const analytics = useMemo(() => {
    const primaryRows = filteredTickets.map((ticket) => ticket.primary);
    const totalUnique = filteredTickets.length;
    const status = countBy(primaryRows, (row) => row.status);
    const severity = countBy(primaryRows, (row) => row.severity);
    const region = countBy(primaryRows, (row) => row.region);
    const impact = countBy(primaryRows, (row) => row.impact);
    const escalation = countBy(primaryRows, (row) => row.escalationLevel);
    const monthlyMap = new Map<string, { name: string; value: number }>();
    primaryRows.forEach((row) => {
      const key = row.openingMonthKey || openingMonthKey(row.observationDate);
      const current = monthlyMap.get(key) ?? { name: openingMonthLabel(key), value: 0 };
      current.value += 1;
      monthlyMap.set(key, current);
    });
    const monthly = Array.from(monthlyMap.entries())
      .sort(([a], [b]) => {
        if (a === "Unknown") return 1;
        if (b === "Unknown") return -1;
        return a.localeCompare(b);
      })
      .map(([, value]) => value);
    const avgHoursSource = primaryRows.map((row) => parseDurationHours(row.duration)).filter((value): value is number => value !== null);
    const avgHours = avgHoursSource.length ? avgHoursSource.reduce((sum, value) => sum + value, 0) / avgHoursSource.length : 0;
    const uniqueSites = new Set(primaryRows.map((row) => row.siteId).filter(Boolean)).size;
    const rootCauseUpdated = primaryRows.filter((row) => row.actionTaken || !rcaNotProvided(row.rca)).length;
    const totalSiteAffected = filteredTickets.reduce((sum, ticket) => sum + Math.max(ticket.siteIds.size, ticket.primary.siteId ? 1 : 0), 0);
    const rcaByCount = countBy(primaryRows, (row) => rcaNotProvided(row.rca) ? "RCA not Provided" : row.rca);
    const topRcaByCount = rcaByCount[0] ?? { name: "", value: 0 };
    const downtimeByRcaMap = new Map<string, { name: string; value: number; count: number }>();
    primaryRows.forEach((row) => {
      const name = rcaNotProvided(row.rca) ? "RCA not Provided" : row.rca;
      const hours = parseDurationHours(row.duration) ?? 0;
      const current = downtimeByRcaMap.get(name) ?? { name, value: 0, count: 0 };
      current.value += hours;
      if (hours) current.count += 1;
      downtimeByRcaMap.set(name, current);
    });
    const downtimeByRca = Array.from(downtimeByRcaMap.values()).sort((a, b) => b.value - a.value || a.name.localeCompare(b.name));
    const mttrByRca = Array.from(downtimeByRcaMap.values())
      .map((item) => ({ name: item.name, value: item.count ? item.value / item.count : 0, count: item.count }))
      .sort((a, b) => b.value - a.value || a.name.localeCompare(b.name));
    const repeatedSiteRcaMap = new Map<string, number>();
    primaryRows.forEach((row) => {
      const site = row.siteId || "Blank";
      const rca = rcaNotProvided(row.rca) ? "RCA not Provided" : row.rca;
      const key = `${site}||${rca}`;
      repeatedSiteRcaMap.set(key, (repeatedSiteRcaMap.get(key) ?? 0) + 1);
    });
    const repeatedRcaSites = Array.from(repeatedSiteRcaMap.values()).filter((value) => value > 1).length;
    const rcaNotProvidedCount = primaryRows.filter((row) => rcaNotProvided(row.rca)).length;
    const preventableCount = primaryRows.filter((row) => (row.preventability || getPreventability(row.rca)) === "Preventable").length;
    const rcaFamily = countBy(primaryRows, (row) => row.rcaFamily || getRcaFamily(row.rca));

    const siteNameById = new Map<string, string>();
    filteredTickets.forEach((ticket) => {
      ticket.rows.forEach((row) => {
        if (row.siteId && row.siteName && !siteNameById.has(row.siteId)) siteNameById.set(row.siteId, row.siteName);
      });
    });
    const siteLabel = (siteId: string) => {
      const siteName = siteNameById.get(siteId);
      if (!siteId || siteId === "Blank") return "Blank";
      return siteName ? `${siteId} -- ${siteName}` : siteId;
    };
    const siteMap = new Map<string, { name: string; value: number; exposure?: number }>();
    filteredTickets.forEach((ticket) => {
      const sites = ticket.siteIds.size ? Array.from(ticket.siteIds) : [ticket.primary.siteId || "Blank"];
      sites.forEach((site) => {
        const current = siteMap.get(site) ?? { name: siteLabel(site), value: 0 };
        current.value += 1;
        siteMap.set(site, current);
      });
    });
    const topSites = Array.from(siteMap.values()).sort((a, b) => b.value - a.value || a.name.localeCompare(b.name)).slice(0, 12);

    const preventableBreakdown = [
      { name: "Preventable", value: preventableCount },
      { name: "Non-preventable", value: primaryRows.length - preventableCount },
    ].filter((item) => item.value > 0);

    const RCA_FAMILIES = ["Power & Environment", "Fiber & Physical", "Transmission & Link", "Hardware & Device", "Configuration / Software", "Human / Process / Planned", "Other / Review"];
    const monthlyRcaFamilyMap = new Map<string, Record<string, string | number>>();
    primaryRows.forEach((row) => {
      const key = row.openingMonthKey || openingMonthKey(row.observationDate);
      if (!key || key === "Unknown") return;
      const family = row.rcaFamily || getRcaFamily(row.rca);
      if (!monthlyRcaFamilyMap.has(key)) {
        const entry: Record<string, string | number> = { monthKey: key, name: openingMonthLabel(key) };
        RCA_FAMILIES.forEach((f) => { entry[f] = 0; });
        monthlyRcaFamilyMap.set(key, entry);
      }
      const entry = monthlyRcaFamilyMap.get(key)!;
      entry[family] = ((entry[family] as number) || 0) + 1;
    });
    const monthlyRcaFamily = Array.from(monthlyRcaFamilyMap.entries())
      .sort(([a], [b]) => a.localeCompare(b))
      .map(([, value]) => value);

    const managedResourceByCount = countBy(primaryRows, (row) => row.managedResource || "Unknown");
    const topManagedResources = managedResourceByCount.slice(0, 12);

    return {
      totalUnique, status, severity, region, impact, escalation, monthly, avgHours,
      uniqueSites, rootCauseUpdated, totalSiteAffected, topSites, rcaByCount, rcaFamily,
      topRcaByCount,
      topRcaByDowntime: downtimeByRca[0] ?? { name: "", value: 0, count: 0 },
      highestMttrRca: mttrByRca[0] ?? { name: "", value: 0, count: 0 },
      repeatedRcaSites, rcaNotProvidedCount, preventableCount,
      rcaByDowntime: downtimeByRca.map((item) => ({ name: item.name, value: Math.round(item.value * 10) / 10 })),
      rcaByMttr: mttrByRca.map((item) => ({ name: item.name, value: Math.round(item.value * 10) / 10 })),
      preventableBreakdown, monthlyRcaFamily, rcaFamilyKeys: RCA_FAMILIES, topManagedResources,
    };
  }, [filteredTickets]);

  const closed = metricValue(analytics.status, "Closed");
  const pending = metricValue(analytics.status, "Pending");
  const resolved = metricValue(analytics.status, "Resolved");
  const critical = metricValue(analytics.severity, "Critical");
  const major = metricValue(analytics.severity, "Major");
  const minor = metricValue(analytics.severity, "Minor");
  const serviceImpact = metricValue(analytics.impact, "Service Impact");
  const nonServiceImpact = metricValue(analytics.impact, "Non-Service Impact");

  const rcaNotProvidedPct = pct(analytics.rcaNotProvidedCount, analytics.totalUnique);
  const preventableRcaPct = pct(analytics.preventableCount, analytics.totalUnique);

  return (
    <main className="dashboard-shell">
      <section
        className="hero-panel"
        style={{
          backgroundImage: `linear-gradient(90deg, rgba(3,7,18,.96) 0%, rgba(3,7,18,.78) 42%, rgba(3,7,18,.26) 100%), url(${HERO_IMAGE})`,
          backgroundSize: "contain",
          backgroundPosition: "top",
          backgroundRepeat: "no-repeat",
        }}
      >
        <nav className="topbar no-print">
          <div className="brand-cluster">
            <div className="brand-mark"><Network size={18} /> DMR Ticketing Dashboard</div>
            <PartnerLogoStrip />
          </div>
          <div className="topbar-actions">
            {data && <button className="ghost-button" onClick={() => addRegionRef.current?.click()}><UploadCloud size={16} /> Add region</button>}
            {data && <button className="ghost-button" onClick={() => inputRef.current?.click()}><RefreshCw size={16} /> New workbook</button>}
            {data && <button className="ghost-button" onClick={() => exportCsv(monthlyExportTickets)}><Download size={16} /> CSV</button>}
            {data && <button className="ghost-button" onClick={() => exportExcel(monthlyExportTickets)}><FileSpreadsheet size={16} /> Excel</button>}
            {data && <button className="ghost-button" onClick={() => exportPdf(monthlyExportTickets, exportMonths[0] ?? "all")}><Printer size={16} /> PDF</button>}
            {data && <button className="primary-button" onClick={() => window.print()}><Printer size={16} /> Dashboard PDF</button>}
          </div>
        </nav>
        <div className="hero-layout">
          <div className="hero-content">
            <div>
              <h1>DMR Ticketing Dashboard</h1>
            </div>
          </div>
        </div>
        {data && regions.length > 0 && (
          <div className="region-badges no-print">
            {regions.map((r, i) => (
              <span key={i} className="region-badge"><FileSpreadsheet size={12} /> {r.fileName}</span>
            ))}
          </div>
        )}
        {data && (
          <>
            {/* Monthly Tickets Export Card */}
            <div className="hero-export-row no-print">
              <aside className="hero-export-card hero-export-card--5col" aria-label="Monthly TT export filter">
                <div className="hero-export-copy">
                  <span>Monthly Tickets export filter</span>
                  <div className="export-count-center"><span className="export-badge-mini">{monthlyExportTickets.length}</span></div>
                  <strong>tickets &mdash; {selectedExportMonthLabel}</strong>
                </div>
                <MultiSelectFilter label="Report Month" value={exportMonths} options={filterOptions.exportMonth} optionLabels={filterOptions.exportMonthLabels} onChange={setExportMonths} showAllOption />
                <MultiSelectFilter label="Region" value={exportRegions} options={filterOptions.region} onChange={setExportRegions} showAllOption />
                <div className="hero-export-actions">
                  <button className="ghost-button" onClick={() => exportCsv(monthlyExportTickets)}><Download size={16} /> CSV</button>
                  <button className="ghost-button" onClick={() => exportExcel(monthlyExportTickets)}><FileSpreadsheet size={16} /> Excel</button>
                  <button className="ghost-button" onClick={() => exportPdf(monthlyExportTickets, exportMonths[0] ?? "all")}><Printer size={16} /> PDF</button>
                </div>
              </aside>
            </div>

            {/* Monthly Performance Export Card */}
            <div className="hero-export-row no-print">
              <aside className="hero-export-card hero-export-card--5col" aria-label="Monthly Performance filter">
                <div className="hero-export-copy">
                  <span>Monthly Performance</span>
                  <div className="export-count-center"><span className="export-badge-mini">{perfRows.filter(r => r.sitesDownHours > 0).length}</span></div>
                  <strong>
                    affected sites &mdash; {perfMonths.length === 0 ? "All Months" : perfMonths.length === 1 ? formatMonthMMMMYYYY(perfMonths[0]) : `${perfMonths.length} months`}
                    {perfRegions.length > 0 ? ` — ${perfRegions.join(", ")}` : ""}
                  </strong>
                </div>
                <MultiSelectFilter label="Report Month" value={perfMonths} options={filterOptions.exportMonth} optionLabels={filterOptions.exportMonthLabels} onChange={setPerfMonths} showAllOption />
                <MultiSelectFilter label="Region" value={perfRegions} options={filterOptions.region} onChange={setPerfRegions} showAllOption />
                <div className="hero-export-actions">
                  <button className="ghost-button" onClick={() => exportPerfCsv(perfRows, perfMonths[0] ?? "all")}><Download size={16} /> CSV</button>
                  <button className="ghost-button" onClick={() => exportPerfExcel(perfRows, perfMonths[0] ?? "all")}><FileSpreadsheet size={16} /> Excel</button>
                  <button className="ghost-button" onClick={() => exportPerfPdf(perfRows, perfMonths[0] ?? "all")}><Printer size={16} /> PDF</button>
                </div>
              </aside>
            </div>

            {/* KPI Summary tiles -- live with selected perfMonths */}
            {perfRows.length > 0 && (() => {
              const kpi = computePerfKPIs(perfRows);
              return (
                <div className="hero-export-row no-print" style={{ paddingBottom: 0 }}>
                  <div className="perf-kpi-row">
                    {([
                      { label: "% Availability", value: kpi.pctAvailability, color: "#22c55e" },
                      { label: "MTTR", value: kpi.mttr, color: "#f59e0b" },
                      { label: "MTBF", value: kpi.mtbf, color: "#3b82f6" },
                      { label: "MTTF", value: kpi.mttf, color: "#a78bfa" },
                      { label: "Affected Sites", value: String(kpi.affectedSites), color: "#f43f5e" },
                      { label: "Total Down", value: kpi.totalDownHrs, color: "#fb923c" },
                    ] as { label: string; value: string; color: string }[]).map(({ label, value, color }) => (
                      <div key={label} className="perf-kpi-tile">
                        <div className="perf-kpi-label">{label}</div>
                        <div className="perf-kpi-value" style={{ color }}>{value || "--"}</div>
                      </div>
                    ))}
                  </div>
                </div>
              );
            })()}
          </>
        )}
      </section>

      <input ref={inputRef} className="sr-only" type="file" accept=".xlsx,.xls,.xlsm" onChange={(event) => handleFile(event.target.files?.[0])} />
      <input ref={addRegionRef} className="sr-only" type="file" accept=".xlsx,.xls,.xlsm" onChange={(event) => handleAddRegion(event.target.files?.[0])} />

      {!data ? (
        <section className="upload-stage no-print">
          <div
            className={`upload-card ${isDragging ? "dragging" : ""}`}
            onDragOver={(event) => { event.preventDefault(); setIsDragging(true); }}
            onDragLeave={() => setIsDragging(false)}
            onDrop={(event) => { event.preventDefault(); setIsDragging(false); handleFile(event.dataTransfer.files?.[0]); }}
          >
            <div className="upload-copy">
              <span className="section-kicker"><UploadCloud size={14} /> Workbooks upload</span>
              <h2>Please Upload the Tickets Data.</h2>
              {error && <div className="error-banner"><AlertTriangle size={16} /> {error}</div>}
              <button className="primary-button large" onClick={() => inputRef.current?.click()}><FileSpreadsheet size={18} /> Select Excel workbook</button>
            </div>
            <img src={UPLOAD_IMAGE} alt="Abstract workbook upload visualization" />
          </div>
        </section>
      ) : (
        <>
          <section className="filters-panel no-print">
            <label className="search-box">
              <Search size={16} />
              <input value={filters.search} onChange={(event) => setFilters((prev) => ({ ...prev, search: event.target.value }))} placeholder="Search TT, site, issue, status, escalation..." />
            </label>
            <MultiSelectFilter label="Status" value={filters.status} options={filterOptions.status} onChange={(value) => setFilters((prev) => ({ ...prev, status: value }))} />
            <MultiSelectFilter label="Severity" value={filters.severity} options={filterOptions.severity} onChange={(value) => setFilters((prev) => ({ ...prev, severity: value }))} />
            <MultiSelectFilter label="Region" value={filters.region} options={filterOptions.region} onChange={(value) => setFilters((prev) => ({ ...prev, region: value }))} />
            <MultiSelectFilter label="Impact" value={filters.impact} options={filterOptions.impact} onChange={(value) => setFilters((prev) => ({ ...prev, impact: value }))} />
            <MultiSelectFilter label="Opening Month" value={filters.openingMonth} options={filterOptions.openingMonth} optionLabels={filterOptions.openingMonthLabels} onChange={(value) => setFilters((prev) => ({ ...prev, openingMonth: value }))} />
            <MultiSelectFilter label="Site" value={filters.site} options={filterOptions.site} onChange={(value) => setFilters((prev) => ({ ...prev, site: value }))} />
            <MultiSelectFilter label="RCA Family" value={filters.rcaFamily} options={filterOptions.rcaFamily} onChange={(value) => setFilters((prev) => ({ ...prev, rcaFamily: value }))} />
            <button className="ghost-button" onClick={() => { setFilters(EMPTY_FILTERS); setTablePage(1); }}><Filter size={16} /> Clear</button>
          </section>

          <section className="stats-grid workbook-cards" style={{ backgroundImage: `linear-gradient(90deg, rgba(4,13,31,.88), rgba(4,13,31,.70)), url(${RIBBON_IMAGE})` }}>
            <StatCard label="Total Tickets" value={analytics.totalUnique.toLocaleString()} note="Total Tickets Opened" icon={Layers3} tone="#22d3ee" onClick={() => { setFilters(EMPTY_FILTERS); setTablePage(1); }} />
            <StatCard label="Closed Tickets" value={closed.toLocaleString()} note={`${pct(closed, analytics.totalUnique)} closed`} icon={CheckCircle2} tone="#34d399" onClick={() => { setFilters({ ...EMPTY_FILTERS, status: ["Closed"] }); setTablePage(1); }} />
            <StatCard label="Pending Tickets" value={pending.toLocaleString()} note="Needs Follow-Up" icon={ShieldAlert} tone="#f59e0b" onClick={() => { setFilters({ ...EMPTY_FILTERS, status: ["Pending"] }); setTablePage(1); }} />
            <StatCard label="Resolved Tickets" value={resolved.toLocaleString()} note="Tickets Resolved" icon={CheckCircle2} tone="#60a5fa" onClick={() => { setFilters({ ...EMPTY_FILTERS, status: ["Resolved"] }); setTablePage(1); }} />
            <StatCard label="Critical Tickets" value={critical.toLocaleString()} note="High Priority Severity" icon={AlertTriangle} tone="#ef4444" onClick={() => { setFilters({ ...EMPTY_FILTERS, severity: ["Critical"] }); setTablePage(1); }} />
            <StatCard label="Major Tickets" value={major.toLocaleString()} note="Medium Priority Severity" icon={Activity} tone="#f59e0b" onClick={() => { setFilters({ ...EMPTY_FILTERS, severity: ["Major"] }); setTablePage(1); }} />
            <StatCard label="Minor Tickets" value={minor ? minor.toLocaleString() : ""} note="Low Priority Severity" icon={CircleDot} tone="#22d3ee" onClick={() => { setFilters({ ...EMPTY_FILTERS, severity: ["Minor"] }); setTablePage(1); }} />
            <StatCard label="Service Impact" value={serviceImpact.toLocaleString()} note="Exact Service Impact" icon={Network} tone="#a78bfa" onClick={() => { setFilters({ ...EMPTY_FILTERS, impact: ["Service Impact"] }); setTablePage(1); }} />
            <StatCard label="Non-Service Impact" value={nonServiceImpact.toLocaleString()} note="No Service Impact" icon={XCircle} tone="#94a3b8" onClick={() => { setFilters({ ...EMPTY_FILTERS, impact: ["Non-Service Impact"] }); setTablePage(1); }} />
            <StatCard label="Regions" value={analytics.region.toLocaleString()} note="Regions" icon={CircleDot} tone="#60a5fa" />
			<StatCard label="Unique Sites" value={analytics.uniqueSites.toLocaleString()} note="Unique Site ID" icon={BarChart3} tone="#60a5fa" />
            <StatCard label="Root Cause Updated" value={analytics.rootCauseUpdated.toLocaleString()} note="Tickets with Alarm Root Cause" icon={FileSpreadsheet} tone="#34d399" />
            <StatCard label="Top RCA by Tickets Count" value={analytics.topRcaByCount.name || ""} note={analytics.topRcaByCount.value ? `${analytics.topRcaByCount.value.toLocaleString()} Tickets` : ""} icon={BarChart3} tone="#22d3ee" />
            <StatCard label="Top RCA by Downtime" value={analytics.topRcaByDowntime.name || ""} note={formatHours(analytics.topRcaByDowntime.value)} icon={Activity} tone="#f59e0b" />
            <StatCard label="Highest MTTR RCA" value={analytics.highestMttrRca.name || ""} note={formatHours(analytics.highestMttrRca.value)} icon={AlertTriangle} tone="#ef4444" />
            <StatCard label="Repeated RCA Sites" value={analytics.repeatedRcaSites.toLocaleString()} note="Site + RCA pairs repeated" icon={Network} tone="#a78bfa" />
            <StatCard label="RCA not Provided %" value={rcaNotProvidedPct} note={`${analytics.rcaNotProvidedCount.toLocaleString()} Tickets missing RCA`} icon={ShieldAlert} tone="#f97316" />
          </section>

          <section className="chart-mosaic">
            <article className="glass-card wide">
              <div className="card-heading"><div><h3>Unique Tickets by month</h3></div><BarChart3 size={18} /></div>
              <ResponsiveContainer width="100%" height={290}>
                <LineChart data={analytics.monthly} margin={{ left: 0, right: 22, top: 16, bottom: 8 }}>
                  <CartesianGrid stroke="rgba(148,163,184,.16)" vertical={false} />
                  <XAxis dataKey="name" stroke="#94a3b8" tickLine={false} axisLine={false} />
                  <YAxis stroke="#94a3b8" tickLine={false} axisLine={false} allowDecimals={false} />
                  <Tooltip contentStyle={{ background: "#071426", border: "1px solid rgba(34,211,238,.25)", borderRadius: 14, color: "#e2e8f0" }} />
                  <Line type="monotone" dataKey="value" stroke="#22d3ee" strokeWidth={3} dot={{ r: 4, fill: "#071426", strokeWidth: 2 }}>
                    <LabelList dataKey="value" position="top" fill="#e2e8f0" fontSize={12} />
                  </Line>
                </LineChart>
              </ResponsiveContainer>
            </article>

            <article className="glass-card tall">
              <div className="card-heading"><div><h3>Top sites by unique Tickets</h3></div></div>
              <ResponsiveContainer width="100%" height={360}>
                <BarChart data={analytics.topSites.slice(0, 10)} layout="vertical" margin={{ left: 18, right: 44, top: 8, bottom: 8 }}>
                  <CartesianGrid stroke="rgba(148,163,184,.12)" horizontal={false} />
                  <XAxis type="number" stroke="#94a3b8" allowDecimals={false} />
                  <YAxis dataKey="name" type="category" stroke="#cbd5e1" width={210} tickLine={false} axisLine={false} tick={{ fontSize: 11 }} interval={0} />
                  <Tooltip contentStyle={{ background: "#071426", border: "1px solid rgba(34,211,238,.25)", borderRadius: 14, color: "#e2e8f0" }} />
                  <Bar dataKey="value" radius={[0, 10, 10, 0]} fill="#22d3ee">
                    <LabelList dataKey="value" position="right" fill="#e2e8f0" fontSize={12} />
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
              <div className="top-site-mini-table">
                {analytics.topSites.slice(0, 10).map((site, index) => (
                  <div key={`${site.name}-${index}`}><span>{index + 1}</span><strong>{site.name}</strong><em>{site.value}</em><small>{pct(site.value, analytics.totalSiteAffected || analytics.totalUnique)}</small></div>
                ))}
              </div>
            </article>

            <article className="glass-card">
              <div className="card-heading"><div><h3>Status</h3></div></div>
              <ResponsiveContainer width="100%" height={250}>
                <PieChart>
                  <Pie data={analytics.status} dataKey="value" nameKey="name" innerRadius={48} outerRadius={76} paddingAngle={3} labelLine={false} label={renderPieLabel}>
                    {analytics.status.map((entry, index) => <Cell key={entry.name} fill={STATUS_COLORS[entry.name] ?? COLORS[index % COLORS.length]} />)}
                  </Pie>
                  <Tooltip contentStyle={{ background: "#071426", border: "1px solid rgba(34,211,238,.25)", borderRadius: 14, color: "#e2e8f0" }} />
                  <Legend />
                </PieChart>
              </ResponsiveContainer>
            </article>

            <article className="glass-card">
              <div className="card-heading"><div><h3>Severity</h3></div></div>
              <ResponsiveContainer width="100%" height={250}>
                <PieChart>
                  <Pie data={analytics.severity} dataKey="value" nameKey="name" innerRadius={48} outerRadius={76} paddingAngle={3} labelLine={false} label={renderPieLabel}>
                    {analytics.severity.map((entry, index) => <Cell key={entry.name} fill={SEVERITY_COLORS[entry.name] ?? COLORS[index % COLORS.length]} />)}
                  </Pie>
                  <Tooltip contentStyle={{ background: "#071426", border: "1px solid rgba(34,211,238,.25)", borderRadius: 14, color: "#e2e8f0" }} />
                  <Legend />
                </PieChart>
              </ResponsiveContainer>
            </article>

            <article className="glass-card">
              <div className="card-heading"><div><h3>Region</h3></div></div>
              <ResponsiveContainer width="100%" height={250}>
                <PieChart>
                  <Pie data={analytics.region} dataKey="value" nameKey="name" innerRadius={48} outerRadius={76} paddingAngle={3} labelLine={false} label={renderPieLabel}>
                    {analytics.region.map((entry, index) => <Cell key={entry.name} fill={COLORS[index % COLORS.length]} />)}
                  </Pie>
                  <Tooltip contentStyle={{ background: "#071426", border: "1px solid rgba(34,211,238,.25)", borderRadius: 14, color: "#e2e8f0" }} />
                  <Legend />
                </PieChart>
              </ResponsiveContainer>
            </article>

            <article className="glass-card">
              <div className="card-heading"><div><h3>Escalation level</h3></div></div>
              <ResponsiveContainer width="100%" height={250}>
                <PieChart>
                  <Pie data={analytics.escalation} dataKey="value" nameKey="name" innerRadius={48} outerRadius={76} paddingAngle={3} labelLine={false} label={renderPieLabel}>
                    {analytics.escalation.map((entry, index) => <Cell key={entry.name} fill={COLORS[index % COLORS.length]} />)}
                  </Pie>
                  <Tooltip contentStyle={{ background: "#071426", border: "1px solid rgba(34,211,238,.25)", borderRadius: 14, color: "#e2e8f0" }} />
                  <Legend />
                </PieChart>
              </ResponsiveContainer>
            </article>

            <article className="glass-card">
              <div className="card-heading"><div><h3>Service impact</h3></div></div>
              <ResponsiveContainer width="100%" height={250}>
                <BarChart data={analytics.impact} margin={{ left: 0, right: 24, top: 18, bottom: 42 }}>
                  <CartesianGrid stroke="rgba(148,163,184,.12)" vertical={false} />
                  <XAxis dataKey="name" stroke="#cbd5e1" tickLine={false} axisLine={false} interval={0} angle={-18} textAnchor="end" height={58} tick={{ fontSize: 11 }} />
                  <YAxis stroke="#94a3b8" tickLine={false} axisLine={false} allowDecimals={false} />
                  <Tooltip contentStyle={{ background: "#071426", border: "1px solid rgba(34,211,238,.25)", borderRadius: 14, color: "#e2e8f0" }} />
                  <Bar dataKey="value" radius={[10, 10, 0, 0]} fill="#a78bfa">
                    <LabelList dataKey="value" position="top" fill="#e2e8f0" fontSize={12} />
                    {analytics.impact.map((entry, index) => <Cell key={entry.name} fill={COLORS[index % COLORS.length]} />)}
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </article>

            <article className="glass-card full">
              <div className="card-heading"><div><h3>Top 10 RCA by unique Tickets count</h3></div></div>
              <ResponsiveContainer width="100%" height={320}>
                <BarChart data={analytics.rcaByCount.slice(0, 10)} layout="vertical" margin={{ left: 18, right: 56, top: 8, bottom: 8 }}>
                  <CartesianGrid stroke="rgba(148,163,184,.12)" horizontal={false} />
                  <XAxis type="number" stroke="#94a3b8" allowDecimals={false} />
                  <YAxis dataKey="name" type="category" stroke="#cbd5e1" width={200} tickLine={false} axisLine={false} tick={{ fontSize: 12 }} interval={0} />
                  <Tooltip contentStyle={{ background: "#071426", border: "1px solid rgba(34,211,238,.25)", borderRadius: 14, color: "#e2e8f0" }} />
                  <Bar dataKey="value" radius={[0, 10, 10, 0]} fill="#22d3ee">
                    <LabelList dataKey="value" position="right" fill="#e2e8f0" fontSize={13} />
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </article>

            <article className="glass-card wide">
              <div className="card-heading"><div><h3>Operational families -- Tickets distribution</h3></div></div>
              <div className="rca-family-layout">
                <ResponsiveContainer width="100%" height={280}>
                  <PieChart>
                    <Pie data={analytics.rcaFamily} dataKey="value" nameKey="name" innerRadius={68} outerRadius={110} paddingAngle={4} labelLine={false}>
                      {analytics.rcaFamily.map((entry, index) => <Cell key={entry.name} fill={COLORS[index % COLORS.length]} />)}
                    </Pie>
                    <Tooltip contentStyle={{ background: "#071426", border: "1px solid rgba(34,211,238,.25)", borderRadius: 14, color: "#e2e8f0" }} />
                  </PieChart>
                </ResponsiveContainer>
                <div className="rca-family-legend">
                  {analytics.rcaFamily.map((entry, index) => (
                    <div key={entry.name} className="rca-legend-row">
                      <span className="rca-legend-dot" style={{ background: COLORS[index % COLORS.length] }} />
                      <span className="rca-legend-name">{entry.name}</span>
                      <strong className="rca-legend-val">{entry.value} <small>({pct(entry.value, analytics.totalUnique)})</small></strong>
                    </div>
                  ))}
                </div>
              </div>
            </article>

            <article className="glass-card wide">
              <div className="card-heading"><div><h3>Top Managed Resources by Ticket Count</h3></div></div>
              <ResponsiveContainer width="100%" height={Math.max(200, analytics.topManagedResources.length * 32 + 24)}>
                <BarChart data={analytics.topManagedResources} layout="vertical" margin={{ left: 18, right: 72, top: 8, bottom: 8 }}>
                  <CartesianGrid stroke="rgba(148,163,184,.12)" horizontal={false} />
                  <XAxis type="number" stroke="#94a3b8" allowDecimals={false} />
                  <YAxis dataKey="name" type="category" stroke="#cbd5e1" width={220} tickLine={false} axisLine={false} tick={{ fontSize: 12 }} interval={0} />
                  <Tooltip contentStyle={{ background: "#071426", border: "1px solid rgba(34,211,238,.25)", borderRadius: 14, color: "#e2e8f0" }} formatter={(v: number) => [v.toLocaleString(), "Tickets"]} />
                  <Bar dataKey="value" radius={[0, 10, 10, 0]} fill="#22d3ee">
                    <LabelList dataKey="value" position="right" fill="#e2e8f0" fontSize={12} />
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </article>

            <article className="glass-card wide">
              <div className="card-heading"><div><h3>Top 10 RCA by total downtime (hrs)</h3></div></div>
              <ResponsiveContainer width="100%" height={300}>
                <BarChart data={analytics.rcaByDowntime.slice(0, 10)} layout="vertical" margin={{ left: 18, right: 72, top: 8, bottom: 8 }}>
                  <CartesianGrid stroke="rgba(148,163,184,.12)" horizontal={false} />
                  <XAxis type="number" stroke="#94a3b8" allowDecimals={false} />
                  <YAxis dataKey="name" type="category" stroke="#cbd5e1" width={200} tickLine={false} axisLine={false} tick={{ fontSize: 12 }} interval={0} />
                  <Tooltip contentStyle={{ background: "#071426", border: "1px solid rgba(34,211,238,.25)", borderRadius: 14, color: "#e2e8f0" }} formatter={(v: number) => [`${v.toLocaleString()} hrs`, "Downtime"]} />
                  <Bar dataKey="value" radius={[0, 10, 10, 0]} fill="#f59e0b">
                    <LabelList dataKey="value" position="right" fill="#e2e8f0" fontSize={12} formatter={(v: number) => `${v.toLocaleString()}h`} />
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </article>

            <article className="glass-card">
              <div className="card-heading"><div><h3>Highest MTTR by RCA</h3></div></div>
              <ResponsiveContainer width="100%" height={300}>
                <BarChart data={analytics.rcaByMttr.slice(0, 8)} layout="vertical" margin={{ left: 18, right: 72, top: 8, bottom: 8 }}>
                  <CartesianGrid stroke="rgba(148,163,184,.12)" horizontal={false} />
                  <XAxis type="number" stroke="#94a3b8" allowDecimals={false} />
                  <YAxis dataKey="name" type="category" stroke="#cbd5e1" width={200} tickLine={false} axisLine={false} tick={{ fontSize: 12 }} interval={0} />
                  <Tooltip contentStyle={{ background: "#071426", border: "1px solid rgba(34,211,238,.25)", borderRadius: 14, color: "#e2e8f0" }} formatter={(v: number) => [`${v.toLocaleString()} hrs avg`, "MTTR"]} />
                  <Bar dataKey="value" radius={[0, 10, 10, 0]} fill="#ef4444">
                    <LabelList dataKey="value" position="right" fill="#e2e8f0" fontSize={12} formatter={(v: number) => `${v.toLocaleString()}h`} />
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </article>

            <article className="glass-card full">
              <div className="card-heading"><div><h3>Monthly RCA Family distribution -- stacked Tickets count</h3></div></div>
              <ResponsiveContainer width="100%" height={340}>
                <BarChart data={analytics.monthlyRcaFamily} margin={{ left: 0, right: 24, top: 8, bottom: 48 }}>
                  <CartesianGrid stroke="rgba(148,163,184,.12)" vertical={false} />
                  <XAxis dataKey="name" stroke="#cbd5e1" tickLine={false} axisLine={false} interval={0} angle={-30} textAnchor="end" height={64} tick={{ fontSize: 11 }} />
                  <YAxis stroke="#94a3b8" tickLine={false} axisLine={false} allowDecimals={false} />
                  <Tooltip contentStyle={{ background: "#071426", border: "1px solid rgba(34,211,238,.25)", borderRadius: 14, color: "#e2e8f0" }} />
                  <Legend wrapperStyle={{ paddingTop: 8, fontSize: 12, color: "#cbd5e1" }} />
                  {analytics.rcaFamilyKeys.map((family, index) => (
                    <Bar key={family} dataKey={family} stackId="a" fill={COLORS[index % COLORS.length]} radius={index === analytics.rcaFamilyKeys.length - 1 ? [4, 4, 0, 0] : [0, 0, 0, 0]} />
                  ))}
                </BarChart>
              </ResponsiveContainer>
            </article>
          </section>

          {/* Monthly Performance Table */}
          {perfRows.length > 0 && (
            <section className="table-card">
              {/* KPI tiles */}
              {(() => {
                const kpi = computePerfKPIs(perfRows);
                return (
                  <div className="perf-kpi-row perf-kpi-row--table">
                    {([
                      { label: "% Availability", value: kpi.pctAvailability, color: "#22c55e" },
                      { label: "MTTR", value: kpi.mttr, color: "#f59e0b" },
                      { label: "MTBF", value: kpi.mtbf, color: "#3b82f6" },
                      { label: "MTTF", value: kpi.mttf, color: "#a78bfa" },
                      { label: "Affected Sites", value: String(kpi.affectedSites), color: "#f43f5e" },
                      { label: "Total Down", value: kpi.totalDownHrs, color: "#fb923c" },
                    ] as { label: string; value: string; color: string }[]).map(({ label, value, color }) => (
                      <div key={label} className="perf-kpi-tile">
                        <div className="perf-kpi-label">{label}</div>
                        <div className="perf-kpi-value" style={{ color }}>{value || "--"}</div>
                      </div>
                    ))}
                  </div>
                );
              })()}
              <div className="table-heading">
                <div>
                  <h2>Monthly Performance -- {perfMonths.length === 0 ? "All Months" : perfMonths.length === 1 ? formatMonthMMMMYYYY(perfMonths[0]) : `${perfMonths.length} months`}{perfRegions.length > 0 ? ` — ${perfRegions.join(", ")}` : ""}</h2>
                  <p style={{ color: "#94a3b8", fontSize: 13, marginTop: 4 }}>
                    {perfRows.length.toLocaleString()} sites &nbsp;|&nbsp;
                    Total hours in month: {perfMonths.length === 1 ? totalHoursInMonth(perfMonths[0]).toLocaleString() : perfMonths.length > 1 ? perfMonths.reduce((s, m) => s + totalHoursInMonth(m), 0).toLocaleString() : "N/A"} hrs
                  </p>
                </div>
                <div style={{ display: "flex", gap: 10 }}>
                  <button className="ghost-button" onClick={() => exportPerfCsv(perfRows, perfMonths[0] ?? "all")}><Download size={16} /> CSV</button>
                  <button className="ghost-button" onClick={() => exportPerfExcel(perfRows, perfMonths[0] ?? "all")}><FileSpreadsheet size={16} /> Excel</button>
                  <button className="ghost-button" onClick={() => exportPerfPdf(perfRows, perfMonths[0] ?? "all")}><Printer size={16} /> PDF</button>
                </div>
              </div>
              <div className="table-scroll">
                <table>
                  <thead>
                    <tr>{PERF_REPORT_HEADERS.map((h) => <th key={h}>{h}</th>)}</tr>
                  </thead>
                  <tbody>
                    {perfRows.map((row, i) => {
                      // Parse reliability % for color-coding
                      const relNum = parseFloat(row.reliability);
                      const relColor = !row.sitesDownHours ? undefined
                        : relNum < 95 ? "#ef4444"   // red below 95%
                        : relNum < 99 ? "#f59e0b"   // amber below 99%
                        : undefined;                 // no highlight at/above 99%
                      return (
                        <tr key={row.siteId}>
                          <td className="mono">{i + 1}</td>
                          <td className="mono">{row.siteId}</td>
                          <td>{row.siteName}</td>
                          <td className="mono">{row.availHours}</td>
                          <td>{row.availDay}</td>
                          <td className="mono"></td>
                          <td className="mono"></td>
                          <td className="mono" style={relColor ? { color: relColor, fontWeight: 700 } : undefined}>{row.reliability}</td>
                          <td className="mono">{row.sitesDownHours}</td>
                        </tr>
                      );
                    })}
                    {/* Total summary row */}
                    {perfRows.length > 0 && (() => {
                      const totalDown = Math.round(perfRows.reduce((s, r) => s + r.sitesDownHours, 0) * 10) / 10;
                      const totalAvail = Math.round(perfRows.reduce((s, r) => s + r.availHours, 0) * 10) / 10;
                      const totalHrs = totalAvail + totalDown;
                      const overallRel = totalHrs > 0 ? ((totalAvail / totalHrs) * 100).toFixed(2) + "%" : "";
                      const relNum = parseFloat(overallRel);
                      const relColor = totalDown === 0 ? undefined
                        : relNum < 95 ? "#ef4444"
                        : relNum < 99 ? "#f59e0b"
                        : "#22c55e";
                      return (
                        <tr style={{ fontWeight: 700, borderTop: "2px solid rgba(148,163,184,0.3)", background: "rgba(255,255,255,0.04)" }}>
                          <td className="mono" colSpan={3} style={{ textAlign: "right", paddingRight: 16, color: "#94a3b8" }}>TOTAL</td>
                          <td className="mono">{totalAvail}</td>
                          <td></td>
                          <td className="mono"></td>
                          <td className="mono"></td>
                          <td className="mono" style={relColor ? { color: relColor, fontWeight: 700 } : undefined}>{overallRel}</td>
                          <td className="mono">{totalDown}</td>
                        </tr>
                      );
                    })()}
                  </tbody>
                </table>
              </div>
            </section>
          )}

          {/* Tickets Table */}
          <section className="table-card">
            <div className="table-heading">
              <div><h2>{filteredTickets.length.toLocaleString()} distinct Ticket records</h2></div>
              <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
                <span style={{ color: "#94a3b8", fontSize: 13 }}>Page {tablePage} of {Math.max(1, Math.ceil(filteredTickets.length / TABLE_PAGE_SIZE))} &mdash; {filteredTickets.length.toLocaleString()} total</span>
              </div>
            </div>
            <div className="table-scroll" id="ticket-table-wrapper">
              <table>
                <thead>
                  <tr>{DISTINCT_REPORT_HEADERS.map((header) => <th key={header}>{header}</th>)}</tr>
                </thead>
                <tbody>
                  {filteredTickets.slice((tablePage - 1) * TABLE_PAGE_SIZE, tablePage * TABLE_PAGE_SIZE).map((ticket, index) => {
                    const row = ticket.primary;
                    const reportRow = distinctReportRow(ticket, index);
                    const siteIds = Array.from(ticket.siteIds).filter(Boolean);
                    const siteNames = Array.from(ticket.siteNames).filter(Boolean);
                    return (
                      <tr key={ticket.tt}>
                        {reportRow.map((cell, cellIndex) => {
                          const header = DISTINCT_REPORT_HEADERS[cellIndex];
                          if (header === "#" || header === "TT") return <td key={header} className="mono">{cell}</td>;
                          if (header === "Site ID") return (
                            <td key={header} className="mono" style={{ whiteSpace: "pre-line", lineHeight: 1.9 }}>
                              {siteIds.length ? siteIds.map((id, i) => <span key={id}>{id}{i < siteIds.length - 1 ? "\n" : ""}</span>) : cell}
                            </td>
                          );
                          if (header === "Site Name") return (
                            <td key={header} style={{ whiteSpace: "pre-line", lineHeight: 1.9 }}>
                              {siteNames.length ? siteNames.map((name, i) => <span key={name}>{name}{i < siteNames.length - 1 ? "\n" : ""}</span>) : cell}
                            </td>
                          );
                          if (header === "Severity") return <td key={header}><span className="pill" style={{ ["--pill" as string]: SEVERITY_COLORS[row.severity] ?? "#64748b" }}>{cell}</span></td>;
                          if (header === "Status") return <td key={header}><span className="pill" style={{ ["--pill" as string]: STATUS_COLORS[row.status] ?? "#64748b" }}>{cell}</span></td>;
                          if (["Issues", "RCA", "Recommended Action", "Responsible Team"].includes(header)) return <td key={header} className="issue-cell">{cell}</td>;
                          return <td key={header}>{cell}</td>;
                        })}
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
            {Math.ceil(filteredTickets.length / TABLE_PAGE_SIZE) > 1 && (
              <div className="pagination-bar no-print">
                <button className="ghost-button" disabled={tablePage <= 1} onClick={() => setTablePage(1)}>«</button>
                <button className="ghost-button" disabled={tablePage <= 1} onClick={() => setTablePage((p) => p - 1)}>‹ Prev</button>
                {Array.from({ length: Math.ceil(filteredTickets.length / TABLE_PAGE_SIZE) }, (_, i) => i + 1)
                  .filter((p) => p === 1 || p === Math.ceil(filteredTickets.length / TABLE_PAGE_SIZE) || Math.abs(p - tablePage) <= 2)
                  .reduce<(number | "...")[]>((acc, p, idx, arr) => {
                    if (idx > 0 && (p as number) - (arr[idx - 1] as number) > 1) acc.push("...");
                    acc.push(p);
                    return acc;
                  }, [])
                  .map((p, idx) =>
                    p === "..." ? <span key={`ellipsis-${idx}`} style={{ color: "#94a3b8", padding: "0 4px" }}>...</span> :
                    <button key={p} className={`ghost-button${p === tablePage ? " active-page" : ""}`} onClick={() => setTablePage(p as number)}>{p}</button>
                  )}
                <button className="ghost-button" disabled={tablePage >= Math.ceil(filteredTickets.length / TABLE_PAGE_SIZE)} onClick={() => setTablePage((p) => p + 1)}>Next ›</button>
                <button className="ghost-button" disabled={tablePage >= Math.ceil(filteredTickets.length / TABLE_PAGE_SIZE)} onClick={() => setTablePage(Math.ceil(filteredTickets.length / TABLE_PAGE_SIZE))}>»</button>
              </div>
            )}
          </section>
        </>
      )}
    </main>
  );
}
