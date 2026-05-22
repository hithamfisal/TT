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

import nascoLogoSrc from "../assets/nascologo.png";
import ngLogoSrc from "../assets/nglogo.png";

const HERO_IMAGE =
  "https://d2xsxph8kpxj0f.cloudfront.net/310419663031216744/LFprMZJgQoCY2omHrJr6xN/followup-hero-network-cockpit-GEqHM9kSYycEMt32RSfRxg.webp";
const UPLOAD_IMAGE =
  "https://d2xsxph8kpxj0f.cloudfront.net/310419663031216744/LFprMZJgQoCY2omHrJr6xN/followup-upload-orb-YcKWLpbfcdm5ofcRiuQymn.webp";
const RIBBON_IMAGE =
  "https://d2xsxph8kpxj0f.cloudfront.net/310419663031216744/LFprMZJgQoCY2omHrJr6xN/followup-network-ribbon-Lv2N5GpYhLW5eJjvLPNzkg.webp";

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

function observationRecoveryOverlapsMonth(row: TicketRecord, selectedMonth: string): boolean {
  const range = selectedMonthRange(selectedMonth);
  if (!range) return false;
  const observation = parseDateValue(row.observationDate);
  const recovery = parseDateValue(row.recoveryDate);
  if (!observation || !recovery) return false;
  const startDate = observation <= recovery ? observation : recovery;
  const endDate = recovery >= observation ? recovery : observation;
  return startDate <= range.end && endDate >= range.start;
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

    // Criterion 3: Status is Pending AND Observation Date ≤ last day of the selected month
    // (ticket was opened any time before or during the month and is still unresolved as of that month)
    const pendingBeforeMonthEnd = isPendingStatus(row.status) && obsDate !== null && obsDate <= range.end;

    // Criterion 4: Ticket was active throughout the entire month
    // (Observation Date before the first day AND Recovery Date after the last day)
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
  // Fetch the pre-formatted template
  const templateUrl = "/manus-storage/DMRMonthlyReportEOA_76429c79.xlsx";
  let templateBuffer: ArrayBuffer;
  try {
    const response = await fetch(templateUrl);
    if (!response.ok) throw new Error("Template not found");
    templateBuffer = await response.arrayBuffer();
  } catch {
    // Fallback: export as plain Excel
    exportExcel(rows);
    return;
  }
  const wb = XLSX.read(templateBuffer, { type: "array", cellStyles: true });
  const ws = wb.Sheets[wb.SheetNames[0]];

  // Write month label at R5 (col 18 = R)
  const monthLabel = formatMonthMMMMYYYY(monthKey);
  if (monthLabel) {
    ws["R5"] = { t: "s", v: monthLabel };
  }

  // Data starts at row 39 (after header rows 37-38)
  const DATA_START_ROW = 39;
  // Column mapping (1-indexed): A=1 No, B=2 Equipment/site, C=3 Site Name, E=5 Managed Resource,
  // F=6 Severity, G=7 Alarm Type, H=8 Escalation Date, I=9 Escalation Time,
  // J=10 Recovery Date, K=11 Recovery Time, L=12 L3 Date, M=13 L3 Time,
  // N=14 Duration, O=15 TT Number, P=16 TT Status, Q=17 TT Owner, R=18 Comments
  const colLetter = (n: number) => {
    let s = "";
    while (n > 0) { s = String.fromCharCode(((n - 1) % 26) + 65) + s; n = Math.floor((n - 1) / 26); }
    return s;
  };
  const setCell = (row: number, col: number, value: string | number) => {
    const addr = `${colLetter(col)}${row}`;
    ws[addr] = { t: typeof value === "number" ? "n" : "s", v: value };
  };

  // Remove existing data rows between DATA_START_ROW and the signature rows (row 60)
  // by clearing cells in those rows
  for (let r = DATA_START_ROW; r < 60; r++) {
    for (let c = 1; c <= 18; c++) {
      delete ws[`${colLetter(c)}${r}`];
    }
  }

  // Write ticket rows
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

  // Update the worksheet ref range
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
  const days = Number(duration.match(/(\d+)\s*days?/i)?.[1] ?? 0);
  const hrs = Number(duration.match(/(\d+)\s*hrs?/i)?.[1] ?? 0);
  const mins = Number(duration.match(/(\d+)\s*mins?/i)?.[1] ?? 0);
  const total = days * 24 + hrs + mins / 60;
  return Number.isFinite(total) ? total : null;
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

function parseRows(workbook: XLSX.WorkBook, fileName: string): DashboardData {
  const preferred = workbook.SheetNames.find((name) => name.toLowerCase().includes("tickets_data"));
  const sheetName = preferred ?? workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  // Use raw:true so date cells come in as Excel serial numbers (reliable for parseDateValue).
  // We immediately convert date fields to dd/mm/yyyy strings so display remains readable.
  const raw = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, { defval: "", raw: true });

  // Helper: convert a raw cell value (serial number, Date, or string) to dd/mm/yyyy string
  function toDateStr(val: unknown): string {
    if (val === null || val === undefined || val === "") return "";
    const parsed = parseDateValue(val);
    if (!parsed) return String(val);
    const d = String(parsed.getDate()).padStart(2, "0");
    const m = String(parsed.getMonth() + 1).padStart(2, "0");
    return `${d}/${m}/${parsed.getFullYear()}`;
  }

  // Helper: convert a raw cell value (Excel time fraction, Date object, or string) to hh:mm string
  function toTimeStr(val: unknown): string {
    if (val === null || val === undefined || val === "") return "";
    // SheetJS raw:true returns time-only cells as Date objects (epoch 1899-12-30, time in UTC)
    if (val instanceof Date) {
      const hh = String(val.getUTCHours()).padStart(2, "0");
      const mm = String(val.getUTCMinutes()).padStart(2, "0");
      return `${hh}:${mm}`;
    }
    // Excel stores time as a fractional day (0.0 = midnight, 0.5 = noon)
    if (typeof val === "number") {
      const totalMinutes = Math.round(val * 24 * 60);
      const hh = String(Math.floor(totalMinutes / 60) % 24).padStart(2, "0");
      const mm = String(totalMinutes % 60).padStart(2, "0");
      return `${hh}:${mm}`;
    }
    const text = String(val).trim();
    // Already hh:mm or hh:mm:ss -- strip seconds
    const hhmmss = text.match(/^(\d{1,2}):(\d{2})(:\d{2})?$/);
    if (hhmmss) return `${hhmmss[1].padStart(2, "0")}:${hhmmss[2]}`;
    // Fallback: return as-is
    return text;
  }

  const rows: TicketRecord[] = raw
    .map((row, index) => {
      const observationDate = toDateStr(getField(row, ["Observation Date", "Observed Date"]) || "");
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
        siteId: getField(row, ["Site ID", "SiteID"]),
        siteName: getField(row, ["Site Name", "SiteName"]),
        managedResource: getField(row, ["Managed Resource", "ManagedResource", "Managed Resources", "Resource", "NE Name", "Network Element"]),
        issue: getField(row, ["Issues", "Issue"]),
        severity: getField(row, ["Severity"]),
        region: getField(row, ["Region"]),
        observationDate,
        observationTime: toTimeStr(getRawField(row, ["Observation Time", "Observed Time", "ObservationTime"])),
        openingMonthKey: monthKey,
        openingMonthLabel: openingMonthLabel(monthKey),
        recoveryDate: toDateStr(getField(row, ["Recovery Date"]) || ""),
        recoveryTime: toTimeStr(getRawField(row, ["Recovery Time", "RecoveryTime"])),
        duration: getField(row, ["Total Duration Days/Hours", "Total Durration Days/Hours", "Duration"]),
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
    { wch: 34 },
    { wch: 26 },
    { wch: 28 },
    { wch: 30 },
    { wch: 58 },
  ];
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Distinct TT Report");
  XLSX.writeFile(workbook, "follow-up-distinct-tt-report.xlsx");
}

function exportPdf(rows: TicketAggregate[], monthKey: string) {
  const doc = new jsPDF({ orientation: "landscape", unit: "mm", format: "a3" });

  // Title
  const monthLabel = monthKey !== "all" ? formatMonthMMMMYYYY(monthKey) : "All Months";
  doc.setFontSize(14);
  doc.setTextColor(10, 30, 60);
  doc.text(`DMR Monthly Tickets Report -- ${monthLabel}`, 14, 16);
  doc.setFontSize(9);
  doc.setTextColor(100);
  doc.text(`Generated: ${new Date().toLocaleString()} | Total Tickets: ${rows.length}`, 14, 22);

  const headers = DISTINCT_REPORT_HEADERS;
  const body = distinctReportRows(rows);

  autoTable(doc, {
    startY: 26,
    head: [headers],
    body,
    styles: {
      fontSize: 7,
      cellPadding: 1.5,
      overflow: "linebreak",
      valign: "middle",
    },
    headStyles: {
      fillColor: [4, 60, 100],
      textColor: 255,
      fontStyle: "bold",
      fontSize: 7.5,
    },
    alternateRowStyles: { fillColor: [240, 246, 255] },
    columnStyles: {
      0:  { cellWidth: 8 },   // #
      1:  { cellWidth: 18 },  // Site ID
      2:  { cellWidth: 22 },  // Site Name
      3:  { cellWidth: 26 },  // Managed Resource
      4:  { cellWidth: 14 },  // Severity
      5:  { cellWidth: 30 },  // Issues
      6:  { cellWidth: 18 },  // Observation Date
      7:  { cellWidth: 16 },  // Observation Time
      8:  { cellWidth: 18 },  // Recovery Date
      9:  { cellWidth: 16 },  // Recovery Time
      10: { cellWidth: 22 },  // L3 Support Date
      11: { cellWidth: 20 },  // L3 Support Time
      12: { cellWidth: 22 },  // Duration
      13: { cellWidth: 14 },  // TT
      14: { cellWidth: 14 },  // Status
      15: { cellWidth: 18 },  // Escalated to
      16: { cellWidth: 36 },  // RCA
    },
    margin: { left: 10, right: 10 },
    didDrawPage: (data) => {
      const pageCount = (doc as jsPDF & { internal: { getNumberOfPages: () => number } }).internal.getNumberOfPages();
      doc.setFontSize(7);
      doc.setTextColor(150);
      doc.text(
        `Page ${data.pageNumber} of ${pageCount}`,
        doc.internal.pageSize.getWidth() - 30,
        doc.internal.pageSize.getHeight() - 6,
      );
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

function SelectFilter({ label, value, options, optionLabels, onChange }: {
  label: string;
  value: string;
  options: string[];
  optionLabels?: Record<string, string>;
  onChange: (value: string) => void;
}) {
  return (
    <label className="filter-field">
      <span>{label}</span>
      <div className="select-wrap">
        <select value={value} onChange={(event) => onChange(event.target.value)}>
          <option value="all">All</option>
          {options.map((option) => (
            <option key={option} value={option}>{optionLabels?.[option] ?? option}</option>
          ))}
        </select>
        <ChevronDown size={14} />
      </div>
    </label>
  );
}

function MultiSelectFilter({ label, value, options, optionLabels, onChange }: {
  label: string;
  value: string[];
  options: string[];
  optionLabels?: Record<string, string>;
  onChange: (value: string[]) => void;
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

  // Reposition on scroll/resize while open
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

  const toggle = (opt: string) => {
    if (value.includes(opt)) onChange(value.filter((v) => v !== opt));
    else onChange([...value, opt]);
  };
  const displayLabel = value.length === 0 ? "All" : value.length === 1 ? (optionLabels?.[value[0]] ?? value[0]) : `${value.length} selected`;

  const dropdown = open ? createPortal(
    <div className="multi-select-dropdown" style={dropdownStyle} ref={dropdownRef}>
      {value.length > 0 && (
        <button type="button" className="multi-select-clear" onClick={() => onChange([])}>× Clear</button>
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

  const [exportMonth, setExportMonth] = useState("all");
  const [filters, setFilters] = useState<Filters>(EMPTY_FILTERS);
  const [tablePage, setTablePage] = useState(1);
  const TABLE_PAGE_SIZE = 50;

  const [perfMonth, setPerfMonth] = useState("all");
  const [perfRegion, setPerfRegion] = useState("ALL");

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
      setExportMonth("all");
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
        // Merge all regions into a combined DashboardData
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
        const merged: DashboardData = {
          fileName: updated.map((r) => r.fileName).join(" + "),
          sheetName: updated[0].sheetName,
          generatedAt: new Date().toLocaleString(),
          rows: allRows,
          uniqueTickets: Array.from(ttMap.values()),
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
        row.responsibleTeam || getResponsibleTeam(row.rcaFamily || getRcaFamily(row.rca)),
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
        row.responsibleTeam || getResponsibleTeam(row.rcaFamily || getRcaFamily(row.rca)),
        row.escalatedForL3SupportDate,
        row.escalatedForL3SupportTime,
      ]
        .join(" ")
        .toLowerCase();
      // NOTE: Status filter is intentionally excluded here.
      // ticketMatchesMonthlyExport applies its own Pending criterion,
      // so filtering by status here would incorrectly exclude Pending tickets.
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

  const monthlyExportTickets = useMemo(() => monthlyExportBaseTickets.filter((ticket) => ticketMatchesMonthlyExport(ticket, exportMonth)), [exportMonth, monthlyExportBaseTickets]);

  const selectedExportMonthLabel = exportMonth === "all" ? "All export-eligible TT" : openingMonthLabel(exportMonth);

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

    // Monthly RCA Family stacked bar data
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
      totalUnique,
      status,
      severity,
      region,
      impact,
      escalation,
      monthly,
      avgHours,
      uniqueSites,
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
      rcaByDowntime: downtimeByRca.map((item) => ({ name: item.name, value: Math.round(item.value * 10) / 10 })),
      rcaByMttr: mttrByRca.map((item) => ({ name: item.name, value: Math.round(item.value * 10) / 10 })),
      preventableBreakdown,
      monthlyRcaFamily,
      rcaFamilyKeys: RCA_FAMILIES,
      topManagedResources,
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
  const filteredSourceRowCount = filteredTickets.reduce((sum, ticket) => sum + ticket.rows.length, 0);
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
            {data && <button className="ghost-button" onClick={() => exportPdf(monthlyExportTickets, exportMonth)}><Printer size={16} /> PDF</button>}
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
        <>
  
  
  
  <div className="hero-export-row no-print">
    <aside className="hero-export-card" aria-label="Monthly TT export filter">
      <div className="hero-export-copy">
        <span>Monthly Tickets export filter</span>
        <strong>
          <span className="export-badge">{monthlyExportTickets.length}</span>{" "}
          tickets for {selectedExportMonthLabel}
        </strong>
      </div>

      <SelectFilter
        label="Report Month"
        value={exportMonth}
        options={filterOptions.exportMonth}
        optionLabels={filterOptions.exportMonthLabels}
        onChange={setExportMonth}
      />

      <div className="hero-export-actions">
        <button className="ghost-button" onClick={() => exportCsv(monthlyExportTickets)}>
          <Download size={16} /> Export CSV
        </button>

        <button className="ghost-button" onClick={() => exportExcel(monthlyExportTickets)}>
          <FileSpreadsheet size={16} /> Export Excel
        </button>

        <button className="ghost-button" onClick={() => exportPdf(monthlyExportTickets, exportMonth)}>
          <Printer size={16} /> Export PDF
        </button>
      </div>
    </aside>
  </div>

<div className="hero-export-row no-print">
  <aside className="hero-export-card" aria-label="Monthly Performance filter">
    <div className="hero-export-copy">
      <span>Monthly Performance</span>
      <strong>
        <span className="export-badge">
          {monthlyExportTickets.filter(
            (t) => perfRegion === "ALL" || t.primary.region === perfRegion
          ).length}
        </span>
        Performance for{" "}
        {perfMonth === "all" ? "All Months" : formatMonthMMMMYYYY(perfMonth)}
      </strong>
    </div>

    <SelectFilter
      label="Report Month"
      value={perfMonth}
      options={filterOptions.exportMonth}
      optionLabels={filterOptions.exportMonthLabels}
      onChange={setPerfMonth}
    />

    <SelectFilter
      label="Region"
      value={perfRegion}
      options={["ALL", ...filterOptions.region]}
      onChange={setPerfRegion}
    />

    <div className="hero-export-actions">
      <button
        className="ghost-button"
        onClick={() =>
          exportCsv(
            monthlyExportTickets.filter(
              (t) => perfRegion === "ALL" || t.primary.region === perfRegion
            )
          )
        }
      >
        <Download size={16} /> Export CSV
      </button>

      <button
        className="ghost-button"
        onClick={() =>
          exportExcel(
            monthlyExportTickets.filter(
              (t) => perfRegion === "ALL" || t.primary.region === perfRegion
            )
          )
        }
      >
        <FileSpreadsheet size={16} /> Export Excel
      </button>

      <button
        className="ghost-button"
        onClick={() =>
          exportPdf(
            monthlyExportTickets.filter(
              (t) => perfRegion === "ALL" || t.primary.region === perfRegion
            ),
            perfMonth
          )
        }
      >
        <Printer size={16} /> Export PDF
      </button>
    </div>
  </aside>
  </div>
  </>
</section>
</main>
);
}