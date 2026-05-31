// src/lib/dateUtils.ts

import * as XLSX from "xlsx";
import type { TicketAggregate, TicketRecord } from "../types/dashboard";

export function clean(value: unknown): string {
  if (value === null || value === undefined) return "";
  return String(value).replace(/\s+/g, " ").trim();
}

export function normalizeHeader(value: string): string {
  return clean(value)
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
}

export function getField(row: Record<string, unknown>, aliases: string[]): string {
  const map = new Map<string, unknown>();
  Object.entries(row).forEach(([key, value]) =>
    map.set(normalizeHeader(key), value)
  );

  for (const alias of aliases) {
    const found = map.get(normalizeHeader(alias));
    if (found !== undefined) return clean(found);
  }

  return "";
}

export function getRawField(
  row: Record<string, unknown>,
  aliases: string[]
): unknown {
  const map = new Map<string, unknown>();
  Object.entries(row).forEach(([key, value]) =>
    map.set(normalizeHeader(key), value)
  );

  for (const alias of aliases) {
    const found = map.get(normalizeHeader(alias));
    if (found !== undefined && found !== null && found !== "") return found;
  }

  return "";
}

export function isRfSiteId(value: unknown): boolean {
  return clean(value).toUpperCase().startsWith("RF");
}

export function parseDateValue(value: unknown): Date | null {
  if (value === null || value === undefined || value === "") return null;
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;

  const numericValue =
    typeof value === "number"
      ? value
      : /^\d+(\.\d+)?$/.test(clean(value))
        ? Number(clean(value))
        : null;

  if (numericValue !== null && numericValue > 20000 && numericValue < 90000) {
    const excelDate = XLSX.SSF.parse_date_code(numericValue);
    if (excelDate) return new Date(excelDate.y, excelDate.m - 1, excelDate.d);
  }

  const text = clean(value);
  const dmyMatch = text.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/);

  if (dmyMatch) {
    const d = Number(dmyMatch[1]);
    const m = Number(dmyMatch[2]);
    const y = Number(dmyMatch[3]);

    if (d > 12) {
      const candidate = new Date(y, m - 1, d);
      if (!Number.isNaN(candidate.getTime())) return candidate;
    } else {
      const candidate = new Date(y, m - 1, d);
      if (!Number.isNaN(candidate.getTime())) return candidate;
    }
  }

  const direct = new Date(text);
  if (!Number.isNaN(direct.getTime())) return direct;

  const match = text.match(
    /(\d{1,2})[-/ ]([A-Za-z]{3,}|\d{1,2})[-/ ](\d{2,4})/
  );

  if (!match) return null;

  const parsed = new Date(text.replace(/-/g, " "));
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

export function normalizeMonthKey(value: unknown): string | null {
  const text = clean(value);
  if (!text) return null;

  const keyMatch = text.match(/^(\d{4})[-/](\d{1,2})$/);
  if (keyMatch) return `${keyMatch[1]}-${keyMatch[2].padStart(2, "0")}`;

  const parsed = parseDateValue(text.startsWith("1 ") ? text : `1 ${text}`);
  if (!parsed) return null;

  return `${parsed.getFullYear()}-${String(parsed.getMonth() + 1).padStart(2, "0")}`;
}

export function openingMonthKey(value: unknown): string {
  const parsed = parseDateValue(value);
  if (!parsed) return "Unknown";

  const month = String(parsed.getMonth() + 1).padStart(2, "0");
  return `${parsed.getFullYear()}-${month}`;
}

export function resolveOpeningMonthKey(
  sourceKey: unknown,
  sourceLabel: unknown,
  observationDate: unknown
): string {
  return (
    normalizeMonthKey(sourceKey) ??
    normalizeMonthKey(sourceLabel) ??
    openingMonthKey(observationDate)
  );
}

export function openingMonthLabel(key: string): string {
  if (!key || key === "Unknown") return "Unknown";

  const parsed = new Date(`${key}-01T00:00:00`);
  if (Number.isNaN(parsed.getTime())) return key;

  return parsed.toLocaleDateString("en", { month: "short", year: "numeric" });
}

export function startOfWeek(date: Date): Date {
  const copy = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  const day = copy.getDay();
  const diff = day === 0 ? -6 : 1 - day;

  copy.setDate(copy.getDate() + diff);
  return copy;
}

export function weekKey(value: unknown): string {
  const parsed = parseDateValue(value);
  if (!parsed) return "Unknown";

  const start = startOfWeek(parsed);

  return `${start.getFullYear()}-${String(start.getMonth() + 1).padStart(
    2,
    "0"
  )}-${String(start.getDate()).padStart(2, "0")}`;
}

export function weekLabel(key: string): string {
  if (!key || key === "Unknown") return "Unknown";

  const parsed = new Date(`${key}T00:00:00`);
  if (Number.isNaN(parsed.getTime())) return key;

  const end = new Date(parsed);
  end.setDate(end.getDate() + 6);

  const sameMonth =
    parsed.getMonth() === end.getMonth() &&
    parsed.getFullYear() === end.getFullYear();

  const startLabel = parsed.toLocaleDateString("en", {
    day: "2-digit",
    month: "short",
  });

  const endLabel = end.toLocaleDateString(
    "en",
    sameMonth ? { day: "2-digit" } : { day: "2-digit", month: "short" }
  );

  return `${startLabel}-${endLabel}`;
}

export function recordDateMonthKey(value: string): string {
  const parsed = parseDateValue(value);
  if (!parsed) return "Unknown";

  return `${parsed.getFullYear()}-${String(parsed.getMonth() + 1).padStart(
    2,
    "0"
  )}`;
}

export function selectedMonthRange(
  selectedMonth: string
): { start: Date; end: Date } | null {
  const match = selectedMonth.match(/^(\d{4})-(\d{2})$/);
  if (!match) return null;

  const year = Number(match[1]);
  const monthIndex = Number(match[2]) - 1;

  const start = new Date(year, monthIndex, 1);
  const end = new Date(year, monthIndex + 1, 0, 23, 59, 59, 999);

  return Number.isNaN(start.getTime()) || Number.isNaN(end.getTime())
    ? null
    : { start, end };
}

export function totalHoursInMonth(monthKey: string): number {
  const range = selectedMonthRange(monthKey);
  if (!range) return 0;

  const daysInMonth = range.end.getDate();
  return daysInMonth * 24;
}

export function coveredMonthKeys(row: TicketRecord): string[] {
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
    keys.add(
      `${cursor.getFullYear()}-${String(cursor.getMonth() + 1).padStart(2, "0")}`
    );

    cursor.setMonth(cursor.getMonth() + 1);
    guard += 1;
  }

  return Array.from(keys);
}

export function isPendingStatus(value: string): boolean {
  return clean(value).toLowerCase() === "pending";
}

export function dateWithinMonth(dateValue: string, selectedMonth: string): boolean {
  const range = selectedMonthRange(selectedMonth);
  if (!range) return false;

  const parsed = parseDateValue(dateValue);
  if (!parsed) return false;

  return parsed >= range.start && parsed <= range.end;
}

export function ticketMatchesMonthlyExport(
  ticket: TicketAggregate,
  selectedMonth: string
): boolean {
  if (selectedMonth === "all") return true;

  const range = selectedMonthRange(selectedMonth);
  if (!range) return false;

  return ticket.rows.some(row => {
    const obsDate = parseDateValue(row.observationDate);
    const recDate = parseDateValue(row.recoveryDate);

    const observationInMonth =
      obsDate !== null && obsDate >= range.start && obsDate <= range.end;

    const recoveryInMonth =
      recDate !== null && recDate >= range.start && recDate <= range.end;

    const pendingBeforeMonthEnd =
      isPendingStatus(row.status) && obsDate !== null && obsDate <= range.end;

    const spansEntireMonth =
      obsDate !== null &&
      recDate !== null &&
      obsDate < range.start &&
      recDate > range.end;

    return (
      observationInMonth ||
      recoveryInMonth ||
      pendingBeforeMonthEnd ||
      spansEntireMonth
    );
  });
}

export function dateKey(value: string): number {
  const parsed = new Date(value);
  return Number.isNaN(parsed.getTime()) ? 0 : parsed.getTime();
}

export function formatDateDDMMYYYY(value: string): string {
  const parsed = parseDateValue(value);
  if (!parsed) return value || "";

  const d = String(parsed.getDate()).padStart(2, "0");
  const m = String(parsed.getMonth() + 1).padStart(2, "0");
  const y = parsed.getFullYear();

  return `${d}/${m}/${y}`;
}

export function formatMonthMMMMYYYY(monthKey: string): string {
  if (!monthKey || monthKey === "all") return "";

  const parsed = new Date(`${monthKey}-01T00:00:00`);
  if (Number.isNaN(parsed.getTime())) return monthKey;

  return parsed.toLocaleDateString("en", { month: "short", year: "numeric" });
}

export function parseDurationHours(duration: string): number | null {
  if (!duration) return null;

  if (/days?/i.test(duration) || /hrs?/i.test(duration)) {
    const days = Number(duration.match(/(\d+)\s*days?/i)?.[1] ?? 0);
    const hrs = Number(duration.match(/(\d+)\s*hrs?/i)?.[1] ?? 0);
    const mins = Number(duration.match(/(\d+)\s*mins?/i)?.[1] ?? 0);

    const total = days * 24 + hrs + mins / 60;
    return Number.isFinite(total) ? total : null;
  }

  const num = Number(duration);

  if (Number.isFinite(num) && num >= 0) {
    if (num > 0 && num < 1) return Math.round(num * 24 * 10) / 10;
    return Math.round(num * 10) / 10;
  }

  return null;
}

export function combineDateTime(dateStr: string, timeStr: string): Date | null {
  const parsed = parseDateValue(dateStr);
  if (!parsed) return null;

  const time = clean(timeStr);
  const match = time.match(/^(\d{1,2}):(\d{2})/);

  if (match) parsed.setHours(Number(match[1]), Number(match[2]), 0, 0);

  return parsed;
}

export function hoursBetween(start: Date | null, end: Date | null): number | null {
  if (!start || !end || end < start) return null;

  return Math.round(((end.getTime() - start.getTime()) / 36e5) * 10) / 10;
}

export function average(values: Array<number | null | undefined>): number {
  const valid = values.filter(
    (value): value is number =>
      typeof value === "number" && Number.isFinite(value)
  );

  return valid.length
    ? valid.reduce((sum, value) => sum + value, 0) / valid.length
    : 0;
}

export function formatHours(value: number): string {
  if (!value || !Number.isFinite(value)) return "";

  return `${value.toFixed(1)} hrs`;
}

export function normalizeSiteId(id: string): string {
  return id
    .replace(
      /(\D+)(\d+)$/,
      (_, prefix, num) => prefix.toUpperCase() + String(parseInt(num, 10))
    )
    .trim();
}
