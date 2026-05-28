import * as XLSX from "xlsx";

import type { DashboardData, TicketAggregate, TicketRecord } from "../types/dashboard";

import {
  clean,
  combineDateTime,
  dateKey,
  getField,
  getRawField,
  hoursBetween,
  openingMonthLabel,
  parseDateValue,
  parseDurationHours,
  resolveOpeningMonthKey,
} from "./dateUtils";

import {
  getPreventability,
  getRcaFamily,
  getRecommendedAction,
  getResponsibleTeam,
} from "./rcaRules";

export function groupTickets(rows: TicketRecord[]): TicketAggregate[] {
  const grouped = new Map<string, TicketAggregate>();
  rows.forEach(row => {
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
      if (!currentDate || (nextDate && nextDate < currentDate))
        existing.primary = row;
    }
  });
  return Array.from(grouped.values());
}

export function parseSiteOrder(
  workbook: XLSX.WorkBook
): { siteId: string; siteName: string }[] {
  const siteSheetName = workbook.SheetNames.find(name => {
    const n = name.toLowerCase().replace(/[^a-z0-9]/g, "");
    return (
      n === "dashboarddata" ||
      n === "siteid" ||
      n === "sites" ||
      n === "sitelist" ||
      n === "siteids" ||
      n === "sitedata"
    );
  });
  if (!siteSheetName) return [];
  const sheet = workbook.Sheets[siteSheetName];
  const raw2d = XLSX.utils.sheet_to_json<unknown[]>(sheet, {
    header: 1,
    defval: null,
  }) as unknown[][];
  let siteIdCol = -1;
  let siteNameCol = -1;
  let headerRowIdx = -1;
  for (let ri = 0; ri < raw2d.length; ri++) {
    const row = raw2d[ri];
    for (let ci = 0; ci < row.length; ci++) {
      const cell = String(row[ci] ?? "")
        .trim()
        .toLowerCase();
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
    if (!id) break;
    if (!seen.has(id)) {
      seen.add(id);
      result.push({ siteId: id, siteName: name });
    }
  }
  return result;
}

export function parseRows(workbook: XLSX.WorkBook, fileName: string): DashboardData {
  const preferred = workbook.SheetNames.find(name =>
    name.toLowerCase().includes("tickets_data")
  );
  const sheetName = preferred ?? workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const raw = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, {
    defval: "",
    raw: true,
  });
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
      const observationDate = toDateStr(
        getRawField(row, ["Observation Date", "Observed Date"]) || ""
      );
      const monthKey = resolveOpeningMonthKey(
        getField(row, ["Opening Month Key", "OpeningMonthKey"]),
        getField(row, ["Opening Month", "OpeningMonth"]),
        observationDate
      );
      const observationTime = toTimeStr(
        getRawField(row, [
          "Observation Time",
          "Observed Time",
          "ObservationTime",
        ])
      );
      const recoveryDate = toDateStr(getRawField(row, ["Recovery Date"]) || "");
      const recoveryTime = toTimeStr(
        getRawField(row, ["Recovery Time", "RecoveryTime"])
      );
      const duration = String(
        getRawField(row, [
          "Total Duration Days/Hours",
          "Total Durration Days/Hours",
          "Duration",
        ]) ?? ""
      );
      const l3Date = toDateStr(
        getField(row, [
          "Escalated for L3 Support Date",
          "Escalated For L3 Support Date",
          "L3 Support Date",
          "L3 Escalation Date",
          "Escalation L3 Date",
          "Escalated L3 Date",
        ]) || ""
      );
      const l3Time = toTimeStr(
        getRawField(row, [
          "Escalated for L3 Support Time",
          "Escalated For L3 Support Time",
          "L3 Support Time",
          "L3 Escalation Time",
          "Escalation L3 Time",
          "Escalated L3 Time",
        ])
      );
      const observedAt = combineDateTime(observationDate, observationTime);
      const l3At = combineDateTime(l3Date, l3Time);
      const recoveredAt = combineDateTime(recoveryDate, recoveryTime);
      const explicitFrt = parseDurationHours(
        String(
          getRawField(row, [
            "FRT",
            "Avg FRT",
            "First Reply Time",
            "First Response Time",
          ]) ?? ""
        )
      );
      const explicitResponse = parseDurationHours(
        String(
          getRawField(row, [
            "Response Time",
            "Avg Response Time",
            "Ticket Response Time",
          ]) ?? ""
        )
      );
      const explicitResolution = parseDurationHours(
        String(
          getRawField(row, [
            "Resolution Time",
            "Avg Resolution Time",
            "Ticket Resolution Time",
          ]) ?? ""
        )
      );
      const rca = getField(row, [
        "RCA",
        "Root Cause Analysis",
        "Root Cause",
        "Action Taken/RCA",
      ]);
      const rcaFamily = getRcaFamily(rca);
      return {
        rowNo: index + 2,
        tt: getField(row, ["TT", "Ticket", "Ticket Number"]),
        siteId: getField(row, ["Site ID", "SiteID", "Site Name", "SiteName"]),
        siteName: getField(row, ["Site Name", "SiteName"]),
        managedResource: getField(row, [
          "Managed Resource",
          "ManagedResource",
          "Managed Resources",
          "Resource",
          "NE Name",
          "Network Element",
        ]),
        issue: getField(row, ["Issues", "Issue"]),
        severity: getField(row, ["Severity"]),
        region: getField(row, ["Region"]),
        observationDate,
        observationTime,
        openingMonthKey: monthKey,
        openingMonthLabel: openingMonthLabel(monthKey),
        recoveryDate,
        recoveryTime,
        duration,
        impact: getField(row, [
          "Service Impaction Status",
          "Service Impact Status",
        ]),
        escalatedTo: getField(row, [
          "Escalated to",
          "Escalated To",
          "Escalated to ",
        ]),
        escalationLevel: getField(row, ["Escalation Level", "Esclation Level"]),
        escalatedForL3SupportDate: l3Date,
        escalatedForL3SupportTime: l3Time,
        frtHours: explicitFrt ?? hoursBetween(observedAt, l3At),
        responseHours: explicitResponse ?? hoursBetween(observedAt, l3At),
        resolutionHours:
          explicitResolution ??
          parseDurationHours(duration) ??
          hoursBetween(observedAt, recoveredAt),
        status: getField(row, ["Status"]),
        rca,
        rcaFamily,
        preventability: getPreventability(rca),
        responsibleTeam: getResponsibleTeam(rcaFamily),
        recommendedAction: getRecommendedAction(rcaFamily),
        actionTaken: getField(row, ["Action"]),
        sourceFile: fileName,
      };
    })
    .filter(row => row.tt || row.siteId || row.siteName || row.issue);

  return {
    fileName,
    sheetName,
    generatedAt: new Date().toLocaleString(),
    rows,
    uniqueTickets: groupTickets(rows),
    siteOrder,
  };
}
