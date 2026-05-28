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
import ngLogoSrc from "../assets/nglogo.png";
import seLogoSrc from "../assets/se.png";

const THEME_IMAGES = {
  dark: "/dark.png",
  light: "/light.png",
} as const;

type DashboardTheme = keyof typeof THEME_IMAGES;


type DashboardSectionId =
  | "reports"
  | "performanceKpis"
  | "kpis"
  | "executive"
  | "deepDive"
  | "overviewCharts"
  | "trendCharts"
  | "ticketsTable"
  | "performanceTable";

type DashboardSectionDefinition = {
  id: DashboardSectionId;
  label: string;
  title: string;
  selector: string;
};

const DASHBOARD_SECTIONS: DashboardSectionDefinition[] = [
  { id: "reports", label: "Reports", title: "Report Export Center", selector: "#section-reports" },
  { id: "performanceKpis", label: "Performance KPIs", title: "Performance KPI Gauge Cards", selector: "#section-performance-kpis" },
  { id: "kpis", label: "KPIs", title: "Operational KPI Cards", selector: "#section-kpis" },
  { id: "executive", label: "Executive", title: "Executive Insights", selector: "#section-executive" },
  { id: "deepDive", label: "RCA / SLA", title: "RCA / Preventability / SLA Deep-Dive", selector: "#section-deep-dive" },
  { id: "overviewCharts", label: "Overview Charts", title: "Operational Overview Charts", selector: "#section-overview-charts" },
  { id: "trendCharts", label: "Trend Charts", title: "Trend & RCA Charts", selector: "#section-trend-charts" },
  { id: "ticketsTable", label: "Tickets", title: "Tickets Data Table", selector: "#section-tickets-table" },
  { id: "performanceTable", label: "Performance", title: "Performance Table", selector: "#section-performance-table" },
];

const INITIAL_COLLAPSED_SECTIONS: Record<DashboardSectionId, boolean> = {
  reports: true,
  performanceKpis: true,
  kpis: true,
  executive: true,
  deepDive: true,
  overviewCharts: true,
  trendCharts: true,
  ticketsTable: true,
  performanceTable: true,
};

const SAVED_DASHBOARD_KEY = "followup-dashboard:last-workbook:v1";

import {
  Activity,
  AlertTriangle,
  ArrowUp,
  BarChart3,
  CheckCircle2,
  ChevronDown,
  CircleDot,
  Download,
  FileSpreadsheet,
  Filter,
  CloudOff,
  Home as HomeIcon,
  Layers3,
  ImageDown,
  Maximize2,
  Minimize2,
  Moon,
  type LucideIcon,
  Network,
  Presentation,
  Printer,
  RefreshCw,
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
import { calculateDeepDiveAnalytics, calculateExecutiveInsights } from "../lib/ticketAnalytics";


type ZipFileMap = Record<string, Uint8Array>;
type ZipTextCodec = {
  strFromU8: (bytes: Uint8Array) => string;
  strToU8: (text: string) => Uint8Array;
};
type ExcelHorizontalAlign = "left" | "center";

const applyExcelCellStyle = (
  xml: string,
  ref: string,
  styleIndex: number
): string => {
  const markerIdx = xml.indexOf(` r="${ref}"`);
  if (markerIdx === -1) return xml;
  const cStart = xml.lastIndexOf("<c", markerIdx);
  const tagClose = xml.indexOf(">", cStart);
  if (cStart === -1 || tagClose === -1) return xml;

  const openTag = xml.slice(cStart, tagClose + 1);
  const styledTag = openTag.includes(' s="')
    ? openTag.replace(/\s+s="[^"]*"/, ` s="${styleIndex}"`)
    : openTag.replace(/\/?>$/, ending => ` s="${styleIndex}"${ending}`);

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
  horizontal: ExcelHorizontalAlign
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
          ""
        )
        .replace("</xf>", `${alignmentTag}</xf>`);

  const existingIndex = xfs.findIndex(xf => xf === alignedXf);
  if (existingIndex !== -1) return existingIndex;

  const nextIndex = xfs.length;
  const beforeClose = stylesXml.slice(0, cellXfsEnd);
  const afterClose = stylesXml.slice(cellXfsEnd);
  stylesXml = `${beforeClose}${alignedXf}${afterClose}`;
  stylesXml = stylesXml.replace(
    /(<cellXfs\b[^>]*\bcount=")\d+(")/,
    `$1${nextIndex + 1}$2`
  );
  files[stylesKey] = codec.strToU8(stylesXml);
  return nextIndex;
};

const applyExcelColumnAlignment = (
  files: ZipFileMap,
  codec: ZipTextCodec,
  xml: string,
  refs: string[],
  horizontal: ExcelHorizontalAlign
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
  horizontal: ExcelHorizontalAlign
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
  rows.map(row =>
    row.map(cell => {
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
    })
  );

// 1. Performance Data Exporter to PPT
const exportPerfPpt = (
  data: PerfRow[],
  monthLabel: string = "All",
  executiveInsights?: ExecutiveInsights
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
    const labels = (data || []).map(r => safePptText(r?.siteName || "N/A", 160));
    const availValues = (data || []).map(r => r?.availHours || 0);
    const downValues = (data || []).map(r => r?.sitesDownHours || 0);

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
      s1.addText(`Network Health Score: ${executiveInsights.healthScore.score} / 100 · ${executiveInsights.healthScore.status}`, {
        x: 0.7,
        y: 2.95,
        w: 9,
        fontSize: 15,
        bold: true,
        color: WHITE,
        fontFace: "Segoe UI",
      });
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
    kpiSlide.addText(`Month: ${monthLabel}  ·  ${data.length} sites`, {
      x: 0.5,
      y: 0.55,
      w: 12,
      fontSize: 11,
      color: MUTED,
      fontFace: "Segoe UI",
    });

    const kpiItems = [
      {
        label: "% Availability",
        value: kpi.pctAvailability,
        color: GREEN,
        icon: "✓",
        labelStyle: { fontSize: "20px", fontWeight: 900, color: "#94a3b8" },
        valueStyle: { fontSize: "20px", fontWeight: 700, color: GREEN },
      },
      {
        label: "MTTR",
        value: kpi.mttr,
        color: AMBER,
        icon: "⏱",
        labelStyle: { fontSize: "20px", fontWeight: 900, color: "#94a3b8" },
        valueStyle: { fontSize: "20px", fontWeight: 700, color: AMBER },
      },
      {
        label: "MTBF",
        value: kpi.mtbf,
        color: CYAN,
        icon: "↻",
        labelStyle: { fontSize: "20px", fontWeight: 900, color: "#94a3b8" },
        valueStyle: { fontSize: "20px", fontWeight: 700, color: AMBER },
      },
      {
        label: "MTTF",
        value: kpi.mttf,
        color: CYAN,
        icon: "↻",
        labelStyle: { fontSize: "20px", fontWeight: 900, color: "#94a3b8" },
        valueStyle: { fontSize: "20px", fontWeight: 700, color: AMBER },
      },
      {
        label: "Affected Sites",
        value: String(kpi.affectedSites),
        color: kpi.affectedSites > 0 ? RED : GREEN,
        icon: "⚠",
        labelStyle: { fontSize: "20px", fontWeight: 900, color: "#94a3b8" },
        valueStyle: { fontSize: "20px", fontWeight: 700, color: AMBER },
      },
      {
        label: "Total Down Time",
        value: kpi.totalDownHrs,
        color: kpi.totalDownHrs === "0.0 hrs" ? GREEN : RED,
        icon: "↓",
        labelStyle: { fontSize: "20px", fontWeight: 900, color: "#94a3b8" },
        valueStyle: { fontSize: "20px", fontWeight: 700, color: AMBER },
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
      slide.addTable(safePptTableRows(tableRows as any) as any, {
        x: 0.5,
        y: 0.65,
        w: 12,
        fontSize: 10,
        fontFace: "Segoe UI",
        border: { type: "solid", color: "1e3a5f", pt: 0.5 },
        fill: { color: CARD_BG },
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
    slide3.addText("Site Availability — Hours", {
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
    slide3.addChart(
      pptx.ChartType.bar,
      [{ name: "Availability (Hrs)", labels, values: availValues }],
      {
        x: 0.4,
        y: 0.85,
        w: 12.2,
        h: 5.8,
        barDir: "col",
        barGapWidthPct: 120,
        chartColors: [GREEN],
        showValue: true,
        dataLabelFontSize: 7,
        dataLabelColor: WHITE,
        catAxisLabelFontSize: 8,
        catAxisLabelColor: MUTED,
        catAxisLabelRotate: 45,
        valAxisLabelFontSize: 9,
        valAxisLabelColor: MUTED,
        valGridLine: { style: "solid", color: "1e3a5f", size: 0.5 },
        ...({ plotAreaFill: { color: CARD_BG } } as any),
        showLegend: false,
      }
    );

    const slide4 = pptx.addSlide();
    slide4.background = { color: BG };
    slide4.addShape(pptx.ShapeType.rect, {
      x: 0,
      y: 0,
      w: PPT_SLIDE_W,
      h: 0.06,
      fill: { color: RED },
    });
    slide4.addText("Site Downtime — Hours", {
      x: 0.5,
      y: 0.18,
      w: 12,
      fontSize: 18,
      bold: true,
      color: WHITE,
      fontFace: "Segoe UI",
    });
    slide4.addText(
      `Month: ${monthLabel}  ·  ${data.filter(r => r.sitesDownHours > 0).length} of ${data.length} sites affected`,
      {
        x: 0.5,
        y: 0.52,
        w: 10,
        fontSize: 11,
        color: MUTED,
        fontFace: "Segoe UI",
      }
    );
    slide4.addChart(
      pptx.ChartType.bar,
      [{ name: "Down (Hrs)", labels, values: downValues }],
      {
        x: 0.4,
        y: 0.85,
        w: 12.2,
        h: 5.8,
        barDir: "col",
        barGapWidthPct: 120,
        chartColors: [RED],
        showValue: true,
        dataLabelFontSize: 7,
        dataLabelColor: WHITE,
        catAxisLabelFontSize: 8,
        catAxisLabelColor: MUTED,
        catAxisLabelRotate: 45,
        valAxisLabelFontSize: 9,
        valAxisLabelColor: MUTED,
        valGridLine: { style: "solid", color: "1e3a5f", size: 0.5 },
        ...({ plotAreaFill: { color: CARD_BG } } as any),
        showLegend: false,
      }
    );

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
  deepDive?: DeepDiveAnalytics
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

      const insightRows = executiveInsights.cards.slice(0, 8).map(card => [
        { text: card.label, options: { color: MUTED, fontSize: 9, bold: true } },
        { text: String(card.value), options: { color: WHITE, fontSize: 12, bold: true } },
        { text: card.note, options: { color: MUTED, fontSize: 8 } },
      ]);
      execSlide.addTable(safePptTableRows(insightRows as any) as any, {
        x: 0.55,
        y: 2.25,
        w: 12,
        colW: [3.1, 2.2, 6.7],
        rowH: 0.35,
        border: { type: "solid", color: "1e3a5f", pt: 0.4 },
        fill: { color: CARD_BG },
        margin: 0.04,
      } as any);
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
        ["Rank", "Site ID", "Site Name", "Tickets", "Downtime", "Reliability", "Top RCA", "Risk"].map(text => ({
          text,
          options: { bold: true, color: "0a1628", fill: { color: "ef4444" }, align: "center" },
        })),
        ...executiveInsights.highRiskSites.slice(0, 10).map(site => [
          { text: String(site.rank), options: { color: MUTED, align: "center" } },
          { text: site.siteId, options: { color: CYAN, bold: true } },
          { text: site.siteName || "-", options: { color: WHITE } },
          { text: String(site.ticketCount), options: { color: WHITE, align: "center" } },
          { text: `${site.downtimeHours} hrs`, options: { color: WHITE, align: "center" } },
          { text: `${site.reliability.toFixed(2)}%`, options: { color: WHITE, align: "center" } },
          { text: site.topRca, options: { color: WHITE } },
          { text: `${site.riskLevel} ${site.riskScore}`, options: { color: "f59e0b", bold: true } },
        ]),
      ];
      riskSlide.addTable(safePptTableRows(riskRows as any) as any, {
        x: 0.35,
        y: 0.9,
        w: 12.65,
        fontSize: 8,
        rowH: 0.34,
        colW: [0.55, 1.2, 2.2, 0.8, 1.0, 1.0, 3.3, 1.6],
        border: { type: "solid", color: "1e3a5f", pt: 0.4 },
        fill: { color: CARD_BG },
      } as any);
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
        ["RCA Family", "Tickets", "Downtime", "Missing RCA", "Preventable", "Owner Team"].map(text => ({
          text,
          options: { bold: true, color: "0a1628", fill: { color: "a78bfa" }, align: "center" },
        })),
        ...deepDive.rcaFamilyDeepDive.slice(0, 7).map(row => [
          { text: row.family, options: { color: WHITE, bold: true } },
          { text: String(row.tickets), options: { color: WHITE, align: "center" } },
          { text: `${row.downtimeHours} hrs`, options: { color: WHITE, align: "center" } },
          { text: String(row.missingRca), options: { color: row.missingRca ? "f59e0b" : "34d399", align: "center" } },
          { text: String(row.preventableTickets), options: { color: WHITE, align: "center" } },
          { text: row.responsibleTeam, options: { color: MUTED } },
        ]),
      ];
      deepSlide.addTable(safePptTableRows(rcaRows as any) as any, {
        x: 0.45,
        y: 0.85,
        w: 12.4,
        fontSize: 8,
        rowH: 0.35,
        colW: [2.7, 0.8, 1.1, 1.1, 1.1, 5.6],
        border: { type: "solid", color: "1e3a5f", pt: 0.4 },
        fill: { color: CARD_BG },
      } as any);
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
      slide.addTable(safePptTableRows(tableRows as any) as any, {
        x: 0.4,
        y: 0.7,
        w: 12.5,
        fontSize: 9,
        fontFace: "Segoe UI",
        border: { type: "solid", color: "1e3a5f", pt: 0.5 },
        fill: { color: CARD_BG },
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
  keyFn: (item: T) => string
): { name: string; value: number }[] {
  const map = new Map<string, number>();
  items.forEach(item => {
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
  expected: string
): number {
  return (
    items.find(item => item.name.toLowerCase() === expected.toLowerCase())
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

type ReportLogos = { ng: HTMLImageElement; nasco: HTMLImageElement };
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
  reportLogoPromise ??= Promise.all([
    loadImageElement(ngLogoSrc),
    loadImageElement(nascoLogoSrc),
  ]).then(([ng, nasco]) => ({ ng, nasco }));
  return reportLogoPromise;
}
function drawPdfReportHeader(
  doc: jsPDF,
  pageW: number,
  title: string,
  subtitle: string,
  logos: ReportLogos,
  titleColor: [number, number, number] = TEMPLATE_PDF_COLORS.title
) {
  doc.addImage(logos.ng, "PNG", 12, 6, 34, 12);
  doc.addImage(logos.nasco, "PNG", pageW - 46, 6, 34, 12);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(15);
  doc.setTextColor(...titleColor);
  doc.text(title, pageW / 2, 12.2, { align: "center" });
  doc.setFont("helvetica", "normal");
  doc.setFontSize(8);
  doc.setTextColor(...TEMPLATE_PDF_COLORS.muted);
  doc.text(subtitle, pageW / 2, 17.2, { align: "center" });
}

function perfSourceLabel(row: TicketRecord): string {
  return clean(row.region) || clean(row.sourceFile) || "Workbook";
}
function perfEntryKey(sourceLabel: string, siteId: string): string {
  return `${sourceLabel.toLowerCase()}||${normalizeSiteId(clean(siteId)).toLowerCase()}`;
}

function computePerfRows(
  allRows: TicketRecord[],
  monthKey: string,
  siteOrder: { siteId: string; siteName: string }[] = []
): PerfRow[] {
  const range = monthKey !== "all" ? selectedMonthRange(monthKey) : null;
  const monthHours = monthKey !== "all" ? totalHoursInMonth(monthKey) : 24 * 30;
  const siteNameMap = new Map<string, string>();
  const siteIdMap = new Map<string, string>();
  const sourceMap = new Map<string, string>();
  const siteTicketCount = new Map<string, number>();
  const siteDownHours = new Map<string, number>();

  allRows.forEach(row => {
    if (!row.siteId) return;
    const sourceLabel = perfSourceLabel(row);
    const key = perfEntryKey(sourceLabel, row.siteId);
    if (!siteIdMap.has(key)) siteIdMap.set(key, row.siteId);
    if (!sourceMap.has(key)) sourceMap.set(key, sourceLabel);
    if (!siteNameMap.has(key) && row.siteName)
      siteNameMap.set(key, row.siteName);
    siteTicketCount.set(key, (siteTicketCount.get(key) ?? 0) + 1);
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

  allRows.forEach(row => {
    if (!row.siteId) return;
    if (clean(row.impact).toLowerCase() !== "service impact") return;
    const mr = clean(row.managedResource).toLowerCase();
    if (mr !== "complete site" && mr !== "link down") return;
    const outageStart = combineDatetime(
      row.observationDate,
      row.observationTime
    );
    if (!outageStart) return;
    let outageEnd: Date | null = combineDatetime(
      row.recoveryDate,
      row.recoveryTime
    );
    const sourceLabel = perfSourceLabel(row);
    if (monthKey === "all") {
      const hours = parseDurationHours(row.duration) ?? 0;
      const key = perfEntryKey(sourceLabel, row.siteId);
      siteDownHours.set(
        key,
        Math.round(((siteDownHours.get(key) ?? 0) + hours) * 10) / 10
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
    const key = perfEntryKey(sourceLabel, row.siteId);
    siteDownHours.set(
      key,
      Math.round(((siteDownHours.get(key) ?? 0) + hours) * 10) / 10
    );
  });

  const allSiteKeys = Array.from(siteIdMap.keys()).filter(Boolean);
  allSiteKeys.sort(
    (a, b) => (siteTicketCount.get(b) ?? 0) - (siteTicketCount.get(a) ?? 0)
  );
  let siteEntries: {
    siteId: string;
    siteName: string;
    sourceLabel: string;
    perfKey: string;
  }[];
  if (allSiteKeys.length > 0) {
    siteEntries = allSiteKeys.map(key => ({
      siteId: siteIdMap.get(key) ?? "",
      siteName: siteNameMap.get(key) ?? "",
      sourceLabel: sourceMap.get(key) ?? "Workbook",
      perfKey: key,
    }));
  } else {
    siteEntries = siteOrder.map(site => ({
      siteId: site.siteId,
      siteName: site.siteName,
      sourceLabel: "Workbook",
      perfKey: perfEntryKey("Workbook", site.siteId),
    }));
  }

  return siteEntries.map(({ siteId, siteName, sourceLabel, perfKey }) => {
    const downHours = siteDownHours.get(perfKey) ?? 0;
    const availHours = Math.max(0, monthHours - downHours);
    const totalHours = availHours + downHours;
    const reliability = totalHours > 0 ? availHours / totalHours : 1;
    const totalMins = Math.round(availHours * 60);
    const dDays = Math.floor(totalMins / (60 * 24));
    const dHrs = Math.floor((totalMins % (60 * 24)) / 60);
    const dMins = Math.round(totalMins % 60);
    const availDay = `${dDays} d, ${dHrs} h, ${dMins} m`;
    return {
      siteId,
      siteName,
      displayName: `${siteId}${sourceLabel !== "Workbook" ? ` (${sourceLabel})` : ""}`,
      sourceLabel,
      perfKey,
      sitesDownHours: downHours,
      availHours: Math.round(availHours * 10) / 10,
      availDay,
      reliability: `${(reliability * 100).toFixed(2)}%`,
      channelBusy: 0,
      mwLinkPerf: 0,
      ticketCount: siteTicketCount.get(perfKey) ?? 0,
    };
  });
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

function computePerfKPIs(rows: PerfRow[]): {
  pctAvailability: string;
  mttr: string;
  mtbf: string;
  mttf: string;
  totalDown: number;
  totalAvail: number;
  affectedSites: number;
  nonAffectedSites: number;
  totalDownHrs: string;
} {
  const totalAvail = rows.reduce((s, r) => s + r.availHours, 0);
  const totalDown = rows.reduce((s, r) => s + r.sitesDownHours, 0);
  const totalSiteIds = new Set(
    rows.map(r => clean(r.siteId)).filter(isRfSiteId)
  );
  const affectedSiteIds = new Set(
    rows
      .filter(r => r.sitesDownHours > 0)
      .map(r => clean(r.siteId))
      .filter(isRfSiteId)
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
  const totalDownHrs = `${tdDays}d ${tdHrs}h ${tdMin}m`;
  return {
    pctAvailability,
    mttr,
    mtbf,
    mttf,
    totalDown: totalDownRounded,
    totalAvail: Math.round(totalAvail * 10) / 10,
    affectedSites: sitesWithDown,
    nonAffectedSites,
    totalDownHrs,
  };
}

// ─────────────────────────────────────────────────────────────────────────────
// Due to artifact size limits, the remainder of the unchanged exporters
// (exportPerfTemplate, exportPerfCsv, exportPerfExcel, exportPerfPdf,
//  exportCsv, exportExcel, exportTicketTemplate, exportPdf) and the helpers
//  uniqueTicketValues, distinctReportRow, distinctReportRows remain
//  IDENTICAL to your current file. KEEP YOUR EXISTING IMPLEMENTATIONS.
//
//  >>> ACTION: copy these functions from your current Home.tsx unchanged:
//      • exportPerfTemplate
//      • exportPerfCsv
//      • exportPerfExcel
//      • exportPerfPdf
//      • uniqueTicketValues
//      • distinctReportRow
//      • distinctReportRows
//      • exportCsv
//      • exportExcel
//      • exportTicketTemplate
//      • exportPdf
//
//  Then keep the rest of the file (StatCard, PartnerLogoStrip, SelectFilter,
//  MultiSelectFilter, the default export Home component) intact EXCEPT for
//  the TWO modifications below.
// ─────────────────────────────────────────────────────────────────────────────

// ─────────────────────────────────────────────────────────────────────────────
// Performance Template Export
// Uses fflate to fill Network_Performance_Report.xlsx preserving all formatting.
// Charts auto-update: their data ranges are updated to match the new sheet name
// and actual site count. Place the template in public/Network_Performance_Report.xlsx
// ─────────────────────────────────────────────────────────────────────────────
async function exportPerfTemplate(
  rows: PerfRow[],
  monthKey: string,
  regions: string[]
) {
  const DATA_START = 7; // first data row
  const LAST_DATA = 40; // last template data row
  const PROTECTED = 41; // totals/formulas row — shift if needed
  const TEMPLATE_N = LAST_DATA - DATA_START + 1; // 34 rows
  const OLD_SHEET = "EOA March 2026"; // original sheet name in template
  const COLS = ["A", "B", "C", "D", "E", "F", "G", "H", "I"];

  const xmlEsc = (s: string) =>
    s
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");

  // Find a cell by ref, replace its value while keeping its style
  const setCell = (
    xml: string,
    ref: string,
    value: string | number
  ): string => {
    const markerIdx = xml.indexOf(` r="${ref}"`);
    if (markerIdx === -1) return xml;
    const cStart = xml.lastIndexOf("<c", markerIdx);
    if (cStart === -1) return xml;
    // End of the opening <c …> tag
    const tagClose = xml.indexOf(">", cStart);
    if (tagClose === -1) return xml;
    // Is the opening <c> tag itself self-closing (i.e. <c …/>)?
    const isSelfClosing = xml[tagClose - 1] === "/";
    let cEnd: number;
    if (isSelfClosing) {
      cEnd = tagClose + 1;
    } else {
      const fcIdx = xml.indexOf("</c>", tagClose);
      if (fcIdx === -1) return xml;
      cEnd = fcIdx + 4;
    }
    // Preserve original attributes (e.g. s="N" style index), strip type — we set it ourselves
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

  try {
    const res = await fetch("/Network_Performance_Report.xlsx");
    if (!res.ok)
      throw new Error(
        `HTTP ${res.status} — place Network_Performance_Report.xlsx in your public/ folder`
      );
    const rawBuf = await res.arrayBuffer();
    const { unzipSync, zipSync, strFromU8, strToU8 } = await import("fflate");
    const files = unzipSync(new Uint8Array(rawBuf));

    // ── Month label & sheet name ──────────────────────────────────────────
    const full = monthKey !== "all" ? formatMonthMMMMYYYY(monthKey) : "All";
    const parts = full.split(" ");
    const mmm =
      parts.length === 2 ? `${parts[0].slice(0, 3)}-${parts[1]}` : full;
    const regionPart = regions.length > 0 ? regions.join(", ") + " " : "";
    const sheetLabel = `${regionPart}${mmm}`
      .replace(/[\/*?[\]:]/g, "-")
      .slice(0, 31);

    // ── Locate worksheet XML ──────────────────────────────────────────────
    const sheetKey = Object.keys(files).find(k =>
      /^xl\/worksheets\/sheet\d+\.xml$/.test(k)
    )!;
    let xml = strFromU8(files[sheetKey]);

    // ── KPI summary (row 3 = labels already in template, row 4 = values) ──
    const kpi = computePerfKPIs(rows);
    xml = setCell(xml, "C3", mmm);
    xml = setCell(xml, "D4", kpi.totalDownHrs);
    xml = setCell(xml, "E4", String(kpi.affectedSites));
    xml = setCell(xml, "F4", kpi.pctAvailability);
    xml = setCell(xml, "G4", kpi.mttr);
    xml = setCell(xml, "H4", kpi.mtbf);
    xml = setCell(xml, "I4", kpi.mttf);

    // ── Row count management ──────────────────────────────────────────────
    const needed = rows.length;

    if (needed > TEMPLATE_N) {
      const extra = needed - TEMPLATE_N;

      // Capture template row 7 for cloning
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

      // Shift rows >= PROTECTED and their cell refs
      xml = xml
        .replace(/(<row[^>]* r=")(\d+)(")/g, (_m, a, n, b) =>
          +n >= PROTECTED ? `${a}${+n + extra}${b}` : _m
        )
        .replace(/(<c[^>]* r=")([A-Z]+)(\d+)(")/g, (_m, a, col, n, b) =>
          +n >= PROTECTED ? `${a}${col}${+n + extra}${b}` : _m
        )
        .replace(
          /(<dimension ref="[A-Z]+\d+:[A-Z]+)(\d+)(")/,
          (_m, pre, n, post) =>
            +n >= PROTECTED ? `${pre}${+n + extra}${post}` : _m
        );

      // Update the SUM formula in the shifted totals row
      const sumOld = `SUM(D${DATA_START}:D${LAST_DATA})`;
      const sumNew = `SUM(D${DATA_START}:D${LAST_DATA + extra})`;
      xml = xml.split(sumOld).join(sumNew);

      // Build and insert blank clone rows
      if (tmplRow) {
        const newRows = Array.from({ length: extra }, (_, i) => {
          const rn = LAST_DATA + 1 + i;
          return (
            tmplRow
              .replace(/ r="(\d+)"/, ` r="${rn}"`)
              .replace(
                /(<c[^>]* r=")([A-Z]+)\d+(")/g,
                (_m2, a, col, b) => `${a}${col}${rn}${b}`
              )
              // Strip values, inline-strings and formulas so cloned cells are truly empty.
              // Leaving <v></v> with t="s" → ref to shared-string "" → "Removed Records" error.
              .replace(/<v>[^<]*<\/v>/g, "")
              .replace(/<is>[\s\S]*?<\/is>/g, "")
              .replace(/<f[^>]*\/>/g, "")
              .replace(/<f[^>]*>[\s\S]*?<\/f>/g, "")
              // Strip the type attribute — setCell will assign the correct one when filling
              .replace(/\s+t="[^"]*"/g, "")
          );
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
        COLS.forEach(col => {
          xml = setCell(xml, `${col}${r}`, "");
        });
      }
    }

    // ── Always fix shared-formula markers (runs regardless of row count) ───
    // Strips t="shared", si="N" and ref="..." from <f> formula elements.
    // Also updates formula body row refs when row insertion happened.
    // Without this, Excel shows "Removed Records: Shared formula" on open.
    xml = xml.split(' t="shared"').join("");
    // Strip si="N" attributes
    {
      let tmp = "";
      let src = xml;
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
    // Strip ref="..." from <f> elements (only valid for shared/array formulas)
    {
      let result = "";
      let src = xml;
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
    // Update formula body row refs for any inserted rows.
    // Keep this scoped to <f>...</f> contents only; changing <c r="A41">
    // cell addresses creates duplicate cells and Excel repairs sheet1.xml.
    if (needed > TEMPLATE_N) {
      const extra2 = needed - TEMPLATE_N;
      xml = xml.replace(
        /(<f[^>]*>)([\s\S]*?)(<\/f>)/g,
        (_m, open, body, close) => {
          const updatedBody = body.replace(
            /(\$?)([A-J])(\$?)(\d+)/g,
            (ref: string, absCol: string, col: string, absRow: string, rowText: string) => {
              const rowNum = Number(rowText);
              return rowNum >= PROTECTED
                ? `${absCol}${col}${absRow}${rowNum + extra2}`
                : ref;
            }
          );
          return `${open}${updatedBody}${close}`;
        }
      );
    }

    // ── Fill data rows ──────────────────────────────────────────────────
    rows.forEach((row, i) => {
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
    const perfDataRows = Array.from(
      { length: rows.length },
      (_, i) => DATA_START + i
    );
    const perfTableRows = [DATA_START - 1, ...perfDataRows];
    xml = applyExcelColumnAlignment(
      files,
      { strFromU8, strToU8 },
      xml,
      perfTableRows.flatMap(r =>
        ["A", "D", "E", "F", "G", "H", "I"].map(col => `${col}${r}`)
      ),
      "center"
    );
    xml = applyExcelColumnAlignment(
      files,
      { strFromU8, strToU8 },
      xml,
      perfTableRows.flatMap(r => ["B", "C"].map(col => `${col}${r}`)),
      "left"
    );
    files[sheetKey] = strToU8(xml);

    // ── Rename sheet in workbook.xml (use string ops — no regex needed) ──
    const wbKey = "xl/workbook.xml";
    if (files[wbKey]) {
      let wbXml = strFromU8(files[wbKey]);
      // Replace name="..." in the <sheet> element
      wbXml = wbXml
        .split(`name="${OLD_SHEET}"`)
        .join(`name="${xmlEsc(sheetLabel)}"`);
      // Replace 'OLD_SHEET'! in definedName formula references
      wbXml = wbXml.split(`'${OLD_SHEET}'!`).join(`'${sheetLabel}'!`);
      files[wbKey] = strToU8(wbXml);
    }

    // ── Update chart XMLs: new sheet name + correct last data row ──────────
    const lastRow = DATA_START + rows.length - 1;
    const chartKeys = Object.keys(files).filter(
      k => k.startsWith("xl/charts/chart") && k.endsWith(".xml")
    );

    // Helper: replace non-contiguous chart range (inside parentheses) with simple range
    const flattenRange = (cxml: string, col: string): string => {
      const openTag = "<c:f>(";
      const closeTag = ")</c:f>";
      let result = cxml;
      let searchFrom = 0;
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

    chartKeys.forEach(ck => {
      let cxml = strFromU8(files[ck]);

      // 1. Replace old sheet name with new sheet name (plain string replace)
      cxml = cxml.split(`'${OLD_SHEET}'!`).join(`'${sheetLabel}'!`);

      // 2. Update end row: $X$40 → $X$lastRow (for each chart column)
      ["B", "D", "F", "G", "I"].forEach(col => {
        cxml = cxml.split(`$${col}$${LAST_DATA}`).join(`$${col}$${lastRow}`);
      });

      // 3. Simplify chart3 non-contiguous ranges to full contiguous range
      cxml = flattenRange(cxml, "B");
      cxml = flattenRange(cxml, "G");

      files[ck] = strToU8(cxml);
    });

    // Delete stale calc chain — Excel will rebuild on open
    delete files["xl/calcChain.xml"];
    if (files["[Content_Types].xml"]) {
      let ct = strFromU8(files["[Content_Types].xml"]);
      ct = ct
        .split(
          `<Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>`
        )
        .join("");
      files["[Content_Types].xml"] = strToU8(ct);
    }

    // ── Download ──────────────────────────────────────────────────────────
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
      "Performance template export failed.\n\n" +
        "Make sure Network_Performance_Report.xlsx is in your public/ folder.\n\n" +
        "Error: " +
        (err?.message ?? String(err))
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
    ...rows.map(r => [r.siteName, String(r.availHours)]),
  ];
  const downtimeChartRows: string[][] = [
    [],
    ["Site Downtime Chart Data"],
    ["Site Name", "Down (Hrs)", "Status"],
    ...rows.map(r => [
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
    .map(line =>
      line.map(cell => `"${String(cell ?? "").replace(/"/g, '""')}"`).join(",")
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
  // Sheet 1: Full performance table + KPIs
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
        horizontal
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
  // Sheet 2: Availability chart data
  const availSheet = XLSX.utils.aoa_to_sheet([
    ["Site Name", "Availability (Hrs)", "Down (Hrs)", "Reliability (%)"],
    ...rows.map(r => [
      r.siteName,
      r.availHours,
      r.sitesDownHours,
      parseFloat(r.reliability) || 100,
    ]),
  ]);
  availSheet["!cols"] = [{ wch: 32 }, { wch: 20 }, { wch: 16 }, { wch: 18 }];
  // Sheet 3: Downtime chart data — sorted worst first
  const downtimeSorted = [...rows].sort(
    (a, b) => b.sitesDownHours - a.sitesDownHours
  );
  const downtimeSheet = XLSX.utils.aoa_to_sheet([
    [
      "Site Name",
      "Down (Hrs)",
      "Availability (Hrs)",
      "Reliability (%)",
      "Status",
    ],
    ...downtimeSorted.map(r => [
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
    const perfTableWidth = 251;
    const perfTableMargin = Math.max(10, (templatePageW - perfTableWidth) / 2);
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
        C.title
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
      didParseCell: d => {
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
      didDrawPage: d => {
        templateDoc.setFont("helvetica", "normal");
        templateDoc.setFontSize(7);
        templateDoc.setTextColor(...C.muted);
        templateDoc.text(
          `Network Performance - ${templateMonthLabel} | Page ${d.pageNumber}`,
          templatePageW / 2,
          templatePageH - 5,
          { align: "center" }
        );
      },
    });

    const drawPerfChartPage = (
      title: string,
      chartRows: PerfRow[],
      valueKey: "availHours" | "sitesDownHours",
      color: [number, number, number]
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
        C.title
      );
      const chartW = 245;
      const chartH = 115;
      const chartX = (templatePageW - chartW) / 2;
      const chartY = 42;
      const items = chartRows.slice(0, 34);
      const maxValue = Math.max(
        ...items.map(row => Number(row[valueKey]) || 0),
        1
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
          { align: "center", angle: 45 }
        );
      });
      templateDoc.setFont("helvetica", "normal");
      templateDoc.setFontSize(7);
      templateDoc.setTextColor(...C.muted);
      templateDoc.text(
        `Network Performance - ${templateMonthLabel} | Page ${templateDoc.getNumberOfPages()}`,
        templatePageW / 2,
        templatePageH - 5,
        { align: "center" }
      );
    };

    drawPerfChartPage("Site Availability Chart", rows, "availHours", C.ok);
    drawPerfChartPage(
      "Site Downtime Chart",
      [...rows].sort((a, b) => b.sitesDownHours - a.sitesDownHours),
      "sitesDownHours",
      C.warn
    );
    templateDoc.save(
      `DMR-Monthly-Performance-${templateMonthLabel.replace(/ /g, "-")}.pdf`
    );
    return;
  }

  // ── Theme (mirrors PPT palette) ───────────────────────────────────────
  const C = {
    bg: [10, 22, 40] as [number, number, number],
    card: [15, 31, 56] as [number, number, number],
    card2: [10, 25, 48] as [number, number, number],
    cyan: [34, 211, 238] as [number, number, number],
    white: [248, 250, 252] as [number, number, number],
    muted: [148, 163, 184] as [number, number, number],
    green: [16, 185, 129] as [number, number, number],
    red: [220, 38, 38] as [number, number, number],
    amber: [217, 119, 6] as [number, number, number],
    border: [30, 58, 95] as [number, number, number],
  };

  // ── helpers ───────────────────────────────────────────────────────────
  const pageBg = () => {
    doc.setFillColor(...C.bg);
    doc.rect(0, 0, pageW, pageH, "F");
  };
  const accentBar = (color: [number, number, number] = C.cyan) => {
    doc.setFillColor(...color);
    doc.rect(0, 0, pageW, 2, "F");
  };
  const pageHeader = (
    title: string,
    sub: string,
    accent: [number, number, number] = C.cyan
  ) => {
    pageBg();
    accentBar(accent);
    doc.setFontSize(14);
    doc.setFont("helvetica", "bold");
    doc.setTextColor(...C.white);
    doc.text(title, 14, 12);
    doc.setFontSize(9);
    doc.setFont("helvetica", "normal");
    doc.setTextColor(...C.cyan);
    doc.text(sub, pageW - 14, 12, { align: "right" });
  };
  const pageFooter = (n: number) => {
    doc.setFontSize(7);
    doc.setFont("helvetica", "normal");
    doc.setTextColor(...C.muted);
    doc.text(
      `DMR Monthly Performance Report — ${monthLabel}  |  Page ${n}`,
      pageW / 2,
      pageH - 4,
      { align: "center" }
    );
  };

  // ── PAGE 1: Data Table ────────────────────────────────────────────────
  pageHeader(
    "DMR Monthly Performance Report",
    `Month: ${monthLabel}  ·  Sites: ${rows.length}`
  );
  doc.setFontSize(8);
  doc.setFont("helvetica", "normal");
  doc.setTextColor(...C.muted);
  doc.text(`Generated: ${new Date().toLocaleString()}`, 14, 19);

  const kpi = computePerfKPIs(rows);
  autoTable(doc, {
    startY: 22,
    head: [PERF_REPORT_HEADERS],
    body: perfReportRows(rows),
    styles: {
      fontSize: 8,
      cellPadding: 2.5,
      overflow: "linebreak",
      valign: "middle",
      fillColor: C.card,
      textColor: C.white,
      lineColor: C.border,
      lineWidth: 0.2,
    },
    headStyles: {
      fillColor: C.cyan,
      textColor: C.bg,
      fontStyle: "bold",
      fontSize: 9,
      cellPadding: 3,
    },
    alternateRowStyles: { fillColor: C.card2 },
    columnStyles: {
      0: { cellWidth: 10, halign: "center" },
      1: { cellWidth: 22, halign: "left" },
      2: { cellWidth: 40, halign: "left" },
      3: { cellWidth: 28, halign: "center" },
      4: { cellWidth: 30, halign: "center" },
      5: { cellWidth: 22, halign: "center" },
      6: { cellWidth: 30, halign: "center" },
      7: { cellWidth: 22, halign: "center", fontStyle: "bold" },
      8: { cellWidth: 22, halign: "center", fontStyle: "bold" },
    },
    margin: { left: 10, right: 10, top: 22 },
    didParseCell: d => {
      d.cell.styles.halign =
        d.column.index === 1 || d.column.index === 2 ? "left" : "center";
      d.cell.styles.valign = "middle";
      if (d.section === "body" && d.column.index === 7) {
        const v = parseFloat(String(d.cell.raw ?? "100"));
        d.cell.styles.textColor = v < 95 ? C.red : v < 99 ? C.amber : C.green;
      }
      if (d.section === "body" && d.column.index === 8) {
        const v = parseFloat(String(d.cell.raw ?? "0"));
        if (v > 24) d.cell.styles.textColor = C.red;
        else if (v > 8) d.cell.styles.textColor = C.amber;
        else if (v > 0) d.cell.styles.textColor = [234, 88, 12];
        else d.cell.styles.textColor = C.muted;
      }
    },
    didDrawPage: d => {
      if (d.pageNumber > 1) {
        // Redraw dark header area on continuation pages (table starts at margin.top:22)
        doc.setFillColor(...C.bg);
        doc.rect(0, 0, pageW, 22, "F");
        accentBar();
        doc.setFontSize(11);
        doc.setFont("helvetica", "bold");
        doc.setTextColor(...C.white);
        doc.text("DMR Monthly Performance Report", 14, 11);
        doc.setFontSize(8);
        doc.setFont("helvetica", "normal");
        doc.setTextColor(...C.cyan);
        doc.text(`${monthLabel} (cont.)`, pageW - 14, 11, { align: "right" });
      }
      pageFooter(d.pageNumber);
    },
  });

  // ── KPI PAGE (always rendered, dedicated page) ──────────────────────
  doc.addPage();
  pageHeader("KPI Summary", `Month: ${monthLabel}  ·  ${rows.length} sites`);

  // 6 KPI cards — 3 columns × 2 rows
  const kpiItems = [
    { label: "% Availability", value: kpi.pctAvailability, color: C.green },
    { label: "MTTR", value: kpi.mttr, color: C.amber },
    { label: "MTBF", value: kpi.mtbf, color: C.cyan },
    { label: "MTTF", value: kpi.mttf, color: C.cyan },
    {
      label: "Affected Sites",
      value: String(kpi.affectedSites),
      color: kpi.affectedSites > 0 ? C.red : C.green,
    },
    {
      label: "Total Down Time",
      value: kpi.totalDownHrs,
      color: kpi.totalDownHrs === "0.0 hrs" ? C.green : C.red,
    },
  ];

  const cols = 3,
    cPad = 6;
  const totalGapX = (cols - 1) * cPad;
  const cW = (pageW - 28 - totalGapX) / cols;
  const cH = 52;
  const cStartX = 14,
    cStartY = 26;

  kpiItems.forEach((item, idx) => {
    const col = idx % cols;
    const row = Math.floor(idx / cols);
    const x = cStartX + col * (cW + cPad);
    const y = cStartY + row * (cH + cPad);

    // Card background
    doc.setFillColor(...C.card);
    doc.roundedRect(x, y, cW, cH, 2, 2, "F");

    // Top colour accent bar
    doc.setFillColor(...(item.color as [number, number, number]));
    doc.roundedRect(x, y, cW, 2.5, 1, 1, "F");

    // Label
    doc.setFontSize(8);
    doc.setFont("helvetica", "normal");
    doc.setTextColor(...C.muted);
    doc.text(item.label.toUpperCase(), x + 7, y + 12);

    // Value — large
    doc.setFontSize(22);
    doc.setFont("helvetica", "bold");
    doc.setTextColor(...(item.color as [number, number, number]));
    doc.text(item.value, x + 7, y + 30);

    // Divider line
    doc.setDrawColor(...C.border);
    doc.setLineWidth(0.2);
    doc.line(x + 7, y + 35, x + cW - 7, y + 35);

    // Month sub-label
    doc.setFontSize(7);
    doc.setFont("helvetica", "normal");
    doc.setTextColor(...C.muted);
    doc.text(`Month: ${monthLabel}`, x + 7, y + 44);
  });

  pageFooter(
    (
      doc as jsPDF & { internal: { getNumberOfPages: () => number } }
    ).internal.getNumberOfPages()
  );

  // ── Chart pages helper ────────────────────────────────────────────────
  const drawChartPage = (
    title: string,
    subtitle: string,
    items: { label: string; value: number }[],
    accentColor: [number, number, number],
    colorFn: (v: number) => [number, number, number]
  ) => {
    doc.addPage();
    pageHeader(title, subtitle, accentColor);

    const mL = 24,
      mR = 18,
      mTop = 24,
      mBot = 36;
    const cW = pageW - mL - mR;
    const cH = pageH - mTop - mBot;
    const maxV = Math.max(...items.map(d => d.value), 1);
    const slotW = cW / Math.max(items.length, 1);
    const barW = Math.max(2, Math.min(slotW * 0.55, 14));

    // Chart area card bg
    doc.setFillColor(...C.card);
    doc.roundedRect(mL - 4, mTop - 4, cW + 8, cH + 8, 2, 2, "F");

    // Grid lines
    doc.setDrawColor(...C.border);
    doc.setLineWidth(0.2);
    for (let i = 0; i <= 5; i++) {
      const y = mTop + cH - (i / 5) * cH;
      doc.line(mL, y, mL + cW, y);
      const lv = Math.round(((maxV * i) / 5) * 10) / 10;
      doc.setFontSize(6.5);
      doc.setTextColor(...C.muted);
      doc.setFont("helvetica", "normal");
      doc.text(String(lv), mL - 3, y + 1.8, { align: "right" });
    }

    // Bars
    items.forEach((d, i) => {
      const bH = maxV > 0 ? (d.value / maxV) * cH : 0;
      const x = mL + i * slotW + (slotW - barW) / 2;
      const y = mTop + cH - bH;
      doc.setFillColor(...colorFn(d.value));
      if (bH > 0.5) doc.roundedRect(x, y, barW, bH, 1, 1, "F");
      if (d.value > 0) {
        doc.setFontSize(5.5);
        doc.setTextColor(...C.white);
        doc.setFont("helvetica", "bold");
        doc.text(String(d.value), x + barW / 2, y - 1.5, { align: "center" });
      }
      doc.setFontSize(5.8);
      doc.setTextColor(...C.muted);
      doc.setFont("helvetica", "normal");
      const lbl =
        d.label.length > 13 ? d.label.substring(0, 12) + "…" : d.label;
      doc.text(lbl, x + barW / 2, mTop + cH + 5, {
        align: "center",
        angle: 42,
      });
    });

    // Axis lines
    doc.setDrawColor(...C.cyan);
    doc.setLineWidth(0.5);
    doc.line(mL, mTop, mL, mTop + cH);
    doc.line(mL, mTop + cH, mL + cW, mTop + cH);

    pageFooter(
      (
        doc as jsPDF & { internal: { getNumberOfPages: () => number } }
      ).internal.getNumberOfPages()
    );
  };

  drawChartPage(
    "Site Availability — Hours",
    `Month: ${monthLabel}  ·  ${rows.length} sites`,
    rows.map(r => ({ label: r.siteName, value: r.availHours })),
    C.cyan,
    () => C.green
  );
  drawChartPage(
    "Site Downtime — Hours",
    `Month: ${monthLabel}  ·  ${rows.filter(r => r.sitesDownHours > 0).length} of ${rows.length} sites affected`,
    rows.map(r => ({ label: r.siteName, value: r.sitesDownHours })),
    C.red,
    v => (v === 0 ? C.muted : v > 24 ? C.red : v > 8 ? C.amber : [234, 88, 12])
  );

  doc.save(`DMR-Monthly-Performance-${monthLabel.replace(/ /g, "-")}.pdf`);
}

function uniqueTicketValues(
  ticket: TicketAggregate,
  field: keyof TicketRecord
): string {
  return Array.from(
    new Set(ticket.rows.map(row => clean(row[field])).filter(Boolean))
  ).join(", ");
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
    .map(line =>
      line.map(cell => `"${String(cell ?? "").replace(/"/g, '""')}"`).join(",")
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
        horizontal
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
    { wch: 34 },
  ];
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Distinct TT Report");
  XLSX.writeFile(workbook, "follow-up-distinct-tt-report.xlsx");
}

async function exportTicketTemplate(
  tickets: TicketAggregate[],
  monthKey: string
) {
  // Pure ZIP+XML approach using fflate (already in your project).
  // Only the worksheet XML is touched — xl/styles.xml, xl/drawings/, xl/media/
  // (logos, borders, fills) are never opened, so they survive byte-for-byte.

  const DATA_START = 39; // first data row
  const PROTECTED = 59; // Remarks row — never modified
  const AVAIL = PROTECTED - DATA_START; // 20 template rows
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

  // ── XML helpers ──────────────────────────────────────────────────────
  const xmlEsc = (s: string) =>
    s
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");

  // Find a cell by ref (e.g. "A39"), replace its value, keep its style attribute
  const setCell = (
    xml: string,
    ref: string,
    value: string | number
  ): string => {
    const markerIdx = xml.indexOf(` r="${ref}"`);
    if (markerIdx === -1) return xml; // cell not in XML, skip
    const cStart = xml.lastIndexOf("<c", markerIdx);
    if (cStart === -1) return xml;

    // Locate end of this cell element
    const scIdx = xml.indexOf("/>", cStart);
    const fcIdx = xml.indexOf("</c>", cStart);
    const cEnd =
      scIdx !== -1 && (fcIdx === -1 || scIdx < fcIdx) ? scIdx + 2 : fcIdx + 4;

    // Extract opening-tag attribute string, strip old type attr
    const tagClose = xml.indexOf(">", cStart);
    let attrs = xml
      .slice(cStart + 2, tagClose) // everything after "<c"
      .replace(/\s*\/$/, "") // strip trailing /
      .replace(/\s+t="[^"]*"/g, ""); // strip old type attr

    const v = String(value ?? "");
    let newCell: string;
    if (!v) {
      newCell = `<c${attrs}/>`;
    } else if (typeof value === "number") {
      newCell = `<c${attrs}><v>${v}</v></c>`;
    } else {
      newCell = `<c${attrs} t="inlineStr"><is><t>${xmlEsc(v)}</t></is></c>`;
    }
    return xml.slice(0, cStart) + newCell + xml.slice(cEnd);
  };

  try {
    // ── 1. Fetch template ───────────────────────────────────────────────
    const res = await fetch("/DMR_Monthly_Report.xlsx");
    if (!res.ok)
      throw new Error(
        `HTTP ${res.status} — make sure DMR_Monthly_Report.xlsx is inside your project's public/ folder`
      );
    const buf = await res.arrayBuffer();

    // ── 2. Unzip with fflate ────────────────────────────────────────────
    const { unzipSync, zipSync, strFromU8, strToU8 } = await import("fflate");
    const files = unzipSync(new Uint8Array(buf));

    // ── 3. Locate worksheet XML ─────────────────────────────────────────
    const sheetKey = Object.keys(files).find(k =>
      /^xl\/worksheets\/sheet\d+\.xml$/.test(k)
    );
    if (!sheetKey)
      throw new Error("Could not find worksheet XML inside template");

    let xml = strFromU8(files[sheetKey]);

    // ── 4. Month label MMM-YYYY ─────────────────────────────────────────
    const full = monthKey !== "all" ? formatMonthMMMMYYYY(monthKey) : "All";
    const parts = full.split(" ");
    const label =
      parts.length === 2 ? `${parts[0].slice(0, 3)}-${parts[1]}` : full;
    const safeName = label.replace(/[\/*?[\]:]/g, "-").slice(0, 31);

    // ── 5. Set Q5 (month value) ─────────────────────────────────────────
    xml = setCell(xml, "Q5", label);

    // ── 6. Rename sheet in workbook.xml (sheet element + definedNames refs) ──
    const wbKey = "xl/workbook.xml";
    if (files[wbKey]) {
      let wbXml = strFromU8(files[wbKey]);
      wbXml = wbXml.replace(
        /(<sheet\b[^>]*\bname=")[^"]*(")/,
        `$1${xmlEsc(safeName)}$2`
      );
      wbXml = wbXml.replace(
        /'Month-Year'!/g,
        `'${safeName.replace(/'/g, "''")}'!`
      );
      files[wbKey] = strToU8(wbXml);
    }

    // ── 7. Row insertion when more than 20 tickets ──────────────────────
    const needed = tickets.length;
    if (needed > AVAIL) {
      const extra = needed - AVAIL;

      // Capture template data row using string search — avoids RegExp escape pitfalls
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

      // Shift rows >= PROTECTED: row elements, cell refs, mergeCells, dimension
      xml = xml
        .replace(/(<row[^>]* r=")(\d+)(")/g, (_m, a, n, b) =>
          +n >= PROTECTED ? `${a}${+n + extra}${b}` : _m
        )
        .replace(/(<c[^>]* r=")([A-Z]+)(\d+)(")/g, (_m, a, col, n, b) =>
          +n >= PROTECTED ? `${a}${col}${+n + extra}${b}` : _m
        )
        .replace(
          /(<mergeCell ref=")([A-Z]+)(\d+)(:)([A-Z]+)(\d+)(")/g,
          (_m, a, c1, r1, sep, c2, r2, b) =>
            `${a}${c1}${+r1 >= PROTECTED ? +r1 + extra : +r1}${sep}${c2}${+r2 >= PROTECTED ? +r2 + extra : +r2}${b}`
        )
        .replace(
          /(<dimension ref="[A-Z]+\d+:[A-Z]+)(\d+)(")/,
          (_m, pre, n, post) =>
            +n >= PROTECTED ? `${pre}${+n + extra}${post}` : _m
        );

      // Build blank clone rows and insert before the shifted protected row
      if (tmplRow) {
        const newRows = Array.from({ length: extra }, (_, i) => {
          const rn = DATA_START + AVAIL + i;
          return (
            tmplRow
              .replace(/ r="(\d+)"/, ` r="${rn}"`) // row r attr (first only)
              .replace(
                /(<c[^>]* r=")([A-Z]+)\d+(")/g,
                (_m2, a, col, b) => `${a}${col}${rn}${b}`
              ) // all cell refs in row
              // Empty cloned cells must not keep stale values, formulas, or shared-string types.
              // Excel can repair the worksheet with "Removed Records: Cell information"
              // when blank template rows contain empty <v></v> nodes or invalid shared refs.
              .replace(/<v>[^<]*<\/v>/g, "")
              .replace(/<is>[\s\S]*?<\/is>/g, "")
              .replace(/<f[^>]*\/>/g, "")
              .replace(/<f[^>]*>[\s\S]*?<\/f>/g, "")
              .replace(/\s+t="[^"]*"/g, "")
          );
        }).join("");

        // Find the shifted protected row using plain string search
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

    // ── 8. Clear unused rows (fewer than 20 tickets) ────────────────────
    for (let r = DATA_START + needed; r < PROTECTED; r++) {
      COLS.forEach(col => {
        xml = setCell(xml, `${col}${r}`, "");
      });
    }

    // ── 9. Write ticket data ────────────────────────────────────────────
    tickets.forEach((ticket, i) => {
      const row = DATA_START + i;
      const p = ticket.primary;
      const set = (col: string, val: string | number) => {
        xml = setCell(xml, `${col}${row}`, val);
      };
      set("A", i + 1);
      set("B", Array.from(ticket.siteIds).join(", "));
      set("C", Array.from(ticket.siteNames).join(", "));
      set("D", p.managedResource || "");
      set("E", p.severity || "");
      set("F", p.issue || "");
      set("G", p.observationDate || "");
      set("H", p.observationTime || "");
      set("I", p.recoveryDate || "");
      set("J", p.recoveryTime || "");
      set("K", p.escalatedForL3SupportDate || "");
      set("L", p.escalatedForL3SupportTime || "");
      set("M", p.duration || "");
      set("N", ticket.tt || "");
      set("O", p.status || "");
      set("P", p.escalatedTo || "");
      set("Q", p.actionTaken || "");
    });

    // ── 10. Repack and download ─────────────────────────────────────────
    const ticketDataRows = Array.from(
      { length: tickets.length },
      (_, i) => DATA_START + i
    );
    xml = applyExcelColumnAlignment(
      files,
      { strFromU8, strToU8 },
      xml,
      ticketDataRows.flatMap(r =>
        ["D", "E", "G", "H", "I", "J"].map(col => `${col}${r}`)
      ),
      "center"
    );
    files[sheetKey] = strToU8(xml);
    const output = zipSync(files, { level: 0 }); // store, no recompression
    const blob = new Blob([output], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `DMR_Monthly_Report_${label}.xlsx`;
    a.click();
    URL.revokeObjectURL(url);
  } catch (err: any) {
    console.error("Template export failed:", err);
    alert(
      "Template export failed.\n\n" +
        "Check:\n" +
        "  • DMR_Monthly_Report.xlsx must be inside the  public/  folder\n" +
        "    (not the project root — the sub-folder named  public)\n\n" +
        "Error: " +
        (err?.message ?? String(err))
    );
  }
}

async function exportPdf(rows: TicketAggregate[], monthKey: string) {
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
      format: "a3",
    });
    const templatePageW = templateDoc.internal.pageSize.getWidth();
    const templatePageH = templateDoc.internal.pageSize.getHeight();
    const C = TEMPLATE_PDF_COLORS;
    const ticketTableWidth = 363;
    const ticketTableMargin = Math.max(
      12,
      (templatePageW - ticketTableWidth) / 2
    );
    const drawTemplateHeader = () => {
      templateDoc.setFillColor(255, 255, 255);
      templateDoc.rect(0, 0, templatePageW, templatePageH, "F");
      drawPdfReportHeader(
        templateDoc,
        templatePageW,
        "MONTHLY REPORT",
        "DMR SYSTEM | DMR Hytera",
        logos,
        C.title
      );
      templateDoc.setFontSize(10);
      templateDoc.setTextColor(...C.text);
      templateDoc.setFont("helvetica", "bold");
      templateDoc.text("DMR SYSTEM", 12, 26);

      const info = [
        ["Region", ""],
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
        16: { cellWidth: 44, halign: "center" },
      },
      margin: {
        left: ticketTableMargin,
        right: ticketTableMargin,
        top: 56,
        bottom: 10,
      },
      willDrawPage: drawTemplateHeader,
      didParseCell: d => {
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
      didDrawPage: d => {
        templateDoc.setFont("helvetica", "normal");
        templateDoc.setFontSize(7);
        templateDoc.setTextColor(...C.muted);
        templateDoc.text(
          `DMR Monthly Report - ${monthLabel} | Page ${d.pageNumber}`,
          templatePageW / 2,
          templatePageH - 5,
          { align: "center" }
        );
      },
    });
    const suffix =
      monthKey !== "all" ? `-${monthLabel.replace(/ /g, "-")}` : "";
    templateDoc.save(`DMR-Monthly-Tickets${suffix}.pdf`);
    return;
  }

  // ── Theme (mirrors PPT palette) ───────────────────────────────────────
  const C = {
    bg: [10, 22, 40] as [number, number, number],
    card: [15, 31, 56] as [number, number, number],
    card2: [10, 25, 48] as [number, number, number],
    cyan: [34, 211, 238] as [number, number, number],
    white: [248, 250, 252] as [number, number, number],
    muted: [148, 163, 184] as [number, number, number],
    green: [16, 185, 129] as [number, number, number],
    red: [220, 38, 38] as [number, number, number],
    amber: [217, 119, 6] as [number, number, number],
    border: [30, 58, 95] as [number, number, number],
  };

  // ── Page 1 background + header ────────────────────────────────────────
  doc.setFillColor(...C.bg);
  doc.rect(0, 0, pageW, pageH, "F");
  doc.setFillColor(...C.cyan);
  doc.rect(0, 0, pageW, 2, "F");
  doc.setFontSize(14);
  doc.setFont("helvetica", "bold");
  doc.setTextColor(...C.white);
  doc.text(`DMR Monthly Tickets Report`, 14, 12);
  doc.setFontSize(9);
  doc.setFont("helvetica", "normal");
  doc.setTextColor(...C.cyan);
  doc.text(
    `Month: ${monthLabel}  ·  Tickets: ${rows.length}  ·  ${new Date().toLocaleDateString()}`,
    pageW - 14,
    12,
    { align: "right" }
  );

  autoTable(doc, {
    startY: 18,
    head: [DISTINCT_REPORT_HEADERS],
    body: distinctReportRows(rows),
    styles: {
      fontSize: 7,
      cellPadding: 1.8,
      overflow: "linebreak",
      valign: "middle",
      fillColor: C.card,
      textColor: C.white,
      lineColor: C.border,
      lineWidth: 0.2,
    },
    headStyles: {
      fillColor: C.cyan,
      textColor: C.bg,
      fontStyle: "bold",
      fontSize: 7.5,
      cellPadding: 2.5,
    },
    alternateRowStyles: { fillColor: C.card2 },
    columnStyles: {
      0: { cellWidth: 8, halign: "center" },
      1: { cellWidth: 18 },
      2: { cellWidth: 22 },
      3: { cellWidth: 26, halign: "center" },
      4: { cellWidth: 14, halign: "center", fontStyle: "bold" },
      5: { cellWidth: 30 },
      6: { cellWidth: 18, halign: "center" },
      7: { cellWidth: 16, halign: "center" },
      8: { cellWidth: 18, halign: "center" },
      9: { cellWidth: 16, halign: "center" },
      10: { cellWidth: 22 },
      11: { cellWidth: 20 },
      12: { cellWidth: 22 },
      13: {
        cellWidth: 14,
        fontStyle: "bold",
        textColor: C.cyan as [number, number, number],
      },
      14: { cellWidth: 14, halign: "center", fontStyle: "bold" },
      15: { cellWidth: 18 },
      16: { cellWidth: 36 },
    },
    margin: { left: 10, right: 10, top: 20 },
    didParseCell: d => {
      if (d.section === "body" && d.column.index === 4) {
        const v = String(d.cell.raw ?? "").toLowerCase();
        if (v.includes("critical") || v.includes("p1"))
          d.cell.styles.textColor = C.red;
        else if (v.includes("major") || v.includes("p2"))
          d.cell.styles.textColor = C.amber;
        else if (v.includes("minor") || v.includes("p3"))
          d.cell.styles.textColor = C.green;
        else d.cell.styles.textColor = C.muted;
      }
      if (d.section === "body" && d.column.index === 14) {
        const v = String(d.cell.raw ?? "").toLowerCase();
        if (v.includes("open") || v.includes("pending"))
          d.cell.styles.textColor = C.red;
        else if (v.includes("progress") || v.includes("in-progress"))
          d.cell.styles.textColor = C.amber;
        else if (v.includes("closed") || v.includes("resolved"))
          d.cell.styles.textColor = C.green;
        else d.cell.styles.textColor = C.muted;
      }
    },
    didDrawPage: data => {
      if (data.pageNumber > 1) {
        // Redraw dark header area on continuation pages (margin.top:20 keeps rows below)
        doc.setFillColor(...C.bg);
        doc.rect(0, 0, pageW, 20, "F");
        doc.setFillColor(...C.cyan);
        doc.rect(0, 0, pageW, 2, "F");
        doc.setFontSize(11);
        doc.setFont("helvetica", "bold");
        doc.setTextColor(...C.white);
        doc.text("DMR Monthly Tickets Report", 14, 11);
        doc.setFontSize(8);
        doc.setFont("helvetica", "normal");
        doc.setTextColor(...C.cyan);
        doc.text(`${monthLabel} (cont.)`, pageW - 14, 11, { align: "right" });
      }
      doc.setFontSize(7);
      doc.setFont("helvetica", "normal");
      doc.setTextColor(...C.muted);
      doc.text(
        `DMR Monthly Tickets Report — ${monthLabel}  |  Page ${data.pageNumber}`,
        pageW / 2,
        pageH - 4,
        { align: "center" }
      );
    },
  });

  const suffix = monthKey !== "all" ? `-${monthLabel.replace(/ /g, "-")}` : "";
  doc.save(`DMR-Monthly-Tickets${suffix}.pdf`);
}


function clampGaugePercent(value: number): number {
  if (!Number.isFinite(value)) return 0;
  return Math.max(0, Math.min(100, value));
}

function extractFirstNumericValue(text: string): number | null {
  const match = clean(text).match(/-?\d+(?:\.\d+)?/);
  return match ? Number(match[0]) : null;
}

type GaugeStatus = "excellent" | "good" | "warning" | "critical";

type GaugeThresholds = {
  excellent: number;
  good: number;
  warning: number;
};

type PerformanceGaugeContext = {
  totalSites: number;
  totalHours: number;
};

type PerformanceGaugeConfig = {
  id: string;
  label: string;
  color: string;
  icon: LucideIcon;
  direction: "higher" | "lower";
  thresholds: GaugeThresholds;
  getValue: (kpi: ReturnType<typeof computePerfKPIs>, ctx: PerformanceGaugeContext) => number;
  getScale: (kpi: ReturnType<typeof computePerfKPIs>, ctx: PerformanceGaugeContext) => { min: number; max: number };
  formatValue: (kpi: ReturnType<typeof computePerfKPIs>) => string;
  caption: (kpi: ReturnType<typeof computePerfKPIs>, ctx: PerformanceGaugeContext) => string;
  helper: (kpi: ReturnType<typeof computePerfKPIs>, ctx: PerformanceGaugeContext) => string;
  sparkline: (rows: PerfRow[], kpi: ReturnType<typeof computePerfKPIs>, ctx: PerformanceGaugeContext) => number[];
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
    getValue: kpi => extractFirstNumericValue(kpi.pctAvailability) ?? 0,
    getScale: () => ({ min: 0, max: 100 }),
    formatValue: kpi => kpi.pctAvailability,
    caption: (kpi) => `${(extractFirstNumericValue(kpi.pctAvailability) ?? 0).toFixed(2)}% healthy window`,
    helper: kpi => `Available ${kpi.totalAvail.toFixed(1)} hrs across selected sites`,
    sparkline: rows => rows.map(r => {
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
    getValue: kpi => extractFirstNumericValue(kpi.mttr) ?? 0,
    getScale: () => ({ min: 0, max: 24 }),
    formatValue: kpi => kpi.mttr,
    caption: () => "Target ≤ 24 hrs mean repair time",
    helper: kpi => `${(extractFirstNumericValue(kpi.mttr) ?? 0).toFixed(2)} hrs average repair`,
    sparkline: rows => rows.filter(r => r.sitesDownHours > 0).map(r => r.sitesDownHours),
  },
  {
    id: "mtbf",
    label: "MTBF",
    color: "#3b82f6",
    icon: Network,
    direction: "higher",
    thresholds: { excellent: 168, good: 72, warning: 24 },
    getValue: kpi => extractFirstNumericValue(kpi.mtbf) ?? 0,
    getScale: () => ({ min: 0, max: 168 }),
    formatValue: kpi => kpi.mtbf,
    caption: () => "Target ≥ 168 hrs between failures",
    helper: kpi => `${(extractFirstNumericValue(kpi.mtbf) ?? 0).toFixed(2)} hrs between failures`,
    sparkline: rows => rows.map(r => r.availHours),
  },
  {
    id: "mttf",
    label: "MTTF",
    color: "#a78bfa",
    icon: ShieldAlert,
    direction: "higher",
    thresholds: { excellent: 192, good: 96, warning: 36 },
    getValue: kpi => extractFirstNumericValue(kpi.mttf) ?? 0,
    getScale: () => ({ min: 0, max: 192 }),
    formatValue: kpi => kpi.mttf,
    caption: () => "Target ≥ 192 hrs expected failure-free time",
    helper: kpi => `${(extractFirstNumericValue(kpi.mttf) ?? 0).toFixed(2)} hrs uptime horizon`,
    sparkline: rows => rows.map(r => r.availHours + Math.max(0, r.sitesDownHours * 0.35)),
  },
  {
    id: "affectedSites",
    label: "Affected Sites",
    color: "#f43f5e",
    icon: AlertTriangle,
    direction: "lower",
    thresholds: { excellent: 0, good: 5, warning: 15 },
    getValue: kpi => kpi.affectedSites,
    getScale: (_, ctx) => ({ min: 0, max: Math.max(1, ctx.totalSites) }),
    formatValue: kpi => String(kpi.affectedSites),
    caption: (kpi, ctx) => `${((kpi.affectedSites / Math.max(1, ctx.totalSites)) * 100).toFixed(1)}% of monitored sites`,
    helper: (kpi, ctx) => `${kpi.affectedSites} impacted / ${ctx.totalSites} total sites`,
    sparkline: rows => rows.map(r => (r.sitesDownHours > 0 ? 1 : 0)),
  },
  {
    id: "nonAffectedSites",
    label: "Non-Affected Sites",
    color: "#10b981",
    icon: CheckCircle2,
    direction: "higher",
    thresholds: { excellent: 90, good: 75, warning: 50 },
    getValue: (kpi, ctx) => (kpi.nonAffectedSites / Math.max(1, ctx.totalSites)) * 100,
    getScale: () => ({ min: 0, max: 100 }),
    formatValue: kpi => String(kpi.nonAffectedSites),
    caption: (kpi, ctx) => `${((kpi.nonAffectedSites / Math.max(1, ctx.totalSites)) * 100).toFixed(1)}% healthy sites`,
    helper: (kpi, ctx) => `${kpi.nonAffectedSites} stable / ${ctx.totalSites} total sites`,
    sparkline: rows => rows.map(r => (r.sitesDownHours <= 0 ? 1 : 0)),
  },
  {
    id: "totalDown",
    label: "Total Down",
    color: "#fb923c",
    icon: CloudOff,
    direction: "lower",
    thresholds: { excellent: 12, good: 48, warning: 120 },
    getValue: kpi => kpi.totalDown,
    getScale: (_, ctx) => ({ min: 0, max: Math.max(1, ctx.totalHours || 1) }),
    formatValue: kpi => kpi.totalDownHrs,
    caption: kpi => `${kpi.totalDown.toFixed(1)} hrs lost during selected window`,
    helper: (kpi, ctx) => `${((kpi.totalDown / Math.max(1, ctx.totalHours)) * 100).toFixed(2)}% downtime share`,
    sparkline: rows => rows.map(r => r.sitesDownHours),
  },
];

function gaugeStatusFromValue(
  value: number,
  direction: PerformanceGaugeConfig["direction"],
  thresholds: GaugeThresholds
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
  max: number
): number {
  const range = Math.max(1, max - min);
  const normalized = ((value - min) / range) * 100;
  return clampGaugePercent(direction === "higher" ? normalized : 100 - normalized);
}

function compactSparkline(values: number[], maxPoints = 12): number[] {
  const cleanValues = values.filter(value => Number.isFinite(value));
  if (cleanValues.length === 0) return [0, 0, 0, 0, 0];
  if (cleanValues.length <= maxPoints) return cleanValues;

  const bucketSize = cleanValues.length / maxPoints;
  return Array.from({ length: maxPoints }, (_, index) => {
    const start = Math.floor(index * bucketSize);
    const end = Math.max(start + 1, Math.floor((index + 1) * bucketSize));
    const slice = cleanValues.slice(start, end);
    return slice.reduce((sum, value) => sum + value, 0) / Math.max(1, slice.length);
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
  rows: PerfRow[]
): PerformanceGaugeCardModel[] {
  const ctx: PerformanceGaugeContext = {
    totalSites: Math.max(1, kpi.affectedSites + kpi.nonAffectedSites),
    totalHours: Math.max(1, kpi.totalAvail + kpi.totalDown),
  };

  return PERFORMANCE_GAUGE_CONFIG.map(config => {
    const value = config.getValue(kpi, ctx);
    const scale = config.getScale(kpi, ctx);
    return {
      id: config.id,
      label: config.label,
      value: config.formatValue(kpi),
      color: config.color,
      progress: gaugeProgressFromScale(value, config.direction, scale.min, scale.max),
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
        <span className={`perf-gauge-card__status perf-gauge-card__status--${status}`}>
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
            <linearGradient id={`gauge-arc-${gaugeId}`} x1="0%" x2="100%" y1="0%" y2="0%">
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
            const x1 = 120 + Math.cos(angle) * 86;
            const y1 = 140 + Math.sin(angle) * 86;
            const x2 = 120 + Math.cos(angle) * 102;
            const y2 = 140 + Math.sin(angle) * 102;
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
          ? e => {
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

function PartnerLogoStrip() {
  return (
    <div
      className="header-logo-group header-logo-group--left"
      aria-label="Saudi Energy and National Grid logos"
    >
      <img
        src={seLogoSrc}
        alt="Saudi Energy"
        className="header-logo-img se-logo"
      />
      <span className="logo-divider" aria-hidden="true" />
      <img
        src={ngLogoSrc}
        alt="National Grid SA"
        className="header-logo-img ng-logo"
      />
    </div>
  );
}

function HeaderRightLogo() {
  return (
    <div
      className="header-logo-group header-logo-group--right"
      aria-label="NASCO logo"
    >
      <img
        src={nascoLogoSrc}
        alt="NASCO"
        className="header-logo-img nasco-logo"
      />
    </div>
  );
}

/**
 * SelectFilter using a portal-based dropdown so it is never clipped
 * by overflow:hidden containers (hero panel, export card).
 */
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
    setOpen(o => !o);
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
          {options.map(opt => (
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
        document.body
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
    setOpen(o => !o);
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
    if (allSelected) {
      onChange([opt]);
      return;
    }
    if (value.includes(opt)) {
      const next = value.filter(v => v !== opt);
      // If deselecting last item, revert to All
      onChange(next);
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
              onClick={() => {
                onChange([]);
              }}
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
              ✕ Clear
            </button>
          )}
          {options.map(opt => (
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
        document.body
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

export default function Home() {
  const inputRef = useRef<HTMLInputElement | null>(null);
  const addRegionRef = useRef<HTMLInputElement | null>(null);
  const [data, setData] = useState<DashboardData | null>(null);
  const [regions, setRegions] = useState<DashboardData[]>([]);
  const [error, setError] = useState("");
  const [isDragging, setIsDragging] = useState(false);
  const statsRef = useRef<HTMLDivElement | null>(null);
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
  const [dashboardTheme, setDashboardTheme] = useState<DashboardTheme>("dark");
  const [collapsedSections, setCollapsedSections] = useState<Record<DashboardSectionId, boolean>>(INITIAL_COLLAPSED_SECTIONS);
  const [savedSnapshotAvailable, setSavedSnapshotAvailable] = useState(false);
  const activeThemeImage = THEME_IMAGES[dashboardTheme];
  const heroThemeOverlay =
    dashboardTheme === "dark"
      ? "linear-gradient(90deg, rgba(3,7,18,.94) 0%, rgba(3,7,18,.70) 42%, rgba(3,7,18,.18) 100%)"
      : "linear-gradient(90deg, rgba(248,250,252,.92) 0%, rgba(248,250,252,.68) 44%, rgba(248,250,252,.18) 100%)";
  const ribbonThemeOverlay =
    dashboardTheme === "dark"
      ? "linear-gradient(90deg, rgba(4,13,31,.88), rgba(4,13,31,.70))"
      : "linear-gradient(90deg, rgba(248,250,252,.88), rgba(226,232,240,.66))";

  // ══════════════════════════════════════════════════════════════════════════
  // TABLE COLUMN WIDTHS — edit the px values below to size each column.
  //
  //   • Each header has a clear `width =` entry.
  //   • Resizable columns: user can drag the right edge to adjust at runtime.
  //   • Wrap columns:      listed in *_WRAP_COLUMNS — no drag handle, text wraps.
  // ══════════════════════════════════════════════════════════════════════════

  // ── Tickets table — one entry per header in DISTINCT_REPORT_HEADERS ───────
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
    Action: 400, // width =
  };

  // ── Performance table — one entry per header in PERF_REPORT_HEADERS ───────
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

  // ── Runtime override maps — populated by drag handles, fall through to config above
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

  // Shared resize handler factory — wires a header name to the right state setter.
  function createResizeHandler(
    setWidths: React.Dispatch<React.SetStateAction<Record<string, number>>>,
    getCurrent: (h: string) => number
  ) {
    return (header: string) => (e: React.MouseEvent) => {
      e.preventDefault();
      e.stopPropagation();
      const startX = e.clientX;
      const startW = getCurrent(header);
      const onMove = (ev: MouseEvent) => {
        const next = Math.max(40, startW + (ev.clientX - startX));
        setWidths(prev => ({ ...prev, [header]: next }));
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
    getTicketColumnWidth
  );
  const startPerfColumnResize = createResizeHandler(
    setPerfColumnWidths,
    getPerfColumnWidth
  );

  async function handleFile(file?: File) {
    if (!file) return;
    setError("");
    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
      const parsed = parseRows(workbook, file.name);
      if (!parsed.rows.length || !parsed.uniqueTickets.length) {
        throw new Error(
          "No ticket rows with TT numbers were found. Please upload the Follow-Up Sheets workbook with the Tickets_Data sheet."
        );
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
      setError(
        err instanceof Error ? err.message : "Unable to read this workbook."
      );
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
        throw new Error(
          "No ticket rows with TT numbers were found in the additional workbook."
        );
      }
      setRegions(prev => {
        const updated = [...prev, parsed];
        const allRows = updated.flatMap(r => r.rows);
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
        // Merge site orders: use first region's order, append any new entries from additional regions
        const mergedSiteOrder = [...updated[0].siteOrder];
        const mergedIds = new Set(mergedSiteOrder.map(s => s.siteId));
        updated.slice(1).forEach(r => {
          r.siteOrder.forEach(s => {
            if (!mergedIds.has(s.siteId)) {
              mergedIds.add(s.siteId);
              mergedSiteOrder.push(s);
            }
          });
        });
        const merged: DashboardData = {
          fileName: updated.map(r => r.fileName).join(" + "),
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
      setError(
        err instanceof Error
          ? err.message
          : "Unable to read the additional workbook."
      );
    }
  }

  const uniqueRows = data?.uniqueTickets ?? [];
  const allDataRows = data?.rows ?? [];
  const filterOptions = useMemo(() => {
    const primaryRows = uniqueRows.map(ticket => ticket.primary);
    const uniq = (field: keyof TicketRecord) =>
      Array.from(
        new Set(primaryRows.map(row => clean(row[field])).filter(Boolean))
      ).sort();
    const openingMonths = Array.from(
      new Set(
        primaryRows
          .map(
            row => row.openingMonthKey || openingMonthKey(row.observationDate)
          )
          .filter(Boolean)
      )
    ).sort((a, b) => {
      if (a === "Unknown") return 1;
      if (b === "Unknown") return -1;
      return a.localeCompare(b);
    });
    const exportMonths = Array.from(
      new Set(
        uniqueRows
          .flatMap(ticket => ticket.rows.flatMap(row => coveredMonthKeys(row)))
          .filter(key => key && key !== "Unknown")
      )
    ).sort((a, b) => a.localeCompare(b));
    const rcaFamilyOptions = Array.from(
      new Set(
        primaryRows
          .map(row => row.rcaFamily || getRcaFamily(row.rca))
          .filter(Boolean)
      )
    ).sort() as string[];
    return {
      status: uniq("status"),
      severity: uniq("severity"),
      region: uniq("region"),
      impact: uniq("impact"),
      site: uniq("siteId"),
      openingMonth: openingMonths,
      openingMonthLabels: Object.fromEntries(
        openingMonths.map(key => [key, openingMonthLabel(key)])
      ),
      exportMonth: exportMonths,
      exportMonthLabels: Object.fromEntries(
        exportMonths.map(key => [key, openingMonthLabel(key)])
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
    return uniqueRows.filter(ticket => {
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
            row.openingMonthKey || openingMonthKey(row.observationDate)
          )) &&
        (!filters.site.length ||
          filters.site.some(s => ticket.siteIds.has(s))) &&
        (!filters.rcaFamily.length ||
          filters.rcaFamily.includes(row.rcaFamily || getRcaFamily(row.rca)))
      );
    });
  }, [filters, uniqueRows]);

  const monthlyExportBaseTickets = useMemo(() => {
    const q = filters.search.toLowerCase().trim();
    return uniqueRows.filter(ticket => {
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
        (!filters.severity.length || filters.severity.includes(row.severity)) &&
        (!filters.region.length || filters.region.includes(row.region)) &&
        (!filters.impact.length || filters.impact.includes(row.impact)) &&
        (!filters.site.length ||
          filters.site.some(s => ticket.siteIds.has(s))) &&
        (!filters.rcaFamily.length ||
          filters.rcaFamily.includes(row.rcaFamily || getRcaFamily(row.rca)))
      );
    });
  }, [
    filters.impact,
    filters.rcaFamily,
    filters.region,
    filters.search,
    filters.severity,
    filters.site,
    uniqueRows,
  ]);

  const monthlyExportTickets = useMemo(() => {
    // Apply export region filter
    const regionFiltered =
      exportRegions.length === 0
        ? monthlyExportBaseTickets
        : monthlyExportBaseTickets.filter(ticket =>
            exportRegions.includes(ticket.primary.region)
          );
    // Apply export month filter
    if (exportMonths.length === 0) return regionFiltered;
    return regionFiltered.filter(ticket =>
      exportMonths.some(m => ticketMatchesMonthlyExport(ticket, m))
    );
  }, [exportMonths, exportRegions, monthlyExportBaseTickets]);

  const selectedExportMonthLabel =
    exportMonths.length === 0
      ? "All export-eligible TT"
      : exportMonths.length === 1
        ? openingMonthLabel(exportMonths[0])
        : `${exportMonths.length} months`;

  // Monthly Performance rows -- computed from raw rows (not aggregated tickets)
  const siteOrder = data?.siteOrder ?? [];
  const perfRows = useMemo(() => {
    // Filter by region
    const sourceRows =
      perfRegions.length === 0
        ? allDataRows
        : allDataRows.filter(r => perfRegions.includes(r.region));
    // For multi-month: pass "all" if none selected, or the first month if one selected,
    // or compute combined rows for multiple months
    if (perfMonths.length === 0) {
      return computePerfRows(sourceRows, "all", siteOrder);
    } else if (perfMonths.length === 1) {
      return computePerfRows(sourceRows, perfMonths[0], siteOrder);
    } else {
      // Sum down hours across all selected months per site
      const combined = new Map<string, PerfRow>();
      perfMonths.forEach(mk => {
        const rows = computePerfRows(sourceRows, mk, siteOrder);
        rows.forEach(r => {
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
        0
      );
      return Array.from(combined.values()).map(r => {
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

  const performanceKpiRows = useMemo(() => {
    const q = filters.search.toLowerCase().trim();
    const sourceRows = allDataRows.filter(row => {
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
            row.openingMonthKey || openingMonthKey(row.observationDate)
          )) &&
        (!filters.site.length || filters.site.includes(row.siteId)) &&
        (!filters.rcaFamily.length ||
          filters.rcaFamily.includes(row.rcaFamily || getRcaFamily(row.rca)))
      );
    });

    if (filters.openingMonth.length === 0) {
      return computePerfRows(sourceRows, "all", siteOrder);
    }
    if (filters.openingMonth.length === 1) {
      return computePerfRows(sourceRows, filters.openingMonth[0], siteOrder);
    }

    const combined = new Map<string, PerfRow>();
    filters.openingMonth.forEach(monthKey => {
      computePerfRows(sourceRows, monthKey, siteOrder).forEach(row => {
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
      0
    );
    return Array.from(combined.values()).map(row => {
      const availHours = Math.max(0, totalMonthHours - row.sitesDownHours);
      const totalHrs = availHours + row.sitesDownHours;
      const reliability = totalHrs > 0 ? availHours / totalHrs : 1;
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
    });
  }, [allDataRows, filters, siteOrder]);

  const analytics = useMemo(() => {
    const primaryRows = filteredTickets.map(ticket => ticket.primary);
    const totalUnique = filteredTickets.length;
    const status = countBy(primaryRows, row => row.status);
    const severity = countBy(primaryRows, row => row.severity);
    const region = countBy(primaryRows, row => row.region);
    const impact = countBy(primaryRows, row => row.impact);
    const escalation = countBy(primaryRows, row => row.escalationLevel);
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
    primaryRows.forEach(row => {
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
    const replyTimeTrend = monthly.map(bucket => {
      const rowsInBucket = primaryRows.filter(row => {
        const key =
          trendGrain === "month"
            ? row.openingMonthKey || openingMonthKey(row.observationDate)
            : weekKey(row.observationDate);
        return key === bucket.key;
      });
      return {
        name: bucket.name,
        frt:
          Math.round(average(rowsInBucket.map(row => row.frtHours)) * 10) / 10,
        response:
          Math.round(average(rowsInBucket.map(row => row.responseHours)) * 10) /
          10,
        resolution:
          Math.round(
            average(rowsInBucket.map(row => row.resolutionHours)) * 10
          ) / 10,
      };
    });

    const monthlyResolutionMap = new Map<
      string,
      { sum: number; count: number; label: string }
    >();
    primaryRows.forEach(row => {
      const s = String(row.status ?? "")
        .toLowerCase()
        .trim();
      if (s !== "closed" && s !== "resolved") return;
      const hours = parseDurationHours(row.duration);
      if (hours === null || !Number.isFinite(hours) || hours <= 0) return;
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

    // Pending-age aggregate — all Pending tickets land in current calendar month.
    const nowDate = new Date();
    const currentMonthKey = `${nowDate.getFullYear()}-${String(nowDate.getMonth() + 1).padStart(2, "0")}`;
    const pendingSamples = primaryRows
      .filter(row => isPendingStatus(row.status))
      .map(row => {
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
        let idx = arr.findIndex(m => m.key === currentMonthKey);
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
          idx = arr.findIndex(m => m.key === currentMonthKey);
        }
        arr[idx].pendingAvg = Math.round(pendingAvg * 100) / 100;
        arr[idx].pendingTotal = Math.round(pendingTotal * 100) / 100;
        arr[idx].pendingCount = pendingSamples.length;
      }
      return arr;
    })();

    // Grand mean retained for any consumer that still wants a single number.
    const resolutionSamples = primaryRows
      .filter(row => {
        const s = String(row.status ?? "")
          .toLowerCase()
          .trim();
        return s === "closed" || s === "resolved";
      })
      .map(row => parseDurationHours(row.duration))
      .filter((v): v is number => v !== null && Number.isFinite(v) && v > 0);

    const avgReplyTime = {
      frt: average(primaryRows.map(row => row.frtHours)),
      response: average(primaryRows.map(row => row.responseHours)),
      resolution: resolutionSamples.length
        ? resolutionSamples.reduce((sum, v) => sum + v, 0) /
          resolutionSamples.length
        : 0,
    };

    const avgHoursSource = primaryRows
      .map(row => parseDurationHours(row.duration))
      .filter((value): value is number => value !== null);
    const avgHours = avgHoursSource.length
      ? avgHoursSource.reduce((sum, value) => sum + value, 0) /
        avgHoursSource.length
      : 0;
    const uniqueSites = new Set(
      performanceKpiRows
        .map(row => normalizeSiteId(clean(row.siteId)).toUpperCase())
        .filter(isRfSiteId)
    ).size;
    const sourceSiteSets = new Map<string, Set<string>>();
    filteredTickets.forEach(ticket => {
      ticket.rows.forEach(row => {
        const siteId = normalizeSiteId(clean(row.siteId)).toUpperCase();
        if (!isRfSiteId(siteId)) return;
        const source = row.region || row.sourceFile || "Workbook";
        if (!sourceSiteSets.has(source)) sourceSiteSets.set(source, new Set());
        sourceSiteSets.get(source)!.add(siteId);
      });
    });
    const regionSiteTotal = Array.from(sourceSiteSets.values()).reduce(
      (sum, siteIds) => sum + siteIds.size,
      0
    );
    const rootCauseUpdated = primaryRows.filter(
      row => row.actionTaken || !rcaNotProvided(row.rca)
    ).length;
    const totalSiteAffected = filteredTickets.reduce(
      (sum, ticket) =>
        sum + Math.max(ticket.siteIds.size, ticket.primary.siteId ? 1 : 0),
      0
    );
    const rcaByCount = countBy(primaryRows, row =>
      rcaNotProvided(row.rca) ? "RCA not Provided" : row.rca
    );
    const topRcaByCount = rcaByCount[0] ?? { name: "", value: 0 };
    const downtimeByRcaMap = new Map<
      string,
      { name: string; value: number; count: number }
    >();
    primaryRows.forEach(row => {
      const name = rcaNotProvided(row.rca) ? "RCA not Provided" : row.rca;
      const hours = parseDurationHours(row.duration) ?? 0;
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
      (a, b) => b.value - a.value || a.name.localeCompare(b.name)
    );
    const mttrByRca = Array.from(downtimeByRcaMap.values())
      .map(item => ({
        name: item.name,
        value: item.count ? item.value / item.count : 0,
        count: item.count,
      }))
      .sort((a, b) => b.value - a.value || a.name.localeCompare(b.name));
    const repeatedSiteRcaMap = new Map<string, number>();
    primaryRows.forEach(row => {
      const site = row.siteId || "Blank";
      const rca = rcaNotProvided(row.rca) ? "RCA not Provided" : row.rca;
      const key = `${site}||${rca}`;
      repeatedSiteRcaMap.set(key, (repeatedSiteRcaMap.get(key) ?? 0) + 1);
    });
    const repeatedRcaSites = Array.from(repeatedSiteRcaMap.values()).filter(
      value => value > 1
    ).length;
    const rcaNotProvidedCount = primaryRows.filter(row =>
      rcaNotProvided(row.rca)
    ).length;
    const preventableCount = primaryRows.filter(
      row =>
        (row.preventability || getPreventability(row.rca)) === "Preventable"
    ).length;
    const rcaFamily = countBy(
      primaryRows,
      row => row.rcaFamily || getRcaFamily(row.rca)
    );
    const siteNameById = new Map<string, string>();
    filteredTickets.forEach(ticket => {
      ticket.rows.forEach(row => {
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
    filteredTickets.forEach(ticket => {
      const sites = ticket.siteIds.size
        ? Array.from(ticket.siteIds)
        : [ticket.primary.siteId || "Blank"];
      sites.forEach(site => {
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
    ].filter(item => item.value > 0);

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
    primaryRows.forEach(row => {
      const key = row.openingMonthKey || openingMonthKey(row.observationDate);
      if (!key || key === "Unknown") return;
      const family = row.rcaFamily || getRcaFamily(row.rca);
      if (!monthlyRcaFamilyMap.has(key)) {
        const entry: Record<string, string | number> = {
          monthKey: key,
          name: openingMonthLabel(key),
        };
        RCA_FAMILIES.forEach(f => {
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
      row => row.managedResource || "Unknown"
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
      rcaByDowntime: downtimeByRca.map(item => ({
        name: item.name,
        value: Math.round(item.value * 10) / 10,
      })),
      rcaByMttr: mttrByRca.map(item => ({
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
    analytics.totalUnique
  );
  const preventableRcaPct = pct(
    analytics.preventableCount,
    analytics.totalUnique
  );

  const executiveInsights = useMemo(
    () =>
      calculateExecutiveInsights({
        tickets: filteredTickets,
        performanceRows: performanceKpiRows,
      }),
    [filteredTickets, performanceKpiRows]
  );

  const deepDiveAnalytics = useMemo(
    () =>
      calculateDeepDiveAnalytics({
        tickets: filteredTickets,
        performanceRows: performanceKpiRows,
      }),
    [filteredTickets, performanceKpiRows]
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
    const closedOrResolved = monthlyExportTickets.filter(ticket => {
      const status = clean(ticket.primary.status).toLowerCase();
      return status === "closed" || status === "resolved";
    }).length;
    const pendingCount = monthlyExportTickets.filter(
      ticket => clean(ticket.primary.status).toLowerCase() === "pending"
    ).length;
    const criticalCount = monthlyExportTickets.filter(
      ticket => clean(ticket.primary.severity).toLowerCase() === "critical"
    ).length;
    return { closedOrResolved, pendingCount, criticalCount };
  }, [monthlyExportTickets]);

  const performanceExportMetrics = useMemo(() => {
    const scopedSourceRows = allDataRows.filter(row => {
      const regionMatch =
        perfRegions.length === 0 || perfRegions.includes(row.region);
      const monthMatch =
        perfMonths.length === 0 ||
        coveredMonthKeys(row).some(monthKey => perfMonths.includes(monthKey));
      return regionMatch && monthMatch;
    });
    const scopedSiteIds = new Set(
      scopedSourceRows.map(row => clean(row.siteId)).filter(isRfSiteId)
    );
    const rfPerfRows = scopedSiteIds.size
      ? perfRows.filter(row => scopedSiteIds.has(clean(row.siteId)))
      : perfRows.filter(row => isRfSiteId(row.siteId));
    const rfSiteIds = new Set(
      rfPerfRows.map(row => clean(row.siteId)).filter(isRfSiteId)
    );
    const affectedIds = new Set(
      rfPerfRows
        .filter(row => row.sitesDownHours > 0)
        .map(row => clean(row.siteId))
        .filter(isRfSiteId)
    );
    return {
      totalSites: rfSiteIds.size,
      affectedSites: affectedIds.size,
      nonAffectedSites: Math.max(0, rfSiteIds.size - affectedIds.size),
      reportRows: rfPerfRows.length,
    };
  }, [allDataRows, perfMonths, perfRegions, perfRows]);


  const allSectionsCollapsed = DASHBOARD_SECTIONS.every(section => collapsedSections[section.id]);

  function makeFileSafeName(value: string) {
    return clean(value).replace(/[^a-z0-9]+/gi, "_").replace(/^_+|_+$/g, "") || "dashboard_export";
  }

  function serializeDashboardData(snapshot: DashboardData) {
    return {
      fileName: snapshot.fileName,
      sheetName: snapshot.sheetName,
      generatedAt: snapshot.generatedAt,
      rows: snapshot.rows,
      siteOrder: snapshot.siteOrder,
    };
  }

  function hydrateDashboardData(snapshot: ReturnType<typeof serializeDashboardData>): DashboardData {
    return {
      ...snapshot,
      uniqueTickets: groupTickets(snapshot.rows),
    };
  }

  function loadSavedDashboardSnapshot() {
    try {
      const raw = localStorage.getItem(SAVED_DASHBOARD_KEY);
      if (!raw) return;
      const parsed = JSON.parse(raw) as {
        data?: ReturnType<typeof serializeDashboardData>;
        regions?: Array<ReturnType<typeof serializeDashboardData>>;
      };
      if (!parsed.data?.rows?.length) return;
      const restoredData = hydrateDashboardData(parsed.data);
      const restoredRegions = (parsed.regions ?? []).map(hydrateDashboardData);
      setData(restoredData);
      setRegions(restoredRegions.length ? restoredRegions : [restoredData]);
      setError("");
      setFilters(EMPTY_FILTERS);
      setTablePage(1);
      setPerfPage(1);
    } catch (err) {
      console.error("Failed to restore previous dashboard snapshot:", err);
      setError("Could not restore the previous workbook session. Please upload the workbook again.");
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
    setSavedSnapshotAvailable(Boolean(localStorage.getItem(SAVED_DASHBOARD_KEY)));
    window.scrollTo({ top: 0, behavior: "smooth" });
  }

  function scrollToDashboardTop() {
    const nav = document.querySelector<HTMLElement>("#dashboard-section-nav");
    if (nav) {
      nav.scrollIntoView({ behavior: "smooth", block: "start" });
      return;
    }

    window.scrollTo({ top: 0, behavior: "smooth" });
  }

  function scrollToDashboardSection(sectionId: DashboardSectionId) {
    const definition = DASHBOARD_SECTIONS.find(section => section.id === sectionId);
    if (!definition) return;

    const sectionHeader = document.querySelector<HTMLElement>(`#section-control-${definition.id}`);
    const target = sectionHeader ?? document.querySelector<HTMLElement>(definition.selector);
    target?.scrollIntoView({ behavior: "smooth", block: "start" });
  }

  function toggleDashboardSection(sectionId: DashboardSectionId) {
    setCollapsedSections(prev => ({ ...prev, [sectionId]: !prev[sectionId] }));
  }

  function setAllDashboardSections(collapsed: boolean) {
    setCollapsedSections({
      reports: collapsed,
      performanceKpis: collapsed,
      kpis: collapsed,
      executive: collapsed,
      deepDive: collapsed,
      overviewCharts: collapsed,
      trendCharts: collapsed,
      ticketsTable: collapsed,
      performanceTable: collapsed,
    });
  }

  async function exportElementToPng(element: HTMLElement | null, fileName: string) {
    if (!element) return;
    const { default: html2canvas } = await import("html2canvas");
    try {
      const canvas = await html2canvas(element, {
        backgroundColor: null,
        scale: Math.min(2, window.devicePixelRatio || 1.5),
        useCORS: true,
        logging: false,
        ignoreElements: node =>
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
      setError("PNG export failed. Please try again after the section finishes rendering.");
    }
  }

  useEffect(() => {
    setSavedSnapshotAvailable(Boolean(localStorage.getItem(SAVED_DASHBOARD_KEY)));
  }, []);

  useEffect(() => {
    if (!data) return;
    try {
      localStorage.setItem(
        SAVED_DASHBOARD_KEY,
        JSON.stringify({
          data: serializeDashboardData(data),
          regions: regions.map(serializeDashboardData),
        })
      );
      setSavedSnapshotAvailable(true);
    } catch (err) {
      console.warn("Could not save dashboard snapshot:", err);
    }
  }, [data, regions]);

  useEffect(() => {
    if (!data) return;
    document.querySelectorAll(".section-control-panel").forEach(panel => panel.remove());
    const cleanup: Array<() => void> = [];

    DASHBOARD_SECTIONS.forEach(section => {
      const target = document.querySelector<HTMLElement>(section.selector);
      if (!target) return;
      target.classList.add("dashboard-section-content-block");
      target.classList.toggle("dashboard-section-collapsed", collapsedSections[section.id]);

      const panel = document.createElement("div");
      panel.id = `section-control-${section.id}`;
      panel.className = `section-control-panel no-print${collapsedSections[section.id] ? " section-control-panel-collapsed" : ""}`;
      panel.innerHTML = `
        <div class="section-control-title"><span>${section.title}</span></div>
        <div class="section-control-actions">
          <button type="button" class="section-tool-button" data-action="png">Export PNG</button>
          <button type="button" class="section-tool-button" data-action="toggle">${collapsedSections[section.id] ? "Expand" : "Collapse"}</button>
          <button type="button" class="section-tool-button" data-action="top">Top Nav</button>
        </div>
      `;

      const pngButton = panel.querySelector<HTMLButtonElement>('[data-action="png"]');
      const toggleButton = panel.querySelector<HTMLButtonElement>('[data-action="toggle"]');
      const topButton = panel.querySelector<HTMLButtonElement>('[data-action="top"]');
      const pngHandler = () => exportElementToPng(target, section.title);
      const toggleHandler = () => toggleDashboardSection(section.id);
      const topHandler = () => scrollToDashboardTop();
      pngButton?.addEventListener("click", pngHandler);
      toggleButton?.addEventListener("click", toggleHandler);
      topButton?.addEventListener("click", topHandler);
      target.parentElement?.insertBefore(panel, target);

      cleanup.push(() => {
        pngButton?.removeEventListener("click", pngHandler);
        toggleButton?.removeEventListener("click", toggleHandler);
        topButton?.removeEventListener("click", topHandler);
        panel.remove();
      });
    });

    return () => cleanup.forEach(remove => remove());
  }, [data, collapsedSections]);

  useEffect(() => {
    if (!data) return;
    document.querySelectorAll(".chart-export-png-button").forEach(button => button.remove());
    const chartHosts = new Set<HTMLElement>();

    const explicitChartSelector = [
      ".dashboard-chart-grid > .glass-card",
      ".chart-mosaic > .glass-card",
      ".deep-dive-chart-grid .deep-dive-panel",
      ".client-delivery-section .deep-dive-panel",
      ".glass-card:has(.recharts-wrapper)",
      ".glass-card:has(.recharts-surface)",
      ".deep-dive-panel:has(.recharts-wrapper)",
      ".deep-dive-panel:has(.recharts-surface)",
    ].join(",");

    document.querySelectorAll<HTMLElement>(explicitChartSelector).forEach(card => {
      chartHosts.add(card);
    });

    document.querySelectorAll<HTMLElement>(".recharts-wrapper, .recharts-surface").forEach(chart => {
      const host = chart.closest<HTMLElement>(
        ".glass-card, .deep-dive-panel, .chart-card, article, section"
      );
      if (host) chartHosts.add(host);
    });

    const cleanup: Array<() => void> = [];
    Array.from(chartHosts).forEach((card, index) => {
      if (card.querySelector(".chart-export-png-button")) return;
      card.classList.add("chart-export-target");
      const title = card.querySelector("h3, h4, strong, .card-heading")?.textContent?.trim() ?? `Chart ${index + 1}`;
      const button = document.createElement("button");
      button.type = "button";
      button.className = "chart-export-png-button no-print";
      button.title = `Export ${title} as PNG`;
      button.textContent = "PNG";
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
    return () => cleanup.forEach(remove => remove());
  }, [data, analytics, executiveInsights, deepDiveAnalytics, collapsedSections]);

  const filtersPanel = data ? (
    <section
      ref={filtersRef}
      className="filters-panel no-print dashboard-filters-panel"
      style={{
        width: "100%",
        display: "grid",
        gridTemplateColumns: "1.5fr repeat(7, 1fr) max-content",
        gap: "10px",
        alignItems: "end",
        paddingBottom: "4px",
      }}
    >
      <label className="search-box" style={{ width: "100%", margin: 0 }}>
        <Search size={16} />
        <input
          value={filters.search}
          onChange={event =>
            setFilters(prev => ({ ...prev, search: event.target.value }))
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
          onChange={value => setFilters(prev => ({ ...prev, status: value }))}
        />
      </div>
      <div style={{ width: "100%", overflow: "hidden" }}>
        <MultiSelectFilter
          label="Severity"
          value={filters.severity}
          options={filterOptions.severity}
          onChange={value => setFilters(prev => ({ ...prev, severity: value }))}
        />
      </div>
      <div style={{ width: "100%", overflow: "hidden" }}>
        <MultiSelectFilter
          label="Region"
          value={filters.region}
          options={filterOptions.region}
          onChange={value => setFilters(prev => ({ ...prev, region: value }))}
        />
      </div>
      <div style={{ width: "100%", overflow: "hidden" }}>
        <MultiSelectFilter
          label="Impact"
          value={filters.impact}
          options={filterOptions.impact}
          onChange={value => setFilters(prev => ({ ...prev, impact: value }))}
        />
      </div>
      <div style={{ width: "100%", overflow: "hidden" }}>
        <MultiSelectFilter
          label="Opening Month"
          value={filters.openingMonth}
          options={filterOptions.openingMonth}
          optionLabels={filterOptions.openingMonthLabels}
          onChange={value =>
            setFilters(prev => ({ ...prev, openingMonth: value }))
          }
        />
      </div>
      <div style={{ width: "100%", overflow: "hidden" }}>
        <MultiSelectFilter
          label="Site"
          value={filters.site}
          options={filterOptions.site}
          onChange={value => setFilters(prev => ({ ...prev, site: value }))}
        />
      </div>
      <div style={{ width: "100%", overflow: "hidden" }}>
        <MultiSelectFilter
          label="RCA Family"
          value={filters.rcaFamily}
          options={filterOptions.rcaFamily}
          onChange={value =>
            setFilters(prev => ({ ...prev, rcaFamily: value }))
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

  return (
    <main className="dashboard-shell" data-dashboard-theme={dashboardTheme}>
      <section
        className="hero-panel"
        style={{
          backgroundImage: `${heroThemeOverlay}, url(${activeThemeImage})`,
          backgroundSize: "contain",
          backgroundPosition: "top",
          backgroundRepeat: "no-repeat",
          paddingTop: "12px",
        }}
      >
        <nav className="topbar no-print" style={{ marginBottom: "8px" }}>
          <div className="topbar-brand-row">
            <PartnerLogoStrip />
            <div className="brand-title-center">
              <span>DMR Ticketing Dashboard</span>
            </div>
            <HeaderRightLogo />
          </div>
          <div className="topbar-actions">
            <div className="theme-toggle no-print" aria-label="Dashboard theme selector">
              <button
                type="button"
                className="active"
                onClick={() =>
                  setDashboardTheme(prev => (prev === "dark" ? "light" : "dark"))
                }
                aria-pressed={dashboardTheme === "light"}
                aria-label={`Switch to ${dashboardTheme === "dark" ? "light" : "dark"} theme`}
                title={`Switch to ${dashboardTheme === "dark" ? "Light" : "Dark"} Theme`}
              >
                {dashboardTheme === "dark" ? (
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
            {data && (
              <button
                type="button"
                className="ghost-button"
                onClick={returnToWelcomeUploadScreen}
                title="Return to the welcome upload screen"
              >
                <HomeIcon size={30} /> Home
              </button>
            )}
            {data && (
              <button
                className="ghost-button"
                onClick={() => addRegionRef.current?.click()}
              >
                <UploadCloud size={30} /> Add region
              </button>
            )}
            {data && (
              <button
                className="ghost-button"
                onClick={() => inputRef.current?.click()}
              >
                <RefreshCw size={30} /> New workbook
              </button>
            )}
            {data && (
              <button className="primary-button" onClick={() => window.print()}>
                <Printer size={30} /> Dashboard PDF
              </button>
            )}
          </div>
        </nav>
        {data && (
          <div id="dashboard-section-nav" className="dashboard-section-nav no-print" aria-label="Dashboard section navigation">
            <div className="section-nav-buttons">
              {DASHBOARD_SECTIONS.map(section => (
                <button
                  key={section.id}
                  type="button"
                  className="section-nav-button"
                  onClick={() => scrollToDashboardSection(section.id)}
                >
                  {section.label}
                </button>
              ))}
            </div>
            <button
              type="button"
              className="section-nav-toggle"
              onClick={() => setAllDashboardSections(!allSectionsCollapsed)}
            >
              {allSectionsCollapsed ? <Maximize2 size={15} /> : <Minimize2 size={15} />}
              {allSectionsCollapsed ? "Expand All" : "Collapse All"}
            </button>
          </div>
        )}
        {filtersPanel}

        {data && (
          <>
            <div
              id="section-reports"
              className="hero-export-row no-print export-row-dual dashboard-section-content-block"
              style={{
                display: "grid",
                gridTemplateColumns: "repeat(auto-fit, minmax(360px, 1fr))",
                gap: "18px",
                justifyContent: "center",
                width: "100%",
                marginBottom: "-12px", // Pulls up the next sibling element (KPIs)
                marginTop: "0px",
              }}
            >
              {/* ===== CARD 1: Monthly Tickets Table ===== */}
              <aside
                className="hero-export-card hero-export-card--5col report-export-card report-export-card--tickets"
                style={{
                  display: "flex",
                  flexDirection: "column",
                  width: "100%",
                }}
              >
                <div
                  className="hero-export-copy"
                  style={{
                    marginBottom: "2px",
                    display: "flex",
                    flexDirection: "column",
                    alignItems: "center",
                    textAlign: "center",
                    width: "100%",
                  }}
                >
                  <span>Monthly Tickets Export</span>

                  <strong>
                    Total Tickets —{" "}
                    {exportMonths.length === 0
                      ? "All Months"
                      : exportMonths.length === 1
                        ? (filterOptions.exportMonthLabels?.[exportMonths[0]] ??
                          exportMonths[0])
                        : `${exportMonths.length} months`}
                  </strong>
                </div>

                {/* Report card chips paused until final layout decision.
<div className="report-export-metrics">
  <span><b>{monthlyExportTickets.length}</b>Total TT</span>
  <span><b>{ticketExportMetrics.closedOrResolved}</b>Resolved</span>
  <span><b>{ticketExportMetrics.pendingCount}</b>Pending</span>
  <span><b>{ticketExportMetrics.criticalCount}</b>Critical</span>
</div>
*/}

                <div className="report-export-filters">
                  <div style={{ flex: 1, minWidth: 0 }}>
                    <MultiSelectFilter
                      label="Report Month"
                      value={exportMonths}
                      options={filterOptions.exportMonth}
                      optionLabels={filterOptions.exportMonthLabels}
                      onChange={setExportMonths}
                      showAllOption
                    />
                  </div>
                  <div style={{ flex: 1, minWidth: 0 }}>
                    <MultiSelectFilter
                      label="Region"
                      value={exportRegions}
                      options={filterOptions.region}
                      onChange={setExportRegions}
                      showAllOption
                    />
                  </div>
                </div>

                {/* Swapped marginTop auto to 4px to force-close the layout gap */}
                <div
                  className="hero-export-actions report-export-actions"
                  style={{
                    display: "flex",
                    justifyContent: "center",
                    alignItems: "center",
                    width: "100%",
                  }}
                >
                  <button
                    className="ghost-button"
                    onClick={() =>
                      exportTicketTemplate(
                        monthlyExportTickets,
                        exportMonths[0] ?? "all"
                      )
                    }
                  >
                    <FileSpreadsheet size={16} /> Excel
                  </button>
                  <button
                    className="ghost-button"
                    onClick={() =>
                      exportPdf(monthlyExportTickets, exportMonths[0] ?? "all")
                    }
                  >
                    <Printer size={16} /> PDF
                  </button>
                  {/* ====== ADD YOUR NEW PPT BUTTON HERE ====== */}
                  <button
                    className="ghost-button"
                    onClick={() => {
                      const currentMonthLabel =
                        exportMonths.length === 1
                          ? (filterOptions.exportMonthLabels?.[
                              exportMonths[0]
                            ] ?? exportMonths[0])
                          : "All";
                      exportTicketsPpt(
                        monthlyExportTickets,
                        currentMonthLabel,
                        executiveInsights,
                        deepDiveAnalytics
                      );
                    }}
                    style={{
                      display: "flex",
                      alignItems: "center",
                      gap: "8px",
                    }}
                  >
                    {" "}
                    <Presentation size={16} /> PPT{" "}
                  </button>
                </div>
              </aside>

              {/* ===== CARD 2: Monthly Performance Table ===== */}
              <aside
                className="hero-export-card hero-export-card--5col report-export-card report-export-card--performance"
                style={{
                  display: "flex",
                  flexDirection: "column",
                  width: "100%",
                }}
              >
                <div
                  className="hero-export-copy"
                  style={{
                    marginBottom: "2px",
                    display: "flex",
                    flexDirection: "column",
                    alignItems: "center",
                    textAlign: "center",
                    width: "100%",
                  }}
                >
                  <span>Monthly Performance</span>
                  <strong>
                    Sites Performance —{" "}
                    {perfMonths.length === 0
                      ? "All Months"
                      : perfMonths.length === 1
                        ? (filterOptions.exportMonthLabels?.[perfMonths[0]] ??
                          perfMonths[0])
                        : `${perfMonths.length} months`}
                  </strong>
                </div>

                {/* Report card chips paused until final layout decision.
    <div className="report-export-metrics">
      <span><b>{performanceExportMetrics.totalSites}</b>Total RF Sites</span>
      <span><b>{performanceExportMetrics.affectedSites}</b>Affected</span>
      <span><b>{performanceExportMetrics.nonAffectedSites}</b>Non-Affected</span>
      <span><b>{performanceExportMetrics.reportRows}</b>Report Rows</span>
    </div>
    */}

                <div className="report-export-filters">
                  <div style={{ flex: 1, minWidth: 0 }}>
                    <MultiSelectFilter
                      label="Report Month"
                      value={perfMonths}
                      options={filterOptions.exportMonth}
                      optionLabels={filterOptions.exportMonthLabels}
                      onChange={setPerfMonths}
                      showAllOption
                    />
                  </div>

                  <div style={{ flex: 1, minWidth: 0 }}>
                    <MultiSelectFilter
                      label="Region"
                      value={perfRegions}
                      options={filterOptions.region}
                      onChange={setPerfRegions}
                      showAllOption
                    />
                  </div>
                </div>

                <div
                  className="hero-export-actions report-export-actions"
                  style={{
                    display: "flex",
                    justifyContent: "center",
                    alignItems: "center",
                    width: "100%",
                  }}
                >
                  <button
                    className="ghost-button"
                    onClick={() =>
                      exportPerfTemplate(
                        perfRows,
                        perfMonths[0] ?? "all",
                        perfRegions
                      )
                    }
                  >
                    <FileSpreadsheet size={16} /> Excel
                  </button>
                  <button
                    className="ghost-button"
                    onClick={() =>
                      exportPerfPdf(perfRows, perfMonths[0] ?? "all")
                    }
                  >
                    <Printer size={16} /> PDF
                  </button>
                  <button
                    className="ghost-button"
                    onClick={() => {
                      const lbl =
                        perfMonths.length === 1
                          ? (filterOptions.exportMonthLabels?.[perfMonths[0]] ??
                            perfMonths[0])
                          : perfMonths.length > 1
                            ? `${perfMonths.length} months`
                            : "All";
                      exportPerfPpt(perfRows, lbl, executiveInsights);
                    }}
                  >
                    <Presentation size={16} /> PPT
                  </button>
                </div>
              </aside>
            </div>

            {/* KPI Summary tiles -- live with dashboard filters */}
            {performanceKpiRows.length > 0 &&
              (() => {
                const kpi = computePerfKPIs(performanceKpiRows);
                return (
                  <div
                    id="section-performance-kpis"
                    className="hero-export-row no-print"
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
                          performanceKpiRows
                        );

                        return gaugeCards.map((card, index) => (
                          <PerformanceGaugeCard key={card.id} {...card} index={index} />
                        ));
                      })()}
                    </div>
                  </div>
                );
              })()}
          </>
        )}
      </section>

      <input
        ref={inputRef}
        className="sr-only"
        type="file"
        accept=".xlsx,.xls,.xlsm"
        onChange={event => handleFile(event.target.files?.[0])}
      />
      <input
        ref={addRegionRef}
        className="sr-only"
        type="file"
        accept=".xlsx,.xls,.xlsm"
        onChange={event => handleAddRegion(event.target.files?.[0])}
      />

      {!data ? (
        <section className="upload-stage no-print">
          <div
            className={`upload-card ${isDragging ? "dragging" : ""}`}
            onDragOver={event => {
              event.preventDefault();
              setIsDragging(true);
            }}
            onDragLeave={() => setIsDragging(false)}
            onDrop={event => {
              event.preventDefault();
              setIsDragging(false);
              handleFile(event.dataTransfer.files?.[0]);
            }}
          >
            <div className="upload-copy">
              <span className="section-kicker">
                <UploadCloud size={14} /> Workbooks Upload
              </span>
              <h2>Load the tickets workbook</h2>
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
                <span>or drop the workbook here</span>
              </div>
              <div className="upload-checks" aria-label="Workbook readiness">
                <div>
                  <CheckCircle2 size={18} />
                  <span>Tickets_Data</span>
                </div>
                <div>
                  <Layers3 size={18} />
                  <span>Regional merge</span>
                </div>
                <div>
                  <BarChart3 size={18} />
                  <span>Charts & reports</span>
                </div>
              </div>
            </div>
            <div className="upload-visual">
              <img src={activeThemeImage} alt="Dashboard workbook preview" />
              <div className="upload-preview-card">
                <span>Ready For</span>
                <strong>Excel · PDF · PPT</strong>
                <small>
                  Tickets, performance, RCA, and site availability reports
                </small>
              </div>
            </div>
          </div>
        </section>
      ) : (
        <>
          <section
            ref={statsRef}
            id="section-kpis"
            className="stats-grid workbook-cards dashboard-section-content-block"
            style={{
              backgroundImage: `${ribbonThemeOverlay}, url(${activeThemeImage})`,
              display: "grid",
              // 48 columns lets every card span 6 tracks: 8 cards per row, 2 rows total.
              gridTemplateColumns: "repeat(48, 1fr)",
              gap: "16px",
              width: "100%",
              boxSizing: "border-box",
              textAlign: "center",
              alignItems: "stretch",
            }}
          >
            {/* Standard StatCards */}
            <StatCard
              label={
                <span
                  style={{
                    fontSize: "30px",
                    fontWeight: 1000,
                    color: "#00ff15",
                  }}
                >{`Total TT's`}</span>
              }
              value={
                <span
                  style={{
                    fontSize: "30px",
                    fontWeight: "bold",
                    color: "#ff0000",
                  }}
                >
                  {analytics.totalUnique.toLocaleString()}
                </span>
              }
              // note={<span style={{ fontSize: "14px", fontWeight: 500, color: "#00ff15" }}> TT's Opened</span>}
              icon={Layers3}
              tone="#fff200"
              onClick={() => {
                setFilters(EMPTY_FILTERS);
                setTablePage(1);
              }}
              style={{ gridColumn: "span 6" }}
            />

            <StatCard
              label={
                <span
                  style={{
                    fontSize: "16px",
                    fontWeight: 1000,
                    color: "#00ff15",
                  }}
                >{`Closed TT's`}</span>
              }
              value={
                <span
                  style={{
                    fontSize: "30px",
                    fontWeight: "bold",
                    color: "#ff0000",
                  }}
                >
                  {closed.toLocaleString()}
                </span>
              }
              // note={<span style={{ fontSize: "14px", fontWeight: 1000, color: "#00ff15" }}>Closed TT's</span>}
              icon={CheckCircle2}
              tone="#34d399"
              onClick={() => {
                setFilters({ ...EMPTY_FILTERS, status: ["Closed"] });
                setTablePage(1);
              }}
              style={{ gridColumn: "span 6" }}
            />

            <StatCard
              label={
                <span
                  style={{
                    fontSize: "16px",
                    fontWeight: 1000,
                    color: "#00ff15",
                  }}
                >{`Pending TT's`}</span>
              }
              value={
                <span
                  style={{
                    fontSize: "30px",
                    fontWeight: "bold",
                    color: "#ff0000",
                  }}
                >
                  {pending.toLocaleString()}
                </span>
              }
              // note={<span style={{ fontSize: "14px", fontWeight: 1000, color: "#00ff15" }}>Pending TT's</span>}
              icon={ShieldAlert}
              tone="#f59e0b"
              onClick={() => {
                setFilters({ ...EMPTY_FILTERS, status: ["Pending"] });
                setTablePage(1);
              }}
              style={{ gridColumn: "span 6" }}
            />

            <StatCard
              label={
                <span
                  style={{
                    fontSize: "16px",
                    fontWeight: 1000,
                    color: "#00ff15",
                  }}
                >{`Resolved TT's`}</span>
              }
              value={
                <span
                  style={{
                    fontSize: "30px",
                    fontWeight: "bold",
                    color: "#ff0000",
                  }}
                >
                  {resolved.toLocaleString()}
                </span>
              }
              // note={<span style={{ fontSize: "14px", fontWeight: 1000, color: "#00ff15" }}>Resolved TT's</span>}
              icon={CheckCircle2}
              tone="#60a5fa"
              onClick={() => {
                setFilters({ ...EMPTY_FILTERS, status: ["Resolved"] });
                setTablePage(1);
              }}
              style={{ gridColumn: "span 6" }}
            />

            <StatCard
              label={
                <span
                  style={{
                    fontSize: "16px",
                    fontWeight: 1000,
                    color: "#00ff15",
                  }}
                >{`Critical TT's`}</span>
              }
              value={
                <span
                  style={{
                    fontSize: "30px",
                    fontWeight: "bold",
                    color: "#ff0000",
                  }}
                >
                  {critical.toLocaleString()}
                </span>
              }
              // note={<span style={{ fontSize: "14px", fontWeight: 1000, color: "#00ff15" }}>Critical TT's</span>}
              icon={AlertTriangle}
              tone="#ef4444"
              onClick={() => {
                setFilters({ ...EMPTY_FILTERS, severity: ["Critical"] });
                setTablePage(1);
              }}
              style={{ gridColumn: "span 6" }}
            />

            <StatCard
              label={
                <span
                  style={{
                    fontSize: "16px",
                    fontWeight: 1000,
                    color: "#00ff15",
                  }}
                >{`Major TT's`}</span>
              }
              value={
                <span
                  style={{
                    fontSize: "30px",
                    fontWeight: "bold",
                    color: "#ff0000",
                  }}
                >
                  {major.toLocaleString()}
                </span>
              }
              // note={<span style={{ fontSize: "14px", fontWeight: 1000, color: "#00ff15" }}>Major TT's</span>}
              icon={Activity}
              tone="#f59e0b"
              onClick={() => {
                setFilters({ ...EMPTY_FILTERS, severity: ["Major"] });
                setTablePage(1);
              }}
              style={{ gridColumn: "span 6" }}
            />

            <StatCard
              label={
                <span
                  style={{
                    fontSize: "16px",
                    fontWeight: 1000,
                    color: "#00ff15",
                  }}
                >{`Minor TT's`}</span>
              }
              value={
                <span
                  style={{
                    fontSize: "30px",
                    fontWeight: "bold",
                    color: "#ff0000",
                  }}
                >
                  {minor !== undefined && minor !== null
                    ? minor.toLocaleString()
                    : "0"}
                </span>
              }
              // note={<span style={{ fontSize: "14px", fontWeight: 1000, color: "#00ff15" }}>Minor TT's</span>}
              icon={CircleDot}
              tone="#22d3ee"
              onClick={() => {
                setFilters({ ...EMPTY_FILTERS, severity: ["Minor"] });
                setTablePage(1);
              }}
              style={{ gridColumn: "span 6" }}
            />

            <StatCard
              label={
                <span
                  style={{
                    fontSize: "16px",
                    fontWeight: 1000,
                    color: "#00ff15",
                  }}
                >{`Region Sites`}</span>
              }
              value={
                <span
                  style={{
                    fontSize: "30px",
                    fontWeight: 1000,
                    color: "#ff0000",
                  }}
                >
                  {analytics.regionSiteTotal.toLocaleString()}
                </span>
              }
              // note={<span style={{ fontSize: "14px", fontWeight: 1000, color: "#00ff15" }}>{`${analytics.uniqueSites.toLocaleString()} Unique RF Sites`}</span>}
              icon={BarChart3}
              tone="#60a5fa"
              style={{ gridColumn: "span 6" }}
            />

            <StatCard
              label={
                <span
                  style={{
                    fontSize: "16px",
                    fontWeight: 1000,
                    color: "#00ff15",
                  }}
                >{`Regions`}</span>
              }
              value={
                <span
                  style={{
                    fontSize: "30px",
                    fontWeight: 1000,
                    color: "#ff0000",
                  }}
                >
                  {analytics.region.length.toLocaleString()}
                </span>
              }
              // note={<span style={{ fontSize: "14px", fontWeight: 1000, color: "#00ff15" }}>Total Regions</span>}
              icon={CircleDot}
              tone="#60a5fa"
              style={{ gridColumn: "span 6" }}
            />

            <StatCard
              label={
                <span
                  style={{
                    fontSize: "16px",
                    fontWeight: 1000,
                    color: "#00ff15",
                  }}
                >{`Non-Service Impact`}</span>
              }
              value={
                <span
                  style={{
                    fontSize: "30px",
                    fontWeight: 1000,
                    color: "#ff0000",
                  }}
                >
                  {nonServiceImpact.toLocaleString()}
                </span>
              }
              // note={<span style={{ fontSize: "14px", fontWeight: 1000, color: "#00ff15" }}>No Service Impact</span>}
              icon={CloudOff}
              tone="#94a3b8"
              onClick={() => {
                setFilters({
                  ...EMPTY_FILTERS,
                  impact: ["Non-Service Impact"],
                });
                setTablePage(1);
              }}
              style={{ gridColumn: "span 6" }}
            />

            <StatCard
              label={
                <span
                  style={{
                    fontSize: "16px",
                    fontWeight: 1000,
                    color: "#00ff15",
                  }}
                >{`Service Impact`}</span>
              }
              value={
                <span
                  style={{
                    fontSize: "30px",
                    fontWeight: 1000,
                    color: "#ff0000",
                  }}
                >
                  {serviceImpact.toLocaleString()}
                </span>
              }
              // note={<span style={{ fontSize: "14px", fontWeight: 1000, color: "#00ff15" }}>Exact Service Impact</span>}
              icon={Network}
              tone="#ef4444"
              onClick={() => {
                setFilters({ ...EMPTY_FILTERS, impact: ["Service Impact"] });
                setTablePage(1);
              }}
              style={{ gridColumn: "span 6" }}
            />

            {/* Top RCA by Tickets Count — direct grid child so it stretches to the same height as siblings */}
            <StatCard
              label={
                <span
                  style={{
                    fontSize: "16px",
                    fontWeight: 1000,
                    color: "#00ff15",
                  }}
                >{`Top RCA / TT's`}</span>
              }
              value={
                <span
                  style={{
                    fontSize: "18px",
                    fontWeight: 1000,
                    color: "#ff0000",
                    display: "block",
                    wordBreak: "keep-all",
                    whiteSpace: "normal",
                    lineHeight: "1.3",
                  }}
                >
                  {analytics.topRcaByCount.name || "N/A"}
                </span>
              }
              note={
                <span
                  style={{
                    fontSize: "14px",
                    fontWeight: 1000,
                    color: "#00ff15",
                    marginTop: "4px",
                    display: "block",
                  }}
                >
                  {analytics.topRcaByCount.value
                    ? `${analytics.topRcaByCount.value.toLocaleString()} Tickets`
                    : ""}
                </span>
              }
              icon={BarChart3}
              tone="#22d3ee"
              className="rca-inline-fix"
              style={{
                gridColumn: "span 6",
                alignSelf: "stretch",
                height: "auto",
              }}
            />

            {/* Top RCA by Downtime */}
            <StatCard
              label={
                <span
                  style={{
                    fontSize: "16px",
                    fontWeight: 1000,
                    color: "#00ff15",
                  }}
                >{`Top RCA / Downtime`}</span>
              }
              value={
                <span
                  style={{
                    fontSize: "16px",
                    fontWeight: 1000,
                    color: "#ff0000",
                    display: "block",
                    wordBreak: "keep-all",
                    whiteSpace: "normal",
                    lineHeight: "1.3",
                  }}
                >
                  {analytics.topRcaByDowntime.name || "N/A"}
                </span>
              }
              note={
                <span
                  style={{
                    fontSize: "14px",
                    fontWeight: 1000,
                    color: "#00ff15",
                    marginTop: "4px",
                    display: "block",
                  }}
                >
                  {formatHours(analytics.topRcaByDowntime.value)}
                </span>
              }
              icon={Activity}
              tone="#f59e0b"
              style={{ gridColumn: "span 6" }}
            />

            {/* Highest MTTR RCA */}
            <StatCard
              label={
                <span
                  style={{
                    fontSize: "16px",
                    fontWeight: 1000,
                    color: "#00ff15",
                  }}
                >{`Highest MTTR RCA`}</span>
              }
              value={
                <span
                  style={{
                    fontSize: "16px",
                    fontWeight: 1000,
                    color: "#ff0000",
                    display: "block",
                    wordBreak: "keep-all",
                    whiteSpace: "normal",
                    lineHeight: "1.3",
                  }}
                >
                  {analytics.highestMttrRca.name || "N/A"}
                </span>
              }
              note={
                <span
                  style={{
                    fontSize: "14px",
                    fontWeight: 1000,
                    color: "#00ff15",
                    marginTop: "4px",
                    display: "block",
                  }}
                >
                  {formatHours(analytics.highestMttrRca.value)}
                </span>
              }
              icon={AlertTriangle}
              tone="#ef4444"
              style={{ gridColumn: "span 6" }}
            />

            <StatCard
              label={
                <span
                  style={{
                    fontSize: "20px",
                    fontWeight: 1000,
                    color: "#00ff15",
                  }}
                >{`Repeated RCA/Sites`}</span>
              }
              value={
                <span
                  style={{
                    fontSize: "30px",
                    fontWeight: 1000,
                    color: "#ff0000",
                    display: "block",
                    wordBreak: "keep-all",
                    whiteSpace: "normal",
                    lineHeight: "1.3",
                  }}
                >
                  {analytics.repeatedRcaSites.toLocaleString()}
                </span>
              }
              note={<span />}
              icon={Network}
              tone="#a78bfa"
              style={{ gridColumn: "span 6" }}
            />

            <StatCard
              label={
                <span
                  style={{
                    fontSize: "16px",
                    fontWeight: 1000,
                    color: "#00ff15",
                  }}
                >{`RCA not Provided %`}</span>
              }
              value={
                <span
                  style={{
                    fontSize: "30px",
                    fontWeight: 1000,
                    color: "#ff0000",
                    display: "block",
                    wordBreak: "keep-all",
                    whiteSpace: "normal",
                    lineHeight: "1.3",
                  }}
                >
                  {analytics.rcaNotProvidedCount.toLocaleString()}
                </span>
              }
              note={<span />}
              icon={ShieldAlert}
              tone="#ff0000"
              style={{ gridColumn: "span 6" }}
            />
          </section>

          <section id="section-executive" className="glass-card executive-insights-section dashboard-section-content-block">
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
              {executiveInsights.cards.map(card => (
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

            <div className="executive-risk-table-card">
              <div className="executive-risk-table-header">
                <div>
                  <strong>High Risk Sites Ranking</strong>
                  <p>
                    Risk score uses ticket count, downtime, service impact, critical severity, missing RCA, and reliability.
                  </p>
                </div>
                <span>Top {executiveInsights.highRiskSites.length}</span>
              </div>

              <div className="table-scroll executive-risk-table-scroll">
                <table className="data-table executive-risk-table">
                  <thead>
                    <tr>
                      <th>Rank</th>
                      <th>Site ID</th>
                      <th>Site Name</th>
                      <th>Tickets</th>
                      <th>Downtime</th>
                      <th>Reliability</th>
                      <th>Top RCA</th>
                      <th>Risk Level</th>
                    </tr>
                  </thead>
                  <tbody>
                    {executiveInsights.highRiskSites.length ? (
                      executiveInsights.highRiskSites.map(site => (
                        <tr key={`${site.rank}-${site.siteId}`}>
                          <td>{site.rank}</td>
                          <td>{site.siteId}</td>
                          <td>{site.siteName || "-"}</td>
                          <td>{site.ticketCount.toLocaleString()}</td>
                          <td>{site.downtimeHours.toLocaleString()} hrs</td>
                          <td>{site.reliability.toFixed(2)}%</td>
                          <td>{site.topRca}</td>
                          <td>
                            <span
                              className={`pill executive-risk-pill executive-risk-pill--${site.riskLevel.toLowerCase()}`}
                            >
                              {site.riskLevel} · {site.riskScore}
                            </span>
                          </td>
                        </tr>
                      ))
                    ) : (
                      <tr>
                        <td colSpan={8} className="empty-table-cell">
                          No high-risk sites found for the selected filters.
                        </td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          </section>

          <section id="section-deep-dive" className="glass-card deep-dive-section client-delivery-section dashboard-section-content-block">
            <div className="card-heading deep-dive-header">
              <div>
                <span className="section-kicker">
                  <ShieldAlert size={14} /> RCA / Preventability / SLA Deep-Dive
                </span>
                <h3>Operational Quality & Follow-Up Priorities</h3>
                <p>
                  This layer converts RCA quality, preventability, SLA response, and repeated-site patterns into management follow-up actions.
                </p>
              </div>
            </div>

            <div className="deep-dive-kpi-grid">
              <div className="deep-dive-kpi-card">
                <span>Avg FRT</span>
                <strong>{deepDiveAnalytics.slaSummary.avgFrtHours.toFixed(1)}h</strong>
                <small>{deepDiveAnalytics.slaSummary.frtBreaches.toLocaleString()} above 1h target</small>
              </div>
              <div className="deep-dive-kpi-card">
                <span>Avg Response</span>
                <strong>{deepDiveAnalytics.slaSummary.avgResponseHours.toFixed(1)}h</strong>
                <small>{deepDiveAnalytics.slaSummary.responseBreaches.toLocaleString()} above 4h target</small>
              </div>
              <div className="deep-dive-kpi-card">
                <span>Avg Resolution</span>
                <strong>{deepDiveAnalytics.slaSummary.avgResolutionHours.toFixed(1)}h</strong>
                <small>{deepDiveAnalytics.slaSummary.resolutionBreaches.toLocaleString()} above 24h target</small>
              </div>
              <div className="deep-dive-kpi-card">
                <span>Repeated Sites</span>
                <strong>{deepDiveAnalytics.repeatedOffenderSites.length.toLocaleString()}</strong>
                <small>Top sites requiring technical follow-up</small>
              </div>
            </div>

            <div className="deep-dive-chart-grid">
              <article className="deep-dive-panel">
                <div className="deep-dive-panel-heading">
                  <strong>Preventability by Tickets</strong>
                  <small>Preventable vs non-preventable events</small>
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
                      label={({ name, percentage }) => `${name}: ${percentage ?? 0}%`}
                    >
                      {deepDiveAnalytics.preventabilityByCount.map((entry, index) => (
                        <Cell key={entry.name} fill={COLORS[index % COLORS.length]} />
                      ))}
                    </Pie>
                    <Tooltip contentStyle={CHART_TOOLTIP_STYLE} />
                    <Legend wrapperStyle={CHART_LEGEND_STYLE} />
                  </PieChart>
                </ResponsiveContainer>
              </article>

              <article className="deep-dive-panel">
                <div className="deep-dive-panel-heading">
                  <strong>Pending Aging Buckets</strong>
                  <small>Open/pending tickets by age</small>
                </div>
                <ResponsiveContainer width="100%" height={230}>
                  <BarChart data={deepDiveAnalytics.slaSummary.pendingAgingBuckets} margin={{ left: 4, right: 20, top: 14, bottom: 4 }}>
                    <CartesianGrid stroke={CHART_GRID_STROKE} vertical={false} />
                    <XAxis dataKey="name" stroke={CHART_AXIS_STROKE} tickLine={false} axisLine={false} />
                    <YAxis stroke={CHART_AXIS_STROKE} allowDecimals={false} tickLine={false} axisLine={false} />
                    <Tooltip contentStyle={CHART_TOOLTIP_STYLE} />
                    <Bar dataKey="value" name="Tickets" radius={COLUMN_BAR_RADIUS} fill="#f59e0b">
                      <LabelList dataKey="value" position="top" fill={CHART_LABEL_FILL} fontSize={11} />
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </article>

              <article className="deep-dive-panel deep-dive-panel--wide">
                <div className="deep-dive-panel-heading">
                  <strong>Top RCA Families by Downtime</strong>
                  <small>Includes RCA quality, preventability, and ownership</small>
                </div>
                <ResponsiveContainer width="100%" height={280}>
                  <BarChart
                    data={deepDiveAnalytics.rcaFamilyDeepDive.slice(0, 8)}
                    layout="vertical"
                    margin={{ left: 18, right: 48, top: 8, bottom: 8 }}
                  >
                    <CartesianGrid stroke={CHART_GRID_STROKE} horizontal={false} />
                    <XAxis type="number" stroke={CHART_AXIS_STROKE} allowDecimals={false} tickLine={false} axisLine={false} />
                    <YAxis dataKey="family" type="category" stroke={CHART_AXIS_STROKE} width={190} tickLine={false} axisLine={false} tick={{ fontSize: 11 }} />
                    <Tooltip contentStyle={CHART_TOOLTIP_STYLE} />
                    <Bar dataKey="downtimeHours" name="Downtime hrs" radius={BAR_RADIUS} fill="#22d3ee">
                      <LabelList dataKey="downtimeHours" position="right" fill={CHART_LABEL_FILL} fontSize={11} />
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </article>
            </div>

            <div className="deep-dive-bottom-grid">
              <div className="deep-dive-table-card">
                <div className="deep-dive-panel-heading">
                  <strong>Repeated Offender Sites</strong>
                  <small>Prioritized by downtime and repeated tickets</small>
                </div>
                <div className="table-scroll deep-dive-table-scroll">
                  <table className="data-table compact-table">
                    <thead>
                      <tr>
                        <th>Site ID</th>
                        <th>Site Name</th>
                        <th>Tickets</th>
                        <th>Downtime</th>
                        <th>Top RCA</th>
                      </tr>
                    </thead>
                    <tbody>
                      {deepDiveAnalytics.repeatedOffenderSites.length ? (
                        deepDiveAnalytics.repeatedOffenderSites.slice(0, 8).map(site => (
                          <tr key={site.siteId}>
                            <td>{site.siteId}</td>
                            <td>{site.siteName || "-"}</td>
                            <td>{site.tickets}</td>
                            <td>{site.downtimeHours.toLocaleString()} hrs</td>
                            <td>{site.topRca}</td>
                          </tr>
                        ))
                      ) : (
                        <tr>
                          <td colSpan={5} className="empty-table-cell">No repeated offender sites found.</td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </div>
              </div>

              <div className="deep-dive-actions-card">
                <div className="deep-dive-panel-heading">
                  <strong>Recommended Management Actions</strong>
                  <small>Auto-generated from current filtered scope</small>
                </div>
                <ol>
                  {deepDiveAnalytics.recommendations.map(item => (
                    <li key={item}>{item}</li>
                  ))}
                </ol>
              </div>
            </div>
          </section>

          {/* ══ Charts ══ */}
          <div id="section-overview-charts" className="chart-2col dashboard-chart-grid dashboard-section-content-block">
            {/* ── Column 1 — five charts, one per row ─────────────────────────── */}
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
                              }
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
                                  : ["—", name];
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

            {/* ── Column 2 — five charts stacked vertically, spans column 1 rows 1-4 ── */}
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
          {/* ══ Full-width + bottom-row charts ═══════════════════════════════════ */}
          <section id="section-trend-charts" className="chart-mosaic dashboard-section-content-block">
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
                          }
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

            <article className="glass-card" style={{ gridColumn: "7 / 13" }}>
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
          <section id="section-tickets-table" className="table-card dashboard-section-content-block">
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
                    Math.ceil(filteredTickets.length / TABLE_PAGE_SIZE)
                  )}{" "}
                  &mdash; {filteredTickets.length.toLocaleString()} total
                </span>
              </div>
            </div>
            <div className="table-scroll" id="ticket-table-wrapper">
              <div ref={tableRef}>
                <div style={{ marginBottom: "12px" }}>
                  <button className="ghost-button" onClick={scrollToTopCards}>
                    {" "}
                    ↑ Back to Summary{" "}
                  </button>
                </div>

                <table
                  style={{
                    tableLayout: "fixed",
                    width: "auto",
                    borderCollapse: "separate",
                  }}
                >
                  <colgroup>
                    {DISTINCT_REPORT_HEADERS.map(header => (
                      <col
                        key={header}
                        style={{ width: `${getTicketColumnWidth(header)}px` }}
                      />
                    ))}
                  </colgroup>
                  <thead>
                    <tr>
                      {DISTINCT_REPORT_HEADERS.map(header => {
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
                              onMouseEnter={e => {
                                (
                                  e.currentTarget as HTMLSpanElement
                                ).style.backgroundSize = "100% 100%";
                              }}
                              onMouseLeave={e => {
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
                        tablePage * TABLE_PAGE_SIZE
                      )
                      .map((ticket, index) => {
                        const row = ticket.primary;
                        const reportRow = distinctReportRow(ticket, index);
                        const siteIds = Array.from(ticket.siteIds).filter(
                          Boolean
                        );
                        const siteNames = Array.from(ticket.siteNames).filter(
                          Boolean
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
                        //   <td>  →  <div width=col-width, flex-column, overflow:hidden>  →  <span overflow:ellipsis>
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
                          header: string
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
                          header: string
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
                                      {items.map(id => (
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
                                      {items.map(name => (
                                        <span key={name} style={spanStyle}>
                                          {name}
                                        </span>
                                      ))}
                                    </div>
                                  </td>
                                );
                              }

                              // Pill: Severity / Status — the pill itself replaces the span.
                              if (
                                header === "Severity" ||
                                header === "Status"
                              ) {
                                const tone =
                                  (header === "Severity"
                                    ? SEVERITY_COLORS
                                    : STATUS_COLORS)[String(cell ?? "")] ?? "#64748b";
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
                  «
                </button>
                <button
                  className="ghost-button"
                  disabled={tablePage <= 1}
                  onClick={() => setTablePage(p => p - 1)}
                >
                  ‹ Prev
                </button>
                {Array.from(
                  {
                    length: Math.ceil(filteredTickets.length / TABLE_PAGE_SIZE),
                  },
                  (_, i) => i + 1
                )
                  .filter(
                    p =>
                      p === 1 ||
                      p ===
                        Math.ceil(filteredTickets.length / TABLE_PAGE_SIZE) ||
                      Math.abs(p - tablePage) <= 2
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
                    )
                  )}
                <button
                  className="ghost-button"
                  disabled={
                    tablePage >=
                    Math.ceil(filteredTickets.length / TABLE_PAGE_SIZE)
                  }
                  onClick={() => setTablePage(p => p + 1)}
                >
                  Next ›
                </button>
                <button
                  className="ghost-button"
                  disabled={
                    tablePage >=
                    Math.ceil(filteredTickets.length / TABLE_PAGE_SIZE)
                  }
                  onClick={() =>
                    setTablePage(
                      Math.ceil(filteredTickets.length / TABLE_PAGE_SIZE)
                    )
                  }
                >
                  »
                </button>
              </div>
            )}
          </section>

          {/* Monthly Performance Table */}
          {perfRows.length > 0 && (
            <section id="section-performance-table" className="table-card dashboard-section-content-block">
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
                      ? ` — ${perfRegions.join(", ")}`
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
                  {" "}
                  ↑ Back to Summary{" "}
                </button>
              </div>
              <div className="table-scroll">
                <table
                  style={{
                    tableLayout: "fixed",
                    width: "auto",
                    borderCollapse: "separate",
                  }}
                >
                  <colgroup>
                    {PERF_REPORT_HEADERS.map(h => (
                      <col
                        key={h}
                        style={{ width: `${getPerfColumnWidth(h)}px` }}
                      />
                    ))}
                  </colgroup>
                  <thead>
                    <tr>
                      {PERF_REPORT_HEADERS.map(header => {
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
                              onMouseEnter={e => {
                                (
                                  e.currentTarget as HTMLSpanElement
                                ).style.backgroundSize = "100% 100%";
                              }}
                              onMouseLeave={e => {
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
                        perfPage * TABLE_PAGE_SIZE
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
                        //   <td>  →  <div width=col-width, flex-column, overflow:hidden>  →  <span overflow:ellipsis>
                        const baseStyle: React.CSSProperties = {
                          verticalAlign: "top",
                          padding: "10px 12px",
                        };
                        const isPerfLeftAligned = (header: string) =>
                          header === "Site ID" || header === "Site Name";
                        const perfCellStyle = (
                          header: string
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
                          header: string
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
                    {/* Total summary row — always computed over the full dataset, shown only on the last page */}
                    {perfPage ===
                      Math.max(
                        1,
                        Math.ceil(perfRows.length / TABLE_PAGE_SIZE)
                      ) &&
                      perfRows.length > 0 &&
                      (() => {
                        const totalDown =
                          Math.round(
                            perfRows.reduce((s, r) => s + r.sitesDownHours, 0) *
                              10
                          ) / 10;
                        const totalAvail =
                          Math.round(
                            perfRows.reduce((s, r) => s + r.availHours, 0) * 10
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
                    «
                  </button>
                  <button
                    className="ghost-button"
                    disabled={perfPage <= 1}
                    onClick={() => setPerfPage(p => p - 1)}
                  >
                    ‹ Prev
                  </button>
                  {Array.from(
                    { length: Math.ceil(perfRows.length / TABLE_PAGE_SIZE) },
                    (_, i) => i + 1
                  )
                    .filter(
                      p =>
                        p === 1 ||
                        p === Math.ceil(perfRows.length / TABLE_PAGE_SIZE) ||
                        Math.abs(p - perfPage) <= 2
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
                      )
                    )}
                  <button
                    className="ghost-button"
                    disabled={
                      perfPage >= Math.ceil(perfRows.length / TABLE_PAGE_SIZE)
                    }
                    onClick={() => setPerfPage(p => p + 1)}
                  >
                    Next ›
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
                    »
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