/*
Design philosophy: Calm Enterprise Glass for a premium telecom network-observability cockpit.
Use deep ink backgrounds, translucent panels, cyan focus accents, tabular TT numerals, and restrained motion.
Does this choice reinforce or dilute our design philosophy?
*/
import { useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import {
  Activity,
  AlertTriangle,
  BarChart3,
  CheckCircle2,
  ChevronDown,
  CircleDot,
  Download,
  FileSpreadsheet,
  HardDrive,
  Trash2,
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

const HERO_IMAGE =
  "https://d2xsxph8kpxj0f.cloudfront.net/310419663031216744/LFprMZJgQoCY2omHrJr6xN/followup-hero-network-cockpit-GEqHM9kSYycEMt32RSfRxg.webp";
const UPLOAD_IMAGE =
  "https://d2xsxph8kpxj0f.cloudfront.net/310419663031216744/LFprMZJgQoCY2omHrJr6xN/followup-upload-orb-YcKWLpbfcdm5ofcRiuQymn.webp";
const RIBBON_IMAGE =
  "https://d2xsxph8kpxj0f.cloudfront.net/310419663031216744/LFprMZJgQoCY2omHrJr6xN/followup-network-ribbon-Lv2N5GpYhLW5eJjvLPNzkg.webp";

const COLORS = ["#22d3ee", "#60a5fa", "#f59e0b", "#ef4444", "#34d399", "#a78bfa", "#f472b6", "#94a3b8"];
const SESSION_KEY = "follow-up-sheets-dashboard-session-v1";
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
  issue: string;
  severity: string;
  region: string;
  observationDate: string;
  recoveryDate: string;
  duration: string;
  impact: string;
  escalatedTo: string;
  escalationLevel: string;
  status: string;
  comments: string;
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

type StoredDashboardSession = {
  version: 1;
  fileName: string;
  sheetName: string;
  generatedAt: string;
  savedAt: string;
  rows: TicketRecord[];
};

type CountMode = "primary" | "exposure";

type Filters = {
  search: string;
  status: string;
  severity: string;
  region: string;
  impact: string;
  site: string;
  openingMonth: string;
};

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

function parseDateValue(value: string): Date | null {
  if (!value) return null;
  const direct = new Date(value);
  if (!Number.isNaN(direct.getTime())) return direct;
  const match = value.match(/(\d{1,2})[-/ ]([A-Za-z]{3,}|\d{1,2})[-/ ](\d{2,4})/);
  if (!match) return null;
  const parsed = new Date(value.replace(/-/g, " "));
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

function openingMonthKey(value: string): string {
  const parsed = parseDateValue(value);
  if (!parsed) return "Unknown";
  const month = String(parsed.getMonth() + 1).padStart(2, "0");
  return `${parsed.getFullYear()}-${month}`;
}

function openingMonthLabel(key: string): string {
  if (!key || key === "Unknown") return "Unknown";
  const parsed = new Date(`${key}-01T00:00:00`);
  if (Number.isNaN(parsed.getTime())) return key;
  return parsed.toLocaleDateString("en", { month: "short", year: "numeric" });
}

function dateToMonth(value: string): string {
  return openingMonthLabel(openingMonthKey(value));
}

function dateKey(value: string): number {
  const parsed = new Date(value);
  return Number.isNaN(parsed.getTime()) ? 0 : parsed.getTime();
}

function parseDurationHours(duration: string): number | null {
  if (!duration) return null;
  const days = Number(duration.match(/(\d+)\s*days?/i)?.[1] ?? 0);
  const hrs = Number(duration.match(/(\d+)\s*hrs?/i)?.[1] ?? 0);
  const mins = Number(duration.match(/(\d+)\s*mins?/i)?.[1] ?? 0);
  const total = days * 24 + hrs + mins / 60;
  return Number.isFinite(total) ? total : null;
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
  const raw = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, { defval: "", raw: false });

  const rows: TicketRecord[] = raw
    .map((row, index) => ({
      rowNo: index + 2,
      tt: getField(row, ["TT", "Ticket", "Ticket Number"]),
      siteId: getField(row, ["Site ID", "SiteID"]),
      siteName: getField(row, ["Site Name", "SiteName"]),
      issue: getField(row, ["Issues", "Issue"]),
      severity: getField(row, ["Severity"]),
      region: getField(row, ["Region"]),
      observationDate: getField(row, ["Observation Date", "Observed Date"]),
      recoveryDate: getField(row, ["Recovery Date"]),
      duration: getField(row, ["Total Durration Days/Hours", "Total Duration Days/Hours", "Duration"]),
      impact: getField(row, ["Service Impaction Status", "Service Impact Status"]),
      escalatedTo: getField(row, ["Escalated to", "Escalated to "]),
      escalationLevel: getField(row, ["Escalation Level", "Esclation Level"]),
      status: getField(row, ["Status"]),
      comments: getField(row, ["Comments-Feedback", "Comments"]),
      actionTaken: getField(row, ["Action Taken/RCA", "Action Taken", "RCA"]),
    }))
    .filter((row) => row.tt || row.siteId || row.siteName || row.issue);

  return {
    fileName,
    sheetName,
    generatedAt: new Date().toLocaleString(),
    rows,
    uniqueTickets: groupTickets(rows),
  };
}

function saveSession(parsed: DashboardData): string {
  const savedAt = new Date().toLocaleString();
  const session: StoredDashboardSession = {
    version: 1,
    fileName: parsed.fileName,
    sheetName: parsed.sheetName,
    generatedAt: parsed.generatedAt,
    savedAt,
    rows: parsed.rows,
  };
  localStorage.setItem(SESSION_KEY, JSON.stringify(session));
  return savedAt;
}

function loadSession(): { data: DashboardData; savedAt: string } | null {
  const rawSession = localStorage.getItem(SESSION_KEY);
  if (!rawSession) return null;
  const session = JSON.parse(rawSession) as StoredDashboardSession;
  if (session.version !== 1 || !Array.isArray(session.rows) || !session.rows.length) return null;
  return {
    savedAt: session.savedAt,
    data: {
      fileName: session.fileName,
      sheetName: session.sheetName,
      generatedAt: session.generatedAt,
      rows: session.rows,
      uniqueTickets: groupTickets(session.rows),
    },
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
  const name = props.name ?? "";
  const value = props.value ?? 0;
  const percent = props.percent ?? 0;
  if (!value) return "";
  return `${name}: ${value} (${(percent * 100).toFixed(0)}%)`;
}

function exportCsv(rows: TicketAggregate[]) {
  const header = ["TT", "Primary Site ID", "Primary Site Name", "All Site IDs", "Severity", "Region", "Status", "Impact", "Escalation Level", "Issue"];
  const body = rows.map((ticket) => {
    const row = ticket.primary;
    return [
      ticket.tt,
      row.siteId,
      row.siteName,
      Array.from(ticket.siteIds).join(" | "),
      row.severity,
      row.region,
      row.status,
      row.impact,
      row.escalationLevel,
      row.issue,
    ];
  });
  const csv = [header, ...body]
    .map((line) => line.map((cell) => `"${String(cell ?? "").replace(/"/g, '""')}"`).join(","))
    .join("\n");
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = "follow-up-unique-tt-register.csv";
  link.click();
  URL.revokeObjectURL(url);
}

function StatCard({ label, value, note, icon: Icon, tone }: { label: string; value: string | number; note: string; icon: typeof Activity; tone: string }) {
  return (
    <div className="stat-card" style={{ ["--tone" as string]: tone }}>
      <div className="stat-icon"><Icon size={18} /></div>
      <span>{label}</span>
      <strong>{value}</strong>
      <small>{note}</small>
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

export default function Home() {
  const inputRef = useRef<HTMLInputElement | null>(null);
  const [data, setData] = useState<DashboardData | null>(null);
  const [error, setError] = useState("");
  const [isDragging, setIsDragging] = useState(false);
  const [countMode, setCountMode] = useState<CountMode>("primary");
  const [savedAt, setSavedAt] = useState("");
  const [filters, setFilters] = useState<Filters>({ search: "", status: "all", severity: "all", region: "all", impact: "all", site: "all", openingMonth: "all" });

  useEffect(() => {
    try {
      const restored = loadSession();
      if (restored) {
        setData(restored.data);
        setSavedAt(restored.savedAt);
      }
    } catch {
      localStorage.removeItem(SESSION_KEY);
    }
  }, []);

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
      setSavedAt(saveSession(parsed));
      setFilters({ search: "", status: "all", severity: "all", region: "all", impact: "all", site: "all", openingMonth: "all" });
      if (inputRef.current) inputRef.current.value = "";
    } catch (err) {
      setError(err instanceof Error ? err.message : "Unable to read this workbook.");
    }
  }

  function clearSavedSession() {
    localStorage.removeItem(SESSION_KEY);
    setSavedAt("");
  }

  const uniqueRows = data?.uniqueTickets ?? [];
  const filterOptions = useMemo(() => {
    const primaryRows = uniqueRows.map((ticket) => ticket.primary);
    const uniq = (field: keyof TicketRecord) => Array.from(new Set(primaryRows.map((row) => clean(row[field])).filter(Boolean))).sort();
    const openingMonths = Array.from(new Set(primaryRows.map((row) => openingMonthKey(row.observationDate)).filter(Boolean))).sort((a, b) => {
      if (a === "Unknown") return 1;
      if (b === "Unknown") return -1;
      return a.localeCompare(b);
    });
    return {
      status: uniq("status"),
      severity: uniq("severity"),
      region: uniq("region"),
      impact: uniq("impact"),
      site: uniq("siteId"),
      openingMonth: openingMonths,
      openingMonthLabels: Object.fromEntries(openingMonths.map((key) => [key, openingMonthLabel(key)])),
    };
  }, [uniqueRows]);

  const filteredTickets = useMemo(() => {
    const q = filters.search.toLowerCase().trim();
    return uniqueRows.filter((ticket) => {
      const row = ticket.primary;
      const allSites = Array.from(ticket.siteIds).join(" ");
      const haystack = [ticket.tt, row.siteId, row.siteName, allSites, row.issue, row.status, row.severity, row.region, row.impact, row.escalationLevel]
        .join(" ")
        .toLowerCase();
      return (
        (!q || haystack.includes(q)) &&
        (filters.status === "all" || row.status === filters.status) &&
        (filters.severity === "all" || row.severity === filters.severity) &&
        (filters.region === "all" || row.region === filters.region) &&
        (filters.impact === "all" || row.impact === filters.impact) &&
        (filters.openingMonth === "all" || openingMonthKey(row.observationDate) === filters.openingMonth) &&
        (filters.site === "all" || (countMode === "primary" ? row.siteId === filters.site : ticket.siteIds.has(filters.site)))
      );
    });
  }, [countMode, filters, uniqueRows]);

  const analytics = useMemo(() => {
    const primaryRows = filteredTickets.map((ticket) => ticket.primary);
    const totalUnique = filteredTickets.length;
    const status = countBy(primaryRows, (row) => row.status);
    const severity = countBy(primaryRows, (row) => row.severity);
    const region = countBy(primaryRows, (row) => row.region);
    const impact = countBy(primaryRows, (row) => row.impact);
    const escalation = countBy(primaryRows, (row) => row.escalationLevel);
    const monthly = countBy(primaryRows, (row) => dateToMonth(row.observationDate)).reverse();
    const avgHoursSource = primaryRows.map((row) => parseDurationHours(row.duration)).filter((value): value is number => value !== null);
    const avgHours = avgHoursSource.length ? avgHoursSource.reduce((sum, value) => sum + value, 0) / avgHoursSource.length : 0;
    const uniqueSites = new Set(primaryRows.map((row) => row.siteId).filter(Boolean)).size;
    const rootCauseUpdated = primaryRows.filter((row) => row.actionTaken).length;
    const totalSiteAffected = filteredTickets.reduce((sum, ticket) => sum + Math.max(ticket.siteIds.size, ticket.primary.siteId ? 1 : 0), 0);

    const siteNameById = new Map<string, string>();
    filteredTickets.forEach((ticket) => {
      ticket.rows.forEach((row) => {
        if (row.siteId && row.siteName && !siteNameById.has(row.siteId)) siteNameById.set(row.siteId, row.siteName);
      });
    });
    const siteLabel = (siteId: string) => {
      const siteName = siteNameById.get(siteId);
      if (!siteId || siteId === "Blank") return "Blank";
      return siteName ? `${siteId} — ${siteName}` : siteId;
    };
    const siteMap = new Map<string, { name: string; value: number; exposure?: number }>();
    if (countMode === "primary") {
      filteredTickets.forEach((ticket) => {
        const site = ticket.primary.siteId || "Blank";
        const current = siteMap.get(site) ?? { name: siteLabel(site), value: 0 };
        current.value += 1;
        siteMap.set(site, current);
      });
    } else {
      filteredTickets.forEach((ticket) => {
        const sites = ticket.siteIds.size ? Array.from(ticket.siteIds) : [ticket.primary.siteId || "Blank"];
        sites.forEach((site) => {
          const current = siteMap.get(site) ?? { name: siteLabel(site), value: 0 };
          current.value += 1;
          siteMap.set(site, current);
        });
      });
    }
    const topSites = Array.from(siteMap.values()).sort((a, b) => b.value - a.value || a.name.localeCompare(b.name)).slice(0, 12);

    return { totalUnique, status, severity, region, impact, escalation, monthly, avgHours, uniqueSites, rootCauseUpdated, totalSiteAffected, topSites };
  }, [countMode, filteredTickets]);

  const closed = metricValue(analytics.status, "Closed");
  const pending = metricValue(analytics.status, "Pending");
  const resolved = metricValue(analytics.status, "Resolved");
  const critical = metricValue(analytics.severity, "Critical");
  const major = metricValue(analytics.severity, "Major");
  const minor = metricValue(analytics.severity, "Minor");
  const serviceImpact = metricValue(analytics.impact, "Service Impact");
  const nonServiceImpact = metricValue(analytics.impact, "Non-Service Impact");
  const duplicateRows = data ? data.rows.filter((row) => row.tt).length - data.uniqueTickets.length : 0;

  return (
    <main className="dashboard-shell">
      <section className="hero-panel" style={{ backgroundImage: `linear-gradient(90deg, rgba(3,7,18,.96) 0%, rgba(3,7,18,.78) 42%, rgba(3,7,18,.26) 100%), url(${HERO_IMAGE})` }}>
        <nav className="topbar no-print">
          <div className="brand-mark"><Network size={18} /> TT Operations Cockpit</div>
          <div className="topbar-actions">
            {data && <button className="ghost-button" onClick={() => inputRef.current?.click()}><RefreshCw size={16} /> New workbook</button>}
            {data && savedAt && <button className="ghost-button" onClick={clearSavedSession}><Trash2 size={16} /> Clear saved session</button>}
            {data && <button className="ghost-button" onClick={() => exportCsv(filteredTickets)}><Download size={16} /> CSV</button>}
            {data && <button className="primary-button" onClick={() => window.print()}><Printer size={16} /> Print / PDF</button>}
          </div>
        </nav>
        <div className="hero-content">
          <div>
            <p className="eyebrow"><CircleDot size={12} /> Excel-powered unique TT intelligence</p>
            <h1>Follow-Up Sheets Dashboard</h1>
            <p className="hero-copy">Upload the workbook and convert the ticket register into an interactive executive dashboard with corrected unique TT counting, operational filters, site ranking, trend analysis, and an export-ready view.</p>
            <div className="hero-meta">
              <span>Source sheet: Tickets_Data</span>
              <span>Rule: one TT number = one ticket</span>
              <span>Runs in browser</span>
            </div>
          </div>
        </div>
      </section>

      <input ref={inputRef} className="sr-only" type="file" accept=".xlsx,.xls,.xlsm" onChange={(event) => handleFile(event.target.files?.[0])} />

      {!data ? (
        <section className="upload-stage no-print">
          <div
            className={`upload-card ${isDragging ? "dragging" : ""}`}
            onDragOver={(event) => { event.preventDefault(); setIsDragging(true); }}
            onDragLeave={() => setIsDragging(false)}
            onDrop={(event) => { event.preventDefault(); setIsDragging(false); handleFile(event.dataTransfer.files?.[0]); }}
          >
            <div className="upload-copy">
              <span className="section-kicker"><UploadCloud size={14} /> Workbook upload</span>
              <h2>Bring the Follow-Up Sheets workbook online.</h2>
              <p>The parser reads the `Tickets_Data` sheet, groups records by TT number, removes duplicated tickets from KPI totals, and keeps a separate site-exposure mode when a single TT is associated with multiple sites.</p>
              {error && <div className="error-banner"><AlertTriangle size={16} /> {error}</div>}
              <button className="primary-button large" onClick={() => inputRef.current?.click()}><FileSpreadsheet size={18} /> Select Excel workbook</button>
            </div>
            <img src={UPLOAD_IMAGE} alt="Abstract workbook upload visualization" />
          </div>
        </section>
      ) : (
        <>
          <section className="control-strip no-print">
            <div>
              <span className="section-kicker"><FileSpreadsheet size={14} /> {data.fileName}</span>
              <h2>Live dashboard from `{data.sheetName}`</h2>
              <p>Last parsed: {data.generatedAt}. The register has {data.rows.length.toLocaleString()} source rows and {data.uniqueTickets.length.toLocaleString()} unique TT numbers.</p>
              {savedAt && <p className="session-note"><HardDrive size={14} /> Saved locally in this browser at {savedAt}; it will reopen automatically until cleared.</p>}
            </div>
            <div className="count-toggle" role="group" aria-label="Site counting mode">
              <button className={countMode === "primary" ? "active" : ""} onClick={() => setCountMode("primary")}>Primary TT allocation</button>
              <button className={countMode === "exposure" ? "active" : ""} onClick={() => setCountMode("exposure")}>Affected-site exposure</button>
            </div>
          </section>

          <section className="stats-grid workbook-cards" style={{ backgroundImage: `linear-gradient(90deg, rgba(4,13,31,.88), rgba(4,13,31,.70)), url(${RIBBON_IMAGE})` }}>
            <StatCard label="Unique TT" value={analytics.totalUnique.toLocaleString()} note="Core Ticket Volume" icon={Layers3} tone="#22d3ee" />
            <StatCard label="Closed TT" value={closed.toLocaleString()} note={`${pct(closed, analytics.totalUnique)} closed`} icon={CheckCircle2} tone="#34d399" />
            <StatCard label="Pending TT" value={pending.toLocaleString()} note="Needs Follow-Up" icon={ShieldAlert} tone="#f59e0b" />
            <StatCard label="Resolved TT" value={resolved.toLocaleString()} note="TT Resolved" icon={CheckCircle2} tone="#60a5fa" />
            <StatCard label="Critical TT" value={critical.toLocaleString()} note="High Priority Severity" icon={AlertTriangle} tone="#ef4444" />
            <StatCard label="Major TT" value={major.toLocaleString()} note="Medium Priority Severity" icon={Activity} tone="#f59e0b" />
            <StatCard label="Minor TT" value={minor ? minor.toLocaleString() : ""} note="Low Priority Severity" icon={CircleDot} tone="#22d3ee" />
            <StatCard label="Service Impact" value={serviceImpact.toLocaleString()} note="Exact Service Impact" icon={Network} tone="#a78bfa" />
            <StatCard label="Non-Service Impact" value={nonServiceImpact.toLocaleString()} note="No Service Impact" icon={XCircle} tone="#94a3b8" />
            <StatCard label="Unique Sites" value={analytics.uniqueSites.toLocaleString()} note="Unique Site ID" icon={BarChart3} tone="#60a5fa" />
            <StatCard label="Root Cause Updated" value={analytics.rootCauseUpdated.toLocaleString()} note="TT with Alarm Root Cause" icon={FileSpreadsheet} tone="#34d399" />
          </section>

          <section className="filters-panel no-print">
            <label className="search-box">
              <Search size={16} />
              <input value={filters.search} onChange={(event) => setFilters((prev) => ({ ...prev, search: event.target.value }))} placeholder="Search TT, site, issue, status, escalation..." />
            </label>
            <SelectFilter label="Status" value={filters.status} options={filterOptions.status} onChange={(value) => setFilters((prev) => ({ ...prev, status: value }))} />
            <SelectFilter label="Severity" value={filters.severity} options={filterOptions.severity} onChange={(value) => setFilters((prev) => ({ ...prev, severity: value }))} />
            <SelectFilter label="Region" value={filters.region} options={filterOptions.region} onChange={(value) => setFilters((prev) => ({ ...prev, region: value }))} />
            <SelectFilter label="Impact" value={filters.impact} options={filterOptions.impact} onChange={(value) => setFilters((prev) => ({ ...prev, impact: value }))} />
            <SelectFilter label="Opening Month" value={filters.openingMonth} options={filterOptions.openingMonth} optionLabels={filterOptions.openingMonthLabels} onChange={(value) => setFilters((prev) => ({ ...prev, openingMonth: value }))} />
            <SelectFilter label="Site" value={filters.site} options={filterOptions.site} onChange={(value) => setFilters((prev) => ({ ...prev, site: value }))} />
            <button className="ghost-button" onClick={() => setFilters({ search: "", status: "all", severity: "all", region: "all", impact: "all", site: "all", openingMonth: "all" })}><Filter size={16} /> Clear</button>
          </section>

          <section className="chart-mosaic">
            <article className="glass-card wide">
              <div className="card-heading"><div><span>Trend</span><h3>Unique TT by month</h3></div><BarChart3 size={18} /></div>
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
              <div className="card-heading"><div><span>{countMode === "primary" ? "Primary allocation" : "Site exposure"}</span><h3>Top sites by unique TT</h3></div></div>
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
              <div className="card-heading"><div><span>Distribution</span><h3>Status</h3></div></div>
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
              <div className="card-heading"><div><span>Distribution</span><h3>Severity</h3></div></div>
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
              <div className="card-heading"><div><span>Distribution</span><h3>Region</h3></div></div>
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
              <div className="card-heading"><div><span>Operational view</span><h3>Escalation level</h3></div></div>
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
              <div className="card-heading"><div><span>Impact</span><h3>Service impact</h3></div></div>
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
          </section>

          <section className="definition-card">
            <h3>Unique TT logic used in this dashboard</h3>
            <p><strong>Primary TT allocation</strong> counts each TT number only once globally and assigns it to the first/earliest site row for that TT. The sum of site counts equals the filtered unique TT total. <strong>Affected-site exposure</strong> counts one TT once per affected site, so totals can exceed global unique TT when the same TT appears in multiple sites.</p>
          </section>

          <section className="table-card">
            <div className="table-heading">
              <div><span className="section-kicker">Unique register</span><h2>{filteredTickets.length.toLocaleString()} distinct TT records</h2></div>
              <p>Showing first 150 filtered tickets. Use search and filters to narrow the register.</p>
            </div>
            <div className="table-scroll" id="ticket-table-wrapper">
              <table>
                <thead>
                  <tr><th>TT</th><th>Primary site</th><th>All affected sites</th><th>Severity</th><th>Status</th><th>Region</th><th>Impact</th><th>Escalation</th><th>Issue</th></tr>
                </thead>
                <tbody>
                  {filteredTickets.slice(0, 150).map((ticket) => {
                    const row = ticket.primary;
                    return (
                      <tr key={ticket.tt}>
                        <td className="mono">{ticket.tt}</td>
                        <td><strong>{row.siteId || ""}</strong><small>{row.siteName}</small></td>
                        <td>{Array.from(ticket.siteIds).join(", ")}</td>
                        <td><span className="pill" style={{ ["--pill" as string]: SEVERITY_COLORS[row.severity] ?? "#64748b" }}>{row.severity}</span></td>
                        <td><span className="pill" style={{ ["--pill" as string]: STATUS_COLORS[row.status] ?? "#64748b" }}>{row.status}</span></td>
                        <td>{row.region}</td>
                        <td>{row.impact}</td>
                        <td>{row.escalationLevel}</td>
                        <td className="issue-cell">{row.issue}</td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </section>
        </>
      )}
    </main>
  );
}
