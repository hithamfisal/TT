from pathlib import Path

path = Path('/home/ubuntu/follow-up-sheets-dashboard-web/client/src/pages/Home.tsx')
text = path.read_text()

replacements = [
    (
        '''function openingMonthLabel(key: string): string {\n  if (!key || key === "Unknown") return "Unknown";\n  const parsed = new Date(`${key}-01T00:00:00`);\n  if (Number.isNaN(parsed.getTime())) return key;\n  return parsed.toLocaleDateString("en", { month: "short", year: "numeric" });\n}\n''',
        '''function openingMonthLabel(key: string): string {\n  if (!key || key === "Unknown") return "Unknown";\n  const parsed = new Date(`${key}-01T00:00:00`);\n  if (Number.isNaN(parsed.getTime())) return key;\n  return parsed.toLocaleDateString("en", { month: "short", year: "numeric" });\n}\n\nfunction recordDateMonthKey(value: string): string {\n  const parsed = parseDateValue(value);\n  if (!parsed) return "Unknown";\n  return `${parsed.getFullYear()}-${String(parsed.getMonth() + 1).padStart(2, "0")}`;\n}\n\nfunction isPendingStatus(value: string): boolean {\n  return clean(value).toLowerCase() === "pending";\n}\n\nfunction ticketMatchesMonthlyExport(ticket: TicketAggregate, selectedMonth: string): boolean {\n  if (selectedMonth === "all") return true;\n  return ticket.rows.some((row) => {\n    const observationMatches = recordDateMonthKey(row.observationDate) === selectedMonth;\n    const recoveryMatches = recordDateMonthKey(row.recoveryDate) === selectedMonth;\n    const pendingMatches = isPendingStatus(row.status);\n    return observationMatches || recoveryMatches || pendingMatches;\n  });\n}\n'''
    ),
    (
        '''  const [savedAt, setSavedAt] = useState("");\n  const [filters, setFilters] = useState<Filters>({ search: "", status: "all", severity: "all", region: "all", impact: "all", site: "all", openingMonth: "all" });\n''',
        '''  const [savedAt, setSavedAt] = useState("");\n  const [exportMonth, setExportMonth] = useState("all");\n  const [filters, setFilters] = useState<Filters>({ search: "", status: "all", severity: "all", region: "all", impact: "all", site: "all", openingMonth: "all" });\n'''
    ),
    (
        '''      setSavedAt(saveSession(parsed));\n      setFilters({ search: "", status: "all", severity: "all", region: "all", impact: "all", site: "all", openingMonth: "all" });\n''',
        '''      setSavedAt(saveSession(parsed));\n      setExportMonth("all");\n      setFilters({ search: "", status: "all", severity: "all", region: "all", impact: "all", site: "all", openingMonth: "all" });\n'''
    ),
    (
        '''    const openingMonths = Array.from(new Set(primaryRows.map((row) => row.openingMonthKey || openingMonthKey(row.observationDate)).filter(Boolean))).sort((a, b) => {\n      if (a === "Unknown") return 1;\n      if (b === "Unknown") return -1;\n      return a.localeCompare(b);\n    });\n    return {\n      status: uniq("status"),\n      severity: uniq("severity"),\n      region: uniq("region"),\n      impact: uniq("impact"),\n      site: uniq("siteId"),\n      openingMonth: openingMonths,\n      openingMonthLabels: Object.fromEntries(openingMonths.map((key) => [key, openingMonthLabel(key)])),\n    };\n''',
        '''    const openingMonths = Array.from(new Set(primaryRows.map((row) => row.openingMonthKey || openingMonthKey(row.observationDate)).filter(Boolean))).sort((a, b) => {\n      if (a === "Unknown") return 1;\n      if (b === "Unknown") return -1;\n      return a.localeCompare(b);\n    });\n    const exportMonths = Array.from(new Set(uniqueRows.flatMap((ticket) =>\n      ticket.rows.flatMap((row) => [recordDateMonthKey(row.observationDate), recordDateMonthKey(row.recoveryDate)]),\n    ).filter((key) => key && key !== "Unknown"))).sort((a, b) => a.localeCompare(b));\n    return {\n      status: uniq("status"),\n      severity: uniq("severity"),\n      region: uniq("region"),\n      impact: uniq("impact"),\n      site: uniq("siteId"),\n      openingMonth: openingMonths,\n      openingMonthLabels: Object.fromEntries(openingMonths.map((key) => [key, openingMonthLabel(key)])),\n      exportMonth: exportMonths,\n      exportMonthLabels: Object.fromEntries(exportMonths.map((key) => [key, openingMonthLabel(key)])),\n    };\n'''
    ),
    (
        '''  const analytics = useMemo(() => {\n''',
        '''  const monthlyExportTickets = useMemo(() => filteredTickets.filter((ticket) => ticketMatchesMonthlyExport(ticket, exportMonth)), [exportMonth, filteredTickets]);\n\n  const selectedExportMonthLabel = exportMonth === "all" ? "All dashboard-filtered TT" : openingMonthLabel(exportMonth);\n\n  const analytics = useMemo(() => {\n'''
    ),
    (
        '''            {data && <button className="ghost-button" onClick={() => exportCsv(filteredTickets)}><Download size={16} /> CSV</button>}\n            {data && <button className="ghost-button" onClick={() => exportExcel(filteredTickets)}><FileSpreadsheet size={16} /> Excel</button>}\n''',
        '''            {data && <button className="ghost-button" onClick={() => exportCsv(monthlyExportTickets)}><Download size={16} /> CSV</button>}\n            {data && <button className="ghost-button" onClick={() => exportExcel(monthlyExportTickets)}><FileSpreadsheet size={16} /> Excel</button>}\n'''
    ),
    (
        '''            <div className="table-heading">\n              <div><span className="section-kicker">Unique register</span><h2>{filteredTickets.length.toLocaleString()} distinct TT records</h2></div>\n              <p>Showing first 150 filtered tickets in the same report order as the source-style register. Site ID and Site Name include all affected sites for each distinct TT, and CSV/Excel exports include every filtered ticket in this same order.</p>\n            </div>\n''',
        '''            <div className="table-heading">\n              <div><span className="section-kicker">Unique register</span><h2>{filteredTickets.length.toLocaleString()} distinct TT records</h2></div>\n              <p>Showing first 150 dashboard-filtered tickets in the same report order as the source-style register. Site ID and Site Name include all affected sites for each distinct TT.</p>\n              <div className="monthly-export-panel no-print">\n                <div>\n                  <strong>Monthly TT export filter</strong>\n                  <span>{monthlyExportTickets.length.toLocaleString()} TT records export for {selectedExportMonthLabel}. Rule: Observation Date in month OR Recovery Date in month OR Status pending.</span>\n                </div>\n                <SelectFilter label="Report Month" value={exportMonth} options={filterOptions.exportMonth} optionLabels={filterOptions.exportMonthLabels} onChange={setExportMonth} />\n                <button className="ghost-button" onClick={() => exportCsv(monthlyExportTickets)}><Download size={16} /> Export CSV</button>\n                <button className="ghost-button" onClick={() => exportExcel(monthlyExportTickets)}><FileSpreadsheet size={16} /> Export Excel</button>\n              </div>\n            </div>\n'''
    ),
]

for old, new in replacements:
    if old not in text:
        raise SystemExit(f'Missing expected text for replacement:\n{old[:300]}')
    text = text.replace(old, new, 1)

path.write_text(text)
