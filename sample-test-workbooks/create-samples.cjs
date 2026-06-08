const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const outDir = path.join(process.cwd(), 'sample-test-workbooks');
fs.mkdirSync(outDir, { recursive: true });

const headers = [
  'SN.', "TT", 'Site ID', 'Site Name', 'Managed Resource ', 'Issues',
  'Severity', 'Region', 'Observation Date', 'Observation Time', 'Recovery Date',
  'Recovery Time', 'Escalated for L3 Support Date', 'Escalated for L3 Support Time',
  'Total Duration/Days/Hours', 'Duration (hrs)', 'NE Detail/Impacted Object',
  'Service Impaction Status', 'No. of correlated Alarms', 'Escalated to ',
  'Escalation Level', 'Status', 'Comments Date', 'Comments-Feedback',
  'Maintenance person/Team', 'Maintenance Contact Details', 'Action Taken/RCA',
  'Action', 'RCA'
];

const actions = [
  ['Auto Restored', 'Power and Environment'],
  ['Power Recycled', 'Power and Environment'],
  ['Fiber Rerouted', 'Transmission and Link'],
  ['Configuration Corrected', 'Software and Configuration'],
  ['Module Replaced', 'Hardware and Device'],
  ['Antenna Alignment', 'Radio RF'],
  ['Vendor Escalated', 'Vendor Third Party'],
  ['Preventive Maintenance', 'Planned Activity'],
  ['No Fault Found', 'Other Review'],
  ['Cable Re-Terminated', 'Transmission and Link'],
  ['Cooling Restored', 'Power and Environment'],
  ['Firmware Updated', 'Software and Configuration']
];

const resources = ['Complete site', 'IP Link Down', 'Repeater', 'Router', 'Switch', 'MW Link', 'Power System', 'Transmission Node', 'Access Node', 'Signal and Coverage issue'];
const issues = ['Site Down', 'Link Down', 'Link Unstable', 'CHU 1 Down', 'High VSWR Alarm', 'Packet Loss', 'Power Alarm', 'Intermittent Connectivity', 'Coverage Degradation', 'Planned Activity'];
const severity = ['Critical', 'Major', 'Minor'];
const statuses = ['Closed', 'Closed', 'Closed', 'Resolved', 'Pending'];
const impacts = ['Service Impact', 'Non Service Impact'];
const escalations = ['L1', 'L2', 'L3', 'L4', 'L5'];
const teams = ['Sample FMD Team', 'Demo Field Team', 'Test NOC', 'Training Maintenance', 'Mock Vendor Team', 'Demo Transmission'];
const siteWords = ['Alpha', 'Bravo', 'Cedar', 'Delta', 'Emerald', 'Falcon', 'Granite', 'Harbor', 'Ivory', 'Jasmine', 'Kingfisher', 'Lagoon', 'Metro', 'Northstar', 'Orchid', 'Palm', 'Quartz', 'River', 'Summit', 'Tide', 'Unity', 'Vertex', 'Westbay', 'Yard', 'Zenith'];

function pad(n) { return String(n).padStart(2, '0'); }
function pick(a, i) { return a[i % a.length]; }
function rand(seed) { const x = Math.sin(seed) * 10000; return x - Math.floor(x); }
function dmy(d) { return `${pad(d.getDate())}/${pad(d.getMonth() + 1)}/${d.getFullYear()}`; }
function hm(d) { return `${pad(d.getHours())}:${pad(d.getMinutes())}`; }
function durationText(hours) {
  if (!Number.isFinite(hours) || hours < 0) return '';
  const mins = Math.round(hours * 60);
  const days = Math.floor(mins / 1440);
  const rem = mins % 1440;
  const h = Math.floor(rem / 60);
  const m = rem % 60;
  return `${days} days ${h} hrs ${m} mins`;
}
function siteName(siteNo, reg) {
  return `Sample ${pick(siteWords, siteNo)} ${pick(['Hub', 'Relay', 'Switching', 'Terminal', 'Node', 'Gateway'], siteNo)} (${reg})`;
}
const TT_PER_REGION = 50;
const SITES_PER_REGION = 10;

function regionSiteNumber(regs, region, siteIndex) {
  const regionOffset = regs.indexOf(region) * SITES_PER_REGION;
  return regionOffset + siteIndex;
}

function makeSites(regs) {
  const rows = [['#', 'Site ID', 'Site Name', 'Region']];
  let sn = 1;
  for (const reg of regs) {
    for (let siteIndex = 1; siteIndex <= SITES_PER_REGION; siteIndex += 1) {
      const siteNo = regionSiteNumber(regs, reg, siteIndex);
      rows.push([sn, `RF Site ${pad(siteNo)}`, siteName(siteNo, reg), reg]);
      sn += 1;
    }
  }
  return rows;
}
function makeRows(regs, ttPrefix) {
  const rows = [headers];
  let sn = 1;
  for (const reg of regs) {
    const regionCode = reg.replace(/[^A-Z0-9]/gi, '').slice(0, 2).toUpperCase();
    for (let regionRow = 1; regionRow <= TT_PER_REGION; regionRow += 1) {
      const i = sn;
      const siteIndex = ((regionRow - 1) % SITES_PER_REGION) + 1;
      const siteNo = regionSiteNumber(regs, reg, siteIndex);
      const siteId = `RF Site ${pad(siteNo)}`;
      const status = pick(statuses, i);
      const start = new Date(2026, 4, 1 + ((i * 3) % 27), (i * 5) % 24, (i * 11) % 60);
      const dur = status === 'Pending' ? null : Math.round((0.25 + rand(i) * 36) * 10) / 10;
      const rec = dur === null ? null : new Date(start.getTime() + dur * 3600000);
      const l3 = i % 4 === 0 ? new Date(start.getTime() + Math.round((0.2 + rand(i + 9) * 2) * 60) * 60000) : null;
      const action = pick(actions, i);
      rows.push([
        sn,
        `${ttPrefix}${regionCode}${pad(regionRow)}`,
        siteId,
        siteName(siteNo, reg),
        pick(resources, i),
        pick(issues, i),
        pick(severity, i),
        reg,
        dmy(start),
        hm(start),
        rec ? dmy(rec) : '',
        rec ? hm(rec) : '',
        l3 ? dmy(l3) : '',
        l3 ? hm(l3) : '',
        dur === null ? 'TT in progress' : durationText(dur),
        dur === null ? '' : dur,
        `Demo impacted object ${pad(i)}`,
        pick(impacts, i),
        String((i * 2) % 5),
        pick(['Demo L2 Desk', 'Sample L3 Support', 'Test Vendor', 'Training Desk'], i),
        pick(escalations, i),
        status,
        status === 'Pending' ? '' : dmy(rec),
        status === 'Pending' ? 'Pending follow-up' : 'Closed after test restoration',
        pick(teams, i),
        `050000${pad(i)}`,
        `${action[0]} completed for sample data`,
        action[0],
        action[1]
      ]);
      sn += 1;
    }
  }
  return rows;
}
function widths(aoa) {
  return aoa[0].map((_, c) => ({
    wch: Math.max(10, Math.min(34, Math.max(...aoa.map((r) => String(r[c] ?? '').length)) + 2))
  }));
}
function addSheet(wb, name, aoa) {
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws['!cols'] = widths(aoa);
  XLSX.utils.book_append_sheet(wb, ws, name);
}

const groups = [
  ['EOA_NEOA_Sample_Test_Workbook.xlsx', ['EOA', 'NEOA'], 'EN'],
  ['SOA_Sample_Test_Workbook.xlsx', ['SOA'], 'SO'],
  ['COA_WOA_Sample_Test_Workbook.xlsx', ['COA', 'WOA'], 'CW']
];

for (const [file, regs, prefix] of groups) {
  const wb = XLSX.utils.book_new();
  addSheet(wb, 'TT-History', makeRows(regs, prefix));
  addSheet(wb, 'Site ID', makeSites(regs));
  addSheet(wb, 'RCA', [['Action', 'RCA'], ...actions]);
  const outputPath = path.join(outDir, file);
  XLSX.writeFile(wb, outputPath, { bookType: 'xlsx' });
  console.log(outputPath);
}

