// src/lib/rcaRules.ts

import { clean } from "./dateUtils";

export const RCA_FAMILY_MAP: Record<string, string> = {
  "Power Issue": "Power & Environment",
  "High Temperature": "Power & Environment",
  "High VSWR": "Power & Environment",
  "Weather Issue": "Power & Environment",
  "DC Charger Faulty": "Power & Environment",

  "Link Down": "Transmission & Link",
  "Link Flapping": "Transmission & Link",
  Transmission: "Transmission & Link",
  "MW Issue": "Transmission & Link",
  "MPLS Issue": "Transmission & Link",
  "SDH Hanged": "Transmission & Link",

  "Hardware Failure": "Hardware & Device",
  "Hardware Faulty": "Hardware & Device",
  "Device Hanged": "Hardware & Device",

  "Fiber Cut": "Fiber & Physical",
  Cabling: "Fiber & Physical",
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

export const NON_PREVENTABLE_RCAS = new Set([
  "Weather Issue",
  "Approved Activity",
  "Planned Activity",
  "RCA not Provided",
]);

export const RESPONSIBLE_TEAM_BY_FAMILY: Record<string, string> = {
  "Power & Environment": "Power / Facilities Team",
  "Transmission & Link": "Transmission / NOC Team",
  "Hardware & Device": "Field Maintenance / Vendor",
  "Fiber & Physical": "Fiber / Physical Maintenance Team",
  "Configuration / Software": "NOC / Configuration Team",
  "Human / Process / Planned": "Process Owner / Project Team",
  "Unknown / Missing": "RCA Owner / Follow-up Required",
  "Other / Review": "Operations Review Team",
};

export const RECOMMENDED_ACTION_BY_FAMILY: Record<string, string> = {
  "Power & Environment":
    "Check power source, rectifier, batteries, grounding, cooling, and repeated environmental alarms.",
  "Transmission & Link":
    "Review link stability, transmission path, MPLS/MW/SDH health, and vendor escalation history.",
  "Hardware & Device":
    "Inspect device health, replace faulty hardware, verify spares, and monitor repeated failures.",
  "Fiber & Physical":
    "Inspect fiber/cabling route, port status, optical levels, patching quality, and civil-work exposure.",
  "Configuration / Software":
    "Review recent changes, configuration backup, software version, rollback records, and approval controls.",
  "Human / Process / Planned":
    "Validate activity approval, handover, method of procedure, and team process compliance.",
  "Unknown / Missing":
    "Complete RCA, assign owner, and update action taken before closure reporting.",
  "Other / Review":
    "Review RCA text manually and assign the correct operational owner.",
};

export function getRcaFamily(rca: string): string {
  const normalized = clean(rca) || "RCA not Provided";
  return RCA_FAMILY_MAP[normalized] ?? "Other / Review";
}

export function getPreventability(rca: string): string {
  const normalized = clean(rca) || "RCA not Provided";

  return NON_PREVENTABLE_RCAS.has(normalized)
    ? "Non-preventable"
    : "Preventable";
}

export function getResponsibleTeam(rcaFamily: string): string {
  return (
    RESPONSIBLE_TEAM_BY_FAMILY[rcaFamily] ??
    RESPONSIBLE_TEAM_BY_FAMILY["Other / Review"]
  );
}

export function getRecommendedAction(rcaFamily: string): string {
  return (
    RECOMMENDED_ACTION_BY_FAMILY[rcaFamily] ??
    RECOMMENDED_ACTION_BY_FAMILY["Other / Review"]
  );
}

export function rcaNotProvided(rca: string): boolean {
  const normalized = clean(rca).toLowerCase();

  return !normalized || normalized === "rca not provided";
}
