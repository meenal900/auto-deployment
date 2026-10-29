
export interface EmployeeData {
  EmployeeID: string;
  EmployeeAlias: string;
  EmployeeName: string;
  BadgeID: string;
  PrimaryPP: string;
  SecondaryPP: string;
  PrimaryLC: number;
  Source: string;
}

export function cleanBadge(v: any): string {
  if (v === null || v === undefined || v === "") return "";
  let s = String(v);
  if (s.includes(".")) s = s.split(".")[0];
  s = s.trim();
  while (s.startsWith("0") && s.length > 1) s = s.substring(1);
  return s;
}

export function extractLCNumber(lcValue: any): number {
  if (lcValue === null || lcValue === undefined || lcValue === "") return 0;
  let lcStr = String(lcValue).trim().toUpperCase();
  if (!isNaN(Number(lcStr))) return parseInt(lcStr, 10);
  if (lcStr.startsWith("LC")) {
    const num = lcStr.substring(2);
    if (!isNaN(Number(num))) return parseInt(num, 10);
  }
  return 0;
}

export interface ALCRow {
  alias: string;
  primaryPP: string;
  secondaryPP: string;
  employeeID: string;
  employeeName: string;
  primaryLC: any;
  badgeID: string;
}

export function getALCMappingByID(alcRows: ALCRow[], employeeID: string): EmployeeData | null {
  const rosterEmpID = cleanBadge(employeeID);
  if (!rosterEmpID) return null;

  for (const row of alcRows) {
    const alcEmpID = cleanBadge(row.employeeID);
    if (alcEmpID === rosterEmpID) {
      return {
        EmployeeID: rosterEmpID,
        EmployeeAlias: (row.alias || "").trim(),
        EmployeeName: (row.employeeName || "").trim(),
        BadgeID: String(row.badgeID || "").trim(),
        PrimaryPP: (row.primaryPP || "").trim(),
        SecondaryPP: (row.secondaryPP || "").trim(),
        PrimaryLC: extractLCNumber(row.primaryLC),
        Source: "Roster",
      };
    }
  }
  return null;
}

export function getALCMappingByAlias(alcRows: ALCRow[], employeeAlias: string): EmployeeData | null {
  const inputAlias = (employeeAlias || "").trim().toLowerCase();
  if (!inputAlias) return null;

  for (const row of alcRows) {
    const alcAlias = (row.alias || "").trim().toLowerCase();
    if (alcAlias === inputAlias) {
      return {
        EmployeeID: cleanBadge(row.employeeID),
        EmployeeAlias: (row.alias || "").trim(),
        EmployeeName: (row.employeeName || "").trim(),
        BadgeID: String(row.badgeID || "").trim(),
        PrimaryPP: (row.primaryPP || "").trim(),
        SecondaryPP: (row.secondaryPP || "").trim(),
        PrimaryLC: extractLCNumber(row.primaryLC),
        Source: "Alfa_OT",
      };
    }
  }
  return null;
}

