
function safe(val: any): string {
  if (val === null || val === undefined || val === "") return "";
  return String(val).trim();
}

export function buildALCLookup(alcRows: any[]): Map<string, any> {
  const map = new Map<string, any>();
  for (const row of alcRows) {
    const alias = safe(row.alias);
    if (alias) map.set(alias.toLowerCase(), row);
    const badge = safe(row.badgeID);
    if (badge) map.set(badge.toLowerCase(), row);
    const empID = safe(row.employeeID);
    if (empID) map.set(empID.toLowerCase(), row);
  }
  return map;
}

export function lookupALC(map: Map<string, any>, key: any): any {
  const k = safe(key).toLowerCase();
  return k ? map.get(k) || null : null;
}

