
function safe(val: any): string {
  if (val === null || val === undefined || val === "") return "";
  return String(val).trim();
}

export function mapPP(alcInfo: any, ppRequirements: Record<string, number>): string {
  if (!alcInfo) return "Unassigned";
  const primary = safe(alcInfo.primaryPP);
  const secondary = safe(alcInfo.secondaryPP);
  if (primary && ppRequirements[primary] !== undefined) return primary;
  if (secondary && ppRequirements[secondary] !== undefined) return secondary;
  return primary || "Unassigned";
}

