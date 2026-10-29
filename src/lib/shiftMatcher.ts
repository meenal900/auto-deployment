
function safe(val: any): string {
  if (val === null || val === undefined || val === "") return "";
  return String(val).trim();
}

export function matchesShift(rosterShift: any, targetShift: any): boolean {
  const r = safe(rosterShift).toLowerCase();
  const t = safe(targetShift).toLowerCase();
  if (!r || !t) return true;
  if (r === t) return true;
  if (r.includes(t) || t.includes(r)) return true;
  const aliases: Record<string, string[]> = {
    "day": ["day", "d", "day shift"],
    "night": ["night", "n", "night shift"],
    "early morning": ["early morning", "em", "early_morning", "earlymorning"],
    "evening pt": ["evening pt", "ept", "evening_pt", "eveningpt", "evening"],
  };
  for (const [key, vals] of Object.entries(aliases)) {
    if (vals.includes(r) && vals.includes(t)) return true;
    if (key === t && vals.includes(r)) return true;
    if (key === r && vals.includes(t)) return true;
  }
  return false;
}

