
function safe(val: any): string {
  if (val === null || val === undefined || val === "") return "";
  return String(val).trim();
}

function getEmpType(person: any): string {
  const raw = safe(person.empType).toUpperCase();
  if (raw === "PT" || raw === "PART TIME" || raw === "PART-TIME" || raw === "PARTTIME") return "PT";
  return "FT";
}

function getLCKey(person: any): string {
  const lc = safe(person.primaryLC).toUpperCase();
  if (lc.match(/^LC\d$/)) return lc;
  return "LC0";
}

export function runDeployment(
  alcRows: any[],
  rosterRows: any[],
  alfaLogins: string[],
  ppRequirements: Record<string, number>,
  date: string,
  shift: string
) {
  const alcByAlias = new Map<string, any>();
  for (const a of alcRows) {
    const alias = safe(a.alias);
    if (alias) alcByAlias.set(alias.toLowerCase(), a);
  }

  const available: any[] = [];
  const usedAliases = new Set<string>();
  const skippedNotInALC: string[] = [];

  for (const r of rosterRows) {
    const alias = safe(r.alias).toLowerCase();
    if (!alias) continue;
    const rShift = safe(r.shift).toLowerCase();
    const dShift = shift.toLowerCase();
    if (rShift && !rShift.includes(dShift) && !dShift.includes(rShift)) continue;
    const alc = alcByAlias.get(alias);
    if (!alc) { skippedNotInALC.push(alias); continue; }
    if (usedAliases.has(alias)) continue;
    usedAliases.add(alias);
    available.push({ ...alc, source: "Roster" });
  }

  for (const a of alfaLogins) {
    const alias = safe(a).toLowerCase();
    if (!alias || usedAliases.has(alias)) continue;
    const alc = alcByAlias.get(alias);
    if (!alc) { skippedNotInALC.push(alias); continue; }
    usedAliases.add(alias);
    available.push({ ...alc, source: "Alfa/OT" });
  }

  const lcOrder: Record<string, number> = { LC5: 0, LC4: 1, LC3: 2, LC2: 3, LC1: 4, LC0: 5 };
  available.sort((a, b) => {
    const aLC = lcOrder[getLCKey(a)] ?? 5;
    const bLC = lcOrder[getLCKey(b)] ?? 5;
    if (aLC !== bLC) return aLC - bLC;
    const aFT = getEmpType(a) === "FT" ? 0 : 1;
    const bFT = getEmpType(b) === "FT" ? 0 : 1;
    return aFT - bFT;
  });

  const ppList = Object.keys(ppRequirements);
  const assigned = new Set<string>();
  const fskData: any[] = [];
  const summary: Record<string, any> = {};

  for (const pp of ppList) {
    summary[pp] = {
      required: ppRequirements[pp] || 0, deployed: 0,
      primaryCount: 0, secondaryCount: 0, overflowCount: 0, shortage: 0,
      lc: { LC5: 0, LC4: 0, LC3: 0, LC2: 0, LC1: 0, LC0: 0 },
    };
  }

  const addToFSK = (person: any, pp: string, assignType: string) => {
    const empType = getEmpType(person);
    const lcKey = getLCKey(person);
    assigned.add(safe(person.alias).toLowerCase());
    summary[pp].deployed++;
    if (assignType === "Primary") summary[pp].primaryCount++;
    else if (assignType === "Secondary") summary[pp].secondaryCount++;
    else summary[pp].overflowCount++;
    if (summary[pp].lc[lcKey] !== undefined) summary[pp].lc[lcKey]++;
    fskData.push({
      BadgeId: safe(person.badgeID),
      EmployeeName: safe(person.empName),
      EmployeeLogin: safe(person.alias),
      EmployeeID: safe(person.employeeID),
      AssignedRole: pp.toLowerCase() + ".lets rock!",
      _empType: empType,
      _source: safe(person.source),
      _assignType: assignType,
      _primaryLC: lcKey,
      _assignedPP: pp,
    });
  };

  for (const pp of ppList) {
    const needed = ppRequirements[pp] || 0;
    if (needed <= 0) continue;
    for (const person of available) {
      if (summary[pp].deployed >= needed) break;
      const alias = safe(person.alias).toLowerCase();
      if (assigned.has(alias)) continue;
      if (safe(person.primaryPP).toLowerCase() === pp.toLowerCase()) addToFSK(person, pp, "Primary");
    }
  }

  for (const pp of ppList) {
    const needed = ppRequirements[pp] || 0;
    if (summary[pp].deployed >= needed) continue;
    for (const person of available) {
      if (summary[pp].deployed >= needed) break;
      const alias = safe(person.alias).toLowerCase();
      if (assigned.has(alias)) continue;
      if (safe(person.secondaryPP).toLowerCase() === pp.toLowerCase()) addToFSK(person, pp, "Secondary");
    }
  }

  for (const pp of ppList) {
    const needed = ppRequirements[pp] || 0;
    if (summary[pp].deployed >= needed) continue;
    for (const person of available) {
      if (summary[pp].deployed >= needed) break;
      const alias = safe(person.alias).toLowerCase();
      if (assigned.has(alias)) continue;
      addToFSK(person, pp, "Overflow");
    }
  }

  for (const pp of ppList) {
    summary[pp].shortage = Math.max(0, summary[pp].required - summary[pp].deployed);
  }

  let rosterCount = 0, alfaCount = 0;
  for (const f of fskData) {
    if (safe(f._source).toLowerCase().includes("alfa") || safe(f._source).toLowerCase().includes("ot")) alfaCount++;
    else rosterCount++;
  }

  return { fskData, summary, totalDeployed: fskData.length, rosterCount, alfaCount, totalAvailable: available.length, skippedNotInALC };
}

