
function safe(val: any): string {
  if (val === null || val === undefined || val === "") return "";
  return String(val).trim();
}

function getEmpType(val: any): string {
  const raw = safe(val).toUpperCase();
  if (raw === "PT" || raw === "PART TIME" || raw === "PART-TIME" || raw === "PARTTIME") return "PT";
  return "FT";
}

function getLCKey(val: any): string {
  const lc = safe(val).toUpperCase();
  if (lc.match(/^LC\d$/)) return lc;
  return "LC0";
}

export function generateActualDeployed(
  fskRows: any[],
  punchRows: any[],
  alcRows: any[],
  ppRequirements: Record<string, number>,
  date: string,
  shift: string
) {
  const alcByBadge = new Map<string, any>();
  const alcByAlias = new Map<string, any>();
  const alcByEmpID = new Map<string, any>();
  for (const a of alcRows) {
    const badge = safe(a.badgeID).toLowerCase();
    const alias = safe(a.alias).toLowerCase();
    const empID = safe(a.employeeID).toLowerCase();
    if (badge) alcByBadge.set(badge, a);
    if (alias) alcByAlias.set(alias, a);
    if (empID) alcByEmpID.set(empID, a);
  }

  const fskByBadge = new Map<string, any>();
  const fskByAlias = new Map<string, any>();
  const fskByEmpID = new Map<string, any>();
  for (const f of fskRows) {
    const badge = safe(f.BadgeId).toLowerCase();
    const alias = safe(f.EmployeeLogin).toLowerCase();
    const empID = safe(f.EmployeeID).toLowerCase();
    if (badge) fskByBadge.set(badge, f);
    if (alias) fskByAlias.set(alias, f);
    if (empID) fskByEmpID.set(empID, f);
  }

  const ppList = Object.keys(ppRequirements);
  const punchedKeys = new Set<string>();
  const auditLog: any[] = [];
  const absentList: any[] = [];
  const skippedNotInALC: string[] = [];

  const summary: Record<string, any> = {};
  const badgesByPP: Record<string, string[]> = {};
  for (const pp of ppList) {
    summary[pp] = {
      required: ppRequirements[pp] || 0,
      deployed: 0, ftDeployed: 0, ptDeployed: 0,
      present: 0, absent: 0, extra: 0,
      shortage: 0, overflow: 0,
      lc: { LC5: 0, LC4: 0, LC3: 0, LC2: 0, LC1: 0, LC0: 0 },
    };
    badgesByPP[pp] = [];
  }

  let totalPresent = 0, totalAbsent = 0, totalExtra = 0, totalFT = 0, totalPT = 0;
  const totalLC: Record<string, number> = { LC5: 0, LC4: 0, LC3: 0, LC2: 0, LC1: 0, LC0: 0 };

  const extractPPFromRole = (role: string): string => {
    const r = safe(role).toLowerCase();
    for (const pp of ppList) { if (r.includes(pp.toLowerCase())) return pp; }
    return "";
  };

  for (const punch of punchRows) {
    const pBadge = safe(punch.badgeID).toLowerCase();
    const pAlias = safe(punch.alias).toLowerCase();
    const pEmpID = safe(punch.empID).toLowerCase();
    if (!pBadge && !pAlias && !pEmpID) continue;

    const punchKey = pBadge || pAlias || pEmpID;
    if (punchedKeys.has(punchKey)) continue;
    punchedKeys.add(punchKey);
    if (pBadge) punchedKeys.add(pBadge);
    if (pAlias) punchedKeys.add(pAlias);
    if (pEmpID) punchedKeys.add(pEmpID);

    const alcMatch = (pBadge && alcByBadge.get(pBadge))
      || (pAlias && alcByAlias.get(pAlias))
      || (pEmpID && alcByEmpID.get(pEmpID));

    if (!alcMatch) {
      skippedNotInALC.push(safe(punch.badgeID) || punchKey);
      auditLog.push({ inFSK: "NO", empID: safe(punch.empID), alias: safe(punch.alias), badgeID: safe(punch.badgeID) || punchKey, status: "SKIPPED", reason: "Not in ALC Mapping", primaryPP: "", assignedPP: "", empType: "", lc: "" });
      continue;
    }

    const fskMatch = (pBadge && fskByBadge.get(pBadge))
      || (pAlias && fskByAlias.get(pAlias))
      || (pEmpID && fskByEmpID.get(pEmpID));

    const empType = getEmpType(alcMatch.empType);
    const lcKey = getLCKey(alcMatch.primaryLC);
    const displayBadge = safe(punch.badgeID) || safe(alcMatch.badgeID) || punchKey;

    let targetPP = "";
    if (fskMatch) {
      const fskPP = safe(fskMatch._assignedPP) || extractPPFromRole(safe(fskMatch.AssignedRole));
      targetPP = (fskPP && summary[fskPP]) ? fskPP : "";
    }
    if (!targetPP) { const p = safe(alcMatch.primaryPP); targetPP = (p && summary[p]) ? p : ""; }
    if (!targetPP) { const s = safe(alcMatch.secondaryPP); targetPP = (s && summary[s]) ? s : ""; }
    if (!targetPP) { let mx = 0; for (const pp of ppList) { const sh = summary[pp].required - summary[pp].deployed; if (sh > mx) { mx = sh; targetPP = pp; } } }
    if (!targetPP && ppList.length > 0) targetPP = ppList[0];
    if (!targetPP) continue;

    if (fskMatch) {
      summary[targetPP].present++;
      summary[targetPP].deployed++;
      if (empType === "FT") { summary[targetPP].ftDeployed++; totalFT++; } else { summary[targetPP].ptDeployed++; totalPT++; }
      if (summary[targetPP].lc[lcKey] !== undefined) summary[targetPP].lc[lcKey]++;
      if (totalLC[lcKey] !== undefined) totalLC[lcKey]++;
      badgesByPP[targetPP].push(displayBadge);
      totalPresent++;
      auditLog.push({ inFSK: "YES", empID: safe(fskMatch.EmployeeID) || safe(punch.empID), alias: safe(fskMatch.EmployeeLogin) || safe(punch.alias), badgeID: displayBadge, status: "Present", reason: "Punched In", primaryPP: safe(alcMatch.primaryPP), assignedPP: targetPP, empType, lc: lcKey });
    } else {
      summary[targetPP].extra++;
      summary[targetPP].deployed++;
      if (empType === "FT") { summary[targetPP].ftDeployed++; totalFT++; } else { summary[targetPP].ptDeployed++; totalPT++; }
      if (summary[targetPP].lc[lcKey] !== undefined) summary[targetPP].lc[lcKey]++;
      if (totalLC[lcKey] !== undefined) totalLC[lcKey]++;
      badgesByPP[targetPP].push(displayBadge);
      totalExtra++;
      auditLog.push({ inFSK: "NO", empID: safe(punch.empID) || safe(alcMatch.employeeID), alias: safe(punch.alias) || safe(alcMatch.alias), badgeID: displayBadge, status: "Extra", reason: "In ALC but not in planned FSK", primaryPP: safe(alcMatch.primaryPP), assignedPP: targetPP, empType, lc: lcKey });
    }
  }

  for (const f of fskRows) {
    const fBadge = safe(f.BadgeId).toLowerCase();
    const fAlias = safe(f.EmployeeLogin).toLowerCase();
    const fEmpID = safe(f.EmployeeID).toLowerCase();
    if (!fBadge && !fAlias && !fEmpID) continue;
    const didPunch = (fBadge && punchedKeys.has(fBadge)) || (fAlias && punchedKeys.has(fAlias)) || (fEmpID && punchedKeys.has(fEmpID));
    if (didPunch) continue;

    const alcMatch = (fBadge && alcByBadge.get(fBadge)) || (fAlias && alcByAlias.get(fAlias)) || (fEmpID && alcByEmpID.get(fEmpID));
    const fskPP = safe(f._assignedPP) || extractPPFromRole(safe(f.AssignedRole));
    const targetPP = (fskPP && summary[fskPP]) ? fskPP : (ppList[0] || "");
    if (!targetPP || !summary[targetPP]) continue;

    const empType = getEmpType(alcMatch?.empType || f._empType);
    const lcKey = getLCKey(alcMatch?.primaryLC || f._primaryLC);
    summary[targetPP].absent++;
    totalAbsent++;
    absentList.push({ badgeID: safe(f.BadgeId), empID: safe(f.EmployeeID), name: safe(f.EmployeeName), alias: safe(f.EmployeeLogin), plannedPP: targetPP, lc: lcKey });
    auditLog.push({ inFSK: "YES", empID: safe(f.EmployeeID), alias: safe(f.EmployeeLogin), badgeID: safe(f.BadgeId), status: "ABSENT", reason: "No punch found", primaryPP: safe(alcMatch?.primaryPP), assignedPP: targetPP, empType, lc: lcKey });
  }

  for (const pp of ppList) { const s = summary[pp]; const d = s.deployed - s.required; if (d < 0) s.shortage = Math.abs(d); else s.overflow = d; }

  const totalHC = totalPresent + totalExtra;
  const totalPlanned = fskRows.length;
  const attendance = totalPlanned > 0 ? `${Math.round((totalPresent / totalPlanned) * 100)}%` : "0%";

  return { summary, badgesByPP, absentList, auditLog, skippedNotInALC, totals: { totalHC, planned: totalPlanned, present: totalPresent, absent: totalAbsent, extra: totalExtra, totalFT, totalPT, attendance, lc: totalLC } };
}

