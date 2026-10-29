
import { extractLCNumber } from "./alcMapping";
import { PP_LIST } from "./ppMapper";

export interface ActualPPStats {
  required: number;
  deployed: number;
  ftDeployed: number;
  ptDeployed: number;
  present: number;
  extra: number;
  shiftMismatch: number;
  absent: number;
  shortage: number;
  overflow: number;
  lc: Record<string, number>;
}

export interface ActualDebugEntry {
  empID: string;
  badgeID: string;
  name: string;
  alias: string;
  manager: string;
  punchShift: string;
  punchTime: string;
  inALC: string;
  inFSK: string;
  primaryPP: string;
  assignedPP: string;
  lc: string;
  status: string;
  reason: string;
  empType: string;
}

export interface AbsentEntry {
  badgeID: string;
  empID: string;
  name: string;
  alias: string;
  plannedPP: string;
  status: string;
  reason: string;
}

export interface ActualDeployResult {
  summary: Record<string, ActualPPStats>;
  badgesByPP: Record<string, string[]>;
  ftBadgesByPP: Record<string, string[]>;
  ptBadgesByPP: Record<string, string[]>;
  absentList: AbsentEntry[];
  notInALCList: ActualDebugEntry[];
  auditLog: ActualDebugEntry[];
  debugLog: ActualDebugEntry[];
  totals: {
    planned: number;
    present: number;
    absent: number;
    extra: number;
    shiftMismatch: number;
    notInALC: number;
    totalFT: number;
    totalPT: number;
    totalRequired: number;
    totalDeployed: number;
    totalShortage: number;
    totalOverflow: number;
    attendance: string;
  };
}

function cleanID(v: any): string {
  if (v === null || v === undefined || v === "") return "";
  let s = String(v);
  if (s.includes(".")) s = s.split(".")[0];
  s = s.trim();
  while (s.startsWith("0") && s.length > 1) s = s.substring(1);
  return s;
}

export function normalizePPName(rawName: string): string {
  let r = (rawName || "").trim().toLowerCase().replace(/\s+/g, " ");

  if (r === "sort slide" || r === "sortslide" || r === "sort_slide" ||
      r.includes("sort slide") || r.includes("diverter")) return "Sort Slide";
  if (r === "bag pulling" || r === "bagpulling" || r === "bag_pulling" ||
      r.includes("bag pulling") || r.includes("bag pull")) return "Bag Pulling";
  if (r === "bagging" || r === "bag sortation" || r === "secondary sort" ||
      r.includes("bagging") || r.includes("bag sortation")) return "Bagging";
  if (r === "bigbox" || r === "big box" || r === "big_box" || r === "non-con" ||
      r.includes("bigbox") || r.includes("big box") || r.includes("non-con") || r.includes("non con")) return "BigBox";
  if (r === "ps" || r === "problem solve" || r === "problem_solve" || r === "ic-qa" ||
      r.includes("problem solve") || r.includes("ic-qa")) return "PS";
  if (r === "zbs" || r === "fluid" || r.includes("zbs") || r.includes("fluid")) return "ZBS";
  if (r === "inbound" || r === "ib" || r.includes("inbound") ||
      r.includes("unload") || r.includes("receive") || r.includes("stow")) return "Inbound";
  if (r === "outbound" || r === "ob" || r.includes("outbound") ||
      r.includes("loading") || r.includes("dispatch")) return "Outbound";

  return rawName.replace(/\b\w/g, (c) => c.toUpperCase());
}

function extractPPName(fskRole: string): string {
  const r = (fskRole || "").trim().toLowerCase();
  if (r.includes("sort slide") || r.includes("diverter")) return "Sort Slide";
  if (r.includes("bag pulling")) return "Bag Pulling";
  if (r.includes("bag sortation") || r.includes("secondary sorting")) return "Bagging";
  if (r.includes("non-con") || r.includes("big box")) return "BigBox";
  if (r.includes("problem solve") || r.includes("ic-qa")) return "PS";
  if (r.includes("fluid")) return "ZBS";
  if (r.includes("inbound") || r.includes("unload")) return "Inbound";
  if (r.includes("outbound") || r.includes("loading container")) return "Outbound";
  return fskRole;
}

function determineCorrectPP(alcPrimaryPP: string, alcPrimaryFunc: string): string {
  const func = (alcPrimaryFunc || "").trim().toLowerCase();
  const pp = (alcPrimaryPP || "").trim().toLowerCase();

  if (func.includes("sort slide") || func.includes("diverter")) return "Sort Slide";
  if (func.includes("bag pulling")) return "Bag Pulling";
  if (func.includes("bag sortation") || func.includes("secondary sort") || func.includes("primary sort")) return "Bagging";
  if (func.includes("fluid") || func.includes("zbs")) return "ZBS";
  if (func.includes("problem solve") || func.includes("ic-qa")) return "PS";
  if (func.includes("big box") || func.includes("non-con") || func.includes("non con")) return "BigBox";
  if (func.includes("unload") || func.includes("receive") || func.includes("stow")) return "Inbound";
  if (func.includes("loading") || func.includes("dispatch")) return "Outbound";

  if (pp.includes("sort slide") || pp.includes("diverter")) return "Sort Slide";
  if (pp.includes("bag pulling")) return "Bag Pulling";
  if (pp.includes("bagging") || pp.includes("bag sortation") || pp.includes("secondary sort")) return "Bagging";
  if (pp.includes("bigbox") || pp.includes("big box") || pp.includes("non-con")) return "BigBox";
  if (pp.includes("fluid") || pp.includes("zbs")) return "ZBS";
  if (pp.includes("problem solve") || pp === "ps") return "PS";
  if (pp.includes("inbound") || pp.includes("receive") || pp.includes("unload")) return "Inbound";
  if (pp.includes("outbound") || pp.includes("loading") || pp.includes("dispatch")) return "Outbound";

  const fallback: Record<string, string> = {
    inbound: "Inbound", outbound: "Outbound", ps: "PS", zbs: "ZBS",
    "sort slide": "Sort Slide", bagging: "Bagging", bigbox: "BigBox", "bag pulling": "Bag Pulling",
  };
  return fallback[pp] || alcPrimaryPP;
}

interface ALCFullRow {
  alias: string;
  primaryPP: string;
  primaryFunc: string;
  secondaryFunc: string;
  employeeID: string;
  employeeName: string;
  empType: string;
  primaryLC: any;
  badgeID: string;
}

interface FSKRow {
  badgeID: string;
  name: string;
  login: string;
  empID: string;
  role: string;
}

interface PunchRow {
  empID: string;
  name: string;
  manager: string;
  shift: string;
  punchType: string;
  punchTime: string;
}

export function generateActualDeployed(
  fskRows: FSKRow[],
  punchRows: PunchRow[],
  alcFullRows: ALCFullRow[],
  summaryPPRequired: Record<string, number>,
  deployDate: string,
  deployShift: string
): ActualDeployResult {

  const fskDict = new Map<string, { name: string; login: string; empID: string; role: string }>();
  const fskEmpToBadge = new Map<string, string>();

  for (const row of fskRows) {
    const badge = cleanID(row.badgeID);
    const empID = cleanID(row.empID);
    if (!badge) continue;
    if (!fskDict.has(badge)) {
      fskDict.set(badge, { name: row.name, login: row.login, empID, role: row.role });
    }
    if (empID && !fskEmpToBadge.has(empID)) fskEmpToBadge.set(empID, badge);
  }

  const punchList: PunchRow[] = [];
  const punchIDSet = new Set<string>();
  for (const row of punchRows) {
    const empID = cleanID(row.empID);
    if (empID && row.punchType.toUpperCase() === "IN" && !punchIDSet.has(empID)) {
      punchIDSet.add(empID);
      punchList.push({ ...row, empID });
    }
  }

  const alcByEmpID = new Map<string, ALCFullRow>();
  for (const row of alcFullRows) {
    const id = cleanID(row.employeeID);
    if (id && !alcByEmpID.has(id)) alcByEmpID.set(id, row);
  }

  const ppStats: Record<string, ActualPPStats> = {};
  const ppBadges: Record<string, string[]> = {};
  const ftBadges: Record<string, string[]> = {};
  const ptBadges: Record<string, string[]> = {};

  for (const pp of PP_LIST) {
    ppStats[pp] = {
      required: summaryPPRequired[pp] || 0,
      deployed: 0, ftDeployed: 0, ptDeployed: 0,
      present: 0, extra: 0, shiftMismatch: 0, absent: 0,
      shortage: 0, overflow: 0,
      lc: { LC0: 0, LC1: 0, LC2: 0, LC3: 0, LC4: 0, LC5: 0 },
    };
    ppBadges[pp] = [];
    ftBadges[pp] = [];
    ptBadges[pp] = [];
  }

  const debugEntries: ActualDebugEntry[] = [];
  let totalPresent = 0, totalAbsent = 0, totalExtra = 0;
  let totalShiftMismatch = 0, totalNotInALC = 0;
  let totalFT = 0, totalPT = 0;
  const punchedBadges = new Set<string>();

  for (const punch of punchList) {
    const alcRow = alcByEmpID.get(cleanID(punch.empID));

    const entry: ActualDebugEntry = {
      empID: punch.empID, name: punch.name, manager: punch.manager,
      punchShift: punch.shift, punchTime: punch.punchTime,
      badgeID: "", alias: "", primaryPP: "", assignedPP: "",
      lc: "", status: "", reason: "", inALC: "NO", inFSK: "NO", empType: "",
    };

    if (alcRow) {
      const badge = cleanID(alcRow.badgeID);
      entry.inALC = "YES";
      entry.badgeID = badge;
      entry.alias = alcRow.alias.trim();
      entry.primaryPP = alcRow.primaryPP.trim();
      entry.lc = String(extractLCNumber(alcRow.primaryLC));

      const rawEmpType = (alcRow.empType || "").toUpperCase().trim();
      const empType = rawEmpType.includes("PART") ? "PT" : "FT";
      entry.empType = empType;

      punchedBadges.add(badge);
      const fskBadge = fskEmpToBadge.get(cleanID(punch.empID));
      if (fskBadge) punchedBadges.add(fskBadge);

      let assignPP = "";
      let inFSK = false;

      if (fskDict.has(badge)) {
        inFSK = true;
        assignPP = extractPPName(fskDict.get(badge)!.role);
      }
      if (!inFSK && fskBadge && fskDict.has(fskBadge)) {
        inFSK = true;
        assignPP = extractPPName(fskDict.get(fskBadge)!.role);
      }
      if (!assignPP) {
        assignPP = determineCorrectPP(alcRow.primaryPP, alcRow.primaryFunc);
      }

      entry.inFSK = inFSK ? "YES" : "NO";
      entry.assignedPP = assignPP;

      const punchShiftUC = (punch.shift || "").toUpperCase().trim();
      const deployShiftUC = (deployShift || "").toUpperCase().trim();
      const knownShifts = ["ALFA", "NIGHT", "DAY", "EARLY MORNING", "EVENING PT"];
      const isShiftMismatch = punchShiftUC !== "" && punchShiftUC !== deployShiftUC && knownShifts.includes(punchShiftUC);

      if (inFSK) {
        if (isShiftMismatch) {
          entry.status = "DEPLOYED (Shift Mismatch)";
          entry.reason = `In ALC + FSK, Shift=${punch.shift} vs Plan=${deployShift}`;
          totalShiftMismatch++;
        } else {
          entry.status = "DEPLOYED";
          entry.reason = "In ALC + FSK + Punched In";
        }
        totalPresent++;
      } else {
        if (isShiftMismatch) {
          entry.status = "DEPLOYED (Shift Mismatch, Not Planned)";
          entry.reason = "In ALC, punched in, shift mismatch, not in FSK";
          totalShiftMismatch++;
        } else {
          entry.status = "DEPLOYED (Extra)";
          entry.reason = "In ALC + Punched In, but NOT in FSK plan";
        }
        totalExtra++;
      }

      if (!ppStats[assignPP]) {
        ppStats[assignPP] = {
          required: 0, deployed: 0, ftDeployed: 0, ptDeployed: 0,
          present: 0, extra: 0, shiftMismatch: 0, absent: 0,
          shortage: 0, overflow: 0,
          lc: { LC0: 0, LC1: 0, LC2: 0, LC3: 0, LC4: 0, LC5: 0 },
        };
        ppBadges[assignPP] = [];
        ftBadges[assignPP] = [];
        ptBadges[assignPP] = [];
      }

      ppStats[assignPP].deployed++;
      if (inFSK) {
        ppStats[assignPP].present++;
        if (isShiftMismatch) ppStats[assignPP].shiftMismatch++;
      } else {
        ppStats[assignPP].extra++;
        if (isShiftMismatch) ppStats[assignPP].shiftMismatch++;
      }

      if (empType === "FT") {
        ppStats[assignPP].ftDeployed++;
        totalFT++;
        ftBadges[assignPP].push(badge);
      } else {
        ppStats[assignPP].ptDeployed++;
        totalPT++;
        ptBadges[assignPP].push(badge);
      }

      const lcNum = extractLCNumber(alcRow.primaryLC);
      const lcKey = `LC${Math.min(Math.max(lcNum, 0), 5)}`;
      ppStats[assignPP].lc[lcKey]++;

      ppBadges[assignPP].push(badge);
    } else {
      entry.status = "SKIPPED";
      entry.reason = "Punched In but NOT in ALC - Not Deployed";
      totalNotInALC++;
    }

    debugEntries.push(entry);
  }

  const absentList: AbsentEntry[] = [];
  for (const [badge, fskData] of fskDict.entries()) {
    let wasPunched = punchedBadges.has(badge);
    if (!wasPunched && fskData.empID && punchIDSet.has(cleanID(fskData.empID))) {
      wasPunched = true;
    }

    if (!wasPunched) {
      totalAbsent++;
      const abPP = extractPPName(fskData.role);
      absentList.push({
        badgeID: badge,
        empID: fskData.empID,
        name: fskData.name,
        alias: fskData.login,
        plannedPP: abPP,
        status: "ABSENT",
        reason: "In FSK but did not punch in",
      });
      if (ppStats[abPP]) ppStats[abPP].absent++;
    }
  }

  let totalShortage = 0, totalOverflow = 0, totalRequired = 0, totalDeployed = 0;
  for (const pp of Object.keys(ppStats)) {
    const s = ppStats[pp];
    if (s.required > s.deployed) { s.shortage = s.required - s.deployed; }
    else if (s.deployed > s.required) { s.overflow = s.deployed - s.required; }
    totalShortage += s.shortage;
    totalOverflow += s.overflow;
    totalRequired += s.required;
    totalDeployed += s.deployed;
  }

  const planned = fskDict.size;
  const attendance = planned > 0 ? ((totalPresent / planned) * 100).toFixed(1) : "0.0";

  return {
    summary: ppStats,
    badgesByPP: ppBadges,
    ftBadgesByPP: ftBadges,
    ptBadgesByPP: ptBadges,
    absentList,
    notInALCList: debugEntries.filter((e) => e.status === "SKIPPED"),
    auditLog: debugEntries,
    debugLog: debugEntries,
    totals: {
      planned, present: totalPresent, absent: totalAbsent,
      extra: totalExtra, shiftMismatch: totalShiftMismatch,
      notInALC: totalNotInALC, totalFT, totalPT,
      totalRequired, totalDeployed, totalShortage, totalOverflow,
      attendance: attendance + "%",
    },
  };
}

