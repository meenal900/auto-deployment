
import { EmployeeData } from "./alcMapping";
import { matchesSecondaryPP, mapPrimaryPPToDeployPP, getPPDescription, PP_LIST } from "./ppMapper";

export interface PPStats {
  required: number;
  deployed: number;
  primaryCount: number;
  secondaryCount: number;
  overflowCount: number;
  shortage: number;
  lc: Record<string, number>;
}

export interface DeployedEmployee {
  badgeID: string;
  employeeName: string;
  employeeAlias: string;
  employeeID: string;
  assignedPP: string;
  deployType: string;
  fskRole: string;
}

export interface DebugEntry {
  source: string;
  employeeID: string;
  employeeAlias: string;
  badgeID: string;
  status: string;
  reason: string;
  primaryPP: string;
  assignedPP: string;
  deployType: string;
}

export interface DeploymentResult {
  summary: Record<string, PPStats>;
  deployment: Record<string, string[]>;
  fskData: DeployedEmployee[];
  debugLog: DebugEntry[];
  rosterCount: number;
  alfaCount: number;
  rosterScanned: number;
  alfaScanned: number;
  rosterUnmatched: number;
  alfaUnmatched: number;
  alfaDuplicates: number;
  totalDeployed: number;
}

export function runDeployment(
  availableAAs: EmployeeData[],
  ppRequirements: Record<string, number>,
  debugLog: DebugEntry[],
  rosterCount: number,
  alfaCount: number,
  rosterScanned: number,
  alfaScanned: number,
  rosterUnmatched: number,
  alfaUnmatched: number,
  alfaDuplicates: number
): DeploymentResult {
  const used = new Set<string>();
  const assignedPPMap = new Map<string, string>();
  const deployTypeMap = new Map<string, string>();

  const ppStats: Record<string, PPStats> = {};
  const ppBadges: Record<string, string[]> = {};
  const ppEmployees: Record<string, DeployedEmployee[]> = {};

  for (const pp of PP_LIST) {
    ppStats[pp] = {
      required: ppRequirements[pp] || 0,
      deployed: 0,
      primaryCount: 0,
      secondaryCount: 0,
      overflowCount: 0,
      shortage: 0,
      lc: { LC0: 0, LC1: 0, LC2: 0, LC3: 0, LC4: 0, LC5: 0 },
    };
    ppBadges[pp] = [];
    ppEmployees[pp] = [];
  }

  const addToStats = (ppName: string, aa: EmployeeData, type: string) => {
    if (!ppStats[ppName]) {
      ppStats[ppName] = {
        required: 0, deployed: 0, primaryCount: 0,
        secondaryCount: 0, overflowCount: 0, shortage: 0,
        lc: { LC0: 0, LC1: 0, LC2: 0, LC3: 0, LC4: 0, LC5: 0 },
      };
      ppBadges[ppName] = [];
      ppEmployees[ppName] = [];
    }

    used.add(aa.EmployeeAlias);
    assignedPPMap.set(aa.EmployeeAlias.toLowerCase(), ppName);
    deployTypeMap.set(aa.EmployeeAlias.toLowerCase(), type);

    ppStats[ppName].deployed++;
    if (type === "Primary") ppStats[ppName].primaryCount++;
    else if (type === "Secondary") ppStats[ppName].secondaryCount++;
    else ppStats[ppName].overflowCount++;

    const lcKey = `LC${Math.min(Math.max(aa.PrimaryLC, 0), 5)}`;
    ppStats[ppName].lc[lcKey]++;

    ppBadges[ppName].push(aa.BadgeID);
    ppEmployees[ppName].push({
      badgeID: aa.BadgeID,
      employeeName: aa.EmployeeName,
      employeeAlias: aa.EmployeeAlias,
      employeeID: aa.EmployeeID,
      assignedPP: ppName,
      deployType: type,
      fskRole: getPPDescription(ppName),
    });
  };

  // PHASE 1: Fill using Primary PP
  for (const pp of PP_LIST) {
    const req = ppStats[pp].required;
    if (req <= 0) continue;
    for (const aa of availableAAs) {
      if (used.has(aa.EmployeeAlias)) continue;
      if (aa.PrimaryPP.toLowerCase().trim() === pp.toLowerCase().trim()) {
        if (ppStats[pp].deployed >= req) break;
        addToStats(pp, aa, "Primary");
      }
    }
  }

  // PHASE 2: Fill shortages using Secondary PP
  for (const pp of PP_LIST) {
    const req = ppStats[pp].required;
    if (ppStats[pp].deployed >= req) continue;
    for (const aa of availableAAs) {
      if (used.has(aa.EmployeeAlias)) continue;
      if (matchesSecondaryPP(aa.SecondaryPP, pp)) {
        addToStats(pp, aa, "Secondary");
        if (ppStats[pp].deployed >= req) break;
      }
    }
  }

  // PHASE 3: Remaining AAs go to their own Primary PP
  for (const aa of availableAAs) {
    if (used.has(aa.EmployeeAlias)) continue;
    const mappedPP = mapPrimaryPPToDeployPP(aa.PrimaryPP);
    addToStats(mappedPP, aa, "Overflow-Own PP");
  }

  // Calculate shortages
  for (const pp of Object.keys(ppStats)) {
    const s = ppStats[pp].required - ppStats[pp].deployed;
    ppStats[pp].shortage = s > 0 ? s : 0;
  }

  // Update debug log
  for (const entry of debugLog) {
    if (entry.status === "DEPLOYED") {
      const key = (entry.employeeAlias || "").toLowerCase().trim();
      entry.assignedPP = assignedPPMap.get(key) || "UNASSIGNED";
      entry.deployType = deployTypeMap.get(key) || "ERROR";
    }
  }

  // Build FSK data
  const fskData: DeployedEmployee[] = [];
  for (const pp of PP_LIST) {
    if (ppEmployees[pp]) fskData.push(...ppEmployees[pp]);
  }
  for (const pp of Object.keys(ppEmployees)) {
    if (!PP_LIST.includes(pp)) fskData.push(...ppEmployees[pp]);
  }

  return {
    summary: ppStats,
    deployment: ppBadges,
    fskData,
    debugLog,
    rosterCount,
    alfaCount,
    rosterScanned,
    alfaScanned,
    rosterUnmatched,
    alfaUnmatched,
    alfaDuplicates,
    totalDeployed: availableAAs.length,
  };
}

