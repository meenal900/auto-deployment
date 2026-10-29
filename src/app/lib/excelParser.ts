
import * as XLSX from "xlsx";
import { ALCRow, EmployeeData, getALCMappingByID, getALCMappingByAlias } from "./alcMapping";
import { isMatchingShift } from "./shiftMatcher";
import { DebugEntry } from "./deployEngine";

export function parseALCMapping(workbook: XLSX.WorkBook): ALCRow[] {
  const ws = workbook.Sheets["ALC Mapping"];
  if (!ws) throw new Error("Sheet 'ALC Mapping' not found");
  const data: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1 });
  const rows: ALCRow[] = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row || row.length === 0) continue;
    rows.push({
      alias: String(row[0] || ""),
      primaryPP: String(row[1] || ""),
      secondaryPP: String(row[3] || ""),
      employeeID: String(row[4] || ""),
      employeeName: String(row[5] || ""),
      primaryLC: row[12] ?? "",
      badgeID: String(row[13] || ""),
    });
  }
  return rows;
}

function findDateColumn(headerRow: any[], targetDate: Date): number {
  for (let c = 0; c < headerRow.length; c++) {
    const cellVal = headerRow[c];
    if (cellVal instanceof Date) {
      if (cellVal.toDateString() === targetDate.toDateString()) return c;
    } else if (typeof cellVal === "number") {
      const d = XLSX.SSF.parse_date_code(cellVal);
      if (d) {
        const parsed = new Date(d.y, d.m - 1, d.d);
        if (parsed.toDateString() === targetDate.toDateString()) return c;
      }
    } else if (typeof cellVal === "string") {
      const parsed = new Date(cellVal);
      if (!isNaN(parsed.getTime()) && parsed.toDateString() === targetDate.toDateString()) return c;
    }
  }
  return -1;
}

export interface CollectionResult {
  availableAAs: EmployeeData[];
  debugLog: DebugEntry[];
  rosterCount: number;
  alfaCount: number;
  rosterScanned: number;
  alfaScanned: number;
  rosterUnmatched: number;
  alfaUnmatched: number;
  alfaDuplicates: number;
}

export function collectAvailableAAs(
  workbook: XLSX.WorkBook,
  alcRows: ALCRow[],
  deployDate: Date,
  shiftType: string,
  alfaLogins: string[] = []
): CollectionResult {
  const availableAAs: EmployeeData[] = [];
  const debugLog: DebugEntry[] = [];
  let rosterCount = 0, alfaCount = 0;
  let rosterScanned = 0, alfaScanned = 0;
  let rosterUnmatched = 0, alfaUnmatched = 0, alfaDuplicates = 0;

  // --- ROSTER_DUMP (from admin-uploaded workbook) ---
  const wsRoster = workbook.Sheets["Roster_Dump"];
  if (!wsRoster) throw new Error("Sheet 'Roster_Dump' not found");

  const rosterData: any[][] = XLSX.utils.sheet_to_json(wsRoster, { header: 1, raw: false, dateNF: "yyyy-mm-dd" });
  const dateCol = findDateColumn(rosterData[0] || [], deployDate);

  if (dateCol === -1) throw new Error(`Date ${deployDate.toLocaleDateString()} not found in Roster_Dump header row`);

  for (let i = 1; i < rosterData.length; i++) {
    const row = rosterData[i];
    if (!row) continue;
    const shiftValue = String(row[dateCol] || "").trim();
    const upperShift = shiftValue.toUpperCase();

    if (shiftValue && upperShift !== "WEEKOFF" && upperShift !== "WO") {
      if (isMatchingShift(shiftValue, shiftType)) {
        rosterScanned++;
        const employeeID = String(row[3] || "").trim();
        const alcData = getALCMappingByID(alcRows, employeeID);

        if (alcData) {
          availableAAs.push(alcData);
          rosterCount++;
          debugLog.push({
            source: "Roster", employeeID: alcData.EmployeeID,
            employeeAlias: alcData.EmployeeAlias, badgeID: alcData.BadgeID,
            status: "DEPLOYED", reason: "In ALC", primaryPP: alcData.PrimaryPP,
            assignedPP: "", deployType: "",
          });
        } else {
          rosterUnmatched++;
          debugLog.push({
            source: "Roster", employeeID, employeeAlias: "", badgeID: "",
            status: "SKIPPED", reason: "NOT in ALC", primaryPP: "",
            assignedPP: "N/A", deployType: "N/A",
          });
        }
      }
    }
  }

  // --- ALFA OT (from user-entered login IDs) ---
  if (alfaLogins && alfaLogins.length > 0) {
    const alfaAliasSet = new Set<string>();

    for (const empAlias of alfaLogins) {
      const trimmedAlias = empAlias.trim().toLowerCase();
      if (!trimmedAlias) continue;

      alfaScanned++;

      // Check duplicate within Alfa list
      if (alfaAliasSet.has(trimmedAlias)) {
        alfaDuplicates++;
        debugLog.push({
          source: "Alfa_OT", employeeID: "", employeeAlias: trimmedAlias, badgeID: "",
          status: "DUPLICATE", reason: "Duplicate in Alfa input", primaryPP: "",
          assignedPP: "N/A", deployType: "N/A",
        });
        continue;
      }

      alfaAliasSet.add(trimmedAlias);
      const alfaEmployee = getALCMappingByAlias(alcRows, trimmedAlias);

      if (alfaEmployee) {
        // Check if already in Roster
        const isDuplicate = availableAAs.some(
          (aa) => aa.EmployeeAlias.toLowerCase().trim() === alfaEmployee.EmployeeAlias.toLowerCase().trim()
        );

        if (isDuplicate) {
          alfaDuplicates++;
          debugLog.push({
            source: "Alfa_OT", employeeID: alfaEmployee.EmployeeID,
            employeeAlias: trimmedAlias, badgeID: alfaEmployee.BadgeID,
            status: "DUPLICATE", reason: "Already from Roster", primaryPP: "",
            assignedPP: "N/A", deployType: "N/A",
          });
        } else {
          availableAAs.push(alfaEmployee);
          alfaCount++;
          debugLog.push({
            source: "Alfa_OT", employeeID: alfaEmployee.EmployeeID,
            employeeAlias: alfaEmployee.EmployeeAlias, badgeID: alfaEmployee.BadgeID,
            status: "DEPLOYED", reason: "In ALC (Alfa OT)", primaryPP: alfaEmployee.PrimaryPP,
            assignedPP: "", deployType: "",
          });
        }
      } else {
        alfaUnmatched++;
        debugLog.push({
          source: "Alfa_OT", employeeID: "", employeeAlias: trimmedAlias, badgeID: "",
          status: "SKIPPED", reason: "NOT in ALC", primaryPP: "",
          assignedPP: "N/A", deployType: "N/A",
        });
      }
    }
  }

  return {
    availableAAs, debugLog, rosterCount, alfaCount,
    rosterScanned, alfaScanned, rosterUnmatched, alfaUnmatched, alfaDuplicates,
  };
}

