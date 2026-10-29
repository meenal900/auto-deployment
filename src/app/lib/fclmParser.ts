
import * as XLSX from "xlsx";
import { normalizePPName } from "./actualDeployEngine";

export interface FSKRow {
  badgeID: string;
  name: string;
  login: string;
  empID: string;
  role: string;
}

export function parseFSKSheet(workbook: XLSX.WorkBook): { rows: FSKRow[]; deployDate: string; deployShift: string } {
  const ws = workbook.Sheets["FSK"];
  if (!ws) throw new Error("Sheet 'FSK' not found");
  const data: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1 });
  const rows: FSKRow[] = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row || !row[0]) continue;
    rows.push({
      badgeID: String(row[0] || "").trim(),
      name: String(row[1] || "").trim(),
      login: String(row[2] || "").trim(),
      empID: String(row[3] || "").trim(),
      role: String(row[4] || "").trim(),
    });
  }

  const deployDate = String(data[0]?.[25] || "").trim();
  const deployShift = String(data[1]?.[25] || "").trim();

  return { rows, deployDate, deployShift };
}

export interface PunchRow {
  empID: string;
  name: string;
  manager: string;
  shift: string;
  punchType: string;
  punchTime: string;
}

export function parsePunchSheet(workbook: XLSX.WorkBook): PunchRow[] {
  const ws = workbook.Sheets["FCLM_Punch"];
  if (!ws) throw new Error("Sheet 'FCLM_Punch' not found");
  const data: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1 });
  const rows: PunchRow[] = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row || !row[0]) continue;
    rows.push({
      empID: String(row[0] || "").trim(),
      name: String(row[1] || "").trim(),
      manager: String(row[2] || "").trim(),
      shift: String(row[3] || "").trim(),
      punchType: String(row[4] || "").trim(),
      punchTime: String(row[5] || "").trim(),
    });
  }
  return rows;
}

export interface ALCFullRow {
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

export function parseALCFull(workbook: XLSX.WorkBook): ALCFullRow[] {
  const ws = workbook.Sheets["ALC Mapping"];
  if (!ws) throw new Error("Sheet 'ALC Mapping' not found");
  const data: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1 });
  const rows: ALCFullRow[] = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row || row.length === 0) continue;
    rows.push({
      alias: String(row[0] || ""),
      primaryPP: String(row[1] || ""),
      primaryFunc: String(row[2] || ""),
      secondaryFunc: String(row[3] || ""),
      employeeID: String(row[4] || ""),
      employeeName: String(row[5] || ""),
      empType: String(row[6] || ""),
      primaryLC: row[12] ?? "",
      badgeID: String(row[13] || ""),
    });
  }
  return rows;
}

export function parseSummaryRequired(workbook: XLSX.WorkBook): Record<string, number> {
  const ws = workbook.Sheets["Summary"];
  if (!ws) throw new Error("Sheet 'Summary' not found");
  const data: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1 });
  const result: Record<string, number> = {};

  for (let i = 4; i < data.length; i++) {
    const row = data[i];
    if (!row) continue;
    const ppRaw = String(row[0] || "").trim();
    const reqVal = Number(row[1]) || 0;
    if (!ppRaw) continue;

    const ppNorm = normalizePPName(ppRaw);
    if (ppNorm) {
      result[ppNorm] = (result[ppNorm] || 0) + reqVal;
    }
  }
  return result;
}

