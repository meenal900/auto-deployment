
import * as XLSX from "xlsx";

function safe(val: any): string {
  if (val === null || val === undefined || val === "") return "";
  return String(val).trim();
}

export function parseSheet(workbook: XLSX.WorkBook, sheetName: string): any[] {
  const ws = workbook.Sheets[sheetName];
  if (!ws) return [];
  const data: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1 });
  if (data.length < 2) return [];

  const headers = (data[0] || []).map((h: any) => safe(h));
  const rows: any[] = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row) continue;
    const obj: Record<string, string> = {};
    let hasData = false;
    for (let j = 0; j < headers.length; j++) {
      const val = safe(row[j]);
      obj[headers[j]] = val;
      if (val) hasData = true;
    }
    if (hasData) rows.push(obj);
  }
  return rows;
}

export function findHeader(data: any[][], searchTerms: string[]): number {
  for (let i = 0; i < Math.min(10, data.length); i++) {
    const row = data[i];
    if (!row) continue;
    for (const cell of row) {
      const val = safe(cell).toLowerCase();
      if (searchTerms.some(t => val.includes(t))) return i;
    }
  }
  return 0;
}

