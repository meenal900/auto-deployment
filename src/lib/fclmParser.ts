
import * as XLSX from "xlsx";

// Safe string converter — prevents "Cannot read properties of undefined (reading 'trim')"
function safe(val: any): string {
  if (val === null || val === undefined || val === "") return "";
  return String(val).trim();
}

// ========== FCLM PUNCH PARSER ==========
export function parsePunchSheet(workbook: XLSX.WorkBook) {
  // Try multiple possible sheet names
  const sheetNames = ["FCLM_Punch", "FCLM Punch", "Punch", "Sheet1"];
  let ws = null;
  let usedName = "";
  for (const name of sheetNames) {
    if (workbook.Sheets[name]) { ws = workbook.Sheets[name]; usedName = name; break; }
  }
  // If none found, use first sheet
  if (!ws) {
    const firstName = workbook.SheetNames[0];
    ws = workbook.Sheets[firstName];
    usedName = firstName;
  }
  if (!ws) throw new Error("No sheet found in FCLM file");

  const data: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1 });
  const rows: any[] = [];
  if (data.length === 0) return rows;

  // Find header row
  let headerIdx = 0;
  for (let i = 0; i < Math.min(10, data.length); i++) {
    const row = data[i];
    if (row && row.some((c: any) => {
      const s = safe(c).toLowerCase();
      return s.includes("fclm") || s.includes("badge") || s.includes("employee") || s.includes("punch");
    })) { headerIdx = i; break; }
  }

  const rawHeaders = data[headerIdx] || [];
  const headers: string[] = rawHeaders.map((h: any) => safe(h).toLowerCase());

  // Find columns by header name
  let badgeCol = headers.findIndex((h: string) => h.includes("fclm") || h.includes("badge"));
  if (badgeCol === -1) badgeCol = 0;

  const nameCol = headers.findIndex((h: string) => h.includes("employee") && h.includes("name"));
  const managerCol = headers.findIndex((h: string) => h.includes("manager"));
  const shiftCol = headers.findIndex((h: string) => h === "shift" || h.includes("shift"));
  const punchTypeCol = headers.findIndex((h: string) => h.includes("punch") && h.includes("type"));
  const punchTimeCol = headers.findIndex((h: string) => h.includes("punch") && h.includes("time"));
  const aliasCol = headers.findIndex((h: string) => h.includes("alias") || h.includes("login"));
  const empIDCol = headers.findIndex((h: string) => h.includes("emp") && h.includes("id") && !h.includes("name"));
  const deptCol = headers.findIndex((h: string) => h.includes("dept") || h.includes("department"));

  for (let i = headerIdx + 1; i < data.length; i++) {
    const row = data[i];
    if (!row) continue;

    const badgeVal = safe(row[badgeCol]);
    if (!badgeVal) continue;

    // Filter: only "In" punch types (or if no punch type column)
    if (punchTypeCol >= 0) {
      const punchType = safe(row[punchTypeCol]).toLowerCase();
      if (punchType && punchType !== "in") continue;
    }

    // Determine shift from shift column or punch time
    let punchShift = "";
    if (shiftCol >= 0) punchShift = safe(row[shiftCol]);

    // Check if it's an actual shift label or a code like "ALFA", "FC007"
    const shiftLabels = ["day", "night", "early morning", "evening pt", "early_morning", "evening_pt"];
    const isActualShift = shiftLabels.some((s) => punchShift.toLowerCase().includes(s));

    if (!punchShift || !isActualShift) {
      // Try to determine shift from punch time
      if (punchTimeCol >= 0) {
        const timeStr = safe(row[punchTimeCol]);
        const hour = extractHour(timeStr);
        if (hour >= 0) {
          if (hour >= 5 && hour < 10) punchShift = "Early Morning";
          else if (hour >= 10 && hour < 16) punchShift = "Day";
          else if (hour >= 16 && hour < 20) punchShift = "Evening PT";
          else punchShift = "Night";
        }
      }
      if (!punchShift) punchShift = "Day";
    }

    rows.push({
      badgeID: badgeVal,
      empID: empIDCol >= 0 ? safe(row[empIDCol]) : badgeVal,
      alias: aliasCol >= 0 ? safe(row[aliasCol]) : "",
      name: nameCol >= 0 ? safe(row[nameCol]) : "",
      punchShift,
      department: deptCol >= 0 ? safe(row[deptCol]) : "",
      manager: managerCol >= 0 ? safe(row[managerCol]) : "",
      shiftLabel: shiftCol >= 0 ? safe(row[shiftCol]) : "",
      punchTime: punchTimeCol >= 0 ? safe(row[punchTimeCol]) : "",
    });
  }
  return rows;
}

function extractHour(timeStr: string): number {
  if (!timeStr) return -1;
  // Handle "10/13/26, 10:00 AM" format (with possible special characters)
  const cleanTime = timeStr.replace(/[^\d:APMapm\s]/g, " ").replace(/\s+/g, " ").trim();

  // Try 12h format: "10:00 AM"
  const match = cleanTime.match(/(\d{1,2}):(\d{2})\s*(AM|PM|am|pm)/i);
  if (match) {
    let hour = parseInt(match[1]);
    const ampm = match[3].toUpperCase();
    if (ampm === "PM" && hour !== 12) hour += 12;
    if (ampm === "AM" && hour === 12) hour = 0;
    return hour;
  }
  // Try 24h format: "10:00"
  const match24 = cleanTime.match(/(\d{1,2}):(\d{2})/);
  if (match24) return parseInt(match24[1]);
  return -1;
}

// ========== ALC MAPPING PARSER — Column M = LC ==========
export function parseALCFull(workbook: XLSX.WorkBook) {
  // Try multiple sheet names
  const sheetNames = ["ALC Mapping", "ALC_Mapping", "ALC", "Mapping"];
  let ws = null;
  for (const name of sheetNames) {
    if (workbook.Sheets[name]) { ws = workbook.Sheets[name]; break; }
  }
  if (!ws) throw new Error("Sheet 'ALC Mapping' not found");

  const data: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1 });
  const rows: any[] = [];
  if (data.length === 0) return rows;

  // Find header row
  let headerIdx = 0;
  for (let i = 0; i < Math.min(5, data.length); i++) {
    const row = data[i];
    if (row && row.some((c: any) => {
      const s = safe(c).toLowerCase();
      return s.includes("alias") || s.includes("login") || s.includes("employee") || s.includes("badge");
    })) { headerIdx = i; break; }
  }

  const headers: string[] = (data[headerIdx] || []).map((h: any) => safe(h).toLowerCase());

  const aliasCol = headers.findIndex((h: string) => h.includes("alias") || h.includes("login"));

  const ppCol = headers.findIndex((h: string) =>
    (h.includes("primary") && (h.includes("pp") || h.includes("process") || h.includes("path")))
    || h === "pp" || h === "primary pp" || h === "primarypp" || h.includes("role")
  );

  const secPPCol = headers.findIndex((h: string) =>
    (h.includes("secondary") && (h.includes("pp") || h.includes("process") || h.includes("path")))
    || h === "secondary pp" || h === "secondarypp"
  );

  // LC column — search by header name first, fallback to Column M (index 12)
  let lcCol = headers.findIndex((h: string) =>
    h === "lc" || h === "primary lc" || h === "primarylc"
    || h.includes("learning curve") || h.includes("lc level")
    || h === "lc_level" || h === "lclevel"
    || (h.includes("lc") && !h.includes("calc") && !h.includes("alc"))
  );
  if (lcCol === -1) lcCol = 12; // Fallback: Column M

  const empTypeCol = headers.findIndex((h: string) =>
    h.includes("type") || h.includes("ft/pt") || h.includes("ft") || h === "emp type" || h === "emptype"
  );

  const empIDCol = headers.findIndex((h: string) =>
    (h.includes("emp") && h.includes("id")) || h === "employee id" || h === "employeeid"
  );

  const badgeCol = headers.findIndex((h: string) => h.includes("badge"));
  const nameCol = headers.findIndex((h: string) => h.includes("name") && !h.includes("manager"));

  for (let i = headerIdx + 1; i < data.length; i++) {
    const row = data[i];
    if (!row) continue;

    const alias = aliasCol >= 0 ? safe(row[aliasCol]) : "";
    const primaryPP = ppCol >= 0 ? safe(row[ppCol]) : "";
    const secondaryPP = secPPCol >= 0 ? safe(row[secPPCol]) : "";
    const empType = empTypeCol >= 0 ? safe(row[empTypeCol]) || "FT" : "FT";
    const employeeID = empIDCol >= 0 ? safe(row[empIDCol]) : "";
    const badgeID = badgeCol >= 0 ? safe(row[badgeCol]) : "";
    const empName = nameCol >= 0 ? safe(row[nameCol]) : "";

    // Read LC from Column M (or wherever lcCol points)
    let rawLC = lcCol >= 0 && lcCol < (row.length || 0) ? safe(row[lcCol]) : "";

    // Normalize LC value
    let primaryLC = "LC0";
    if (rawLC) {
      const upper = rawLC.toUpperCase();
      if (upper.match(/^LC\d$/)) primaryLC = upper;
      else if (rawLC.match(/^\d$/)) primaryLC = `LC${rawLC}`;
      else {
        const num = parseInt(rawLC.replace(/\D/g, ""));
        primaryLC = isNaN(num) || num < 0 || num > 9 ? "LC0" : `LC${num}`;
      }
    }

    if (!alias && !employeeID && !badgeID) continue;

    rows.push({ alias, primaryPP, secondaryPP, empType, employeeID, badgeID, empName, primaryLC });
  }
  return rows;
}

// ========== SUMMARY REQUIRED PARSER ==========
export function parseSummaryRequired(workbook: XLSX.WorkBook): Record<string, number> {
  const ws = workbook.Sheets["Summary"];
  if (!ws) return {};
  const data: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1 });
  const result: Record<string, number> = {};
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row) continue;
    const pp = safe(row[0]);
    const req = parseInt(safe(row[1]) || "0") || 0;
    if (pp && req > 0) result[pp] = req;
  }
  return result;
}

