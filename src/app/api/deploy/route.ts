
import { NextRequest, NextResponse } from "next/server";
import * as XLSX from "xlsx";
import * as fs from "fs";
import * as path from "path";
import { parseALCFull, parsePunchSheet } from "../../../lib/fclmParser";
import { runDeployment } from "../../../lib/deploymentEngine";
import { generateActualDeployed } from "../../../lib/actualDeployEngine";

// ===== PERSISTENT STORAGE — data survives server restarts =====
const DATA_DIR = path.join(process.cwd(), "data");
if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });

function getFilePath(site: string): string {
  return path.join(DATA_DIR, `${site}_data.json`);
}

function getSiteData(site: string): { alcRows: any[]; rosterRows: any[]; lastUpdated: string } {
  const fp = getFilePath(site);
  if (fs.existsSync(fp)) {
    try {
      const raw = fs.readFileSync(fp, "utf-8");
      return JSON.parse(raw);
    } catch (e) { /* corrupt file, reset */ }
  }
  return { alcRows: [], rosterRows: [], lastUpdated: "" };
}

function saveSiteData(site: string, data: any): void {
  const fp = getFilePath(site);
  fs.writeFileSync(fp, JSON.stringify(data), "utf-8");
}

// ===== API HANDLER =====
export async function POST(req: NextRequest) {
  try {
    const formData = await req.formData();
    const action = formData.get("action") as string;
    const site = (formData.get("site") as string) || "SBCZ";

    // ==================== STATUS ====================
    if (action === "status") {
      const sd = getSiteData(site);
      return NextResponse.json({
        alcMapping: sd.alcRows.length > 0,
        roster: sd.rosterRows.length > 0,
        alcCount: sd.alcRows.length,
        rosterCount: sd.rosterRows.length,
        lastUpdated: sd.lastUpdated || null,
      });
    }

    // ==================== ADMIN UPLOAD ====================
    if (action === "admin-upload") {
      const file = formData.get("file") as File;
      if (!file) return NextResponse.json({ error: "No file uploaded" }, { status: 400 });
      const buffer = Buffer.from(await file.arrayBuffer());
      const workbook = XLSX.read(buffer, { type: "buffer" });
      const sd = getSiteData(site);

      // Parse ALC Mapping
      try {
        const alcRows = parseALCFull(workbook);
        if (alcRows.length > 0) sd.alcRows = alcRows;
      } catch (e: any) {
        console.log("No ALC Mapping sheet:", e.message);
      }

      // Parse Roster
      try {
        const rosterSheet = workbook.Sheets["Roster_Dump"]
          || workbook.Sheets["Roster"]
          || workbook.Sheets["roster"]
          || workbook.Sheets["Roster Dump"];
        if (rosterSheet) {
          const data: any[][] = XLSX.utils.sheet_to_json(rosterSheet, { header: 1 });
          const rows: any[] = [];
          let headerIdx = 0;
          for (let i = 0; i < Math.min(5, data.length); i++) {
            if (data[i] && data[i].some((c: any) =>
              String(c || "").toLowerCase().includes("alias") || String(c || "").toLowerCase().includes("login")
            )) { headerIdx = i; break; }
          }
          const headers = (data[headerIdx] || []).map((h: any) => String(h || "").toLowerCase().trim());
          const aliasCol = headers.findIndex((h: string) => h.includes("alias") || h.includes("login"));
          const shiftCol = headers.findIndex((h: string) => h.includes("shift"));
          for (let i = headerIdx + 1; i < data.length; i++) {
            const row = data[i];
            if (!row) continue;
            const alias = aliasCol >= 0 ? String(row[aliasCol] || "").trim() : "";
            if (!alias) continue;
            rows.push({
              alias,
              shift: shiftCol >= 0 ? String(row[shiftCol] || "").trim() : "",
            });
          }
          if (rows.length > 0) sd.rosterRows = rows;
        }
      } catch (e: any) {
        console.log("No Roster sheet:", e.message);
      }

      sd.lastUpdated = new Date().toISOString();
      saveSiteData(site, sd);
      return NextResponse.json({
        alcRows: sd.alcRows.length,
        rosterRows: sd.rosterRows.length,
      });
    }

    // ==================== SEARCH ALC ====================
    if (action === "search-alc") {
      const query = (formData.get("query") as string || "").toLowerCase().trim();
      if (!query) return NextResponse.json({ results: [] });
      const sd = getSiteData(site);
      const results = sd.alcRows.filter((r: any) => {
        return (r.alias && r.alias.toLowerCase().includes(query))
          || (r.badgeID && String(r.badgeID).toLowerCase().includes(query))
          || (r.employeeID && String(r.employeeID).toLowerCase().includes(query))
          || (r.empName && r.empName.toLowerCase().includes(query));
      }).slice(0, 50);
      return NextResponse.json({ results });
    }

    // ==================== EDIT ALC ====================
    if (action === "edit-alc") {
      const dataStr = formData.get("data") as string;
      if (!dataStr) return NextResponse.json({ error: "No data provided" }, { status: 400 });
      const editData = JSON.parse(dataStr);
      const sd = getSiteData(site);

      let found = false;
      for (let i = 0; i < sd.alcRows.length; i++) {
        const r = sd.alcRows[i];
        const match =
          (editData.originalAlias && r.alias && r.alias.toLowerCase() === editData.originalAlias.toLowerCase())
          || (editData.originalBadge && r.badgeID && String(r.badgeID) === String(editData.originalBadge))
          || (editData.originalEmpID && r.employeeID && String(r.employeeID) === String(editData.originalEmpID));
        if (match) {
          sd.alcRows[i] = {
            ...sd.alcRows[i],
            alias: editData.alias !== undefined ? editData.alias : sd.alcRows[i].alias,
            badgeID: editData.badgeID !== undefined ? editData.badgeID : sd.alcRows[i].badgeID,
            employeeID: editData.employeeID !== undefined ? editData.employeeID : sd.alcRows[i].employeeID,
            empName: editData.empName !== undefined ? editData.empName : sd.alcRows[i].empName,
            primaryPP: editData.primaryPP !== undefined ? editData.primaryPP : sd.alcRows[i].primaryPP,
            secondaryPP: editData.secondaryPP !== undefined ? editData.secondaryPP : sd.alcRows[i].secondaryPP,
            primaryLC: editData.primaryLC !== undefined ? editData.primaryLC : sd.alcRows[i].primaryLC,
            empType: editData.empType !== undefined ? editData.empType : sd.alcRows[i].empType,
          };
          found = true;
          break;
        }
      }
      if (!found) return NextResponse.json({ error: "AA not found in ALC data" }, { status: 404 });
      sd.lastUpdated = new Date().toISOString();
      saveSiteData(site, sd);
      return NextResponse.json({ success: true });
    }

    // ==================== ADD ALC ====================
    if (action === "add-alc") {
      const dataStr = formData.get("data") as string;
      if (!dataStr) return NextResponse.json({ error: "No data provided" }, { status: 400 });
      const newData = JSON.parse(dataStr);
      const sd = getSiteData(site);

      // Check for duplicate
      const dup = sd.alcRows.find((r: any) =>
        (newData.alias && r.alias && r.alias.toLowerCase() === newData.alias.toLowerCase())
        || (newData.badgeID && r.badgeID && String(r.badgeID) === String(newData.badgeID))
      );
      if (dup) return NextResponse.json({ error: `AA already exists: ${dup.alias || dup.badgeID}` }, { status: 400 });

      sd.alcRows.push({
        alias: newData.alias || "",
        badgeID: newData.badgeID || "",
        employeeID: newData.employeeID || "",
        empName: newData.empName || "",
        primaryPP: newData.primaryPP || "",
        secondaryPP: newData.secondaryPP || "",
        primaryLC: newData.primaryLC || "LC0",
        empType: newData.empType || "FT",
      });
      sd.lastUpdated = new Date().toISOString();
      saveSiteData(site, sd);
      return NextResponse.json({ success: true, totalRows: sd.alcRows.length });
    }

    // ==================== SCAN ====================
    if (action === "scan") {
      const shiftVal = formData.get("shift") as string;
      const alfaRaw = formData.get("alfaLogins") as string || "";
      const sd = getSiteData(site);

      if (sd.alcRows.length === 0) {
        return NextResponse.json({ error: "No ALC data. Upload via Admin Panel first." }, { status: 400 });
      }

      const alfaLogins = alfaRaw.split(/[\n,;]+/).map((s: string) => s.trim()).filter(Boolean);
      const alcByAlias = new Map<string, any>();
      for (const a of sd.alcRows) {
        if (a.alias) alcByAlias.set(a.alias.toLowerCase(), a);
      }

      // Filter roster by shift
      let rosterCount = 0;
      let rosterUnmatched = 0;
      const rosterAliases = new Set<string>();
      for (const r of sd.rosterRows) {
        const rShift = String(r.shift || "").toLowerCase();
        const dShift = shiftVal.toLowerCase();
        if (rShift.includes(dShift) || dShift.includes(rShift) || !rShift) {
          if (alcByAlias.has(r.alias.toLowerCase())) {
            rosterAliases.add(r.alias.toLowerCase());
            rosterCount++;
          } else {
            rosterUnmatched++;
          }
        }
      }

      // Alfa / OT logins
      let alfaCount = 0;
      let alfaUnmatched = 0;
      let alfaDuplicates = 0;
      for (const a of alfaLogins) {
        const key = a.toLowerCase();
        if (rosterAliases.has(key)) { alfaDuplicates++; continue; }
        if (alcByAlias.has(key)) { rosterAliases.add(key); alfaCount++; }
        else { alfaUnmatched++; }
      }

      return NextResponse.json({
        totalAvailable: rosterCount + alfaCount,
        rosterCount,
        alfaCount,
        rosterScanned: sd.rosterRows.length,
        rosterUnmatched,
        alfaScanned: alfaLogins.length,
        alfaUnmatched,
        alfaDuplicates,
      });
    }

    // ==================== DEPLOY ====================
    if (action === "deploy") {
      const dateVal = formData.get("date") as string;
      const shiftVal = formData.get("shift") as string;
      const alfaRaw = formData.get("alfaLogins") as string || "";
      const ppReqs = JSON.parse(formData.get("ppRequirements") as string || "{}");
      const sd = getSiteData(site);

      if (sd.alcRows.length === 0) {
        return NextResponse.json({ error: "No ALC data. Upload via Admin Panel first." }, { status: 400 });
      }

      const alfaLogins = alfaRaw.split(/[\n,;]+/).map((s: string) => s.trim()).filter(Boolean);
      const deploy = runDeployment as (...args: any[]) => ReturnType<typeof runDeployment>;
      const result = deploy(sd.alcRows, sd.rosterRows, alfaLogins, ppReqs, dateVal, shiftVal);
      return NextResponse.json(result);
    }

    // ==================== ACTUAL DEPLOYED ====================
    if (action === "actual") {
      const dateVal = formData.get("date") as string;
      const shiftVal = formData.get("shift") as string;
      const fclmFile = formData.get("fclmFile") as File;
      const fskDataRaw = formData.get("fskData") as string || "[]";
      const ppReqs = JSON.parse(formData.get("ppRequirements") as string || "{}");
      const sd = getSiteData(site);

      if (!fclmFile) {
        return NextResponse.json({ error: "No FCLM file uploaded" }, { status: 400 });
      }

      const buffer = Buffer.from(await fclmFile.arrayBuffer());
      const workbook = XLSX.read(buffer, { type: "buffer" });
      const punchRows = parsePunchSheet(workbook);
      const fskRows = JSON.parse(fskDataRaw);

      const result = generateActualDeployed(fskRows, punchRows, sd.alcRows, ppReqs, dateVal, shiftVal);
      return NextResponse.json(result);
    }

    return NextResponse.json({ error: "Unknown action: " + action }, { status: 400 });
  } catch (err: any) {
    console.error("API Error:", err);
    return NextResponse.json({ error: err.message || "Server error" }, { status: 500 });
  }
}

