
import * as fs from "fs";
import * as path from "path";

// Storage directory for admin-uploaded files
const STORAGE_DIR = path.join(process.cwd(), ".data");

// Ensure storage directory exists
if (!fs.existsSync(STORAGE_DIR)) {
  fs.mkdirSync(STORAGE_DIR, { recursive: true });
}

// Simple user credentials (in production, use a real database)
export const USERS: Record<string, { password: string; role: "admin" | "user" }> = {
  admin: { password: "admin123", role: "admin" },
  meenal: { password: "[PASSWORD]", role: "admin" },
  user1: { password: "user123", role: "user" },
  user2: { password: "user123", role: "user" },
  user3: { password: "user123", role: "user" },
};

// Save uploaded file buffer to disk
export function saveAdminFile(filename: string, buffer: Buffer): void {
  const filePath = path.join(STORAGE_DIR, filename);
  fs.writeFileSync(filePath, buffer);
}

// Load admin file from disk
export function loadAdminFile(filename: string): Buffer | null {
  const filePath = path.join(STORAGE_DIR, filename);
  if (fs.existsSync(filePath)) {
    return fs.readFileSync(filePath);
  }
  return null;
}

// Check if admin has uploaded the required files
export function getAdminFileStatus(): {
  alcMapping: boolean;
  rosterDump: boolean;
  alfaOT: boolean;
  fclmPunch: boolean;
  lastUpdated: string | null;
} {
  const statusFile = path.join(STORAGE_DIR, "status.json");
  if (fs.existsSync(statusFile)) {
    const data = JSON.parse(fs.readFileSync(statusFile, "utf-8"));
    return {
      alcMapping: fs.existsSync(path.join(STORAGE_DIR, "workbook.xlsx")),
      rosterDump: fs.existsSync(path.join(STORAGE_DIR, "workbook.xlsx")),
      alfaOT: fs.existsSync(path.join(STORAGE_DIR, "workbook.xlsx")),
      fclmPunch: fs.existsSync(path.join(STORAGE_DIR, "fclm.xlsx")),
      lastUpdated: data.lastUpdated || null,
    };
  }
  return {
    alcMapping: false,
    rosterDump: false,
    alfaOT: false,
    fclmPunch: false,
    lastUpdated: null,
  };
}

// Save upload status
export function saveUploadStatus(): void {
  const statusFile = path.join(STORAGE_DIR, "status.json");
  fs.writeFileSync(
    statusFile,
    JSON.stringify({
      lastUpdated: new Date().toISOString(),
    })
  );
}

