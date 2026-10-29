
// Store is now handled by route.ts with persistent JSON files
// This file is kept for compatibility — no trim calls needed

function safe(val: any): string {
  if (val === null || val === undefined || val === "") return "";
  return String(val).trim();
}

export const store: Record<string, any> = {};

export function getStore(key: string): any {
  return store[safe(key)] || null;
}

export function setStore(key: string, value: any): void {
  store[safe(key)] = value;
}

