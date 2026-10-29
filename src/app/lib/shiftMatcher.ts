
// 4 SHIFTS matching logic
export function isMatchingShift(rosterShiftValue: string, selectedShift: string): boolean {
  const sv = (rosterShiftValue || "").trim().toUpperCase();
  if (!sv || sv === "WEEKOFF" || sv === "WO") return false;

  switch (selectedShift) {
    case "Early Morning":
      return sv.includes("06:00-10:00") || sv.includes("06:00-16:00");
    case "Day":
      return sv.includes("10:00-20:00");
    case "Evening PT":
      return sv.includes("16:00-20:00");
    case "Night":
      return sv.includes("20:00-00:00") || sv.includes("20:00-06:00") || sv.includes("20:15-06:15");
    default:
      return false;
  }
}

