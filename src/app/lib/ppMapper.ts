
export function matchesSecondaryPP(secondaryText: any, targetPP: any): boolean {
  const t = String(secondaryText || "").trim().toLowerCase();
  const pp = String(targetPP || "").trim().toLowerCase();
  if (!t || !pp) return false;
  if (t === pp) return true;

  const rules: Record<string, string[]> = {
    bagging: ["primary sort", "secondary sort", "bag", "sortation"],
    zbs: ["fluid", "zbs"],
    inbound: ["receive", "receiving", "inbound", "stow"],
    outbound: ["dispatch", "load", "outbound"],
    ps: ["problem solve", "ps"],
    "sort slide": ["sort slide", "slide"],
    bigbox: ["big box", "bigbox", "non con", "non-con"],
    "bag pulling": ["bag pull", "bp"],
  };

  const keywords = rules[pp];
  if (!keywords) return false;
  return keywords.some((kw) => t.includes(kw));
}

export function mapPrimaryPPToDeployPP(alcPrimaryPP: string): string {
  const pp = (alcPrimaryPP || "").trim().toLowerCase();

  if (pp === "inbound" || pp.includes("receive") || pp.includes("stow") || pp.includes("unload"))
    return "Inbound";
  if (pp === "outbound" || pp.includes("dispatch") || pp.includes("loading"))
    return "Outbound";
  if (pp === "ps" || pp.includes("problem solve"))
    return "PS";
  if (pp === "zbs" || pp.includes("fluid"))
    return "ZBS";
  if (pp === "sort slide" || pp.includes("slide") || pp.includes("diverter"))
    return "Sort Slide";
  if (pp === "bagging" || pp.includes("bag sortation") || pp.includes("secondary sort") || pp.includes("primary sort"))
    return "Bagging";
  if (pp === "bigbox" || pp.includes("big box") || pp.includes("non con") || pp.includes("non-con"))
    return "BigBox";
  if (pp === "bag pulling" || pp.includes("bag pull"))
    return "Bag Pulling";

  return alcPrimaryPP;
}

export function getPPDescription(assignedPP: string): string {
  const pp = (assignedPP || "").trim().toLowerCase();
  const descriptions: Record<string, string> = {
    "bag pulling": "bag pulling | bag pulling sorting.lets rock!",
    bagging: "bag sortation | secondary sorting.lets rock!",
    inbound: "inbound | unload.lets rock!",
    "sort slide": "inbound | sort slide diverter.lets rock!",
    outbound: "outbound | loading container.lets rock!",
    ps: "ic-qa master session | problem solve.lets rock!",
    zbs: "fluid sortation | fluid sorting.lets rock!",
    bigbox: "non-con processing | big box sortation and movement.lets rock!",
  };
  return descriptions[pp] || `${pp} | ${pp}.lets rock!`;
}

export const PP_LIST = ["Inbound", "Outbound", "PS", "ZBS", "Sort Slide", "Bagging", "BigBox", "Bag Pulling"];

