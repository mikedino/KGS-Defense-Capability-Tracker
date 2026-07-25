import { IConfigItem } from "../common/props";

const seed = (configType: IConfigItem["configType"], values: string[]): IConfigItem[] =>
  values.map((value, index) => ({
    Id: 0,
    Title: value,
    configType,
    configValue: value,
    sortOrder: index + 1,
    isActive: true,
    infoText: ""
  }));

export const configSeedItems: IConfigItem[] = [
  ...seed("backend", ["SharePoint", "SQL", "Oracle"]),
  ...seed("capabilityStatus", ["Active", "In Dev", "Pending"]),
  ...seed("codingLanguage", ["C#", "Java", "TypeScript", "Power Fx"]),
  ...seed("compliance", ["FedRAMP", "FISMA", "N/A"]),
  ...seed("connectivity", ["Internal", "External", "Hybrid"]),
  ...seed("customer", [
    "Air Force",
    "Army",
    "ATF",
    "BEP",
    "CMS",
    "DEA",
    "DHA",
    "DHS",
    "DoD",
    "DoE",
    "DoEd",
    "DOI",
    "DoS",
    "FBI",
    "FDA",
    "FEMA",
    "GSA",
    "HHS",
    "IHS",
    "Marines",
    "NARA",
    "NIH",
    "NOAA",
    "SSA",
    "USAID",
    "USCIS",
    "USDA",
    "DOW",
    "USN"
  ]),
  ...seed("partner", ["Accenture Federal Services LLC", "DecisionPoint Corporation", "Imagineeer, LLC", "LMI", "Spry Methods"]),
  ...seed("hostingEnvironment", ["SharePoint Online", "Azure", "On-Premises"]),
  ...seed("platform", ["SharePoint", "Web", "Desktop"])
];
