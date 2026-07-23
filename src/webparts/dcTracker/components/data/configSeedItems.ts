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
  ...seed("codingLanguage", ["C#", "Java", "TypeScript"]),
  ...seed("compliance", ["FedRAMP", "FISMA", "N/A"]),
  ...seed("connectivity", ["Internal", "External", "Hybrid"]),
  ...seed("customer", ["DOW", "DOS", "USN"]),
  ...seed("partner", ["LMI", "Guidehouse"]),
  ...seed("hostingEnvironment", ["SharePoint Online", "Azure", "On-Premises"]),
  ...seed("platform", ["SharePoint", "Web", "Desktop"])
];
