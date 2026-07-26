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
  ...seed("backend", ["SharePoint Lists", "Dataverse", "SQL Server", "Azure SQL", "Oracle", "REST API", "Graph API", "Cosmos DB", "File Share", "No Backend", "External System"]),
  ...seed("capabilityStatus", ["Active", "In Dev", "Pending"]),
  ...seed("codingLanguage", ["TypeScript", "JavaScript", "C#", "Python", "Java", "Power Fx", "SQL", "DAX", "HTML/CSS", "PowerShell", "None / Low-Code"]),
  ...seed("compliance", ["FedRAMP", "FISMA", "N/A"]),
  ...seed("connectivity", ["Internal", "External", "Hybrid", "Internet Available", "VPN Required", "CAC/PIV Required", "NIPR", "SIPR", "Offline Capable", "API Integration Required"]),
  ...seed("customer", ["Air Force", "Army", "ATF", "BEP", "CMS", "DEA", "DHA", "DHS", "DoD", "DoE", "DoEd", "DOI", "DoS", "FBI", "FDA",
    "FEMA", "GSA", "HHS", "IHS", "Marines", "NARA", "NIH", "NOAA", "SSA", "USAID", "USCIS", "USDA", "DOW", "USN"
  ]),
  ...seed("partner", ["Accenture Federal Services LLC", "DecisionPoint Corporation", "Imagineeer, LLC", "LMI", "Spry Methods"]),
  ...seed("hostingEnvironment", ["SharePoint Online", "Azure App Service", "Azure Function", "Dataverse", "Power Platform", "On-Premises Server", "Client Device", "SaaS Vendor Hosted", "Hybrid"]),
  ...seed("platform", ["SharePoint", "Power Platform", "Azure", "Microsoft 365", "Standalone", "AWS", "On-Premises", "Hybrid Cloud"]),
  ...seed("solutionType", ["SPFx", "Power App", "Power Automate Flow", "Power BI Report/Dashboard", "Web App", "AI Solution", "Mobile App", "Desktop App", "SharePoint Site/Solution", "Data Integration", "API/Service", "Automation", "Reporting/Analytics", "Document/Knowledge Management", "Workflow/Application Modernization"]),
  ...seed("documentType", ["Screenshot", "Technical", "Requirements", "Testing", "508 Compliance", "User Guides/User Manuals", "FAQ"])
];
