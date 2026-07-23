import { Helper, SPTypes } from "gd-sprest-bs";
import Strings from "../common/strings";

/** SharePoint assets for the current site - installed on first run */
export const Configuration = Helper.SPConfig({
    ListCfg: [
        {
            ListInformation: {
                Title: Strings.Sites.main.lists.Configuration,
                Description: "*DO NOT DELETE* Library containing configuration for the Defense Capabilities Tracker",
                BaseTemplate: SPTypes.ListTemplateType.GenericList,
                OnQuickLaunch: false,
                Hidden: true
            },
            TitleFieldDisplayName: "Display Text",
            CustomFields: [
                {
                    name: "configType",
                    title: "Config Type",
                    description: "Configuration category",
                    type: Helper.SPCfgFieldType.Text
                },
                {
                    name: "configValue",
                    title: "Config Value",
                    description: "Compared value/saved value",
                    type: Helper.SPCfgFieldType.Text
                },
                {
                    name: "sortOrder",
                    title: "Sort Order",
                    type: Helper.SPCfgFieldType.Number
                } as Helper.IFieldInfoNumber,
                {
                    name: "infoText",
                    title: "Informational Text",
                    description: "Describe the value or purpose (if necessary)",
                    type: Helper.SPCfgFieldType.Text
                },
                {
                    name: "isActive",
                    title: "Is Active?",
                    type: Helper.SPCfgFieldType.Boolean,
                    defaultValue: "1",
                }
            ],
            ViewInformation: [
                {
                    ViewName: "All Items",
                    Default: true,
                    ViewQuery: '<OrderBy><FieldRef Name="configType" Ascending="TRUE"/><FieldRef Name="sortOrder" Ascending="TRUE"/><FieldRef Name="Title" Ascending="TRUE"/></OrderBy>',
                    ViewFields: ['LinkTitle', 'configType', 'configValue', 'sortOrder', 'infoText', 'isActive']
                }
            ]
        },
        {
            ListInformation: {
                Title: Strings.Sites.main.lists.Contracts,
                Description: "*DO NOT DELETE* Library containing contract info for the Defense Capabilities Tracker",
                BaseTemplate: SPTypes.ListTemplateType.GenericList,
                OnQuickLaunch: false,
                Hidden: true
            },
            TitleFieldDisplayName: "Contract Title",
            CustomFields: [
                {
                    name: "capability",
                    title: "Capability",
                    description: "Capability this contract relationship belongs to",
                    type: Helper.SPCfgFieldType.Lookup,
                    listName: Strings.Sites.main.lists.Capabilities,
                    showField: "Title"
                } as Helper.IFieldInfoLookup,
                {
                    name: "contractId",
                    title: "Contract ID",
                    type: Helper.SPCfgFieldType.Text,
                    description: "Lookup: Jamis_Data_API => ContractEndPoint => Contract ID"
                },
                {
                    name: "customerContractCode",
                    title: "Customer Contract Code",
                    type: Helper.SPCfgFieldType.Text,
                    description: "Lookup: Jamis_Data_API => ContractEndPoint => Customer Contract Code"
                },
                {
                    name: "customer",
                    title: "Customer",
                    type: Helper.SPCfgFieldType.Text,
                    description: "Choices from config list"
                },
                {
                    name: "startDate",
                    title: "Capability Start Date",
                    type: Helper.SPCfgFieldType.Date,
                    format: SPTypes.DateFormat.DateOnly
                } as Helper.IFieldInfoDate,
                {
                    name: "endDate",
                    title: "Capability End Date",
                    type: Helper.SPCfgFieldType.Date,
                    format: SPTypes.DateFormat.DateOnly
                } as Helper.IFieldInfoDate,
                {
                    name: "contractPm",
                    title: "KGS Contract Project Manager",
                    type: Helper.SPCfgFieldType.User
                } as Helper.IFieldInfoUser,
                {
                    name: "partner",
                    title: "Relevant Partner Tag",
                    type: Helper.SPCfgFieldType.Text,
                    description: "Choices from config list"
                },
                {
                    name: "infoLink",
                    title: "Contract Info Link/URL",
                    type: Helper.SPCfgFieldType.Text
                }
            ],
            ViewInformation: [
                {
                    ViewName: "All Items",
                    Default: true,
                    ViewQuery: '<OrderBy><FieldRef Name="contractId" Ascending="TRUE"/><FieldRef Name="Title" Ascending="TRUE"/></OrderBy>',
                    ViewFields: [
                        'LinkTitle',
                        'capability',
                        'contractId',
                        'customerContractCode',
                        'customer',
                        'startDate',
                        'endDate',
                        'contractPm',
                        'partner',
                        'infoLink'
                    ]
                }
            ]
        },
        {
            ListInformation: {
                Title: Strings.Sites.main.lists.Capabilities,
                Description: "*DO NOT DELETE* List to track KGS solutions.",
                BaseTemplate: SPTypes.ListTemplateType.GenericList,
                OnQuickLaunch: false,
                Hidden: true
            },
            TitleFieldRequired: true,
            TitleFieldDisplayName: "Capability Title",
            CustomFields: [
                {
                    name: "description",
                    title: "Capability Description",
                    type: Helper.SPCfgFieldType.Note,
                    noteType: SPTypes.FieldNoteType.TextOnly,
                    numberOfLines: 6
                } as Helper.IFieldInfoNote,
                {
                    name: "capabilities",
                    title: "Technical Capabilities",
                    type: Helper.SPCfgFieldType.Note,
                    noteType: SPTypes.FieldNoteType.TextOnly,
                    numberOfLines: 6
                } as Helper.IFieldInfoNote,
                {
                    name: "link",
                    title: "Link/URL",
                    type: Helper.SPCfgFieldType.Text
                },
                {
                    name: "capStatus",
                    title: "Capability Status",
                    description: "Choices from config list",
                    type: Helper.SPCfgFieldType.Text
                },
                {
                    name: "notes",
                    title: "Additional Notes",
                    type: Helper.SPCfgFieldType.Note,
                    noteType: SPTypes.FieldNoteType.TextOnly,
                    numberOfLines: 6
                } as Helper.IFieldInfoNote,
                {
                    name: "platform",
                    title: "Platform",
                    description: "Choices from config list",
                    type: Helper.SPCfgFieldType.Text
                },
                {
                    name: "hostingEnv",
                    title: "Hosting Environment",
                    description: "Choices from config list",
                    type: Helper.SPCfgFieldType.Text
                },
                {
                    name: "connectivity",
                    title: "Connectivity",
                    description: "Choices from config list",
                    type: Helper.SPCfgFieldType.Text
                },
                {
                    name: "compliance",
                    title: "Compliance",
                    description: "Choices from config list",
                    type: Helper.SPCfgFieldType.Text
                },
                {
                    name: "licenseReqd",
                    title: "License Required?",
                    type: Helper.SPCfgFieldType.Choice,
                    choices: ["Yes", "No"],
                    defaultValue: "No",
                    multi: false
                } as Helper.IFieldInfoChoice,
                {
                    name: "licenseReqmts",
                    title: "Licensing Requirements",
                    type: Helper.SPCfgFieldType.Note,
                    noteType: SPTypes.FieldNoteType.TextOnly,
                    numberOfLines: 6
                } as Helper.IFieldInfoNote,
                {
                    name: "extensibility",
                    title: "APIs/Extensibility",
                    type: Helper.SPCfgFieldType.Note,
                    noteType: SPTypes.FieldNoteType.TextOnly,
                    numberOfLines: 6
                } as Helper.IFieldInfoNote,
                {
                    name: "serverReqmts",
                    title: "Server Requirements",
                    type: Helper.SPCfgFieldType.Text
                },
                {
                    name: "codeLanguage",
                    title: "Coding Language",
                    description: "Choices from config list",
                    type: Helper.SPCfgFieldType.Text
                },
                {
                    name: "backend",
                    title: "Backend",
                    description: "Choices from config list",
                    type: Helper.SPCfgFieldType.Text
                },
                {
                    name: "primaryPoc",
                    title: "Capability Primary POC",
                    type: Helper.SPCfgFieldType.User
                } as Helper.IFieldInfoUser,
                {
                    name: "stakeholders",
                    title: "KGS Stakeholders",
                    type: Helper.SPCfgFieldType.User,
                    multi: true
                } as Helper.IFieldInfoUser,
                {
                    name: "oppNetTagsJson",
                    title: "OppNet Tags",
                    description: "JSON-backed selected OppNet tags",
                    type: Helper.SPCfgFieldType.Note,
                    noteType: SPTypes.FieldNoteType.TextOnly,
                    numberOfLines: 6
                } as Helper.IFieldInfoNote,
            ],
            ViewInformation: [
                {
                    ViewName: "All Items",
                    Default: true,
                    ViewQuery: '<OrderBy><FieldRef Name="Modified" Ascending="FALSE" /></OrderBy>',
                    ViewFields: [
                        "LinkTitle",
                        "capStatus",
                        "platform",
                        "hostingEnv",
                        "connectivity",
                        "compliance",
                        "licenseReqd",
                        "codeLanguage",
                        "backend",
                        "primaryPoc",
                        "stakeholders"
                    ]
                }
            ]
        },
        {
            ListInformation: {
                Title: Strings.Sites.main.lists.Documents,
                Description: "Library containing Defense Capabilities Tracker documentation.",
                BaseTemplate: SPTypes.ListTemplateType.DocumentLibrary,
                OnQuickLaunch: false,
                Hidden: true
            },
            CustomFields: [
                {
                    name: "capability",
                    title: "Capability",
                    type: Helper.SPCfgFieldType.Lookup,
                    listName: Strings.Sites.main.lists.Capabilities,
                    showField: "ID"
                } as Helper.IFieldInfoLookup,
                {
                    name: "docType",
                    title: "Document Type",
                    type: Helper.SPCfgFieldType.Choice,
                    choices: ["Screenshot", "Other"],
                } as Helper.IFieldInfoChoice
            ],
            ViewInformation: [
                {
                    ViewName: "All Items",
                    Default: true,
                    ViewQuery: '<OrderBy><FieldRef Name="capability" /><FieldRef Name="FileLeafRef" /></OrderBy>',
                    ViewFields: [
                        'capability', 'DocIcon', 'LinkFilename', 'docType', 'Modified', 'Editor', 'FileSizeDisplay'
                    ]
                }
            ]
        },
    ]
})
