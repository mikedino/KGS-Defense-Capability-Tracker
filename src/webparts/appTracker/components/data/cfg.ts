import { Helper, SPTypes } from "gd-sprest-bs";
import Strings from "../../strings";

/** SharePoint assets for the current site - installed on first run */
export const Configuration = Helper.SPConfig({
    ListCfg: [
        {
            ListInformation: {
                Title: Strings.Lists.Configuration,
                Description: "*DO NOT DELETE* Library containing configuration for the Application Tracker",
                BaseTemplate: SPTypes.ListTemplateType.GenericList,
                OnQuickLaunch: false,
                Hidden: true
            },
            TitleFieldDisplayName: "Value",
            CustomFields: [
                {
                    name: "isFor",
                    title: "Is For",
                    type: Helper.SPCfgFieldType.Text
                },
                {
                    name: "isForDisplayName",
                    title: "Is For Display Name",
                    description: "Friendly name of the configuration type",
                    type: Helper.SPCfgFieldType.Text
                },
                {
                    name: "isActive",
                    title: "Is Active?",
                    type: Helper.SPCfgFieldType.Boolean,
                    defaultValue: "1",
                } as Helper.IFieldInfoChoice,
                {
                    name: "infoText",
                    title: "Informational Text",
                    description: "Describe the value or purpose (if necessary)",
                    type: Helper.SPCfgFieldType.Text
                }
            ],
            ViewInformation: [
                {
                    ViewName: "All Items",
                    Default: true,
                    ViewQuery: '<OrderBy><FieldRef Name="isFor" Ascending="TRUE"/></OrderBy>',
                    ViewFields: ['LinkTitle', 'isFor', 'isForDisplayName', 'infoText', 'isActive']
                }
            ]
        },
        {
            ListInformation: {
                Title: Strings.Lists.Contracts,
                Description: "*DO NOT DELETE* Library containing contract info for the Application Tracker",
                BaseTemplate: SPTypes.ListTemplateType.GenericList,
                OnQuickLaunch: false,
                Hidden: true
            },
            TitleFieldDisplayName: "Contract Name/Title",
            CustomFields: [
                {
                    name: "contractTeam",
                    title: "Contract Lead",
                    type: Helper.SPCfgFieldType.Text,
                    description: "Choices from config list"
                } as Helper.IFieldInfoChoice,
                {
                    name: "contractValue",
                    title: "Total Contract Value",
                    type: Helper.SPCfgFieldType.Currency,
                    decimals: 2,
                    defaultValue: "0"
                } as Helper.IFieldInfoCurrency,
                {
                    name: "start",
                    title: "Start Date",
                    type: Helper.SPCfgFieldType.Date,
                    format: SPTypes.DateFormat.DateOnly
                } as Helper.IFieldInfoDate,
                {
                    name: "end",
                    title: "End Date",
                    type: Helper.SPCfgFieldType.Date,
                    format: SPTypes.DateFormat.DateOnly
                } as Helper.IFieldInfoDate,
                {
                    name: "cor",
                    title: "COR",
                    type: Helper.SPCfgFieldType.User
                } as Helper.IFieldInfoUser,
                {
                    name: "acor",
                    title: "ACOR",
                    type: Helper.SPCfgFieldType.User
                } as Helper.IFieldInfoUser,
                {
                    name: "ko",
                    title: "Contracting Officer",
                    type: Helper.SPCfgFieldType.User
                } as Helper.IFieldInfoUser,
                {
                    name: "primaryPoc",
                    title: "Contract PM",
                    type: Helper.SPCfgFieldType.User
                } as Helper.IFieldInfoUser,
                {
                    name: "secondaryPoc",
                    title: "Assistant Contract PM",
                    type: Helper.SPCfgFieldType.User
                } as Helper.IFieldInfoUser
            ],
            ViewInformation: [
                {
                    ViewName: "All Items",
                    Default: true,
                    ViewQuery: '<OrderBy><FieldRef Name="Title" Ascending="TRUE"/></OrderBy>',
                    ViewFields: ['LinkTitle', 'contractValue', 'cor', 'acor', 'ko', 'primaryPoc', 'secondaryPoc']
                }
            ]
        },
        {
            ListInformation: {
                Title: Strings.Lists.Apps,
                Description: "*DO NOT DELETE* List to track KGS solutions.",
                BaseTemplate: SPTypes.ListTemplateType.GenericList,
                OnQuickLaunch: false,
                Hidden: true
            },
            TitleFieldRequired: true,
            TitleFieldDisplayName: "Application Title",
            CustomFields: [
                {
                    name: "description",
                    title: "Description",
                    type: Helper.SPCfgFieldType.Note,
                    noteType: SPTypes.FieldNoteType.TextOnly
                } as Helper.IFieldInfoNote,
                {
                    name: "highlights",
                    title: "Highlights &amp; Features",
                    description: "Working Title",
                    type: Helper.SPCfgFieldType.Note,
                    noteType: SPTypes.FieldNoteType.RichText
                } as Helper.IFieldInfoNote,
                {
                    name: "appUrl",
                    title: "App URL",
                    type: Helper.SPCfgFieldType.Text
                },
                {
                    name: "relatedInfo",
                    title: "Core Capabilities & Relevant Software",
                    type: Helper.SPCfgFieldType.Note,
                    noteType: SPTypes.FieldNoteType.RichText
                } as Helper.IFieldInfoNote,
                {
                    name: "primaryPoc",
                    title: "Primary POC",
                    type: Helper.SPCfgFieldType.User
                } as Helper.IFieldInfoUser,
                {
                    name: "stakeholders",
                    title: "Stakeholders",
                    type: Helper.SPCfgFieldType.User,
                    multi: true
                } as Helper.IFieldInfoUser,
                {
                    name: "systemOwner",
                    title: "System Owner",
                    type: Helper.SPCfgFieldType.User
                } as Helper.IFieldInfoUser,
                {
                    name: "managingGroup",
                    title: "Managing Group",
                    description: "Choices from config list",
                    type: Helper.SPCfgFieldType.Text
                },
                {
                    name: "supportTeam",
                    title: "Contractor Name",
                    description: "Choices from config list",
                    type: Helper.SPCfgFieldType.Text
                },
                {
                    name: "licenseReqd",
                    title: "License Required?",
                    type: Helper.SPCfgFieldType.Choice,
                    choices: ["Yes", "No", "Depends"],
                    defaultValue: "No",
                    multi: false
                } as Helper.IFieldInfoChoice,
                {
                    name: "connectivity",
                    title: "Connectivity",
                    description: "Choices from config list",
                    type: Helper.SPCfgFieldType.Text
                },
                {
                    name: "integrationsInput",
                    title: "Integrations (Input)",
                    type: Helper.SPCfgFieldType.Note
                } as Helper.IFieldInfoNote,
                {
                    name: "integrationsOutput",
                    title: "Integrations (Output)",
                    type: Helper.SPCfgFieldType.Note
                } as Helper.IFieldInfoNote,
                {
                    name: "synonyms",
                    title: "Synomyms",
                    type: Helper.SPCfgFieldType.Text
                },
                {
                    name: "platform",
                    title: "Platform",
                    description: "Choices from config list",
                    type: Helper.SPCfgFieldType.Text
                },
                {
                    name: "environment",
                    title: "Hosting Environment",
                    description: "Choices from config list",
                    type: Helper.SPCfgFieldType.Text
                },
                {
                    name: "contract",
                    title: "Contract",
                    description: "Choices from lookup list",
                    type: Helper.SPCfgFieldType.Lookup,
                    listName: Strings.Lists.Contracts,
                    showField: "Title"
                } as Helper.IFieldInfoLookup,
                {
                    name: "appStatus",
                    title: "App Status",
                    description: "Choices from config list",
                    type: Helper.SPCfgFieldType.Text
                },
                {
                    name: "appLaunchDate",
                    title: "App Launch Date",
                    type: Helper.SPCfgFieldType.Date,
                    format: SPTypes.DateFormat.DateOnly
                } as Helper.IFieldInfoDate,
                {
                    name: "userCount",
                    title: "User Count",
                    type: Helper.SPCfgFieldType.Number,
                    decimals: 0
                } as Helper.IFieldInfoNumber
            ],
            ViewInformation: [
                {
                    ViewName: "All Items",
                    Default: true,
                    ViewQuery: '<OrderBy><FieldRef Name="Modified" Ascending="FALSE" /></OrderBy>',
                    ViewFields: [
                        "LinkTitle",
                        "platform",
                        "connectivity",
                        "primaryPoc",
                        "stakeholders",
                        "managingGroup",
                        "appStatus",
                        "appLaunchDate",
                        "contract"
                    ]
                }
            ]
        },
        {
            ListInformation: {
                Title: Strings.Lists.Documents,
                Description: "Library containing Application Tracker documentation.",
                BaseTemplate: SPTypes.ListTemplateType.DocumentLibrary,
                OnQuickLaunch: false,
                Hidden: true
            },
            CustomFields: [
                {
                    name: "application",
                    title: "Application",
                    type: Helper.SPCfgFieldType.Lookup,
                    listName: Strings.Lists.Apps,
                    showField: "Title"
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
                    ViewQuery: '<OrderBy><FieldRef Name="application" /><FieldRef Name="FileLeafRef" /></OrderBy>',
                    ViewFields: [
                        'application', 'DocIcon', 'LinkFilename', 'docType', 'Modified', 'Editor', 'FileSizeDisplay'
                    ]
                }
            ]
        },
    ]
})
