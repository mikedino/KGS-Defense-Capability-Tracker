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
                    name: "synonyms",
                    title: "Synonyms",
                    description: "Additional abbreviations or terms used to find this contract",
                    type: Helper.SPCfgFieldType.Text
                },
                {
                    name: "capability",
                    title: "Capability",
                    description: "Capability this contract relationship belongs to",
                    type: Helper.SPCfgFieldType.Lookup,
                    listName: Strings.Sites.main.lists.Capabilities,
                    multi: true,
                    showField: "ID"
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
                    name: "isFlagged",
                    title: "Flag",
                    type: Helper.SPCfgFieldType.Boolean,
                    defaultValue: "0",
                    description: "Flag this contract to hide details from general users"
                },
                {
                    name: "clearance",
                    title: "Clearance Level",
                    type: Helper.SPCfgFieldType.Text,
                    description: "Choices from the contractClearance configuration type"
                },
                {
                    name: "customer",
                    title: "Customer",
                    type: Helper.SPCfgFieldType.Text,
                    description: "Choices from config list"
                },
                {
                    name: "contractType",
                    title: "Contract Type",
                    type: Helper.SPCfgFieldType.Text,
                    description: "Choices from config list"
                },
                {
                    name: "ogTitle",
                    title: "OG",
                    type: Helper.SPCfgFieldType.Text,
                    description: "Lookup: Jamis_Data_API => ContractEndPoint => OG"
                },
                {
                    name: "lobTitle",
                    title: "LOB",
                    type: Helper.SPCfgFieldType.Text,
                    description: "Derived from the selected OG lookup"
                },
                {
                    name: "startDate",
                    title: "Start Date",
                    type: Helper.SPCfgFieldType.Date,
                    format: SPTypes.DateFormat.DateOnly
                } as Helper.IFieldInfoDate,
                {
                    name: "endDate",
                    title: "End Date",
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
                    name: "contractValue",
                    title: "Contract Value",
                    type: Helper.SPCfgFieldType.Currency,
                    decimals: 2,
                    defaultValue: "0",
                    description: "Derived from the selected OG lookup"
                } as Helper.IFieldInfoCurrency,
                {
                    name: "infoLink",
                    title: "Contract Info Link/URL",
                    type: Helper.SPCfgFieldType.Text
                },
                {
                    name: "city",
                    title: "City",
                    type: Helper.SPCfgFieldType.Text
                },
                {
                    name: "state",
                    title: "State",
                    type: Helper.SPCfgFieldType.Text,
                    description: "Two-letter state abbreviation selected from the state configuration type"
                },
                {
                    name: "country",
                    title: "Country",
                    type: Helper.SPCfgFieldType.Text,
                    defaultValue: "US",
                    description: "ISO 3166-1 alpha-2 country code"
                },
                {
                    name: "location",
                    title: "Location",
                    type: Helper.SPCfgFieldType.Calculated,
                    resultType: SPTypes.FieldResultType.Text,
                    formula: '=IF(AND([city]&lt;&gt;"",[state]&lt;&gt;""),[city]&amp;", "&amp;[state],IF([city]&lt;&gt;"",[city],[state]))&amp;IF([country]&lt;&gt;""," ("&amp;[country]&amp;")","")',
                    fieldRefs: ["city", "state", "country"],
                    readOnly: true
                } as Helper.IFieldInfoCalculated
            ],
            ViewInformation: [
                {
                    ViewName: "All Items",
                    Default: true,
                    ViewQuery: '<OrderBy><FieldRef Name="contractId" Ascending="TRUE"/><FieldRef Name="Title" Ascending="TRUE"/></OrderBy>',
                    ViewFields: [
                        'LinkTitle',
                        'synonyms',
                        'capability',
                        'contractId',
                        'isFlagged',
                        'clearance',
                        'contractType',
                        'customerContractCode',
                        'ogTitle',
                        'lobTitle',
                        'customer',
                        'startDate',
                        'endDate',
                        'contractPm',
                        'partner',
                        'infoLink',
                        'location'
                    ]
                }
            ]
        },
        {
            ListInformation: {
                Title: Strings.Sites.main.lists.ContractCapabilitySummary,
                Description: "*DO NOT DELETE* List containing contract-specific capability summaries.",
                BaseTemplate: SPTypes.ListTemplateType.GenericList,
                OnQuickLaunch: false,
                Hidden: true
            },
            TitleFieldDisplayName: "Contract Capability",
            CustomFields: [
                {
                    name: "contract",
                    title: "Contract",
                    type: Helper.SPCfgFieldType.Lookup,
                    listName: Strings.Sites.main.lists.Contracts,
                    showField: "ID"
                } as Helper.IFieldInfoLookup,
                {
                    name: "capability",
                    title: "Capability",
                    type: Helper.SPCfgFieldType.Lookup,
                    listName: Strings.Sites.main.lists.Capabilities,
                    showField: "ID"
                } as Helper.IFieldInfoLookup,
                {
                    name: "summary",
                    title: "Capability Summary",
                    type: Helper.SPCfgFieldType.Note,
                    noteType: SPTypes.FieldNoteType.TextOnly,
                    numberOfLines: 6
                } as Helper.IFieldInfoNote,
                {
                    name: "poc",
                    title: "POC",
                    type: Helper.SPCfgFieldType.User
                } as Helper.IFieldInfoUser
            ],
            ViewInformation: [
                {
                    ViewName: "All Items",
                    Default: true,
                    ViewQuery: '<OrderBy><FieldRef Name="contract" /><FieldRef Name="capability" /></OrderBy>',
                    ViewFields: [
                        'LinkTitle',
                        'contract',
                        'capability',
                        'summary',
                        'poc'
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
                    name: "synonyms",
                    title: "Synonyms",
                    description: "Additional abbreviations or terms used to find this capability",
                    type: Helper.SPCfgFieldType.Text
                },
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
                    description: "Primary ecosystem or technology family the capability is built on, such as SharePoint, Power Platform, Azure, Microsoft 365, AWS, or on-premises.",
                    type: Helper.SPCfgFieldType.Text
                },
                {
                    name: "capabilityTypeTier1",
                    title: "Capability Type - Tier 1",
                    description: "Primary capability category. Choices are managed in the configuration list.",
                    type: Helper.SPCfgFieldType.Text
                },
                {
                    name: "capabilityTypeTier2",
                    title: "Capability Type - Tier 2",
                    description: "Secondary capability category. Choices are managed in the configuration list.",
                    type: Helper.SPCfgFieldType.Text
                },
                {
                    name: "hostingEnv",
                    title: "Hosting Environment",
                    description: "Where the capability runs or is hosted, such as SharePoint Online, Azure App Service, Dataverse, on-premises server, client device, vendor SaaS, or hybrid.",
                    type: Helper.SPCfgFieldType.Text
                },
                {
                    name: "connectivity",
                    title: "Connectivity",
                    description: "Customer network, access, or integration requirement needed to use the capability, such as internal, external, VPN, CAC/PIV, NIPR/SIPR, offline, or API integration.",
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
                    description: "Primary implementation language when custom code matters. Use None / Low-Code for Power Platform, BI, or configuration-focused solutions.",
                    type: Helper.SPCfgFieldType.Text
                },
                {
                    name: "backend",
                    title: "Backend",
                    description: "Primary data store, processing layer, or external system powering the capability, such as SharePoint Lists, Dataverse, SQL, APIs, Graph, file share, or no backend.",
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
                {
                    name: "pastPerformanceTagsJson",
                    title: "Past Performance Tags",
                    description: "JSON-backed selected Past Performance tags",
                    type: Helper.SPCfgFieldType.Note,
                    noteType: SPTypes.FieldNoteType.TextOnly,
                    numberOfLines: 6
                } as Helper.IFieldInfoNote,
                {
                    name: "proposalTagsJson",
                    title: "Proposal Tags",
                    description: "JSON-backed selected Proposal tags",
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
                        "synonyms",
                        "capStatus",
                        "capabilityTypeTier1",
                        "capabilityTypeTier2",
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
                    choices: ["Screenshot", "Technical", "Requirements", "Testing", "508 Compliance", "User Guides/User Manuals", "FAQ"],
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
        {
            ListInformation: {
                Title: Strings.Sites.main.lists.ContractDocuments,
                Description: "Library containing Defense Capabilities Tracker contract documentation.",
                BaseTemplate: SPTypes.ListTemplateType.DocumentLibrary,
                OnQuickLaunch: false,
                Hidden: true
            },
            CustomFields: [
                {
                    name: "contract",
                    title: "Contract",
                    type: Helper.SPCfgFieldType.Lookup,
                    listName: Strings.Sites.main.lists.Contracts,
                    showField: "ID"
                } as Helper.IFieldInfoLookup,
                {
                    name: "cdocType",
                    title: "Contract Document Type",
                    type: Helper.SPCfgFieldType.Text,
                    description: "Choices from config list"
                }
            ],
            ViewInformation: [
                {
                    ViewName: "All Items",
                    Default: true,
                    ViewQuery: '<OrderBy><FieldRef Name="contract" /><FieldRef Name="FileLeafRef" /></OrderBy>',
                    ViewFields: [
                        'contract', 'DocIcon', 'LinkFilename', 'cdocType', 'Modified', 'Editor', 'FileSizeDisplay'
                    ]
                }
            ]
        },
    ]
})
