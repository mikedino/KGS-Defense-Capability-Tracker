import { Web } from "gd-sprest-bs";
import { WebPartContext } from '@microsoft/sp-webpart-base';
import Strings, { setContext } from "../common/strings";
import { ConfigType, ICapabilityItem, ICMSContractItem, IConfigItem, IContractCapabilitySummaryItem, IContractDocumentItem, IContractEndPointItem, IContractItem, IContractSourceItem, IDocumentItem, IOgItem, IOpportunityItem, IPastPerformanceItem, IProposalItem } from "../common/props";
import { formatError } from "../common/utils";
import { ConfigService } from "../services/ConfigService";
import { parseJsonTagField } from "../common/tagUtils";

export interface IConfigOption {
    key: string;
    text: string;
}

export class DataSource {
    //prevent this from being initialized twice
    static initialized: boolean = false;

    // Initializes the capability
    public static init(override: boolean, context?: WebPartContext): PromiseLike<void> {

        // verify the page context exists
        if (context) {
            // Set the context
            setContext(context);
        }

        if (!this.initialized || override) { //ensure this was not already initialized and not manually being refreshed

            return this.getConfig().then(async () => {
                const addedConfigSeedItems = await ConfigService.ensureSeedItems();
                if (addedConfigSeedItems > 0) {
                    console.log(`[${Strings.ProjectName}] Seeded Config items. Added: ${addedConfigSeedItems}.`);
                    await this.getConfig();
                }

                await Promise.all([
                    this.getContracts(),
                    this.getCapabilities(),
                    this.getContractCapabilitySummaries(),
                    this.getOGs(),
                    this._jamisContractsLoaded ? Promise.resolve(this._jamisContracts) : this.getJamisContracts(),
                    this._cmsContractsLoaded ? Promise.resolve(this._cmsContracts) : this.getCmsContracts()
                ]);

                // Rebuild the normalized search collection after regular refreshes without re-fetching cached source lists.
                this.buildContractSources();
            }).then(() => {
                this.initialized = true;
            }).catch((error) => {
                const errorMessage = formatError(error);
                console.error(errorMessage);
                throw error;
            });

        } else {
            console.log(Strings.ProjectName, "tried to init the datasource again");
            return Promise.resolve(); //already initialized once
        }
    }

    // Store raw config for debugging and admin screen
    private static _config: IConfigItem[] = [];
    static get Config(): IConfigItem[] { return this._config; }

    // Grouped items
    private static _configByType = new Map<string, IConfigItem[]>();
    static get ConfigByType(): ReadonlyMap<string, IConfigItem[]> { return this._configByType; }

    // Grouped values (configValue)
    private static _configValuesByType = new Map<string, string[]>();
    static get ConfigValuesByType(): ReadonlyMap<string, string[]> { return this._configValuesByType; }

    // Convenience helper
    static getConfigValues(configType: ConfigType | string): string[] {
        return this._configValuesByType.get(configType) ?? [];
    }

    static getConfigOptions(configType: ConfigType | string): IConfigOption[] {
        return (this._configByType.get(configType) ?? [])
            .filter(item => item.isActive !== false)
            .sort((a, b) => (a.sortOrder ?? 9999) - (b.sortOrder ?? 9999) || (a.Title ?? "").localeCompare(b.Title ?? ""))
            .map(item => ({
                key: item.configValue,
                text: item.Title || item.configValue
            }));
    }

    // get/set config
    static configSelect = ["Id", "Title", "configType", "configValue", "sortOrder", "isActive", "infoText"];
    static getConfig(): Promise<IConfigItem[]> {
        return new Promise<IConfigItem[]>((resolve, reject) => {

            // clear the items
            this._config = [];
            this._configByType = new Map<string, IConfigItem[]>();
            this._configValuesByType = new Map<string, string[]>();

            // load the data
            Web().Lists(Strings.Sites.main.lists.Configuration).Items().query({
                GetAllItems: true,
                OrderBy: ["configType", "sortOrder", "Title"],
                Select: this.configSelect,
                Filter: "isActive eq 1"
            }).execute(
                (items) => {
                    if (items?.results?.length) {
                        const results = items.results as unknown as IConfigItem[];
                        this._config = results;

                        // Group + build value arrays
                        for (const item of results) {
                            const typeKey = (item.configType ?? "").trim();
                            const value = (item.configValue ?? "").trim();
                            if (!typeKey || !value || item.isActive === false) continue;

                            const list = this._configByType.get(typeKey) ?? [];
                            list.push(item);
                            this._configByType.set(typeKey, list);

                            const values = this._configValuesByType.get(typeKey) ?? [];
                            values.push(value);
                            this._configValuesByType.set(typeKey, values);
                        }

                        // Optional: dedupe + sort values per key
                        for (const [k, vals] of this._configValuesByType.entries()) {
                            const unique = Array.from(new Set(vals.map(v => v.trim()).filter(Boolean)));
                            unique.sort((a, b) => a.localeCompare(b));
                            this._configValuesByType.set(k, unique);
                        }

                        resolve(results);
                    } else {
                        //none found - resolve with empty array
                        resolve([])
                    }
                },
                (error) => reject(new Error(`Error fetching Config items: ${formatError(error)}`))
            );
        });
    }

    // Load the Contracts
    static contractQuerySelect: string[] = [
        "Id", "Title", "capability/Id", "capability/Title", "contractId", "customerContractCode", "customer",
        "startDate", "endDate", "partner", "infoLink", "ogTitle", "lobTitle", "contractType","contractValue",
        "contractPm/Id", "contractPm/Title", "contractPm/EMail", "contractPm/JobTitle", "contractPm/Department"
    ];
    static contractQueryExpand: string[] = ["capability", "contractPm"];
    private static _contracts: IContractItem[] = [];
    static get Contracts(): IContractItem[] { return this._contracts; }
    // Refresh only DCTContracts so local duplicate checks stay current without reloading source contract systems.
    static refreshContracts(): Promise<IContractItem[]> {
        return this.getContracts();
    }

    private static getContracts(): Promise<IContractItem[]> {
        return new Promise<IContractItem[]>((resolve, reject) => {

            // clear the items
            this._contracts = [];

            // load the data
            Web().Lists(Strings.Sites.main.lists.Contracts).Items().query({
                GetAllItems: true,
                Select: this.contractQuerySelect,
                OrderBy: ["Title"],
                Expand: this.contractQueryExpand
            }).execute(
                // Success
                (items) => {
                    if (items?.results?.length) {
                        this._contracts = items.results as unknown as IContractItem[];
                        resolve(this._contracts);
                    } else {
                        //none found - resolve with empty array
                        resolve([])
                    }
                },
                // Error
                error => {
                    reject(new Error(`Error fetching Contracts: ${formatError(error)}`));
                }
            )

        });
    }

    // Load contract-capability relationship summaries.
    static contractCapabilitySummaryQuerySelect: string[] = [
        "Id", "Title", "contract/Id", "contract/Title", "capability/Id", "capability/Title",
        "summary", "poc/Id", "poc/Title", "poc/EMail", "poc/JobTitle", "poc/Department"
    ];
    static contractCapabilitySummaryQueryExpand: string[] = ["contract", "capability", "poc"];
    private static _contractCapabilitySummaries: IContractCapabilitySummaryItem[] = [];
    static get ContractCapabilitySummaries(): IContractCapabilitySummaryItem[] { return this._contractCapabilitySummaries; }
    static refreshContractCapabilitySummaries(): Promise<IContractCapabilitySummaryItem[]> {
        return this.getContractCapabilitySummaries();
    }

    static getContractCapabilitySummary(contractId: number, capabilityId: number): IContractCapabilitySummaryItem | undefined {
        return this._contractCapabilitySummaries.find((item) =>
            item.contract?.Id === contractId && item.capability?.Id === capabilityId
        );
    }

    private static getContractCapabilitySummaries(): Promise<IContractCapabilitySummaryItem[]> {
        return new Promise<IContractCapabilitySummaryItem[]>((resolve, reject) => {
            this._contractCapabilitySummaries = [];

            Web().Lists(Strings.Sites.main.lists.ContractCapabilitySummary).Items().query({
                GetAllItems: true,
                Select: this.contractCapabilitySummaryQuerySelect,
                OrderBy: ["Title"],
                Expand: this.contractCapabilitySummaryQueryExpand
            }).execute(
                (items) => {
                    this._contractCapabilitySummaries = (items?.results ?? []) as unknown as IContractCapabilitySummaryItem[];
                    resolve(this._contractCapabilitySummaries);
                },
                (error) => reject(new Error(`Error fetching Contract Capability Summaries: ${formatError(error)}`))
            );
        });
    }

    // Load the Capabilities
    static capabilityQuerySelect: string[] = [
        "Id", "Title", "description", "capabilities", "link", "capStatus", "notes",
        "solutionType", "platform", "hostingEnv", "connectivity", "compliance", "licenseReqd",
        "licenseReqmts", "extensibility", "serverReqmts", "codeLanguage", "backend",
        "oppNetTagsJson", "pastPerformanceTagsJson", "proposalTagsJson", "Modified",
        "primaryPoc/Id", "primaryPoc/Title", "primaryPoc/EMail",
        "stakeholders/Id", "stakeholders/Title", "stakeholders/EMail",
        "Author/Title", "Author/EMail", "Author/Id"
    ]
    static capabilityQueryExpand: string[] = ["primaryPoc", "stakeholders", "Author"];
    private static _capabilities: ICapabilityItem[] = [];
    static get Capabilities(): ICapabilityItem[] { return this._capabilities; }
    private static getCapabilities(): Promise<ICapabilityItem[]> {
        return new Promise<ICapabilityItem[]>((resolve, reject) => {

            // clear the items
            this._capabilities = [];

            // load the data
            Web().Lists(Strings.Sites.main.lists.Capabilities).Items().query({
                GetAllItems: true,
                OrderBy: ["Modified desc"],
                Select: this.capabilityQuerySelect,
                Expand: this.capabilityQueryExpand
            }).execute(
                // Success
                (items) => {
                    if (items?.results?.length) {
                        this._capabilities = (items.results as unknown as ICapabilityItem[]).map((item) => ({
                            ...item,
                            oppNetTags: parseJsonTagField(item.oppNetTagsJson),
                            pastPerformanceTags: parseJsonTagField(item.pastPerformanceTagsJson),
                            proposalTags: parseJsonTagField(item.proposalTagsJson)
                        }));
                        resolve(this._capabilities);
                    } else {
                        //none found - resolve with empty array
                        resolve([])
                    }
                },
                (error) => reject(new Error(`Error fetching Capability items: ${formatError(error)}`))
            )
        });
    }

    // GET DOCUMENTS BY APPLICATION
    static getDocumentsByCapability(appId: number): Promise<IDocumentItem[]> {

        return new Promise<IDocumentItem[]>((resolve, reject) => {

            // load the data
            Web().Lists(Strings.Sites.main.lists.Documents).Items().query({
                Select: ["File_x0020_Type", "UniqueId", "Id", "ServerRedirectedEmbedUrl", "EncodedAbsUrl", "FileLeafRef", "capability/Id",
                    "Modified", "Editor/Id", "Editor/EMail", "Editor/Title", "docType"],
                Filter: `capability/Id eq ${appId}`,
                Expand: ["capability", "Editor"]
            }).execute(
                // Success
                items => {
                    if (items?.results?.length) {
                        const docs = items.results as unknown as IDocumentItem[];
                        resolve(docs);
                    } else {
                        // resolve with empty array
                        resolve([]);
                    }
                },
                // Error
                (error) => {
                    reject(new Error(`Error getting Documents: ${formatError(error)}`));
                }
            )
        });
    }

    // GET CONTRACT DOCUMENTS BY CONTRACT ITEM
    static getDocumentsByContract(contractId: number): Promise<IContractDocumentItem[]> {

        return new Promise<IContractDocumentItem[]>((resolve, reject) => {

            // Load only documents linked to the selected DCTContracts item.
            Web().Lists(Strings.Sites.main.lists.ContractDocuments).Items().query({
                Select: ["File_x0020_Type", "UniqueId", "Id", "ServerRedirectedEmbedUrl", "EncodedAbsUrl", "FileLeafRef", "contract/Id",
                    "Modified", "Editor/Id", "Editor/EMail", "Editor/Title", "cdocType", "Title"],
                Filter: `contract/Id eq ${contractId}`,
                Expand: ["contract", "Editor"]
            }).execute(
                // Success
                items => {
                    if (items?.results?.length) {
                        const docs = items.results as unknown as IContractDocumentItem[];
                        resolve(docs);
                    } else {
                        // Resolve with an empty array so callers can render the empty state.
                        resolve([]);
                    }
                },
                // Error
                (error) => {
                    reject(new Error(`Error getting Contract Documents: ${formatError(error)}`));
                }
            )
        });
    }

    // GET ONLY SCREENSHOTS / DOCUMENTS FOR PDF EXPORT
    static getScreenshotsForBook(): Promise<IDocumentItem[]> {

        return new Promise<IDocumentItem[]>((resolve, reject) => {

            // load the data
            Web().Lists(Strings.Sites.main.lists.Documents).Items().query({
                Select: ["File_x0020_Type", "UniqueId", "Id", "ServerRedirectedEmbedUrl", "EncodedAbsUrl",
                    "FileLeafRef", "capability/Id", "docType", "Modified", "Editor/Id", "Editor/EMail", "Editor/Title"],
                Filter: `docType eq 'Screenshot'`,
                Expand: ["capability", "Editor"]
            }).execute(
                // Success
                items => {
                    if (items && items.results && Array.isArray(items.results)) {
                        const docs = items.results as unknown as IDocumentItem[];
                        resolve(docs);
                    } else {
                        reject(new Error("No screenshots found or unexpected data structure"));
                    }
                },
                // Error
                (error) => {
                    reject(new Error(`Error getting screenshots: ${formatError(error)}`));
                }
            )
        });
    }

    //GET CONTRACTS FROM JAMIS ENDPOINT
    private static _jamisContracts: IContractEndPointItem[] = [];
    private static _jamisContractsLoaded: boolean = false;
    static get JamisContracts(): IContractEndPointItem[] { return this._jamisContracts; }
    static getJamisContracts(): Promise<IContractEndPointItem[]> {
        return new Promise<IContractEndPointItem[]>((resolve, reject) => {
            this._jamisContracts = [];

            Web(Strings.Sites.jamis.url)
                .Lists(Strings.Sites.jamis.lists.ContractEP)
                .Items()
                .query({
                    GetAllItems: true,
                    Select: ["Id", "Title", "field_19", "field_20", "field_35", "field_21", "field_23", "field_75"],
                    OrderBy: ["field_20"]
                })
                .execute(
                    (items) => {
                        this._jamisContracts = (items?.results ?? []) as unknown as IContractEndPointItem[];
                        this._jamisContractsLoaded = true;
                        resolve(this._jamisContracts);
                    },
                    (error) => reject(new Error(`Error fetching Jamis Contracts: ${formatError(error)}`))
                );
        });
    }

    //GET CONTRACTS FROM CMS SITE
    private static _cmsContracts: ICMSContractItem[] = [];
    private static _cmsContractsLoaded: boolean = false;
    static get CmsContracts(): ICMSContractItem[] { return this._cmsContracts; }
    static getCmsContracts(): Promise<ICMSContractItem[]> {
        return new Promise<ICMSContractItem[]>((resolve, reject) => {
            this._cmsContracts = [];

            Web(Strings.Sites.cms.url)
                .Lists(Strings.Sites.cms.lists.Contracts)
                .Items()
                .query({
                    GetAllItems: true,
                    Select: [
                        "Id", "Title", "ContractNumber", "ProjectID", "OperatingGroup",
                        "ProjectManager/Id", "ProjectManager/Title", "ProjectManager/EMail"
                    ],
                    Expand: ["ProjectManager"],
                    OrderBy: ["Title"]
                })
                .execute(
                    (items) => {
                        this._cmsContracts = (items?.results ?? []) as unknown as ICMSContractItem[];
                        this._cmsContractsLoaded = true;
                        resolve(this._cmsContracts);
                    },
                    (error) => reject(new Error(`Error fetching CMS Contracts: ${formatError(error)}`))
                );
        });
    }

    // Store the normalized, de-duplicated contract source rows used by the contract form lookups.
    private static _contractSources: IContractSourceItem[] = [];
    static get ContractSources(): IContractSourceItem[] { return this._contractSources; }

    // Normalize text before comparing source-list rows so casing, whitespace, and punctuation are consistent.
    private static getContractSourceKey(value?: string): string {
        return (value ?? "").trim().toLowerCase().replace(/\s+/g, " ");
    }

    // Build every comparison key that should count as the same contract across source systems.
    private static getContractSourceKeys(source: Pick<IContractSourceItem, "Title" | "customerContractCode">): string[] {
        return [
            source.customerContractCode ? `code:${DataSource.getContractSourceKey(source.customerContractCode)}` : "",
            source.Title ? `title:${DataSource.getContractSourceKey(source.Title)}` : ""
        ].filter(Boolean);
    }

    // Convert a Jamis endpoint row into the common source shape consumed by ContractForm.
    private static normalizeJamisContractSource(contract: IContractEndPointItem): IContractSourceItem {
        const ogTitle = (contract.field_75 ?? "").trim();
        const ogItem = DataSource.OGs.find((og) => og.Title.toLowerCase() === ogTitle.toLowerCase());

        return {
            source: "jamis",
            sourceId: contract.Id,
            sourceLabel: "JAMIS",
            contractId: contract.field_19 ?? "",
            Title: contract.field_20 ?? "",
            customerContractCode: contract.field_35 ?? "",
            projectManagerEmail: contract.field_21 ?? "",
            projectManagerName: contract.field_23 ?? "",
            ogTitle,
            lobTitle: ogItem?.lob?.Title ?? ""
        };
    }

    // Convert a CMS contract row into the same source shape while preserving CMS-specific identifiers.
    private static normalizeCmsContractSource(contract: ICMSContractItem): IContractSourceItem {
        // CMS stores the OG title directly, so use the same OG lookup table as Jamis to derive LOB.
        const ogTitle = (contract.OperatingGroup ?? "").trim();
        const ogItem = DataSource.OGs.find((og) => og.Title.toLowerCase() === ogTitle.toLowerCase());

        return {
            source: "cms",
            sourceId: contract.Id,
            sourceLabel: "CMS",
            contractId: contract.ProjectID ?? "",
            Title: contract.Title ?? "",
            customerContractCode: contract.ContractNumber ?? "",
            projectManager: contract.ProjectManager?.Id ? {
                Id: contract.ProjectManager.Id,
                EMail: contract.ProjectManager.EMail,
                Title: contract.ProjectManager.Title
            } : undefined,
            projectManagerEmail: contract.ProjectManager?.EMail ?? "",
            projectManagerName: contract.ProjectManager?.Title ?? "",
            ogTitle,
            lobTitle: ogItem?.lob?.Title ?? ""
        };
    }

    // Add a normalized source row when none of its contract identity keys have been seen yet.
    private static addContractSource(sourceMap: Map<string, IContractSourceItem>, source: IContractSourceItem): void {
        const keys = DataSource.getContractSourceKeys(source);
        if (!keys.length) return;
        if (keys.some((key) => sourceMap.has(key))) return;

        keys.forEach((key) => sourceMap.set(key, source));
    }

    // Merge Jamis and CMS into one searchable collection, giving Jamis priority when duplicates exist.
    private static buildContractSources(): void {
        const sourceMap = new Map<string, IContractSourceItem>();

        this._jamisContracts
            .map((contract) => DataSource.normalizeJamisContractSource(contract))
            .forEach((source) => DataSource.addContractSource(sourceMap, source));

        this._cmsContracts
            .map((contract) => DataSource.normalizeCmsContractSource(contract))
            .forEach((source) => DataSource.addContractSource(sourceMap, source));

        this._contractSources = Array.from(new Set(sourceMap.values()))
            .sort((a, b) => (a.Title ?? "").localeCompare(b.Title ?? ""));
    }

    //GET ALL OG's
    private static _ogs: IOgItem[] = [];
    static get OGs(): IOgItem[] { return this._ogs; }
    static getOGs(): Promise<IOgItem[]> {
        return new Promise<IOgItem[]>((resolve, reject) => {
            this._ogs = [];

            Web(Strings.Sites.orgLookups.url)
                .Lists(Strings.Sites.orgLookups.lists.OGs)
                .Items()
                .query({
                    GetAllItems: true,
                    OrderBy: ["Title"],
                    Filter: "isActive eq 1",
                    Select: [
                        "Id", "Title", "lob/Id",
                        "lob/Title", "president/Id", "president/Title",
                        "president/EMail", "ogType", "parentOg/Id",
                        "parentOg/Title", "isActive", "isSelectable"
                    ],
                    Expand: ["lob", "president", "parentOg"]
                })
                .execute(
                    (items) => {
                        this._ogs = (items?.results ?? []) as unknown as IOgItem[];
                        resolve(this._ogs);
                    },
                    (error) => reject(new Error(`Error fetching OGs: ${formatError(error)}`))
                );
        });
    }

    //GET OPPORTUNITIES FROM OPPNET SITE
    private static _opportunities: IOpportunityItem[] = [];
    static get Opportunities(): IOpportunityItem[] { return this._opportunities; }
    static getOpportunities(): Promise<IOpportunityItem[]> {
        return new Promise<IOpportunityItem[]>((resolve, reject) => {
            Web(Strings.Sites.oppNet.url)
                .Lists(Strings.Sites.oppNet.lists.Opportunities)
                .Items()
                .query({
                    //GetAllItems: true,
                    Select: ["Id", "Title", "Customer", "Status"],
                    OrderBy: ["Title"],
                    Filter: "Status eq 'Won'", //only show opps where we won the opp
                    Top: 5000
                })
                .execute(
                    (items) => {
                        this._opportunities = (items?.results ?? []) as unknown as IOpportunityItem[];
                        resolve(this._opportunities);
                    },
                    (error) => reject(new Error(`Error fetching Opportunities: ${formatError(error)}`))
                );
        });
    }

    //GET PAST PERFORMANCE FROM PROPOSAL SITE
    private static _pastPerformance: IPastPerformanceItem[] = [];
    static get PastPerformance(): IPastPerformanceItem[] { return this._pastPerformance; }
    static getPastPerformance(): Promise<IPastPerformanceItem[]> {
        return new Promise<IPastPerformanceItem[]>((resolve, reject) => {
            Web(Strings.Sites.proposals.url)
                .Lists(Strings.Sites.proposals.lists.PastPerformance)
                .Items()
                .query({
                    //GetAllItems: true,
                    Select: ["Id", "Contract_x0023_", "Customer_x0020_Agency", "Doc_x0020_Type", "Capability_x0020_Area"],
                    OrderBy: ["Title"],
                    Top: 5000
                })
                .execute(
                    (items) => {
                        this._pastPerformance = (items?.results ?? []) as unknown as IPastPerformanceItem[];
                        resolve(this._pastPerformance);
                    },
                    (error) => reject(new Error(`Error fetching Past Performance: ${formatError(error)}`))
                );
        });
    }

    //GET PROPOSALS FROM PROPOSAL SITE
    private static _proposals: IProposalItem[] = [];
    static get Proposals(): IProposalItem[] { return this._proposals; }
    static getProposals(): Promise<IProposalItem[]> {
        return new Promise<IProposalItem[]>((resolve, reject) => {
            Web(Strings.Sites.proposals.url)
                .Lists(Strings.Sites.proposals.lists.Proposals)
                .Items()
                .query({
                    //GetAllItems: true,
                    Select: ["Id", "Title", "OpportunityStage", "TypeOfOpportunity", "Entity"],
                    OrderBy: ["Title"],
                    Filter: "OpportunityStage eq 'Awarded (won)'", //only show proposals where we won the contract
                    Top: 5000
                })
                .execute(
                    (items) => {
                        this._proposals = (items?.results ?? []) as unknown as IProposalItem[];
                        resolve(this._proposals);
                    },
                    (error) => reject(new Error(`Error fetching Proposals: ${formatError(error)}`))
                );
        });
    }

}
