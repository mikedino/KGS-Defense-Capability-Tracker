import { Web } from "gd-sprest-bs";
import { WebPartContext } from '@microsoft/sp-webpart-base';
import Strings, { setContext } from "../common/strings";
import { ConfigType, ICapabilityItem, IConfigItem, IContractEndPointItem, IContractItem, IDocumentItem, IOpportunityItem } from "../common/props";
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
                    this.getJamisContracts()
                ]);
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
        "startDate", "endDate", "partner", "infoLink",
        "contractPm/Id", "contractPm/Title", "contractPm/EMail", "contractPm/JobTitle", "contractPm/Department"
    ];
    static contractQueryExpand: string[] = ["capability", "contractPm"];
    private static _contracts: IContractItem[] = [];
    static get Contracts(): IContractItem[] { return this._contracts; }
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

    // Load the Capabilities
    static capabilityQuerySelect: string[] = [
        "Id", "Title", "description", "capabilities", "link", "capStatus", "notes",
        "platform", "hostingEnv", "connectivity", "compliance", "licenseReqd",
        "licenseReqmts", "extensibility", "serverReqmts", "codeLanguage", "backend",
        "oppNetTagsJson", "Modified",
        "primaryPoc/Id", "primaryPoc/Title", "primaryPoc/EMail", "primaryPoc/JobTitle", "primaryPoc/Department",
        "stakeholders/Id", "stakeholders/Title", "stakeholders/EMail", "stakeholders/JobTitle", "stakeholders/Department",
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
                            oppNetTags: parseJsonTagField(item.oppNetTagsJson)
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
    static get JamisContracts(): IContractEndPointItem[] { return this._jamisContracts; }
    static getJamisContracts(): Promise<IContractEndPointItem[]> {
        return new Promise<IContractEndPointItem[]>((resolve, reject) => {
            this._jamisContracts = [];

            const today = new Date();
            today.setHours(0, 0, 0, 0);

            Web(Strings.Sites.jamis.url)
                .Lists(Strings.Sites.jamis.lists.ContractEP)
                .Items()
                .query({
                    GetAllItems: true,
                    OrderBy: ["field_20"],
                    Select: [
                        "Id", "Title", "field_19", "field_20", "field_35", "field_21", "field_23"
                    ]
                })
                .execute(
                    (items) => {
                        this._jamisContracts = (items?.results ?? []) as unknown as IContractEndPointItem[];
                        resolve(this._jamisContracts);
                    },
                    (error) => reject(new Error(`Error fetching Jamis Contracts: ${formatError(error)}`))
                );
        });
    }

    //GET OPPORTUNITIES FROM OPPNET SITE
    private static _opportunities: IOpportunityItem[] = [];
    static get Opportunities(): IOpportunityItem[] { return this._opportunities; }
    static getOpportunities(): Promise<IOpportunityItem[]> {
        return new Promise<IOpportunityItem[]>((resolve, reject) => {
            this._opportunities = [];

            const today = new Date();
            today.setHours(0, 0, 0, 0);

            Web(Strings.Sites.oppNet.url)
                .Lists(Strings.Sites.oppNet.lists.Opportunities)
                .Items()
                .query({
                    GetAllItems: true,
                    OrderBy: ["Title"],
                    Select: ["Id", "Title", "Customer", "Status"]
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

}
