import { Web } from "gd-sprest-bs";
import { WebPartContext } from '@microsoft/sp-webpart-base';
import Strings, { setContext } from "../common/strings";
import { ICapabilityItem, IConfigurationItem, IContractItem, IDocumentItem } from "../common/props";
import { formatError } from "../common/utils";

export class DataSource {
    private static readonly defaultConfigValuesByFor = new Map<string, string[]>([
        ["capStatus", ["Active", "In Dev", "Pending"]],
        ["platform", ["SharePoint", "Web", "Desktop"]],
        ["hostingEnv", ["SharePoint Online", "Azure", "On-Premises"]],
        ["connectivity", ["Internal", "External", "Hybrid"]],
        ["compliance", ["FedRAMP", "FISMA", "N/A"]],
        ["codeLanguage", ["C#", "Java", "TypeScript"]],
        ["backend", ["SharePoint", "SQL", "Oracle"]],
        ["customer", ["DOW", "DOS", "USN"]],
        ["partner", ["LMI", "Guidehouse"]]
    ]);

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

            return Promise.all([
                this.getConfig(),
                this.getContracts(),
                this.getCapabilities()
            ]).then(() => {
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
    private static _config: IConfigurationItem[] = [];
    static get Config(): IConfigurationItem[] { return this._config; }

    // Grouped items
    private static _configByFor = new Map<string, IConfigurationItem[]>();
    static get ConfigByFor(): ReadonlyMap<string, IConfigurationItem[]> { return this._configByFor; }

    // Grouped values (Title)
    private static _configValuesByFor = new Map<string, string[]>();
    static get ConfigValuesByFor(): ReadonlyMap<string, string[]> { return this._configValuesByFor; }

    // Convenience helper
    static getConfigValues(isFor: string): string[] {
        const configured = this._configValuesByFor.get(isFor) ?? [];
        return configured.length ? configured : (this.defaultConfigValuesByFor.get(isFor) ?? []);
    }

    // get/set config
    static getConfig(): Promise<IConfigurationItem[]> {
        return new Promise<IConfigurationItem[]>((resolve, reject) => {

            // clear the items
            this._config = [];
            this._configByFor = new Map<string, IConfigurationItem[]>();
            this._configValuesByFor = new Map<string, string[]>();

            // load the data
            Web().Lists(Strings.Lists.Configuration).Items().query({
                GetAllItems: true,
                OrderBy: ["isFor"],
                Select: ["Id", "Title", "isFor", "isForDisplayName", "infoText", "isActive"],
                //Filter: "isActive eq 1",
                Top: 5000
            }).execute(
                (items) => {
                    if (items?.results?.length) {
                        const results = items.results as unknown as IConfigurationItem[];
                        this._config = results;

                        // Group + build value arrays
                        for (const item of results) {
                            const forKey = (item.isFor ?? "").trim();
                            if (!forKey) continue;

                            const list = this._configByFor.get(forKey) ?? [];
                            list.push(item);
                            this._configByFor.set(forKey, list);

                            const values = this._configValuesByFor.get(forKey) ?? [];
                            if (item.Title) values.push(item.Title);
                            this._configValuesByFor.set(forKey, values);
                        }

                        // Optional: dedupe + sort values per key
                        for (const [k, vals] of this._configValuesByFor.entries()) {
                            const unique = Array.from(new Set(vals.map(v => v.trim()).filter(Boolean)));
                            unique.sort((a, b) => a.localeCompare(b));
                            this._configValuesByFor.set(k, unique);
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
        "Id", "Title", "contractId", "invoice", "customerContractCode", "customer",
        "popStart", "popEnd", "partner", "infoLink",
        "contractPm/Id", "contractPm/Title", "contractPm/EMail", "contractPm/JobTitle", "contractPm/Department",
        "primaryPoc/Id", "primaryPoc/Title", "primaryPoc/EMail", "primaryPoc/JobTitle",
        "primaryPoc/Department",
        "stakeholders/Id", "stakeholders/Title", "stakeholders/EMail", "stakeholders/JobTitle",
        "stakeholders/Department"
    ];
    static contractQueryExpand: string[] = ["contractPm", "primaryPoc", "stakeholders"];
    private static _contracts: IContractItem[] = [];
    static get Contracts(): IContractItem[] { return this._contracts; }
    private static getContracts(): Promise<IContractItem[]> {
        return new Promise<IContractItem[]>((resolve, reject) => {

            // clear the items
            this._contracts = [];

            // load the data
            Web().Lists(Strings.Lists.Contracts).Items().query({
                GetAllItems: true,
                Select: this.contractQuerySelect,
                OrderBy: ["Title"],
                Expand: this.contractQueryExpand,
                Top: 5000
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
        "contract/Id", "contract/Title", "contract/contractId", "Modified",
        "Author/Title", "Author/EMail", "Author/Id"
    ]
    static capabilityQueryExpand: string[] = ["contract", "Author"];
    private static _capabilities: ICapabilityItem[] = [];
    static get Capabilities(): ICapabilityItem[] { return this._capabilities; }
    private static getCapabilities(): Promise<ICapabilityItem[]> {
        return new Promise<ICapabilityItem[]>((resolve, reject) => {

            // clear the items
            this._capabilities = [];

            // load the data
            Web().Lists(Strings.Lists.Capabilities).Items().query({
                GetAllItems: true,
                OrderBy: ["Modified desc"],
                Select: this.capabilityQuerySelect,
                Expand: this.capabilityQueryExpand,
                Top: 5000
            }).execute(
                // Success
                (items) => {
                    if (items?.results?.length) {
                        this._capabilities = items.results as unknown as ICapabilityItem[];
                        resolve(this._capabilities);
                    } else {
                        //none found - resolve with empty array
                        resolve([])
                    }
                },
                (error) => reject(new Error(`Error fetching App items: ${formatError(error)}`))
            )
        });
    }

    // GET DOCUMENTS BY APPLICATION
    static getDocumentsByCapability(appId: number): Promise<IDocumentItem[]> {

        return new Promise<IDocumentItem[]>((resolve, reject) => {

            // load the data
            Web().Lists(Strings.Lists.Documents).Items().query({
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
            Web().Lists(Strings.Lists.Documents).Items().query({
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

}
