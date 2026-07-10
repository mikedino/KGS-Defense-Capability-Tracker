import { Web } from "gd-sprest-bs";
import { WebPartContext } from '@microsoft/sp-webpart-base';
import Strings, { setContext } from "../../strings";
import { IApplicationItem, IConfigurationItem, IContractItem, IDocumentItem } from "./props";
import { formatError } from "../utils";

export class DataSource {
    private static readonly defaultConfigValuesByFor = new Map<string, string[]>([
        ["appStatus", ["Backlog", "Reqs Gathering", "In Development", "Active", "Enhancing"]]
    ]);

    //prevent this from being initialized twice
    static initialized: boolean = false;

    // Initializes the application
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
                this.getApps()
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
        "Id", "Title", "contractTeam", "start", "end", "contractValue",
        "cor/Id", "cor/Title", "cor/EMail", "cor/JobTitle",
        "acor/Id", "acor/Title", "acor/EMail", "acor/JobTitle",
        "ko/Id", "ko/Title", "ko/EMail", "ko/JobTitle",
        "primaryPoc/Id", "primaryPoc/Title", "primaryPoc/EMail", "primaryPoc/JobTitle",
        "secondaryPoc/Id", "secondaryPoc/Title", "secondaryPoc/EMail", "secondaryPoc/JobTitle"
    ];
    static contractQueryExpand: string[] = ["cor", "acor", "ko", "primaryPoc", "secondaryPoc"];
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

    // Load the Applications
    static appQuerySelect: string[] = [
        "Id", "Title", "description", "appStatus", "appLaunchDate", "appUrl", "relatedInfo",
        "contract/Id", "contract/Title", "contract/contractTeam", "highlights", "platform",
        "licenseReqd", "integrationsInput", "integrationsOutput",
        "userCount", "managingGroup", "environment",
        "primaryPoc/Id", "primaryPoc/Title", "primaryPoc/EMail", "primaryPoc/JobTitle", "primaryPoc/Department", "supportTeam",
        "connectivity", "synonyms",
        "stakeholders/Id", "stakeholders/Title", "stakeholders/EMail", "Modified",
        "Author/Title", "Author/EMail", "Author/Id",
        "systemOwner/Title", "systemOwner/EMail", "systemOwner/Id"
    ]
    static appQueryExpand: string[] = ["primaryPoc", "stakeholders", "contract", "Author", "systemOwner"];
    private static _apps: IApplicationItem[] = [];
    static get Apps(): IApplicationItem[] { return this._apps; }
    private static getApps(): Promise<IApplicationItem[]> {
        return new Promise<IApplicationItem[]>((resolve, reject) => {

            // clear the items
            this._apps = [];

            // load the data
            Web().Lists(Strings.Lists.Apps).Items().query({
                GetAllItems: true,
                OrderBy: ["Modified desc"],
                Select: this.appQuerySelect,
                Expand: this.appQueryExpand,
                Top: 5000
            }).execute(
                // Success
                (items) => {
                    if (items?.results?.length) {
                        this._apps = items.results as unknown as IApplicationItem[];
                        resolve(this._apps);
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
    static getDocumentsByApplication(appId: number): Promise<IDocumentItem[]> {

        return new Promise<IDocumentItem[]>((resolve, reject) => {

            // load the data
            Web().Lists(Strings.Lists.Documents).Items().query({
                Select: ["File_x0020_Type", "UniqueId", "Id", "ServerRedirectedEmbedUrl", "EncodedAbsUrl", "FileLeafRef", "application/Id",
                    "Modified", "Editor/Id", "Editor/EMail", "Editor/Title", "docType"],
                Filter: `application/Id eq ${appId}`,
                Expand: ["application", "Editor"]
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
                    "FileLeafRef", "application/Id", "docType", "Modified", "Editor/Id", "Editor/EMail", "Editor/Title"],
                Filter: `docType eq 'Screenshot'`,
                Expand: ["application", "Editor"]
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
