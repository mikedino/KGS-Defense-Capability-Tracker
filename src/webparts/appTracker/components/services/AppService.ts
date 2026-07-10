import { IApplicationItem } from "../data/props";
import { Web } from "gd-sprest";
import Strings from "../../strings";
import { formatError, encodeListName } from "../utils";
import { DataSource } from "../data/ds";


export class AppService {
    
    private static getListItemType(): string {
        return `SP.Data.${encodeListName(Strings.Lists.Apps)}ListItem`;
    }

    private static normalizeDateValue(value?: string): string | null {
        return value && value !== "" ? value : null;
    }

    private static buildPayload(item: IApplicationItem): Record<string, unknown> {
        return {
            __metadata: { type: AppService.getListItemType() },
            Title: item.Title,
            description: item.description,
            appUrl: item.appUrl,
            relatedInfo: item.relatedInfo,
            //highlights: item.highlights,
            primaryPocId: item.primaryPoc?.Id ?? null,
            systemOwnerId: item.systemOwner?.Id ?? null,
            stakeholdersId: { results: item.stakeholders?.results.map((bo) => bo.Id) ?? [] },
            managingGroup: item.managingGroup,
            supportTeam: item.supportTeam,
            licenseReqd: item.licenseReqd,
            connectivity: item.connectivity,
            integrationsInput: item.integrationsInput,
            integrationsOutput: item.integrationsOutput,
            synonyms: item.synonyms,
            platform: item.platform,
            environment: item.environment,
            contractId: item.contract?.Id ?? null,
            appStatus: item.appStatus,
            appLaunchDate: AppService.normalizeDateValue(item.appLaunchDate),
            userCount: item.userCount
        };
    }

    private static loadAppById(itemId: number): IApplicationItem {
        const item = Web().Lists(Strings.Lists.Apps).Items(itemId.toString())
            .query({
                Select: DataSource.appQuerySelect,
                Expand: DataSource.appQueryExpand
            })
            .executeAndWait();

        return item as unknown as IApplicationItem;
    }

    static create(item: IApplicationItem): Promise<IApplicationItem> {
        return new Promise<IApplicationItem>((resolve, reject) => {
            Web().Lists(Strings.Lists.Apps).Items().add(AppService.buildPayload(item)).execute(
                (resp) => {
                    if (resp && resp.existsFl && resp.Id) {
                        const newItem: IApplicationItem = AppService.loadAppById(resp.Id);
                        resolve(newItem);
                        return;
                    }

                    reject("Item was created but there was a problem refreshing the data. Please refresh manually.");
                },
                (error) => {
                    const err: string = formatError(error);
                    console.error(`Error creating new App Entry ${err}`);
                    reject(error);
                }
            );
        });
    }

    static edit(item: IApplicationItem): Promise<IApplicationItem> {
        return new Promise<IApplicationItem>((resolve, reject) => {
            Web().Lists(Strings.Lists.Apps).Items().getById(item.Id).update(AppService.buildPayload(item)).execute(
                (resp) => {
                    if (resp && resp.existsFl) {
                        const updatedItem: IApplicationItem = AppService.loadAppById(item.Id);
                        resolve(updatedItem);
                        return;
                    }

                    reject("Item was edited but there was a problem refreshing the data. Please refresh manually.");
                },
                (error) => {
                    const err: string = formatError(error);
                    console.error(`Error updating App ${err}`);
                    reject(error);
                }
            );
        });
    }


    static delete(itemId: number): Promise<void> {
        return new Promise<void>((resolve, reject) => {
            Web().Lists(Strings.Lists.Apps).Items(itemId).delete().execute(
                () => {
                    console.info(`Deleted App item ${itemId} !`)
                    resolve();
                },
                (error) => {
                    const err = formatError(error);
                    console.error(`Error deleting App item: ${err}`);
                    reject(error);
                }
            )
        })
    }

}
