import { ICapabilityItem } from "../common/props";
import { Web } from "gd-sprest";
import Strings from "../common/strings";
import { formatError, encodeListName } from "../common/utils";
import { DataSource } from "../data/ds";


export class CapabilityService {
    
    private static getListItemType(): string {
        return `SP.Data.${encodeListName(Strings.Lists.Capabilities)}ListItem`;
    }

    private static buildPayload(item: ICapabilityItem): Record<string, unknown> {
        return {
            __metadata: { type: CapabilityService.getListItemType() },
            Title: item.Title,
            description: item.description,
            capabilities: item.capabilities,
            link: item.link,
            capStatus: item.capStatus,
            notes: item.notes,
            platform: item.platform,
            hostingEnv: item.hostingEnv,
            compliance: item.compliance,
            licenseReqd: item.licenseReqd,
            licenseReqmts: item.licenseReqmts,
            connectivity: item.connectivity,
            extensibility: item.extensibility,
            serverReqmts: item.serverReqmts,
            codeLanguage: item.codeLanguage,
            backend: item.backend,
            contractId: item.contract?.Id ?? null,
        };
    }

    private static loadAppById(itemId: number): ICapabilityItem {
        const item = Web().Lists(Strings.Lists.Capabilities).Items(itemId.toString())
            .query({
                Select: DataSource.capabilityQuerySelect,
                Expand: DataSource.capabilityQueryExpand
            })
            .executeAndWait();

        return item as unknown as ICapabilityItem;
    }

    static create(item: ICapabilityItem): Promise<ICapabilityItem> {
        return new Promise<ICapabilityItem>((resolve, reject) => {
            Web().Lists(Strings.Lists.Capabilities).Items().add(CapabilityService.buildPayload(item)).execute(
                (resp) => {
                    if (resp && resp.existsFl && resp.Id) {
                        const newItem: ICapabilityItem = CapabilityService.loadAppById(resp.Id);
                        resolve(newItem);
                        return;
                    }

                    reject("Item was created but there was a problem refreshing the data. Please refresh manually.");
                },
                (error) => {
                    const err: string = formatError(error);
                    console.error(`Error creating new Capability Entry ${err}`);
                    reject(error);
                }
            );
        });
    }

    static edit(item: ICapabilityItem): Promise<ICapabilityItem> {
        return new Promise<ICapabilityItem>((resolve, reject) => {
            Web().Lists(Strings.Lists.Capabilities).Items().getById(item.Id).update(CapabilityService.buildPayload(item)).execute(
                (resp) => {
                    if (resp && resp.existsFl) {
                        const updatedItem: ICapabilityItem = CapabilityService.loadAppById(item.Id);
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
            Web().Lists(Strings.Lists.Capabilities).Items(itemId).delete().execute(
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
