import { ICapabilityContractDraft, IContractItem } from "../common/props";
import { DataSource } from "../data/ds";
import { Web } from "gd-sprest-bs";
import Strings from "../common/strings";
import { encodeListName, formatError } from "../common/utils";

export class ContractService {
    private static getListItemType(): string {
        return `SP.Data.${encodeListName(Strings.Sites.main.lists.Contracts)}ListItem`;
    }

    private static normalizeDateValue(value?: string): string | null {
        return value && value !== "" ? value : null;
    }

    private static buildPayload(item: IContractItem): Record<string, unknown> {
        return {
            __metadata: { type: ContractService.getListItemType() },
            Title: item.Title,
            capabilityId: item.capability?.Id,
            contractId: item.contractId,
            customerContractCode: item.customerContractCode,
            customer: item.customer,
            startDate: ContractService.normalizeDateValue(item.startDate),
            endDate: ContractService.normalizeDateValue(item.endDate),
            contractPmId: item.contractPm?.Id ?? null,
            partner: item.partner,
            infoLink: item.infoLink
        };
    }

    private static loadContractById(itemId: number): IContractItem {
        const item = Web().Lists(Strings.Sites.main.lists.Contracts).Items(itemId.toString())
            .query({
                Select: DataSource.contractQuerySelect,
                Expand: DataSource.contractQueryExpand
            })
            .executeAndWait();

        return item as unknown as IContractItem;
    }

    static create(item: IContractItem): Promise<IContractItem> {
        return new Promise<IContractItem>((resolve, reject) => {
            Web().Lists(Strings.Sites.main.lists.Contracts).Items().add(ContractService.buildPayload(item)).execute(
                (resp) => {
                    if (resp && resp.Id) {
                        resolve(ContractService.loadContractById(resp.Id));
                        return;
                    }

                    reject("Item was created but there was a problem refreshing the data. Please refresh manually.");
                },
                (error) => {
                    console.error(`Error creating new Contract ${formatError(error)}`);
                    reject(error);
                }
            );
        });
    }

    static edit(item: IContractItem): Promise<IContractItem> {
        return new Promise<IContractItem>((resolve, reject) => {
            Web().Lists(Strings.Sites.main.lists.Contracts).Items(item.Id).update(ContractService.buildPayload(item)).execute(
                (resp) => {
                    if (resp) {
                        resolve(ContractService.loadContractById(item.Id));
                        return;
                    }

                    reject("Item was edited but there was a problem refreshing the data. Please refresh manually.");
                },
                (error) => {
                    console.error(`Error updating Contract ${formatError(error)}`);
                    reject(error);
                }
            );
        });
    }

    static delete(itemId: number): Promise<void> {
        return new Promise<void>((resolve, reject) => {
            Web().Lists(Strings.Sites.main.lists.Contracts).Items(itemId).delete().execute(
                () => {
                    console.info(`Deleted Contract ${itemId} !`);
                    resolve();
                },
                (error) => {
                    console.error(`Error deleting Contract: ${formatError(error)}`);
                    reject(error);
                }
            );
        });
    }

    static async saveForCapability(
        capabilityId: number,
        contracts: ICapabilityContractDraft[],
        deletedContractIds: number[] = []
    ): Promise<IContractItem[]> {
        await Promise.all(deletedContractIds.map((itemId) => ContractService.delete(itemId)));

        const savedContracts: IContractItem[] = [];
        for (const draft of contracts) {
            const relationship: IContractItem = {
                ...draft,
                capability: {
                    Id: capabilityId,
                    Title: draft.capability?.Title ?? ""
                }
            };

            if (relationship.Id > 0) {
                savedContracts.push(await ContractService.edit(relationship));
            } else {
                savedContracts.push(await ContractService.create(relationship));
            }
        }

        return savedContracts;
    }
}
