import { IContractItem } from "../common/props";
import { DataSource } from "../data/ds";
import { Web } from "gd-sprest-bs";
import Strings from "../common/strings";
import { encodeListName, formatError } from "../common/utils";

export class ContractService {
    private static getListItemType(): string {
        return `SP.Data.${encodeListName(Strings.Lists.Contracts)}ListItem`;
    }

    private static normalizeDateValue(value?: string): string | null {
        return value && value !== "" ? value : null;
    }

    private static buildPayload(item: IContractItem): Record<string, unknown> {
        return {
            __metadata: { type: ContractService.getListItemType() },
            Title: item.Title,
            contractId: item.contractId,
            invoice: item.invoice,
            customerContractCode: item.customerContractCode,
            customer: item.customer,
            popStart: ContractService.normalizeDateValue(item.popStart),
            popEnd: ContractService.normalizeDateValue(item.popEnd),
            contractPmId: item.contractPm?.Id ?? null,
            primaryPocId: item.primaryPoc?.Id ?? null,
            stakeholdersId: { results: item.stakeholders?.results.map((person) => person.Id) ?? [] },
            partner: item.partner,
            infoLink: item.infoLink
        };
    }

    private static loadContractById(itemId: number): IContractItem {
        const item = Web().Lists(Strings.Lists.Contracts).Items(itemId.toString())
            .query({
                Select: DataSource.contractQuerySelect,
                Expand: DataSource.contractQueryExpand
            })
            .executeAndWait();

        return item as unknown as IContractItem;
    }

    static create(item: IContractItem): Promise<IContractItem> {
        return new Promise<IContractItem>((resolve, reject) => {
            Web().Lists(Strings.Lists.Contracts).Items().add(ContractService.buildPayload(item)).execute(
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
            Web().Lists(Strings.Lists.Contracts).Items(item.Id).update(ContractService.buildPayload(item)).execute(
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
            Web().Lists(Strings.Lists.Contracts).Items(itemId).delete().execute(
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
}
