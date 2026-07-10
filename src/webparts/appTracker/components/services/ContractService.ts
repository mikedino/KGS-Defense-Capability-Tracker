import { IContractItem } from "../data/props";
import { DataSource } from "../data/ds";
import { Web } from "gd-sprest-bs";
import Strings from "../../strings";
import { formatError, encodeListName } from "../utils";

export class ContractService {

    static create(item: IContractItem): Promise<IContractItem> {
        return new Promise<IContractItem>((resolve, reject) => {
            Web().Lists(Strings.Lists.Contracts).Items().add({
                __metadata: { type: `SP.Data.${encodeListName(Strings.Lists.Contracts)}ListItem` },
                Title: item.Title,
                contractTeam: item.contractTeam,
                start: item.start && item.start !== "" ? item.start : null,
                end: item.end && item.end !== "" ? item.end : null,
                contractValue: item.contractValue,
                corId: item.cor?.Id,
                acorId: item.acor?.Id,
                koId: item.ko?.Id,
                primaryPocId: item.primaryPoc?.Id,
                secondaryPocId: item.secondaryPoc?.Id
            }).execute(
                //success
                resp => {
                    if (resp && resp.Id) {
                        //get the item and all correct metadata to return to sender
                        const newItem = Web().Lists(Strings.Lists.Contracts).Items(resp.Id.toString())
                            .query({
                                Select: DataSource.contractQuerySelect,
                                Expand: DataSource.contractQueryExpand
                            }).executeAndWait();
                        console.info(`Created Contract ${item?.Id} !`)
                        resolve(newItem as unknown as IContractItem);
                    } else {
                        reject("Item was created but there was a problem refreshing the data. Please refresh manually.")
                    }
                },
                //error
                (error) => {
                    const err = formatError(error);
                    console.error(`Error creating new Contract ${err}`);
                    reject(error);
                }
            )
        })
    }

    static edit(item: IContractItem): Promise<IContractItem> {
        return new Promise<IContractItem>((resolve, reject) => {
            Web().Lists(Strings.Lists.Contracts).Items(item.Id).update({
                __metadata: { type: `SP.Data.${encodeListName(Strings.Lists.Contracts)}ListItem` },
                Title: item.Title,
                contractTeam: item.contractTeam,
                start: item.start && item.start !== "" ? item.start : null,
                end: item.end && item.end !== "" ? item.end : null,
                contractValue: item.contractValue,
                corId: item.cor?.Id ?? null,
                acorId: item.acor?.Id ?? null,
                koId: item.ko?.Id ?? null,
                primaryPocId: item.primaryPoc?.Id ?? null,
                secondaryPocId: item.secondaryPoc?.Id ?? null
            }).execute(
                //success
                resp => {
                    if (resp) {
                        //get the item and all correct metadata to return to sender
                        const editItem = Web().Lists(Strings.Lists.Contracts).Items(item.Id.toString())
                            .query({
                                Select: DataSource.contractQuerySelect,
                                Expand: DataSource.contractQueryExpand
                            }).executeAndWait();
                        resolve(editItem as unknown as IContractItem);
                    } else {
                        reject("Item was edited but there was a problem refreshing the data. Please refresh manually.")
                    }
                },
                //error
                (error) => {
                    const err = formatError(error);
                    console.error(`Error updating Contract ${err}`);
                    reject(error);
                }
            )
        })
    }

    static delete(itemId: number): Promise<void> {
        return new Promise<void>((resolve, reject) => {
            Web().Lists(Strings.Lists.Contracts).Items(itemId).delete().execute(
                //success
                () => {
                    console.info(`Deleted Contract ${itemId} !`)
                    resolve();
                },
                //error
                (error) => {
                    const err = formatError(error);
                    console.error(`Error deleting Contract: ${err}`);
                    reject(error);
                }
            )
        })
    }

}