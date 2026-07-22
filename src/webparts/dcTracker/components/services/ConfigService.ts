import { IConfigurationItem } from "../common/props";
import { Web } from "gd-sprest";
import Strings from "../common/strings";
import { formatError, encodeListName } from "../common/utils";

export class ConfigService {

    static create(item: IConfigurationItem): Promise<IConfigurationItem> {
        return new Promise<IConfigurationItem>((resolve, reject) => {
            Web().Lists(Strings.Lists.Configuration).Items().add({
                __metadata: { type: `SP.Data.${encodeListName(Strings.Lists.Configuration)}ListItem` },
                Title: item.Title,
                isFor: item.isFor,
                isForDisplayName: item.isForDisplayName,
                isActive: item.isActive,
                infoText: item.infoText
            }).execute(
                (resp) => {
                    if (resp && resp.existsFl && resp.Id) {
                        //get the item and all correct metadata to return to sender
                        const newItem = Web().Lists(Strings.Lists.Configuration).Items(resp.Id.toString())
                            .query({
                                Select: ["Id", "Title", "isFor", "isForDisplayName", "isActive", "infoText"]
                            }).executeAndWait();
                        resolve(newItem as unknown as IConfigurationItem);
                    } else {
                        reject("Item was edited but there was a problem refreshing the data. Please refresh manually.")
                    }
                },
                (error) => {
                    const err = formatError(error);
                    console.error(`Error creating new Config Entry ${err}`);
                    reject(error);
                }
            )
        })
    }

    static edit(item: IConfigurationItem): Promise<IConfigurationItem> {
        return new Promise<IConfigurationItem>((resolve, reject) => {
            Web().Lists(Strings.Lists.Configuration).Items().getById(item.Id).update({
                __metadata: { type: `SP.Data.${encodeListName(Strings.Lists.Configuration)}ListItem` },
                Title: item.Title,
                isFor: item.isFor,
                isForDisplayName: item.isForDisplayName,
                isActive: item.isActive,
                infoText: item.infoText
            }).execute(
                //success
                (resp) => {
                    if (resp && resp.existsFl) {
                        //get the item and all correct metadata to return to sender
                        const editItem = Web().Lists(Strings.Lists.Configuration).Items(item.Id.toString())
                            .query({
                                Select: ["Id", "Title", "isFor", "isForDisplayName", "isActive", "infoText"],
                            }).executeAndWait();
                        resolve(editItem as unknown as IConfigurationItem);
                    } else {
                        reject("Item was edited but there was a problem refreshing the data. Please refresh manually.")
                    }
                },
                //error
                (error) => {
                    const err = formatError(error);
                    console.error(`Error updating Config ${err}`);
                    reject(error);
                }
            )
        })
    }

    static delete(itemId: number): Promise<void> {
        return new Promise<void>((resolve, reject) => {
            Web().Lists(Strings.Lists.Configuration).Items(itemId).delete().execute(
                () => {
                    console.info(`Deleted Config item ${itemId} !`)
                    resolve();
                },
                (error) => {
                    const err = formatError(error);
                    console.error(`Error deleting Config item: ${err}`);
                    reject(error);
                }
            )
        })
    }

}
