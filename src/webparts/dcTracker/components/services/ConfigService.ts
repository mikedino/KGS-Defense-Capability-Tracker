import { IConfigItem } from "../common/props";
import { Web } from "gd-sprest";
import Strings from "../common/strings";
import { encodeListName, formatError } from "../common/utils";
import { configSeedItems } from "../data/configSeedItems";

interface IConfigPayload extends Omit<IConfigItem, "Id"> {
    __metadata: { type: string };
}

export class ConfigService {
    private static readonly configSelect = ["Id", "Title", "configType", "configValue", "sortOrder", "isActive", "infoText"];

    private static getListItemType(): string {
        return `SP.Data.${encodeListName(Strings.Sites.main.lists.Configuration)}ListItem`;
    }

    private static async getExistingItemCount(): Promise<number> {
        const list = Web().Lists(Strings.Sites.main.lists.Configuration);

        return new Promise<number>((resolve, reject) => {
            list.Items().query({
                Select: ["Id"],
                Top: 1
            }).execute(
                (results) => {
                    resolve(results?.results?.length ?? 0);
                },
                (error) => reject(new Error(`Error checking Config seed items: ${formatError(error)}`))
            );
        });
    }

    private static buildPayload(item: IConfigItem): IConfigPayload {
        return {
            __metadata: { type: ConfigService.getListItemType() },
            Title: item.Title,
            configType: item.configType,
            configValue: item.configValue,
            sortOrder: item.sortOrder,
            isActive: item.isActive,
            infoText: item.infoText
        };
    }

    static create(item: IConfigItem): Promise<IConfigItem> {
        return new Promise<IConfigItem>((resolve, reject) => {
            Web().Lists(Strings.Sites.main.lists.Configuration).Items().add(ConfigService.buildPayload(item)).execute(
                (resp) => {
                    if (resp && resp.existsFl && resp.Id) {
                        const newItem = Web().Lists(Strings.Sites.main.lists.Configuration).Items(resp.Id.toString())
                            .query({ Select: ConfigService.configSelect })
                            .executeAndWait();
                        resolve(newItem as unknown as IConfigItem);
                        return;
                    }

                    reject("Item was created but there was a problem refreshing the data. Please refresh manually.");
                },
                (error) => {
                    console.error(`Error creating new Config Entry ${formatError(error)}`);
                    reject(error);
                }
            );
        });
    }

    static edit(item: IConfigItem): Promise<IConfigItem> {
        return new Promise<IConfigItem>((resolve, reject) => {
            Web().Lists(Strings.Sites.main.lists.Configuration).Items().getById(item.Id).update(ConfigService.buildPayload(item)).execute(
                (resp) => {
                    if (resp && resp.existsFl) {
                        const editItem = Web().Lists(Strings.Sites.main.lists.Configuration).Items(item.Id.toString())
                            .query({ Select: ConfigService.configSelect })
                            .executeAndWait();
                        resolve(editItem as unknown as IConfigItem);
                        return;
                    }

                    reject("Item was edited but there was a problem refreshing the data. Please refresh manually.");
                },
                (error) => {
                    console.error(`Error updating Config ${formatError(error)}`);
                    reject(error);
                }
            );
        });
    }

    static delete(itemId: number): Promise<void> {
        return new Promise<void>((resolve, reject) => {
            Web().Lists(Strings.Sites.main.lists.Configuration).Items(itemId).delete().execute(
                () => {
                    console.info(`Deleted Config item ${itemId} !`);
                    resolve();
                },
                (error) => {
                    console.error(`Error deleting Config item: ${formatError(error)}`);
                    reject(error);
                }
            );
        });
    }

    static async requiresSeedItems(): Promise<boolean> {
        const existingCount = await this.getExistingItemCount();
        return existingCount === 0;
    }

    static async ensureSeedItems(): Promise<number> {
        const items: IConfigItem[] = configSeedItems;

        if (!items.length) {
            return 0;
        }

        if (!(await this.requiresSeedItems())) {
            return 0;
        }

        const list = Web().Lists(Strings.Sites.main.lists.Configuration);

        for (const item of items) {
            const payload: IConfigPayload = this.buildPayload(item);
            list.Items().add(payload).batch();
        }

        await new Promise<void>((resolve, reject) => {
            list.execute(
                () => resolve(),
                (error: unknown) => reject(new Error(`Error adding Config seed items: ${formatError(error)}`))
            );
        });

        return items.length;
    }
}
