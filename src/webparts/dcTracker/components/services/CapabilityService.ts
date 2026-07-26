import { ICapabilityItem } from "../common/props";
import { Web } from "gd-sprest";
import Strings from "../common/strings";
import { formatError, encodeListName } from "../common/utils";
import { DataSource } from "../data/ds";
import { normalizeOppNetTag, normalizePastPerformanceTag, normalizeProposalTag, parseJsonTagField, serializeJsonTagField } from "../common/tagUtils";

interface ICapabilityPayload extends Omit<ICapabilityItem, "Id"> {
    __metadata: { type: string };
    primaryPocId: number | undefined;
    stakeholdersId: { results: number[] };
    oppNetTagsJson: string;
    pastPerformanceTagsJson: string;
    proposalTagsJson: string;
}

export class CapabilityService {
    
    private static getListItemType(): string {
        return `SP.Data.${encodeListName(Strings.Sites.main.lists.Capabilities)}ListItem`;
    }

    private static buildPayload(item: ICapabilityItem): ICapabilityPayload {
        return {
            __metadata: { type: CapabilityService.getListItemType() },
            Title: item.Title,
            description: item.description,
            capabilities: item.capabilities,
            link: item.link,
            capStatus: item.capStatus,
            notes: item.notes,
            solutionType: item.solutionType,
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
            primaryPocId: item.primaryPoc?.Id,
            stakeholdersId: { results: item.stakeholders?.results.map((person) => person.Id) ?? [] },
            oppNetTagsJson: serializeJsonTagField(item.oppNetTags?.map(normalizeOppNetTag)),
            pastPerformanceTagsJson: serializeJsonTagField(item.pastPerformanceTags?.map(normalizePastPerformanceTag)),
            proposalTagsJson: serializeJsonTagField(item.proposalTags?.map(normalizeProposalTag))
        };
    }

    private static loadCapById(itemId: number): ICapabilityItem {
        const item = Web().Lists(Strings.Sites.main.lists.Capabilities).Items(itemId.toString())
            .query({
                Select: DataSource.capabilityQuerySelect,
                Expand: DataSource.capabilityQueryExpand
            })
            .executeAndWait();

        const capability = item as unknown as ICapabilityItem;

        return {
            ...capability,
            oppNetTags: parseJsonTagField(capability.oppNetTagsJson),
            pastPerformanceTags: parseJsonTagField(capability.pastPerformanceTagsJson),
            proposalTags: parseJsonTagField(capability.proposalTagsJson)
        };
    }

    static create(item: ICapabilityItem): Promise<ICapabilityItem> {
        return new Promise<ICapabilityItem>((resolve, reject) => {
            Web().Lists(Strings.Sites.main.lists.Capabilities).Items().add(CapabilityService.buildPayload(item)).execute(
                (resp) => {
                    if (resp && resp.existsFl && resp.Id) {
                        const newItem: ICapabilityItem = CapabilityService.loadCapById(resp.Id);
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
            Web().Lists(Strings.Sites.main.lists.Capabilities).Items().getById(item.Id).update(CapabilityService.buildPayload(item)).execute(
                (resp) => {
                    if (resp && resp.existsFl) {
                        const updatedItem: ICapabilityItem = CapabilityService.loadCapById(item.Id);
                        resolve(updatedItem);
                        return;
                    }

                    reject("Item was edited but there was a problem refreshing the data. Please refresh manually.");
                },
                (error) => {
                    const err: string = formatError(error);
                    console.error(`Error updating Capability ${err}`);
                    reject(error);
                }
            );
        });
    }


    static delete(itemId: number): Promise<void> {
        return new Promise<void>((resolve, reject) => {
            Web().Lists(Strings.Sites.main.lists.Capabilities).Items(itemId).delete().execute(
                () => {
                    console.info(`Deleted Capability item ${itemId} !`)
                    resolve();
                },
                (error) => {
                    const err = formatError(error);
                    console.error(`Error deleting Capability item: ${err}`);
                    reject(error);
                }
            )
        })
    }

}
