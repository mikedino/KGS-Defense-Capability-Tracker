import { Web } from "gd-sprest-bs";
import { ICapabilityContractDraft, IContractCapabilitySummaryItem, IContractItem } from "../common/props";
import Strings from "../common/strings";
import { encodeListName, formatError } from "../common/utils";
import { DataSource } from "../data/ds";

export class ContractCapabilitySummaryService {
    private static getListItemType(): string {
        return `SP.Data.${encodeListName(Strings.Sites.main.lists.ContractCapabilitySummary)}ListItem`;
    }

    // Keep the generated title readable while still making the relationship obvious in list views.
    static buildTitle(contract: IContractItem, capabilityTitle: string): string {
        const contractKey = contract.contractId || contract.customerContractCode || contract.Title || `${contract.Id}`;
        return `${contractKey}-${capabilityTitle}`;
    }

    private static buildPayload(
        contract: IContractItem,
        capabilityId: number,
        capabilityTitle: string,
        summary?: string,
        pocId?: number
    ): Record<string, unknown> {
        return {
            __metadata: { type: ContractCapabilitySummaryService.getListItemType() },
            Title: ContractCapabilitySummaryService.buildTitle(contract, capabilityTitle),
            contractId: contract.Id,
            capabilityId,
            summary: summary ?? "",
            pocId: pocId ?? null
        };
    }

    private static findExisting(contractId: number, capabilityId: number): IContractCapabilitySummaryItem | undefined {
        return DataSource.getContractCapabilitySummary(contractId, capabilityId);
    }

    // Create or update the one summary row that belongs to a contract-capability relationship.
    static upsertForRelationship(
        contract: IContractItem,
        capabilityId: number,
        capabilityTitle: string,
        draft?: Pick<ICapabilityContractDraft, "capabilitySummary" | "capabilitySummaryPoc">
    ): Promise<IContractCapabilitySummaryItem> {
        return new Promise<IContractCapabilitySummaryItem>((resolve, reject) => {
            const existing = ContractCapabilitySummaryService.findExisting(contract.Id, capabilityId);
            const payload = ContractCapabilitySummaryService.buildPayload(
                contract,
                capabilityId,
                capabilityTitle,
                draft?.capabilitySummary,
                draft?.capabilitySummaryPoc?.Id
            );

            const onSuccess = async (): Promise<void> => {
                await DataSource.refreshContractCapabilitySummaries();
                const refreshed = DataSource.getContractCapabilitySummary(contract.Id, capabilityId);
                if (refreshed) {
                    resolve(refreshed);
                    return;
                }

                reject(new Error("Summary was saved but could not be refreshed."));
            };

            const onError = (error: unknown): void => {
                reject(new Error(`Error saving Contract Capability Summary: ${formatError(error)}`));
            };

            if (existing?.Id) {
                Web().Lists(Strings.Sites.main.lists.ContractCapabilitySummary).Items(existing.Id).update(payload).execute(
                    () => { onSuccess().catch(onError); },
                    onError
                );
                return;
            }

            Web().Lists(Strings.Sites.main.lists.ContractCapabilitySummary).Items().add(payload).execute(
                () => { onSuccess().catch(onError); },
                onError
            );
        });
    }

    // Remove the summary row when a capability is unlinked from a contract.
    static deleteForRelationship(contractId: number, capabilityId: number): Promise<void> {
        const existing = ContractCapabilitySummaryService.findExisting(contractId, capabilityId);
        if (!existing?.Id) return Promise.resolve();

        return new Promise<void>((resolve, reject) => {
            Web().Lists(Strings.Sites.main.lists.ContractCapabilitySummary).Items(existing.Id).delete().execute(
                async () => {
                    await DataSource.refreshContractCapabilitySummaries();
                    resolve();
                },
                (error) => reject(new Error(`Error deleting Contract Capability Summary: ${formatError(error)}`))
            );
        });
    }
}
