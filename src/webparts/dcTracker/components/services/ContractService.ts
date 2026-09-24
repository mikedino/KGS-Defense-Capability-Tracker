import { ICapabilityContractDraft, IContractItem, ILookupItem } from "../common/props";
import { DataSource } from "../data/ds";
import { Web } from "gd-sprest-bs";
import Strings from "../common/strings";
import { encodeListName, formatError } from "../common/utils";
import { DocumentService } from "./DocumentService";
import { ContractCapabilitySummaryService } from "./ContractCapabilitySummaryService";

export class ContractService {
    private static getListItemType(): string {
        return `SP.Data.${encodeListName(Strings.Sites.main.lists.Contracts)}ListItem`;
    }

    private static normalizeDateValue(value?: string): string | null {
        return value && value !== "" ? value : null;
    }

    private static normalizeKey(value?: string): string {
        return (value ?? "").trim().toLowerCase();
    }

    private static toCapabilityLookup(capabilityId: number, capabilityTitle?: string): ILookupItem {
        return {
            Id: capabilityId,
            Title: capabilityTitle ?? ""
        };
    }

    private static mergeCapabilityLookup(
        capabilities: ILookupItem[] = [],
        capabilityId: number,
        capabilityTitle?: string
    ): ILookupItem[] {
        const existing = capabilities.filter((capability) => !!capability?.Id);

        if (existing.some((capability) => capability.Id === capabilityId)) {
            return existing.map((capability) =>
                capability.Id === capabilityId && capabilityTitle
                    ? { ...capability, Title: capabilityTitle }
                    : capability
            );
        }

        return [...existing, ContractService.toCapabilityLookup(capabilityId, capabilityTitle)];
    }

    private static removeCapabilityLookup(capabilities: ILookupItem[] = [], capabilityId: number): ILookupItem[] {
        return capabilities.filter((capability) => capability?.Id !== capabilityId);
    }

    static getCapabilityLookups(contract?: IContractItem): ILookupItem[] {
        return contract?.capability?.results ?? [];
    }

    static isLinkedToCapability(contract: IContractItem | undefined, capabilityId: number): boolean {
        return ContractService.getCapabilityLookups(contract).some((capability) => capability.Id === capabilityId);
    }

    static getContractIdentityKey(contract?: Pick<IContractItem, "Id" | "contractId" | "customerContractCode" | "Title">): string {
        if (!contract) return "";

        const contractId = ContractService.normalizeKey(contract.contractId);
        if (contractId) return `contractId:${contractId}`;

        const customerContractCode = ContractService.normalizeKey(contract.customerContractCode);
        if (customerContractCode) return `customerContractCode:${customerContractCode}`;

        const title = ContractService.normalizeKey(contract.Title);
        if (title) return `title:${title}`;

        return contract.Id > 0 ? `id:${contract.Id}` : "";
    }

    static contractsMatch(
        a?: Pick<IContractItem, "Id" | "contractId" | "customerContractCode" | "Title">,
        b?: Pick<IContractItem, "Id" | "contractId" | "customerContractCode" | "Title">
    ): boolean {
        if (!a || !b) return false;
        if (a.Id > 0 && b.Id > 0 && a.Id === b.Id) return true;

        const aContractId = ContractService.normalizeKey(a.contractId);
        const bContractId = ContractService.normalizeKey(b.contractId);
        if (aContractId && bContractId && aContractId === bContractId) return true;

        const aCode = ContractService.normalizeKey(a.customerContractCode);
        const bCode = ContractService.normalizeKey(b.customerContractCode);
        if (aCode && bCode && aCode === bCode) return true;

        const aTitle = ContractService.normalizeKey(a.Title);
        const bTitle = ContractService.normalizeKey(b.Title);
        return !!aTitle && !!bTitle && aTitle === bTitle;
    }

    static findMatchingContract<T extends Pick<IContractItem, "Id" | "contractId" | "customerContractCode" | "Title">>(
        contract: Pick<IContractItem, "Id" | "contractId" | "customerContractCode" | "Title">,
        contracts: T[] = DataSource.Contracts as unknown as T[]
    ): T | undefined {
        return contracts.find((candidate) => ContractService.contractsMatch(contract, candidate));
    }

    private static buildPayload(item: IContractItem): Record<string, unknown> {
        return {
            __metadata: { type: ContractService.getListItemType() },
            Title: item.Title,
            synonyms: item.synonyms,
            capabilityId: { results: ContractService.getCapabilityLookups(item).map(c => c.Id) },
            contractId: item.contractId,
            customerContractCode: item.customerContractCode,
            isFlagged: item.isFlagged ?? false,
            clearance: item.clearance,
            contractType: item.contractType,
            contractValue: item.contractValue ?? 0,
            customer: item.customer,
            ogTitle: item.ogTitle,
            lobTitle: item.lobTitle,
            startDate: ContractService.normalizeDateValue(item.startDate),
            endDate: ContractService.normalizeDateValue(item.endDate),
            contractPmId: item.contractPm?.Id ?? null,
            partner: item.partner,
            infoLink: item.infoLink,
            city: item.city,
            state: item.state,
            country: (item.country ?? "US").trim().toUpperCase().slice(0, 2)
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
                async (resp) => {
                    if (resp && resp.Id) {
                        // Each new DCT contract gets its own document folder so filenames remain unique per contract.
                        await DocumentService.createContractFolder(resp.Id).catch((error) =>
                            console.warn(`Contract ${resp.Id} was created, but its document folder could not be created.`, error)
                        );
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

    static async linkCapability(
        contract: IContractItem,
        capabilityId: number,
        capabilityTitle: string,
        summary: string
    ): Promise<IContractItem> {
        const updated = await ContractService.edit({
            ...contract,
            capability: {
                results: ContractService.mergeCapabilityLookup(
                    ContractService.getCapabilityLookups(contract),
                    capabilityId,
                    capabilityTitle
                )
            }
        });

        await ContractCapabilitySummaryService.upsertForRelationship(
            updated,
            capabilityId,
            capabilityTitle,
            { capabilitySummary: summary }
        );
        return updated;
    }

    static updateCapabilitySummary(
        contract: IContractItem,
        capabilityId: number,
        capabilityTitle: string,
        summary: string
    ): Promise<unknown> {
        return ContractCapabilitySummaryService.upsertForRelationship(
            contract,
            capabilityId,
            capabilityTitle,
            { capabilitySummary: summary }
        );
    }

    static async unlinkCapability(contract: IContractItem, capabilityId: number): Promise<IContractItem> {
        const updated = await ContractService.edit({
            ...contract,
            capability: {
                results: ContractService.removeCapabilityLookup(
                    ContractService.getCapabilityLookups(contract),
                    capabilityId
                )
            }
        });

        await ContractCapabilitySummaryService.deleteForRelationship(contract.Id, capabilityId);
        return updated;
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
        capabilityTitle: string,
        contracts: ICapabilityContractDraft[],
        deletedContractIds: number[] = []
    ): Promise<IContractItem[]> {
        const workingContracts = [...DataSource.Contracts];
        const savedContracts: IContractItem[] = [];

        for (const itemId of deletedContractIds) {
            const existing = workingContracts.find((contract) => contract.Id === itemId);
            if (!existing) continue;
            if (contracts.some((draft) => ContractService.contractsMatch(draft, existing))) continue;

            await ContractCapabilitySummaryService.deleteForRelationship(existing.Id, capabilityId);

            const remainingCapabilities = ContractService.removeCapabilityLookup(
                ContractService.getCapabilityLookups(existing),
                capabilityId
            );

            if (remainingCapabilities.length) {
                const updated = await ContractService.edit({
                    ...existing,
                    capability: { results: remainingCapabilities }
                });
                savedContracts.push(updated);
                const index = workingContracts.findIndex((contract) => contract.Id === updated.Id);
                if (index >= 0) workingContracts[index] = updated;
            } else {
                await ContractService.delete(existing.Id);
                const index = workingContracts.findIndex((contract) => contract.Id === existing.Id);
                if (index >= 0) workingContracts.splice(index, 1);
            }
        }

        for (const draft of contracts) {
            const existing = ContractService.findMatchingContract(draft, workingContracts);
            const baseContract = existing ?? draft;
            const capabilityResults = ContractService.mergeCapabilityLookup(
                ContractService.getCapabilityLookups(baseContract),
                capabilityId
            );

            const relationship: IContractItem = {
                ...baseContract,
                ...draft,
                Id: existing?.Id ?? draft.Id,
                capability: { results: capabilityResults }
            };

            if (existing?.Id || relationship.Id > 0) {
                const updated = await ContractService.edit(relationship);
                await ContractCapabilitySummaryService.upsertForRelationship(updated, capabilityId, capabilityTitle, draft);
                savedContracts.push(updated);
                const index = workingContracts.findIndex((contract) => contract.Id === updated.Id);
                if (index >= 0) {
                    workingContracts[index] = updated;
                } else {
                    workingContracts.push(updated);
                }
            } else {
                const created = await ContractService.create(relationship);
                await ContractCapabilitySummaryService.upsertForRelationship(created, capabilityId, capabilityTitle, draft);
                savedContracts.push(created);
                workingContracts.push(created);
            }
        }

        return savedContracts;
    }

    static async removeCapabilityFromAllContracts(capabilityId: number): Promise<void> {
        const relatedContracts = DataSource.Contracts.filter((contract) =>
            ContractService.isLinkedToCapability(contract, capabilityId)
        );

        for (const contract of relatedContracts) {
            const remainingCapabilities = ContractService.removeCapabilityLookup(
                ContractService.getCapabilityLookups(contract),
                capabilityId
            );

            await ContractCapabilitySummaryService.deleteForRelationship(contract.Id, capabilityId);

            if (remainingCapabilities.length) {
                await ContractService.edit({
                    ...contract,
                    capability: { results: remainingCapabilities }
                });
            } else {
                await ContractService.delete(contract.Id);
            }
        }
    }
}
