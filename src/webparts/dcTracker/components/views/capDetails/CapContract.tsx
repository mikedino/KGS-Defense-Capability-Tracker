import * as React from "react";
import { Stack, Text } from "@fluentui/react";
import { ICapabilityItem, IContractDocumentItem, IContractItem } from "../../common/props";
import { formatError } from "../../common/utils";
import { DataSource } from "../../data/ds";
import styles from "../../Dct.module.scss";
import { ContractDetailCard } from "../contracts/ContractDetailCard";
import { ContractDocumentsPanel } from "../contracts/ContractDocumentsPanel";

export interface IContractInfoProps {
    capability: ICapabilityItem;
    contracts: IContractItem[];
    isLoading?: boolean;
}

export const ContractInfo: React.FC<IContractInfoProps> = ({ capability, contracts, isLoading }) => {
    const [contractDocumentsById, setContractDocumentsById] = React.useState<Map<number, IContractDocumentItem[]>>(new Map());

    React.useEffect(() => {
        if (!contracts.length) {
            setContractDocumentsById(new Map());
            return;
        }

        // Load documents for each visible supporting contract so the shared detail card can show its document list.
        Promise.all(
            contracts.map((contract) =>
                DataSource.getDocumentsByContract(contract.Id)
                    .then((documents) => ({ contractId: contract.Id, documents }))
            )
        )
            .then((results) => {
                const nextDocumentsById = new Map<number, IContractDocumentItem[]>();
                results.forEach((result) => nextDocumentsById.set(result.contractId, result.documents));
                setContractDocumentsById(nextDocumentsById);
            })
            .catch((error) => console.error(`Error loading supporting contract documents: ${formatError(error)}`));
    }, [contracts]);

    return (
        <Stack tokens={{ childrenGap: 16 }} className={styles.contractDetailList}>
            {isLoading && (
                <Text styles={{ root: { color: "gray", fontStyle: "italic" } }}>
                    Loading contract details...
                </Text>
            )}

            {!isLoading && !contracts.length && (
                <Text styles={{ root: { color: "gray", fontStyle: "italic" } }}>
                    No contracts assigned.
                </Text>
            )}

            {!isLoading && contracts.map((contract, index) => (
                <ContractDetailCard
                    key={contract.Id}
                    contract={contract}
                    eyebrow={`Contract ${index + 1} of ${contracts.length}`}
                    collapsible={true}
                    defaultExpanded={false}
                    showHeaderSummary={true}
                    documentsContent={(
                        <ContractDocumentsPanel
                            documents={contractDocumentsById.get(contract.Id) ?? []}
                            canEdit={false}
                            showAddButton={false}
                        />
                    )}
                    footerContent={(
                        <div className={styles.contractRelationshipSummary}>
                            <span className={styles.contractRelationshipSummaryTitle}>Capability Summary</span>
                            <span className={styles.contractRelationshipSummaryText}>
                                {DataSource.getContractCapabilitySummary(contract.Id, capability.Id)?.summary || "No contract-specific summary has been added."}
                            </span>
                        </div>
                    )}
                />
            ))}
        </Stack>
    );
};
