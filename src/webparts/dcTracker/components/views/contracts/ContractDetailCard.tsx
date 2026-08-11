import * as React from "react";
import { IconButton, Link, Text } from "@fluentui/react";
import { ICapabilityItem, IContractItem, IPeoplePickerExtended } from "../../common/props";
import { formatCurrency, formatDate } from "../../common/utils";
import styles from "../../Dct.module.scss";
import { PeoplePersona } from "../../ui/Persona";

export interface IContractDetailCardProps {
    contract: IContractItem;
    eyebrow: string;
    relatedCapabilities?: ICapabilityItem[];
    documentsContent?: React.ReactNode;
    footerContent?: React.ReactNode;
    onSelectCapability?: (capability: ICapabilityItem) => void;
    onViewCapabilityRelationship?: (capability: ICapabilityItem) => void;
}

const FieldDisplay: React.FC<{ label: string; value?: string }> = ({ label, value }) => (
    <div className={styles.contractDetailField}>
        <span className={styles.contractDetailLabel}>{label}</span>
        <span className={styles.contractDetailValue}>{value || "-"}</span>
    </div>
);

const PersonDisplay: React.FC<{ label: string; person?: IPeoplePickerExtended }> = ({ label, person }) => (
    <div className={styles.contractDetailField}>
        <span className={styles.contractDetailLabel}>{label}</span>
        {person?.Id ? (
            <PeoplePersona person={person} showDetails={true} fallbackText="Not assigned" />
        ) : (
            <Text styles={{ root: { color: "gray", fontStyle: "italic" } }}>Not assigned</Text>
        )}
    </div>
);

export const ContractDetailCard: React.FC<IContractDetailCardProps> = ({
    contract,
    eyebrow,
    relatedCapabilities,
    documentsContent,
    footerContent,
    onSelectCapability,
    onViewCapabilityRelationship
}) => (
    <div className={styles.contractDetailCard}>
        <div className={styles.contractDetailHeader}>
            <div>
                <span className={styles.contractDetailEyebrow}>{eyebrow}</span>
                <span className={styles.contractDetailTitle}>{contract.Title || "Untitled contract"}</span>
            </div>
            <span className={styles.contractDetailId}>{contract.contractId || "No Contract ID"}</span>
        </div>

        <div className={styles.contractDetailBody}>
            <div className={styles.contractDetailGridThree}>
                <FieldDisplay label="Customer Contract Code" value={contract.customerContractCode} />
                <FieldDisplay label="Contract Type" value={contract.contractType} />
                <FieldDisplay label="Contract Value" value={formatCurrency(contract.contractValue)} />
            </div>

            <div className={styles.contractDetailGridThree}>
                <FieldDisplay label="OG" value={contract.ogTitle} />
                <FieldDisplay label="LOB" value={contract.lobTitle} />
                <FieldDisplay label="Customer" value={contract.customer} />
            </div>

            <div className={styles.contractDetailGrid}>
                <FieldDisplay label="Partner Tag" value={contract.partner} />
                <FieldDisplay label="Start" value={contract.startDate ? formatDate(contract.startDate) : undefined} />
                <FieldDisplay label="End" value={contract.endDate ? formatDate(contract.endDate) : undefined} />
            </div>

            <div className={styles.contractDetailGridWide}>
                <div className={styles.contractDetailField}>
                    <span className={styles.contractDetailLabel}>Contract Info Link/URL</span>
                    {contract.infoLink ? (
                        <Link className={styles.contractDetailValue} href={contract.infoLink} target="_blank" rel="noopener noreferrer">
                            {contract.infoLink}
                        </Link>
                    ) : (
                        <span className={styles.contractDetailValue}>-</span>
                    )}
                </div>

                <PersonDisplay label="KGS Contract Project Manager" person={contract.contractPm} />
            </div>

            {(relatedCapabilities || documentsContent) && (
                <div className={`${styles.contractRelatedGrid} ${relatedCapabilities && documentsContent ? styles.contractRelatedGridSplit : ""}`}>
                    {documentsContent && (
                        <div className={styles.contractDocumentsSection}>
                            {documentsContent}
                        </div>
                    )}

                    {relatedCapabilities && (
                        <div className={styles.contractCapabilitiesSection}>
                            <span className={styles.contractCapabilitiesTitle}>
                                Related Capabilities ({relatedCapabilities.length})
                            </span>
                            {relatedCapabilities.length ? (
                                <div className={styles.contractCapabilityLinks}>
                                    {relatedCapabilities.map((capability) => (
                                        <div key={capability.Id} className={styles.contractCapabilityRow}>
                                            <Link
                                                className={styles.listTitleLink}
                                                onClick={(ev) => {
                                                    ev.preventDefault();
                                                    onSelectCapability?.(capability);
                                                }}
                                            >
                                                {capability.Title}
                                            </Link>
                                            {onViewCapabilityRelationship && (
                                                <IconButton
                                                    iconProps={{ iconName: "Info" }}
                                                    className={styles.contractCapabilityInfoButton}
                                                    title="View relationship summary"
                                                    ariaLabel={`View relationship summary for ${capability.Title}`}
                                                    onClick={() => onViewCapabilityRelationship(capability)}
                                                />
                                            )}
                                        </div>
                                    ))}
                                </div>
                            ) : (
                                <Text styles={{ root: { color: "gray", fontStyle: "italic" } }}>
                                    No capabilities are linked to this contract.
                                </Text>
                            )}
                        </div>
                    )}
                </div>
            )}

            {footerContent}
        </div>
    </div>
);
