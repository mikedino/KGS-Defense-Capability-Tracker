import * as React from "react";
import { Icon, IconButton, Link, Text } from "@fluentui/react";
import { ICapabilityItem, IContractItem, IPeoplePickerExtended } from "../../common/props";
import { formatCurrency, formatDate } from "../../common/utils";
import styles from "../../Dct.module.scss";
import { PeoplePersona } from "../../ui/Persona";

export interface IContractDetailCardProps {
    contract: IContractItem;
    eyebrow: string;
    collapsible?: boolean;
    defaultExpanded?: boolean;
    showHeaderSummary?: boolean;
    relatedCapabilities?: ICapabilityItem[];
    documentsContent?: React.ReactNode;
    footerContent?: React.ReactNode;
    onSelectCapability?: (capability: ICapabilityItem) => void;
    onViewCapabilityRelationship?: (capability: ICapabilityItem) => void;
    onEditCapabilityRelationship?: (capability: ICapabilityItem) => void;
    onDeleteCapabilityRelationship?: (capability: ICapabilityItem) => void;
}

const FieldDisplay: React.FC<{ label: string; value?: string }> = ({ label, value }) => (
    <div className={styles.contractDetailField}>
        <span className={styles.contractDetailLabel}>{label}</span>
        <span className={styles.contractDetailValue}>{value || "-"}</span>
    </div>
);

const HeaderFieldDisplay: React.FC<{ label: string; value?: string }> = ({ label, value }) => (
    <span className={styles.contractDetailHeaderField}>
        <span className={styles.contractDetailHeaderLabel}>{label}</span>
        <span className={styles.contractDetailHeaderValue}>{value || "-"}</span>
    </span>
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
    collapsible = false,
    defaultExpanded = true,
    showHeaderSummary = false,
    relatedCapabilities,
    documentsContent,
    footerContent,
    onSelectCapability,
    onViewCapabilityRelationship,
    onEditCapabilityRelationship,
    onDeleteCapabilityRelationship
}) => {
    const [isExpanded, setIsExpanded] = React.useState<boolean>(!collapsible || defaultExpanded);
    const headerContent = (
        <>
            <span className={styles.contractDetailHeaderMain}>
                <span className={styles.contractDetailEyebrow}>{eyebrow}</span>
                <span className={styles.contractDetailTitleRow}>
                    {contract.isFlagged && (
                        <Icon
                            iconName="Flag"
                            className={styles.contractFlagIconLarge}
                            title="Flagged contract"
                            ariaLabel="Flagged contract"
                        />
                    )}
                    <span className={styles.contractDetailTitle}>{contract.Title || "Untitled contract"}</span>
                </span>
            </span>

            {showHeaderSummary && (
                <span className={styles.contractDetailHeaderSummary}>
                    <HeaderFieldDisplay label="Customer Contract Code" value={contract.customerContractCode} />
                    <HeaderFieldDisplay label="Contract Value" value={formatCurrency(contract.contractValue)} />
                    <HeaderFieldDisplay label="Clearance Level" value={contract.clearance} />
                </span>
            )}

            <span className={styles.contractDetailHeaderEnd}>
                <span className={styles.contractDetailId}>{contract.contractId || "No Contract ID"}</span>
                {collapsible && (
                    <Icon
                        iconName={isExpanded ? "ChevronUp" : "ChevronDown"}
                        className={styles.contractDetailChevron}
                        aria-hidden="true"
                    />
                )}
            </span>
        </>
    );

    return (
        <div className={styles.contractDetailCard}>
            {collapsible ? (
                <button
                    type="button"
                    className={`${styles.contractDetailHeader} ${styles.contractDetailHeaderButton}`}
                    aria-expanded={isExpanded}
                    onClick={() => setIsExpanded((expanded) => !expanded)}
                >
                    {headerContent}
                </button>
            ) : (
                <div className={styles.contractDetailHeader}>{headerContent}</div>
            )}

            {isExpanded && (
                <div className={styles.contractDetailBody}>
                    {showHeaderSummary ? (
                        <>
                            <div className={styles.contractDetailGrid}>
                                <FieldDisplay label="Contract Type" value={contract.contractType} />
                                <FieldDisplay label="OG" value={contract.ogTitle} />
                                <FieldDisplay label="LOB" value={contract.lobTitle} />
                                <FieldDisplay label="Customer" value={contract.customer} />
                            </div>

                            <div className={styles.contractDetailGridFive}>
                                <FieldDisplay label="Synonyms" value={contract.synonyms} />
                                <FieldDisplay label="Partner Tag" value={contract.partner} />
                                <FieldDisplay label="Start" value={contract.startDate ? formatDate(contract.startDate) : undefined} />
                                <FieldDisplay label="End" value={contract.endDate ? formatDate(contract.endDate) : undefined} />
                                <FieldDisplay label="Location" value={contract.location} />
                            </div>
                        </>
                    ) : (
                        <>
                            <div className={styles.contractDetailGrid}>
                                <FieldDisplay label="Customer Contract Code" value={contract.customerContractCode} />
                                <FieldDisplay label="Contract Type" value={contract.contractType} />
                                <FieldDisplay label="Contract Value" value={formatCurrency(contract.contractValue)} />
                                <FieldDisplay label="Clearance Level" value={contract.clearance} />
                            </div>

                            <div className={styles.contractDetailGrid}>
                                <FieldDisplay label="OG" value={contract.ogTitle} />
                                <FieldDisplay label="LOB" value={contract.lobTitle} />
                                <FieldDisplay label="Customer" value={contract.customer} />
                                <FieldDisplay label="Synonyms" value={contract.synonyms} />
                            </div>

                            <div className={styles.contractDetailGrid}>
                                <FieldDisplay label="Partner Tag" value={contract.partner} />
                                <FieldDisplay label="Start" value={contract.startDate ? formatDate(contract.startDate) : undefined} />
                                <FieldDisplay label="End" value={contract.endDate ? formatDate(contract.endDate) : undefined} />
                                <FieldDisplay label="Location" value={contract.location} />
                            </div>
                        </>
                    )}

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
                                                    {onEditCapabilityRelationship && (
                                                        <IconButton
                                                            iconProps={{ iconName: "Edit" }}
                                                            className={styles.contractCapabilityInfoButton}
                                                            title="Edit relationship summary"
                                                            ariaLabel={`Edit relationship summary for ${capability.Title}`}
                                                            onClick={() => onEditCapabilityRelationship(capability)}
                                                        />
                                                    )}
                                                    {onDeleteCapabilityRelationship && (
                                                        <IconButton
                                                            iconProps={{ iconName: "Delete" }}
                                                            className={styles.contractCapabilityInfoButton}
                                                            title="Unlink capability"
                                                            ariaLabel={`Unlink ${capability.Title} from this contract`}
                                                            onClick={() => onDeleteCapabilityRelationship(capability)}
                                                        />
                                                    )}
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
            )}
        </div>
    );
};
