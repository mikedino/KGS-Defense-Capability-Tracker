import * as React from "react";
import { Label, Link, Stack, Text } from "@fluentui/react";
import { IContractItem, IPeoplePickerExtended } from "../../common/props";
import styles from "../../Dct.module.scss";
import { formatDate } from "../../common/utils";
import { PeoplePersona } from "../../ui/Persona";

export interface IContractInfoProps {
    contracts: IContractItem[];
    isLoading?: boolean;
}

const contractDetailsCardStyle: React.CSSProperties = {
    flex: "1 1 620px",
    minWidth: 0
};

const FieldDisplay: React.FC<{ label: string; value?: string }> = ({ label, value }) => (
    <Stack style={{ minWidth: 180, flexGrow: 1 }}>
        <Label>{label}</Label>
        <Text>{value || "—"}</Text>
    </Stack>
);

const PersonDisplay: React.FC<{ label: string; person?: IPeoplePickerExtended }> = ({ label, person }) => (
    <Stack style={{ minWidth: 220, flexGrow: 1 }} tokens={{ childrenGap: 4 }}>
        <Label>{label}</Label>
        {person?.Id ? (
            <PeoplePersona person={person} showDetails={true} fallbackText="Not assigned" />
        ) : (
            <Text styles={{ root: { color: "gray", fontStyle: "italic" } }}>Not assigned</Text>
        )}
    </Stack>
);

export const ContractInfo: React.FC<IContractInfoProps> = ({ contracts, isLoading }) => {
    return (
        <Stack tokens={{ childrenGap: 16 }}>
            <Stack horizontal wrap tokens={{ childrenGap: 16 }} styles={{ root: { alignItems: "stretch" } }}>
                <Stack tokens={{ childrenGap: 16 }} className={styles.detailCard} style={contractDetailsCardStyle}>
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

                    {!isLoading && contracts.map((contract) => (
                        <Stack key={contract.Id} tokens={{ childrenGap: 16 }} styles={{ root: { borderBottom: "1px solid #edebe9", paddingBottom: 16 } }}>
                            <Stack horizontal wrap tokens={{ childrenGap: 24 }}>
                                <FieldDisplay label="Contract ID" value={contract.contractId} />
                                <FieldDisplay label="Contract Title" value={contract.Title} />
                            </Stack>

                            <Stack horizontal wrap tokens={{ childrenGap: 24 }}>
                                <FieldDisplay label="Customer Contract Code" value={contract.customerContractCode} />
                                <FieldDisplay label="OG" value={contract.ogTitle} />
                                <FieldDisplay label="LOB" value={contract.lobTitle} />
                                <FieldDisplay label="Customer" value={contract.customer} />
                                <FieldDisplay label="Relevant Partner Tag" value={contract.partner} />
                            </Stack>

                            <Stack horizontal wrap tokens={{ childrenGap: 24 }}>
                                <FieldDisplay label="Start" value={contract.startDate ? formatDate(contract.startDate) : undefined} />
                                <FieldDisplay label="End" value={contract.endDate ? formatDate(contract.endDate) : undefined} />
                                <Stack style={{ minWidth: 240, flexGrow: 1 }}>
                                    <Label>Contract Info Link/URL</Label>
                                    {contract.infoLink ? (
                                        <Link href={contract.infoLink} target="_blank" rel="noopener noreferrer">
                                            {contract.infoLink}
                                        </Link>
                                    ) : (
                                        <Text>—</Text>
                                    )}
                                </Stack>
                            </Stack>

                            <Stack horizontal wrap tokens={{ childrenGap: 24 }}>
                                <PersonDisplay label="KGS Contract Project Manager" person={contract.contractPm} />
                            </Stack>
                        </Stack>
                    ))}
                </Stack>
            </Stack>
        </Stack>
    );
};
