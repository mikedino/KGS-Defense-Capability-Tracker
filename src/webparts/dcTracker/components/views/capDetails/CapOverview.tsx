import * as React from "react";
import { Label, Stack, Text } from "@fluentui/react";
import { ICapabilityItem } from "../../common/props";
import { DataSource } from "../../data/ds";
import styles from "../../Dct.module.scss";
import { PeoplePersona } from "../../ui/Persona";
import { ContractService } from "../../services/ContractService";

export interface ICapabilityOverviewProps {
    capState: ICapabilityItem;
    rightContent?: React.ReactNode;
}

const MultilineDisplay: React.FC<{ label: string; value?: string; emptyText?: string }> = ({
    label,
    value,
    emptyText = "Not provided"
}) => (
    <Stack>
        <Label>{label}</Label>
        {value?.trim() ? (
            <Text styles={{ root: { whiteSpace: "pre-wrap" } }}>{value}</Text>
        ) : (
            <Text styles={{ root: { color: "gray", fontStyle: "italic" } }}>{emptyText}</Text>
        )}
    </Stack>
);

export const CapabilityOverview: React.FC<ICapabilityOverviewProps> = ({ capState, rightContent }) => {
    const relatedContracts = DataSource.Contracts
        .filter((contract) => ContractService.isLinkedToCapability(contract, capState.Id))
        .map((contract) => contract.Title)
        .filter(Boolean);

    const stakeholders = capState.stakeholders?.results ?? [];

    return (
        <Stack tokens={{ childrenGap: 16 }}>
            <Stack horizontal wrap tokens={{ childrenGap: 16 }} className={styles.detailCard}>
                <Stack grow styles={{ root: { flexGrow: 1, flexBasis: "47%", minWidth: 0 } }} tokens={{ childrenGap: 16 }}>
                    <MultilineDisplay label="Capability Description" value={capState.description} />
                    <MultilineDisplay label="Technical Capabilities" value={capState.capabilities} />

                    <Stack>
                        <Label>Link/URL</Label>
                        {capState.link ? (
                            <Text>
                                <a href={capState.link} className={styles.hyperlink} target="_blank" rel="noopener noreferrer">
                                    {capState.link}
                                </a>
                            </Text>
                        ) : (
                            <Text styles={{ root: { color: "gray", fontStyle: "italic" } }}>No link provided</Text>
                        )}
                    </Stack>

                    <Stack horizontal wrap tokens={{ childrenGap: 24 }}>
                        <Stack style={{ width: 200 }}>
                            <Label>Capability Status</Label>
                            <Text>{capState.capStatus || "Not set"}</Text>
                        </Stack>

                        <Stack style={{ minWidth: 240 }}>
                            <Label>Contracts</Label>
                            <Text>{relatedContracts.length ? relatedContracts.join(", ") : "Not assigned"}</Text>
                        </Stack>
                    </Stack>

                    {/* PEOPLE ROW */}
                    <Stack horizontal wrap tokens={{ childrenGap: 24 }}>
                        <Stack tokens={{ childrenGap: 4 }} style={{ width: 200 }}>
                            <Label>Primary POC</Label>
                            {capState.primaryPoc?.Id ? (
                                <PeoplePersona person={capState.primaryPoc} showDetails={true} fallbackText="Not assigned" />
                            ) : (
                                <Text styles={{ root: { color: "gray", fontStyle: "italic" } }}>Not assigned</Text>
                            )}
                        </Stack>
                        <Stack tokens={{ childrenGap: 4 }} style={{ minWidth: 240 }}>
                            <Label>Stakeholder(s)</Label>
                            <Stack horizontal wrap tokens={{ childrenGap: 4, padding: "0 2px" }}>
                                {stakeholders.length > 0 ? (
                                    stakeholders.map((p) => <PeoplePersona key={p.Id} person={p} />)
                                ) : (
                                    <Text styles={{ root: { color: "gray", fontStyle: "italic" } }}>None identified</Text>
                                )}
                            </Stack>
                        </Stack>
                    </Stack>

                    <MultilineDisplay label="Additional Notes" value={capState.notes} emptyText="None provided" />
                </Stack>

                {rightContent && (
                    <Stack styles={{ root: { flex: "1 1 520px", minWidth: 320, maxWidth: "100%" } }} tokens={{ childrenGap: 8 }}>
                        {rightContent}
                    </Stack>
                )}
            </Stack>
        </Stack>
    );
};
