import * as React from "react";
import { Label, Stack, Text } from "@fluentui/react";
import { ICapabilityItem } from "../../common/props";
import { DataSource } from "../../data/ds";
import styles from "../../Dct.module.scss";

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
        .filter((contract) => contract.capability?.Id === capState.Id)
        .map((contract) => contract.Title)
        .filter(Boolean);

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
                        <Stack style={{ width: 220 }}>
                            <Label>Capability Status</Label>
                            <Text>{capState.capStatus || "Not set"}</Text>
                        </Stack>

                        <Stack style={{ width: 260 }}>
                            <Label>Contracts</Label>
                            <Text>{relatedContracts.length ? relatedContracts.join(", ") : "Not assigned"}</Text>
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
