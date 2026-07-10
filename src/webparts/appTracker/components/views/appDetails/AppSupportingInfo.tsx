import * as React from "react";
import { Stack, Text, Label } from "@fluentui/react";
import styles from "../../AppTracker.module.scss";
import { formatDate } from "../../utils";
import { IApplicationItem } from "../../data/props";

interface ISupportingInfoProps {
    appState: IApplicationItem;
}

export const SupportingInfo: React.FC<ISupportingInfoProps> = ({ appState }) => {

    return (
        <Stack tokens={{ childrenGap: 16 }}>
            <Stack horizontal wrap tokens={{ childrenGap: 8 }} className={styles.detailCard}>

                <Stack grow styles={{ root: { flexBasis: "100%", minWidth: 0 } }} tokens={{ childrenGap: 16 }}>

                    {/* Row 1 */}
                    <Stack horizontal wrap tokens={{ childrenGap: 24 }}>
                        <Stack style={{ width: 200 }}>
                            <Label>App Status</Label>
                            <Text>{appState.appStatus || "—"}</Text>
                        </Stack>

                        <Stack style={{ width: 200 }}>
                            <Label>App Launch Date</Label>
                            <Text>{formatDate(appState.appLaunchDate)}</Text>
                        </Stack>
                    </Stack>


                    {/* Row 2 */}
                    <Stack horizontal wrap tokens={{ childrenGap: 24 }}>
                        <Stack style={{ width: 220 }}>
                            <Label>Platform</Label>
                            <Text>{appState.platform || "—"}</Text>
                        </Stack>

                        <Stack style={{ width: 220 }}>
                            <Label>Hosting Environment</Label>
                            <Text>{appState.environment || "—"}</Text>
                        </Stack>

                        <Stack style={{ width: 220 }}>
                            <Label>Connectivity</Label>
                            <Text>{appState.connectivity || "—"}</Text>
                        </Stack>

                    </Stack>

                    {/* Synonyms */}
                    <Stack>
                        <Label>Synonyms</Label>
                        <Text styles={{ root: { whiteSpace: "pre-wrap" } }}>
                            {appState.synonyms || "None provided"}
                        </Text>
                    </Stack>

                    <Stack>
                        <Label>Integrations (Input)</Label>
                        <Text styles={{ root: { whiteSpace: "pre-wrap" } }}>
                            {appState.integrationsInput || "None documented"}
                        </Text>
                    </Stack>

                    <Stack>
                        <Label>Integrations (Output)</Label>
                        <Text styles={{ root: { whiteSpace: "pre-wrap" } }}>
                            {appState.integrationsOutput || "None documented"}
                        </Text>
                    </Stack>

                </Stack>

            </Stack>
        </Stack>
    );
};
