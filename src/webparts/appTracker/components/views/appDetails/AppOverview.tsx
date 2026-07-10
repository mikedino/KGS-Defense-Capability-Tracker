import * as React from "react";
import { Stack, Text, Label } from "@fluentui/react";
import styles from "../../AppTracker.module.scss";
import { IApplicationItem } from "../../data/props";
import { PeoplePersona } from "../../ui/Persona";

export interface IApplicationOverviewProps {
    appState: IApplicationItem;
    // Optional right-side content hook (screenshots, preview, KPI panel, etc.)
    rightContent?: React.ReactNode;
}

/** Renders SPO rich text HTML (string) */
const RichTextFieldDisplay: React.FC<{ label: string; html?: string; emptyText?: string }> = ({
    label,
    html,
    emptyText = "—"
}) => {
    const hasHtml = !!(html && html.trim());

    return (
        <Stack>
            <Label>{label}</Label>
            {hasHtml ? (
                <div
                    // If you sanitize, sanitize before passing into this component.
                    dangerouslySetInnerHTML={{ __html: html as string }}
                />
            ) : (
                <Text styles={{ root: { color: "gray", fontStyle: "italic" } }}>{emptyText}</Text>
            )}
        </Stack>
    );
};

export const ApplicationOverview: React.FC<IApplicationOverviewProps> = ({ appState, rightContent }) => {

    const stakeholders = appState.stakeholders?.results ?? [];

    return (
        <Stack tokens={{ childrenGap: 16 }}>
            <Stack horizontal wrap tokens={{ childrenGap: 16 }} className={styles.detailCard}>
                {/******** Left content - 48% ********/}
                <Stack grow styles={{ root: { flexGrow: 1, flexBasis: "47%", minWidth: 0 } }} tokens={{ childrenGap: 16 }}>
                    {/* Description (plain text) */}
                    <Stack>
                        <Label>Description</Label>
                        {appState.description ? (
                            <Text styles={{ root: { whiteSpace: "pre-wrap" } }}>{appState.description}</Text>
                        ) : (
                            <Text styles={{ root: { color: "gray", fontStyle: "italic" } }}>no description provided</Text>
                        )}
                    </Stack>

                    {/* Rich text fields (HTML strings) */}
                    {/* <RichTextFieldDisplay label="Highlights" html={appState.highlights} emptyText="no highlights provided" /> */}
                    <RichTextFieldDisplay label="Core Capabilities & Relevant Software" html={appState.relatedInfo} emptyText="none identified" />

                    {/* App URL */}
                    <Stack>
                        <Label>App URL</Label>
                        {appState.appUrl ? (
                            <Text>
                                <a href={appState.appUrl} className={styles.hyperlink} target="_blank" rel="noopener noreferrer">
                                    {appState.appUrl}
                                </a>
                            </Text>
                        ) : (
                            <Text styles={{ root: { color: "gray", fontStyle: "italic" } }}>no link provided</Text>
                        )}
                    </Stack>

                    {/* Metadata row */}
                    <Stack horizontal wrap tokens={{ childrenGap: 24 }}>
                        <Stack style={{ width: 240 }}>
                            <Label>Managing Group</Label>
                            <Text>{appState.managingGroup ?? "Not set"}</Text>
                        </Stack>

                        <Stack style={{ width: 160 }}>
                            <Label>Contractor Name</Label>
                            <Text>{appState.supportTeam ?? "Not set"}</Text>
                        </Stack>
                    </Stack>

                    {/* Choices row */}
                    <Stack horizontal wrap tokens={{ childrenGap: 24 }}>
                        <Stack style={{ width: 160 }}>
                            <Label>License Required</Label>
                            <Text>{appState.licenseReqd ?? "Not set"}</Text>
                        </Stack>

                        <Stack style={{ width: 120 }}>
                            <Label>User Count</Label>
                            <Text>{Number.isFinite(appState.userCount) ? appState.userCount : "—"}</Text>
                        </Stack>

                    </Stack>

                    {/* People row */}
                    <Stack horizontal wrap tokens={{ childrenGap: 24 }}>
                        <Stack tokens={{ childrenGap: 4 }} style={{ width: 240 }}>
                            <Label>Primary POC</Label>
                            {appState.primaryPoc?.Id ? (
                                <PeoplePersona person={appState.primaryPoc} showDetails={true} fallbackText="Not assigned" />
                            ) : (
                                <Text styles={{ root: { color: "gray", fontStyle: "italic" } }}>Not assigned</Text>
                            )}
                        </Stack>

                        <Stack tokens={{ childrenGap: 4 }} style={{ minWidth: 120 }}>
                            <Label>Stakeholders</Label>
                            <Stack horizontal wrap tokens={{ childrenGap: 4, padding: "0 2px" }}>
                                {stakeholders.length > 0 ? (
                                    stakeholders.map((p) => <PeoplePersona key={p.Id} person={p} />)
                                ) : (
                                    <Text styles={{ root: { color: "gray", fontStyle: "italic" } }}>None identified</Text>
                                )}
                            </Stack>
                        </Stack>

                        <Stack tokens={{ childrenGap: 4 }} style={{ width: 240 }}>
                            <Label>System Owner</Label>
                            {appState.systemOwner?.Id ? (
                                <PeoplePersona person={appState.systemOwner} showDetails={true} fallbackText="Not assigned" />
                            ) : (
                                <Text styles={{ root: { color: "gray", fontStyle: "italic" } }}>Not assigned</Text>
                            )}
                        </Stack>
                    </Stack>
                </Stack>

                {/******** Right content - 52% ********/}
                {rightContent && (
                    <Stack styles={{ root: { flex: "1 1 520px", minWidth: 320, maxWidth: "100%" } }} tokens={{ childrenGap: 8 }}>
                        {rightContent}
                    </Stack>
                )}
            </Stack>
        </Stack>
    );
};
