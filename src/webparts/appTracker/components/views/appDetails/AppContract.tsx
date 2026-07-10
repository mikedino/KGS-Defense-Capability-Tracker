import * as React from "react";
import { Link, Stack, Text, Label } from "@fluentui/react";
import { useHistory } from "react-router-dom";
import { IApplicationItem, IContractItem } from "../../data/props";
import styles from "../../AppTracker.module.scss";
import { formatCurrency, formatDate } from "../../utils";
import { PeoplePersona } from "../../ui/Persona";
import { routes } from "../../routing/routes";

export interface IContractInfoProps {
    contract: IContractItem | undefined;
    relatedApplications?: IApplicationItem[];
    isLoading?: boolean;
}

const contractDetailsCardStyle: React.CSSProperties = {
    flex: "1 1 620px",
    minWidth: 0
};

const relatedAppsCardStyle: React.CSSProperties = {
    flex: "1 1 320px",
    minWidth: 280,
    maxWidth: "100%"
};

const relatedAppRowStyle: React.CSSProperties = {
    padding: "8px 0",
    borderBottom: "1px solid #edebe9"
};

const relatedAppLinkStyles = {
    root: {
        display: "block",
        overflowWrap: "anywhere" as const
    }
};

export const ContractInfo: React.FC<IContractInfoProps> = ({ contract, relatedApplications = [], isLoading }) => {
    const history = useHistory();

    return (
        <Stack tokens={{ childrenGap: 16 }}>
            <Stack horizontal wrap tokens={{ childrenGap: 16 }} styles={{ root: { alignItems: "stretch" } }}>
                <Stack tokens={{ childrenGap: 16 }} className={styles.detailCard} style={contractDetailsCardStyle}>
                    {isLoading && (
                        <Text styles={{ root: { color: "gray", fontStyle: "italic" } }}>
                            Loading contract details…
                        </Text>
                    )}

                    {!isLoading && !contract && (
                        <Text styles={{ root: { color: "gray", fontStyle: "italic" } }}>
                            No contract assigned.
                        </Text>
                    )}

                    {!isLoading && contract && (
                        <>
                            {/* Top row */}
                            <Stack style={{ minWidth: 200, flexGrow: 1 }}>
                                <Label>Contract</Label>
                                <Text>{contract.Title || "—"}</Text>
                            </Stack>

                            <Stack style={{ width: 250 }}>
                                <Label>Contract Lead</Label>
                                <Text>{contract.contractTeam || "—"}</Text>
                            </Stack>

                            <Stack horizontal wrap tokens={{ childrenGap: 24 }}>
                                <Stack style={{ width: 100 }}>
                                    <Label>Start</Label>
                                    <Text>{formatDate(contract.start)}</Text>
                                </Stack>

                                <Stack style={{ width: 100 }}>
                                    <Label>End</Label>
                                    <Text>{formatDate(contract.end)}</Text>
                                </Stack>

                                <Stack style={{ width: 150 }}>
                                    <Label>Contract Value</Label>
                                    <Text>{formatCurrency(contract.contractValue)}</Text>
                                </Stack>
                            </Stack>

                            {/* People row */}
                            <Stack horizontal wrap tokens={{ childrenGap: 24 }}>
                                <Stack style={{ width: "30%" }} grow={1} tokens={{ childrenGap: 4 }}>
                                    <Label>COR</Label>
                                    {contract.cor?.Id ? (
                                        <PeoplePersona person={contract.cor} showDetails={true} fallbackText="Not assigned" />
                                    ) : (
                                        <Text styles={{ root: { color: "gray", fontStyle: "italic" } }}>Not assigned</Text>
                                    )}
                                </Stack>

                                <Stack style={{ width: "30%" }} grow={1} tokens={{ childrenGap: 4 }}>
                                    <Label>ACOR</Label>
                                    {contract.acor?.Id ? (
                                        <PeoplePersona person={contract.acor} showDetails={true} fallbackText="Not assigned" />
                                    ) : (
                                        <Text styles={{ root: { color: "gray", fontStyle: "italic" } }}>Not assigned</Text>
                                    )}
                                </Stack>

                                <Stack style={{ width: "30%" }} grow={1} tokens={{ childrenGap: 4 }}>
                                    <Label>Contracting Officer</Label>
                                    {contract.ko?.Id ? (
                                        <PeoplePersona person={contract.ko} showDetails={true} fallbackText="Not assigned" />
                                    ) : (
                                        <Text styles={{ root: { color: "gray", fontStyle: "italic" } }}>Not assigned</Text>
                                    )}
                                </Stack>
                            </Stack>

                            <Stack horizontal wrap tokens={{ childrenGap: 24 }}>
                                <Stack style={{ width: "40%" }} tokens={{ childrenGap: 4 }}>
                                    <Label>Contract PM</Label>
                                    {contract.primaryPoc?.Id ? (
                                        <PeoplePersona person={contract.primaryPoc} showDetails={true} fallbackText="Not assigned" />
                                    ) : (
                                        <Text styles={{ root: { color: "gray", fontStyle: "italic" } }}>Not assigned</Text>
                                    )}
                                </Stack>

                                <Stack style={{ width: "40%" }} tokens={{ childrenGap: 4 }}>
                                    <Label>Assistant Contract PM</Label>
                                    {contract.secondaryPoc?.Id ? (
                                        <PeoplePersona person={contract.secondaryPoc} showDetails={true} fallbackText="Not assigned" />
                                    ) : (
                                        <Text styles={{ root: { color: "gray", fontStyle: "italic" } }}>Not assigned</Text>
                                    )}
                                </Stack>
                            </Stack>
                        </>
                    )}
                </Stack>

                {!isLoading && contract && (
                    <Stack tokens={{ childrenGap: 8 }} className={styles.detailCard} style={relatedAppsCardStyle}>
                        <Stack>
                            <Text variant="mediumPlus" styles={{ root: { fontWeight: 600 } }}>
                                Other Applications supported by {contract.contractTeam || "this Contract Lead"}
                            </Text>
                        </Stack>

                        <Stack styles={{ root: { maxHeight: 360, overflowY: "auto" } }}>
                            {relatedApplications.length ? (
                                relatedApplications.map((app) => (
                                    <div key={app.Id} style={relatedAppRowStyle}>
                                        <Link onClick={() => history.push(routes.app(app.Id, "overview"))} styles={relatedAppLinkStyles}>
                                            {app.Title}
                                        </Link>
                                    </div>
                                ))
                            ) : (
                                <Text styles={{ root: { color: "gray", fontStyle: "italic", paddingTop: 8 } }}>
                                    No other applications found.
                                </Text>
                            )}
                        </Stack>
                    </Stack>
                )}
            </Stack>
        </Stack>
    );
};
