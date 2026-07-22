import * as React from "react";
import { Label, Link, Stack, Text } from "@fluentui/react";
import { useHistory } from "react-router-dom";
import { ICapabilityItem, IContractItem, IPeoplePickerExtended } from "../../common/props";
import styles from "../../Dct.module.scss";
import { formatDate } from "../../common/utils";
import { PeoplePersona } from "../../ui/Persona";
import { routes } from "../../routing/routes";

export interface IContractInfoProps {
    contract: IContractItem | undefined;
    relatedCapabilities?: ICapabilityItem[];
    isLoading?: boolean;
}

const contractDetailsCardStyle: React.CSSProperties = {
    flex: "1 1 620px",
    minWidth: 0
};

const relatedCapabilitiesCardStyle: React.CSSProperties = {
    flex: "1 1 320px",
    minWidth: 280,
    maxWidth: "100%"
};

const relatedCapabilityRowStyle: React.CSSProperties = {
    padding: "8px 0",
    borderBottom: "1px solid #edebe9"
};

const relatedCapabilityLinkStyles = {
    root: {
        display: "block",
        overflowWrap: "anywhere" as const
    }
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

const PeopleDisplay: React.FC<{ label: string; people?: { results: IPeoplePickerExtended[] } }> = ({ label, people }) => (
    <Stack style={{ minWidth: 260, flexGrow: 1 }} tokens={{ childrenGap: 4 }}>
        <Label>{label}</Label>
        <Stack horizontal wrap tokens={{ childrenGap: 4, padding: "0 2px" }}>
            {people?.results?.length ? (
                people.results.map((p) => <PeoplePersona key={p.Id} person={p} />)
            ) : (
                <Text styles={{ root: { color: "gray", fontStyle: "italic" } }}>None identified</Text>
            )}
        </Stack>
    </Stack>
);

export const ContractInfo: React.FC<IContractInfoProps> = ({ contract, relatedCapabilities = [], isLoading }) => {
    const history = useHistory();

    return (
        <Stack tokens={{ childrenGap: 16 }}>
            <Stack horizontal wrap tokens={{ childrenGap: 16 }} styles={{ root: { alignItems: "stretch" } }}>
                <Stack tokens={{ childrenGap: 16 }} className={styles.detailCard} style={contractDetailsCardStyle}>
                    {isLoading && (
                        <Text styles={{ root: { color: "gray", fontStyle: "italic" } }}>
                            Loading contract details...
                        </Text>
                    )}

                    {!isLoading && !contract && (
                        <Text styles={{ root: { color: "gray", fontStyle: "italic" } }}>
                            No contract assigned.
                        </Text>
                    )}

                    {!isLoading && contract && (
                        <>
                            <Stack horizontal wrap tokens={{ childrenGap: 24 }}>
                                <FieldDisplay label="Contract ID" value={contract.contractId} />
                                <FieldDisplay label="Task Order/Invoice ID" value={contract.invoice} />
                                <FieldDisplay label="Contract Title" value={contract.Title} />
                            </Stack>

                            <Stack horizontal wrap tokens={{ childrenGap: 24 }}>
                                <FieldDisplay label="Customer Contract Code" value={contract.customerContractCode} />
                                <FieldDisplay label="Customer" value={contract.customer} />
                                <FieldDisplay label="Relevant Partner Tag" value={contract.partner} />
                            </Stack>

                            <Stack horizontal wrap tokens={{ childrenGap: 24 }}>
                                <FieldDisplay label="PoP Start Date" value={contract.popStart ? formatDate(contract.popStart) : undefined} />
                                <FieldDisplay label="PoP End Date" value={contract.popEnd ? formatDate(contract.popEnd) : undefined} />
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
                                <PersonDisplay label="Capability Primary POC" person={contract.primaryPoc} />
                                <PeopleDisplay label="KGS Stakeholders" people={contract.stakeholders} />
                            </Stack>
                        </>
                    )}
                </Stack>

                {!isLoading && contract && (
                    <Stack tokens={{ childrenGap: 8 }} className={styles.detailCard} style={relatedCapabilitiesCardStyle}>
                        <Stack>
                            <Text variant="mediumPlus" styles={{ root: { fontWeight: 600 } }}>
                                Other Capabilities on {contract.Title || "this contract"}
                            </Text>
                        </Stack>

                        <Stack styles={{ root: { maxHeight: 360, overflowY: "auto" } }}>
                            {relatedCapabilities.length ? (
                                relatedCapabilities.map((cap) => (
                                    <div key={cap.Id} style={relatedCapabilityRowStyle}>
                                        <Link onClick={() => history.push(routes.app(cap.Id, "overview"))} styles={relatedCapabilityLinkStyles}>
                                            {cap.Title}
                                        </Link>
                                    </div>
                                ))
                            ) : (
                                <Text styles={{ root: { color: "gray", fontStyle: "italic", paddingTop: 8 } }}>
                                    No other capabilities found.
                                </Text>
                            )}
                        </Stack>
                    </Stack>
                )}
            </Stack>
        </Stack>
    );
};
