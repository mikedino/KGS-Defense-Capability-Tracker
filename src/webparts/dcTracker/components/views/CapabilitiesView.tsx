import * as React from "react";
import { DefaultButton, IColumn, Link, SelectionMode, Stack, Text } from "@fluentui/react";
import { ICapabilityItem } from "../common/props";
import { formatDate } from "../common/utils";
import { getAtoStatusFill } from "../ui/StatusColors";
import { Pill } from "../ui/Pill";
import PaginatedDetailsList from "../ui/PaginatedDetailsList";
import styles from "../Dct.module.scss";
import { DataSource } from "../data/ds";
import Strings from "../common/strings";

export interface ICapabilityContractSummary {
    titles: string[];
    searchText: string;
}

interface ICapabilitiesListProps {
    capabilities: ICapabilityItem[];
    contractSummaryByCapabilityId: Map<number, ICapabilityContractSummary>;
    viewMode: string;
    onSelectCap: (capability: ICapabilityItem) => void;
}

export const CapabilitiesList: React.FunctionComponent<ICapabilitiesListProps> = ({
    capabilities,
    contractSummaryByCapabilityId,
    viewMode,
    onSelectCap
}) => {
    const [sortColumnKey, setSortColumnKey] = React.useState<string | null>("title");
    const [isSortedDescending, setIsSortedDescending] = React.useState(false);

    const sortedCapabilities = React.useMemo(() => {
        if (!sortColumnKey) return capabilities;

        return [...capabilities].sort((a, b) => {
            let aVal: string | number = "";
            let bVal: string | number = "";

            switch (sortColumnKey) {
                case "title":
                    aVal = (a.Title || "").toLowerCase();
                    bVal = (b.Title || "").toLowerCase();
                    break;
                case "platform":
                    aVal = (a.platform || "").toLowerCase();
                    bVal = (b.platform || "").toLowerCase();
                    break;
                case "capabilityType":
                    aVal = DataSource.getConfigText("capabilityType", a.capabilityTypeTier1 || a.capabilityTypeTier2).toLowerCase();
                    bVal = DataSource.getConfigText("capabilityType", b.capabilityTypeTier1 || b.capabilityTypeTier2).toLowerCase();
                    break;
                case "capStatus":
                    aVal = (a.capStatus || "").toLowerCase();
                    bVal = (b.capStatus || "").toLowerCase();
                    break;
                case "hostingEnv":
                    aVal = (a.hostingEnv || "").toLowerCase();
                    bVal = (b.hostingEnv || "").toLowerCase();
                    break;
                case "modified":
                    aVal = a.Modified ? new Date(a.Modified).getTime() : 0;
                    bVal = b.Modified ? new Date(b.Modified).getTime() : 0;
                    break;
                default:
                    return 0;
            }

            if (aVal < bVal) return isSortedDescending ? 1 : -1;
            if (aVal > bVal) return isSortedDescending ? -1 : 1;
            return 0;
        });
    }, [capabilities, sortColumnKey, isSortedDescending]);

    const onColumnClick = (ev: React.MouseEvent<HTMLElement>, column: IColumn): void => {
        const newIsSortedDescending = sortColumnKey === column.key ? !isSortedDescending : false;
        setSortColumnKey(column.key);
        setIsSortedDescending(newIsSortedDescending);
    };

    const stripHtml = (value?: string): string => (value ?? "").replace(/<[^>]*>/g, " ").replace(/\s+/g, " ").trim();

    const renderTileMetaRow = (label: string, value?: string | number): JSX.Element => (
        <div className={styles.capabilityTileMetaRow}>
            <span className={styles.capabilityTileLabel}>{label}</span>
            <span className={styles.capabilityTileValue} title={value ? String(value) : undefined}>
                {value || "-"}
            </span>
        </div>
    );

    const getCapabilityTypeSummary = (capability: ICapabilityItem): string => [
        capability.capabilityTypeTier1
            ? `Tier 1: ${DataSource.getConfigText("capabilityType", capability.capabilityTypeTier1)}`
            : "",
        capability.capabilityTypeTier2
            ? `Tier 2: ${DataSource.getConfigText("capabilityType", capability.capabilityTypeTier2)}`
            : ""
    ].filter(Boolean).join(" • ");

    if (viewMode === "tile") {
        return (
            <Stack tokens={{ childrenGap: 16 }}>
                <div className={styles.tileView}>
                    {sortedCapabilities.map((capability) => {
                        const relatedContracts = contractSummaryByCapabilityId.get(capability.Id)?.titles ?? [];
                        const primaryContract = relatedContracts[0];
                        const contractText = relatedContracts.length > 1
                            ? `${primaryContract} +${relatedContracts.length - 1}`
                            : primaryContract;
                        const statusProps = getAtoStatusFill(capability.capStatus);

                        return (
                            <div
                                key={capability.Id}
                                className={styles.capabilityTile}
                                role="button"
                                tabIndex={0}
                                onClick={() => onSelectCap(capability)}
                                onKeyDown={(ev) => {
                                    if (ev.key === "Enter" || ev.key === " ") {
                                        ev.preventDefault();
                                        onSelectCap(capability);
                                    }
                                }}
                            >
                                <Stack tokens={{ childrenGap: 12 }}>
                                    <div className={styles.capabilityTileHeader}>
                                        <Text className={styles.capabilityTileTitle}>{capability.Title}</Text>
                                        <Pill text={capability.capStatus} backgroundColor={statusProps.backgroundColor} textColor={statusProps.textColor} />
                                    </div>

                                    <Text className={styles.capabilityTileDescription}>
                                        {stripHtml(capability.description) || "No description provided."}
                                    </Text>

                                    <div className={styles.capabilityTileMeta}>
                                        {renderTileMetaRow("Platform:", capability.platform)}
                                        {renderTileMetaRow("Capability Type:", getCapabilityTypeSummary(capability))}
                                        {renderTileMetaRow("Hosting:", capability.hostingEnv)}
                                        {renderTileMetaRow("Primary POC:", capability.primaryPoc?.Title)}
                                        {renderTileMetaRow("Contract:", contractText)}
                                    </div>

                                    <DefaultButton
                                        className={styles.capabilityTileButton}
                                        text="View Details"
                                        iconProps={{ iconName: "ChevronRight" }}
                                        onClick={(ev) => {
                                            ev.stopPropagation();
                                            onSelectCap(capability);
                                        }}
                                    />
                                </Stack>
                            </div>
                        );
                    })}
                </div>
            </Stack>
        );
    }

    const columns: IColumn[] = [
        {
            key: "title",
            name: "Capability",
            fieldName: "Title",
            minWidth: 220,
            maxWidth: 340,
            isResizable: true,
            isSorted: sortColumnKey === "title",
            isSortedDescending,
            onColumnClick,
            onRender: (item: ICapabilityItem) => (
                <Stack>
                    <Link
                        className={styles.listTitleLink}
                        onClick={(ev) => {
                            ev?.preventDefault();
                            ev?.stopPropagation();
                            onSelectCap(item);
                        }}
                    >
                        {item.Title}
                    </Link>
                </Stack>
            )
        },
        {
            key: "capStatus",
            name: "Status",
            fieldName: "capStatus",
            minWidth: 100,
            maxWidth: 120,
            isResizable: true,
            isSorted: sortColumnKey === "capStatus",
            isSortedDescending,
            onColumnClick,
            onRender: (item: ICapabilityItem) => {
                const { backgroundColor, textColor } = getAtoStatusFill(item.capStatus);
                return <Pill text={item.capStatus} backgroundColor={backgroundColor} textColor={textColor} />;
            }
        },
        {
            key: "capabilityType",
            name: "Capability Type",
            fieldName: "capabilityTypeTier1",
            minWidth: 220,
            maxWidth: 320,
            isResizable: true,
            isSorted: sortColumnKey === "capabilityType",
            isSortedDescending,
            onColumnClick,
            onRender: (item: ICapabilityItem) => (
                <Stack tokens={{ childrenGap: 2 }}>
                    {item.capabilityTypeTier1 && (
                        <div className={styles.capabilityTypeCellRow}>
                            <span
                                className={styles.capabilityTypeTierBadge}
                                style={{
                                    backgroundColor: Strings.PillStyles.LilacFill,
                                    borderColor: Strings.PillStyles.LilacColor,
                                    color: Strings.PillStyles.LilacColor
                                }}
                                title="Tier 1"
                                aria-label="Tier 1"
                            >
                                1
                            </span>
                            <Text variant="small">{DataSource.getConfigText("capabilityType", item.capabilityTypeTier1)}</Text>
                        </div>
                    )}
                    {item.capabilityTypeTier2 && (
                        <div className={styles.capabilityTypeCellRow}>
                            <span
                                className={styles.capabilityTypeTierBadge}
                                style={{
                                    backgroundColor: Strings.PillStyles.PurpleFill,
                                    borderColor: Strings.PillStyles.PurpleColor,
                                    color: Strings.PillStyles.PurpleColor
                                }}
                                title="Tier 2"
                                aria-label="Tier 2"
                            >
                                2
                            </span>
                            <Text variant="small">{DataSource.getConfigText("capabilityType", item.capabilityTypeTier2)}</Text>
                        </div>
                    )}
                </Stack>
            )
        },
        {
            key: "platform",
            name: "Platform",
            fieldName: "platform",
            minWidth: 110,
            maxWidth: 150,
            isResizable: true,
            isSorted: sortColumnKey === "platform",
            isSortedDescending,
            onColumnClick,
            onRender: (item: ICapabilityItem) => <Text>{item.platform || ""}</Text>
        },
        {
            key: "hostingEnv",
            name: "Hosting Environment",
            fieldName: "hostingEnv",
            minWidth: 150,
            maxWidth: 200,
            isResizable: true,
            isSorted: sortColumnKey === "hostingEnv",
            isSortedDescending,
            onColumnClick,
            onRender: (item: ICapabilityItem) => <Text>{item.hostingEnv || ""}</Text>
        },
        {
            key: "contract",
            name: "Contracts",
            fieldName: "contract",
            minWidth: 180,
            maxWidth: 280,
            isResizable: true,
            onRender: (item: ICapabilityItem) => {
                const relatedContracts = contractSummaryByCapabilityId.get(item.Id)?.titles ?? [];
                return <Text>{relatedContracts.join(", ")}</Text>;
            }
        },
        {
            key: "modified",
            name: "Modified",
            fieldName: "Modified",
            minWidth: 110,
            maxWidth: 130,
            isResizable: false,
            isSorted: sortColumnKey === "modified",
            isSortedDescending,
            onColumnClick,
            headerClassName: styles.centeredHeader,
            className: styles.centeredColumn,
            onRender: (item: ICapabilityItem) => <Text>{item.Modified ? formatDate(item.Modified) : "-"}</Text>
        }
    ];

    return (
        <Stack tokens={{ childrenGap: 16 }}>
            <PaginatedDetailsList
                items={sortedCapabilities}
                columns={columns}
                selectionMode={SelectionMode.none}
                layoutMode={1}
                isHeaderVisible={true}
                onItemInvoked={(item) => onSelectCap(item as ICapabilityItem)}
                pageSizeOptions={[10, 25, 50]}
                defaultPageSizeOption={25}
                showFirstLastButtons={true}
            />
        </Stack>
    );
};
