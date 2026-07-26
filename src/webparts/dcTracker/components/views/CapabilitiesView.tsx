import * as React from "react";
import { IColumn, SelectionMode, Stack, Text } from "@fluentui/react";
import { ICapabilityItem } from "../common/props";
import { formatDate } from "../common/utils";
import { DataSource } from "../data/ds";
import { getAtoStatusFill } from "../ui/StatusColors";
import { Pill } from "../ui/Pill";
import PaginatedDetailsList from "../ui/PaginatedDetailsList";
import styles from "../Dct.module.scss";

interface ICapabilitiesListProps {
    capabilities: ICapabilityItem[];
    viewMode: string;
    onSelectCap: (capability: ICapabilityItem) => void;
}

export const CapabilitiesList: React.FunctionComponent<ICapabilitiesListProps> = ({
    capabilities,
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
                case "solutionType":
                    aVal = (a.solutionType || "").toLowerCase();
                    bVal = (b.solutionType || "").toLowerCase();
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
                    <Text variant="medium" style={{ fontWeight: 600 }}>
                        {item.Title}
                    </Text>
                </Stack>
            )
        },
        {
            key: "capStatus",
            name: "Status",
            fieldName: "capStatus",
            minWidth: 120,
            maxWidth: 160,
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
            key: "solutionType",
            name: "Solution Type",
            fieldName: "solutionType",
            minWidth: 140,
            maxWidth: 210,
            isResizable: true,
            isSorted: sortColumnKey === "solutionType",
            isSortedDescending,
            onColumnClick,
            onRender: (item: ICapabilityItem) => <Text>{item.solutionType || ""}</Text>
        },
        {
            key: "platform",
            name: "Platform",
            fieldName: "platform",
            minWidth: 130,
            maxWidth: 180,
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
            maxWidth: 220,
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
            minWidth: 140,
            maxWidth: 220,
            isResizable: true,
            onRender: (item: ICapabilityItem) => {
                const relatedContracts = DataSource.Contracts
                    .filter((contract) => contract.capability?.Id === item.Id)
                    .map((contract) => contract.Title)
                    .filter(Boolean);
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
                pageSizeOptions={[5, 10, 25, 50]}
                defaultPageSizeOption={10}
                showFirstLastButtons={true}
            />
        </Stack>
    );
};
