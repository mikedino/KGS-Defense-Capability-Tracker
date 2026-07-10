import * as React from "react"
import { Stack, Text, IColumn, SelectionMode } from "@fluentui/react"
import { Pill } from "../ui/Pill"
import { IApplicationItem, IPeoplePicker } from "../data/props"
import PaginatedDetailsList from "../ui/PaginatedDetailsList"
import { formatDate } from "../utils"
import { getAtoStatusFill } from "../ui/StatusColors"
import styles from "../AppTracker.module.scss"
import Strings from "../../strings"

interface IAppsListProps {
    apps: IApplicationItem[]
    viewMode: string
    onSelectApp: (app: IApplicationItem) => void
}

const renderPeople = (people: IPeoplePicker[] ): string => {
    const arr = Array.isArray(people) ? people : [];
    if (!arr?.length) return "";
    return arr
        .map((p) => p.Title)
        .filter(Boolean)
        .join(", ");
};

export const AppsList: React.FunctionComponent<IAppsListProps> = ({
    apps,
    viewMode,
    onSelectApp
}) => {
    // SORTING
    const [sortColumnKey, setSortColumnKey] = React.useState<string | null>("title");
    const [isSortedDescending, setIsSortedDescending] = React.useState(false);

    const sortedApps = React.useMemo(() => {
        if (!sortColumnKey) return apps;

        return [...apps].sort((a, b) => {
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
                case "managingGroup":
                    aVal = (a.managingGroup || "").toLowerCase();
                    bVal = (b.managingGroup || "").toLowerCase();
                    break;
                case "appStatus":
                    aVal = (a.appStatus || "").toLowerCase();
                    bVal = (b.appStatus || "").toLowerCase();
                    break;
                case "appLaunchDate":
                    aVal = a.appLaunchDate ? new Date(a.appLaunchDate).getTime() : 0;
                    bVal = b.appLaunchDate ? new Date(b.appLaunchDate).getTime() : 0;
                    break;
                case "supportTeam":
                    aVal = (a.supportTeam || "").toLowerCase();
                    bVal = (b.supportTeam || "").toLowerCase();
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
    }, [apps, sortColumnKey, isSortedDescending]);

    const onColumnClick = (ev: React.MouseEvent<HTMLElement>, column: IColumn): void => {
        const newIsSortedDescending = sortColumnKey === column.key ? !isSortedDescending : false;
        setSortColumnKey(column.key);
        setIsSortedDescending(newIsSortedDescending);
    };

    const columns: IColumn[] = [
        {
            key: "title",
            name: "Application",
            fieldName: "Title",
            minWidth: 220,
            maxWidth: 320,
            isResizable: true,
            isSorted: sortColumnKey === "title",
            isSortedDescending,
            onColumnClick,
            onRender: (item: IApplicationItem) => (
                <Stack>
                    <Text variant="medium" style={{ fontWeight: 600 }}>
                        {item.Title}
                    </Text>
                </Stack>
            )
        },
        {
            key: "platform",
            name: "Platform",
            fieldName: "platform",
            minWidth: 110,
            maxWidth: 160,
            isResizable: true,
            isSorted: sortColumnKey === "platform",
            isSortedDescending,
            onColumnClick,
            onRender: (item: IApplicationItem) => <Text>{item.platform || ""}</Text>
        },
        {
            key: "primaryPoc",
            name: "Primary POC",
            fieldName: "primaryPoc",
            minWidth: 130,
            maxWidth: 180,
            isResizable: true,
            onRender: (item: IApplicationItem) => <Text>{item.primaryPoc?.Title || ""}</Text>
        },
        {
            key: "stakeholders",
            name: "Stakeholders",
            fieldName: "stakeholders",
            minWidth: 160,
            maxWidth: 260,
            isResizable: true,
            onRender: (item: IApplicationItem) => <Text>{renderPeople(item.stakeholders?.results ?? [])}</Text>
        },
        {
            key: "managingGroup",
            name: "Managing Group",
            fieldName: "managingGroup",
            minWidth: 120,
            maxWidth: 170,
            isResizable: true,
            isSorted: sortColumnKey === "managingGroup",
            isSortedDescending,
            onColumnClick,
            onRender: (item: IApplicationItem) => <Text>{item.managingGroup || ""}</Text>
        },
        {
            key: "appStatus",
            name: "App Status",
            fieldName: "appStatus",
            minWidth: 120,
            maxWidth: 160,
            isResizable: true,
            isSorted: sortColumnKey === "appStatus",
            isSortedDescending,
            onColumnClick,
            onRender: (item: IApplicationItem) => {
                const { backgroundColor, textColor } = getAtoStatusFill(item.appStatus);
                return <Pill text={item.appStatus} backgroundColor={backgroundColor} textColor={textColor} />;
            }
        },
        {
            key: "appLaunchDate",
            name: "Launch Date",
            fieldName: "appLaunchDate",
            minWidth: 110,
            maxWidth: 130,
            isResizable: true,
            isSorted: sortColumnKey === "appLaunchDate",
            isSortedDescending,
            onColumnClick,
            headerClassName: styles.centeredHeader,
            className: styles.centeredColumn,
            onRender: (item: IApplicationItem) => {
                const rawDate = item.appLaunchDate;

                if (!rawDate) {
                    return <Text>{formatDate(rawDate)}</Text>;
                }

                const expDate = new Date(rawDate);

                const today = new Date();
                today.setHours(0, 0, 0, 0); // normalize so today isn't considered "past"

                const isPast = !isNaN(expDate.getTime()) && expDate < today;

                return (
                    <Text color={isPast ? Strings.PillStyles.RedColor : "inherit" }>
                        {formatDate(rawDate)}
                    </Text>
                );
            }
        },
        {
            key: "supportTeam",
            name: "Contractor Name",
            fieldName: "supportTeam",
            minWidth: 130,
            maxWidth: 180,
            isResizable: true,
            isSorted: sortColumnKey === "supportTeam",
            isSortedDescending,
            onColumnClick,
            onRender: (item: IApplicationItem) => <Text>{item.supportTeam || ""}</Text>
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
            onRender: (item: IApplicationItem) => <Text>{item.Modified ? formatDate(item.Modified) : "-"}</Text>
        }
    ];

    return (
        <Stack tokens={{ childrenGap: 16 }}>
            <PaginatedDetailsList
                items={sortedApps}
                columns={columns}
                selectionMode={SelectionMode.none}
                layoutMode={1}
                isHeaderVisible={true}
                onItemInvoked={(item) => onSelectApp(item as IApplicationItem)}
                pageSizeOptions={[5, 10, 25, 50]}
                defaultPageSizeOption={10}
                showFirstLastButtons={true}
            />
        </Stack>
    );
};
