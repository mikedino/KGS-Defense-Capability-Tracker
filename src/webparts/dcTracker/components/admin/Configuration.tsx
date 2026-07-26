import * as React from "react";
import {
    Checkbox,
    DefaultButton,
    DetailsList,
    DetailsListLayoutMode,
    DetailsRow,
    Dialog,
    DialogFooter,
    DialogType,
    IColumn,
    IDetailsRowProps,
    INavLinkGroup,
    IconButton,
    Nav,
    PrimaryButton,
    SelectionMode,
    Spinner,
    SpinnerSize,
    Stack,
    Text,
    TextField
} from "@fluentui/react";

import type { ConfigType, IConfigItem } from "../common/props";
import { ConfigService } from "../services/ConfigService";
import { formatError } from "../common/utils";
import { DataSource } from "../data/ds";

interface IConfigManagementProps {
    onChanged?: (items: IConfigItem[]) => void;
}

type DialogMode = "add" | "edit" | "delete" | "error" | undefined;

const knownConfigTypes: { key: ConfigType; displayName: string }[] = [
    { key: "backend", displayName: "Backend" },
    { key: "capabilityStatus", displayName: "Capability Status" },
    { key: "codingLanguage", displayName: "Coding Language" },
    { key: "compliance", displayName: "Compliance" },
    { key: "connectivity", displayName: "Connectivity" },
    { key: "customer", displayName: "Customer" },
    { key: "partner", displayName: "Relevant Partner Tag" },
    { key: "hostingEnvironment", displayName: "Hosting Environment" },
    { key: "platform", displayName: "Platform" },
    { key: "solutionType", displayName: "Solution Type" },
    { key: "documentType", displayName: "Document Type" }
];

const getTypeKeys = (items: IConfigItem[]): string[] => {
    const unique = Array.from(new Set([
        ...knownConfigTypes.map(type => type.key),
        ...items.map(i => (i.configType ?? "").trim()).filter(Boolean)
    ]));

    return unique.sort((a, b) => {
        const aIndex = knownConfigTypes.findIndex(type => type.key === a);
        const bIndex = knownConfigTypes.findIndex(type => type.key === b);

        if (aIndex >= 0 && bIndex >= 0) return aIndex - bIndex;
        if (aIndex >= 0) return -1;
        if (bIndex >= 0) return 1;
        return a.localeCompare(b);
    });
};

const getTypeDisplayName = (key: string): string =>
    knownConfigTypes.find(type => type.key === key)?.displayName ?? key;

export const ConfigManagement: React.FC<IConfigManagementProps> = ({ onChanged }) => {
    const [allItems, setAllItems] = React.useState<IConfigItem[]>([]);
    const [typeKeys, setTypeKeys] = React.useState<string[]>([]);
    const [selectedType, setSelectedType] = React.useState<string>("");
    const [dialogMode, setDialogMode] = React.useState<DialogMode>(undefined);
    const [dialogTitle, setDialogTitle] = React.useState<string>("");
    const [dialogMessage, setDialogMessage] = React.useState<string>("");
    const [busy, setBusy] = React.useState<boolean>(false);
    const [working, setWorking] = React.useState<IConfigItem | undefined>(undefined);

    const refreshLocal = React.useCallback((next: IConfigItem[]): void => {
        const sorted = [...next].sort((a, b) =>
            (a.configType ?? "").localeCompare(b.configType ?? "")
            || (a.sortOrder ?? 9999) - (b.sortOrder ?? 9999)
            || (a.Title ?? "").localeCompare(b.Title ?? "")
        );

        setAllItems(sorted);
        const keys = getTypeKeys(sorted);
        setTypeKeys(keys);
        setSelectedType(prev => (prev && keys.includes(prev) ? prev : (keys[0] ?? "")));
        onChanged?.(sorted);
    }, [onChanged]);

    const getConfig = async (): Promise<void> => {
        setBusy(true);
        try {
            refreshLocal(await DataSource.getConfig());
        } catch (error) {
            const errorMessage = formatError(error);
            console.error("Error getting Config items:", errorMessage);
            setDialogMode("error");
            setDialogTitle("Error getting Config items");
            setDialogMessage(errorMessage);
        } finally {
            setBusy(false);
        }
    };

    React.useEffect(() => {
        getConfig().catch((error) => {
            console.error("Unhandled promise rejection:", error);
        });
    }, []);

    const listForSelected = React.useMemo(() => {
        const key = (selectedType ?? "").trim();
        return allItems
            .filter(i => (i.configType ?? "").trim() === key)
            .sort((a, b) => (a.sortOrder ?? 9999) - (b.sortOrder ?? 9999) || (a.Title ?? "").localeCompare(b.Title ?? ""));
    }, [allItems, selectedType]);

    const openAdd = (): void => {
        if (!selectedType) return;

        setWorking({
            Id: 0,
            Title: "",
            configType: selectedType,
            configValue: "",
            sortOrder: listForSelected.length + 1,
            isActive: true,
            infoText: ""
        });

        setDialogMode("add");
        setDialogTitle(`Add: ${getTypeDisplayName(selectedType)}`);
        setDialogMessage("");
    };

    const openEdit = (item: IConfigItem): void => {
        setWorking({ ...item });
        setDialogMode("edit");
        setDialogTitle(`Edit: ${item.Title}`);
        setDialogMessage("");
    };

    const openDelete = (item: IConfigItem): void => {
        setWorking({ ...item });
        setDialogMode("delete");
        setDialogTitle(`Delete: ${item.Title}`);
        setDialogMessage("Are you sure you want to delete this configuration entry?");
    };

    const closeDialog = (): void => {
        setDialogMode(undefined);
        setWorking(undefined);
        setDialogTitle("");
        setDialogMessage("");
    };

    const validateWorking = (item: IConfigItem): string | undefined => {
        if (!item.Title?.trim()) return "Display Text is required.";
        if (!item.configValue?.trim()) return "Saved Value is required.";
        if (!item.configType?.trim()) return "Config Type is required.";
        return undefined;
    };

    const onSave = async (): Promise<void> => {
        if (!working) return;

        try {
            const validationMessage = validateWorking(working);
            if (validationMessage) {
                setDialogMessage(validationMessage);
                return;
            }

            setBusy(true);
            const itemToSave: IConfigItem = {
                ...working,
                Title: working.Title.trim(),
                configType: working.configType.trim(),
                configValue: working.configValue.trim(),
                infoText: working.infoText?.trim() ?? ""
            };

            if (dialogMode === "add") {
                const created = await ConfigService.create(itemToSave);
                refreshLocal([...allItems, created]);
            }

            if (dialogMode === "edit") {
                const edited = await ConfigService.edit(itemToSave);
                refreshLocal(allItems.map(i => (i.Id === edited.Id ? edited : i)));
            }

            closeDialog();
        } catch (e) {
            console.error(e);
            setDialogMessage(formatError(e));
        } finally {
            setBusy(false);
        }
    };

    const onConfirmDelete = async (): Promise<void> => {
        if (!working) return;

        try {
            setBusy(true);
            await ConfigService.delete(working.Id);
            refreshLocal(allItems.filter(i => i.Id !== working.Id));
            closeDialog();
        } catch (e) {
            console.error(e);
            setDialogMessage(formatError(e));
        } finally {
            setBusy(false);
        }
    };

    const columns: IColumn[] = [
        {
            key: "title",
            name: "Display Text",
            fieldName: "Title",
            minWidth: 110,
            maxWidth: 190,
            currentWidth: 160,
            isResizable: true,
            isMultiline: true,
            onRender: (item: IConfigItem) => (
                <Stack styles={{ root: { whiteSpace: "normal", wordBreak: "break-word" } }}>
                    <Text variant="medium" style={{ fontWeight: 600 }}>{item.Title}</Text>
                    {!!item.infoText && (
                        <Text variant="small" styles={{ root: { color: "#666" } }}>
                            {item.infoText}
                        </Text>
                    )}
                </Stack>
            )
        },
        {
            key: "configValue",
            name: "Saved Value",
            fieldName: "configValue",
            minWidth: 110,
            maxWidth: 180,
            currentWidth: 150,
            isResizable: true,
            isMultiline: true,
            onRender: (item: IConfigItem) => (
                <Text styles={{ root: { whiteSpace: "normal", wordBreak: "break-word" } }}>
                    {item.configValue}
                </Text>
            )
        },
        {
            key: "sortOrder",
            name: "Sort",
            fieldName: "sortOrder",
            minWidth: 48,
            maxWidth: 54,
            currentWidth: 52,
            isResizable: false,
            onRender: (item: IConfigItem) => <Text>{item.sortOrder ?? ""}</Text>
        },
        {
            key: "isActive",
            name: "Active",
            fieldName: "isActive",
            minWidth: 64,
            maxWidth: 70,
            currentWidth: 66,
            isResizable: false,
            onRender: (item: IConfigItem) => <Text>{item.isActive ? "Yes" : "No"}</Text>
        },
        {
            key: "actions",
            name: "",
            minWidth: 70,
            maxWidth: 74,
            currentWidth: 72,
            isResizable: false,
            onRender: (item: IConfigItem) => (
                <Stack horizontal tokens={{ childrenGap: 6 }}>
                    <IconButton
                        iconProps={{ iconName: "Edit" }}
                        title="Edit"
                        ariaLabel="Edit"
                        styles={{ root: { width: 28, height: 28 } }}
                        onClick={() => openEdit(item)}
                    />
                    <IconButton
                        iconProps={{ iconName: "Delete" }}
                        title="Delete"
                        ariaLabel="Delete"
                        styles={{ root: { width: 28, height: 28 } }}
                        onClick={() => openDelete(item)}
                    />
                </Stack>
            )
        }
    ];

    const navGroups: INavLinkGroup[] = [
        {
            links: typeKeys.map((k) => ({
                key: k,
                name: `${getTypeDisplayName(k)} (${allItems.filter(i => i.configType === k).length})`,
                url: "",
                onClick: () => setSelectedType(k)
            }))
        }
    ];

    return (
        <Stack tokens={{ childrenGap: 16 }} styles={{ root: { width: "100%", marginTop: 24 } }}>
            <Stack horizontalAlign="space-between" verticalAlign="center">
                <Text variant="xLarge">Configuration Management</Text>
                <Text variant="medium">Adjust dropdown values within the app</Text>
            </Stack>

            <Stack horizontal wrap tokens={{ childrenGap: 20 }}>
                <Stack styles={{ root: { flex: "0 1 240px", minWidth: 180 } }}>
                    <Nav
                        selectedKey={selectedType}
                        groups={navGroups}
                        styles={{
                            root: {
                                borderRight: "1px solid #e1dfdd"
                            }
                        }}
                    />
                </Stack>

                <Stack grow tokens={{ childrenGap: 12 }} styles={{ root: { minWidth: 0 } }}>
                    <Stack horizontal horizontalAlign="space-between" verticalAlign="center" wrap tokens={{ childrenGap: 12 }}>
                        <Text variant="xLarge">{getTypeDisplayName(selectedType)}</Text>

                        <PrimaryButton
                            text="Add Value"
                            iconProps={{ iconName: "Add" }}
                            onClick={openAdd}
                        />
                    </Stack>

                    <DetailsList
                        setKey={`config-${selectedType}`}
                        items={listForSelected}
                        columns={columns}
                        selectionMode={SelectionMode.none}
                        layoutMode={DetailsListLayoutMode.fixedColumns}
                        compact
                        onRenderRow={(props?: IDetailsRowProps) => props ? (
                            <DetailsRow
                                {...props}
                                styles={{
                                    root: { minHeight: 40 },
                                    cell: {
                                        alignItems: "center",
                                        display: "flex",
                                        height: "auto",
                                        minHeight: 40,
                                        paddingBottom: 4,
                                        paddingTop: 4
                                    }
                                }}
                            />
                        ) : null}
                        onItemInvoked={(item) => openEdit(item as IConfigItem)}
                    />
                </Stack>
            </Stack>

            <Dialog
                hidden={dialogMode !== "add" && dialogMode !== "edit"}
                onDismiss={closeDialog}
                dialogContentProps={{
                    type: DialogType.normal,
                    title: dialogTitle,
                    subText: dialogMessage || undefined
                }}
            >
                {working && (
                    <Stack tokens={{ childrenGap: 12 }}>
                        <TextField
                            label="Display Text"
                            value={working.Title ?? ""}
                            onChange={(_, v) => setWorking({ ...working, Title: v ?? "" })}
                            required
                        />

                        <TextField
                            label="Saved Value"
                            value={working.configValue ?? ""}
                            onChange={(_, v) => setWorking({ ...working, configValue: v ?? "" })}
                            required
                        />

                        <TextField
                            label="Sort Order"
                            type="number"
                            value={working.sortOrder !== undefined ? String(working.sortOrder) : ""}
                            onChange={(_, v) => {
                                const sortOrder = v ? Number(v) : undefined;
                                setWorking({ ...working, sortOrder: Number.isFinite(sortOrder) ? sortOrder : undefined });
                            }}
                        />

                        <TextField
                            label="Info Text"
                            placeholder="Optional note that explains when to use this configuration value or what it means to users."
                            value={working.infoText ?? ""}
                            onChange={(_, v) => setWorking({ ...working, infoText: v ?? "" })}
                            multiline
                            autoAdjustHeight
                        />

                        <Checkbox
                            label="Active"
                            checked={!!working.isActive}
                            onChange={(_, checked) => setWorking({ ...working, isActive: !!checked })}
                        />

                        <TextField
                            label="Config Type"
                            value={working.configType ?? ""}
                            disabled
                        />
                    </Stack>
                )}

                <DialogFooter>
                    <PrimaryButton text={dialogMode === "add" ? "Create" : "Save"} onClick={onSave} disabled={busy} />
                    <DefaultButton text="Cancel" onClick={closeDialog} disabled={busy} />
                </DialogFooter>
            </Dialog>

            <Dialog
                hidden={dialogMode !== "delete"}
                onDismiss={closeDialog}
                dialogContentProps={{
                    type: DialogType.normal,
                    title: dialogTitle,
                    subText: dialogMessage
                }}
            >
                <DialogFooter>
                    <PrimaryButton text="Delete" onClick={onConfirmDelete} disabled={busy} />
                    <DefaultButton text="Cancel" onClick={closeDialog} disabled={busy} />
                </DialogFooter>
            </Dialog>

            <Dialog
                hidden={dialogMode !== "error"}
                onDismiss={closeDialog}
                dialogContentProps={{
                    type: DialogType.normal,
                    title: dialogTitle,
                    subText: dialogMessage
                }}
            >
                <DialogFooter>
                    <DefaultButton text="OK" onClick={closeDialog} disabled={busy} />
                </DialogFooter>
            </Dialog>

            <Dialog
                hidden={!busy}
                onDismiss={() => { setBusy(false); }}
                dialogContentProps={{
                    type: DialogType.normal,
                    title: "Loading...",
                    closeButtonAriaLabel: "Close"
                }}
            >
                <Spinner size={SpinnerSize.large} label="Working, please wait..." />
            </Dialog>
        </Stack>
    );
};
