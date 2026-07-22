import * as React from "react";
import {
    Stack, Text, IColumn, DetailsList, DetailsListLayoutMode, SelectionMode,
    PrimaryButton, DefaultButton, IconButton, Dialog, DialogType, DialogFooter, TextField, Checkbox,
    Spinner, SpinnerSize, Nav, INavLinkGroup, DetailsRow, IDetailsRowProps
} from "@fluentui/react";

import type { IConfigurationItem } from "../common/props";
import { ConfigService } from "../services/ConfigService";
import { formatError } from "../common/utils";
import { DataSource } from "../data/ds";

interface IConfigManagementProps {
    onChanged?: (items: IConfigurationItem[]) => void; // optional callback to parent if you want
}

type DialogMode = "add" | "edit" | "delete" | "error" | undefined;

const knownConfigTypes: { key: string; displayName: string }[] = [
    { key: "capStatus", displayName: "Capability Status" },
    { key: "platform", displayName: "Platform" },
    { key: "hostingEnv", displayName: "Hosting Environment" },
    { key: "connectivity", displayName: "Connectivity" },
    { key: "compliance", displayName: "Compliance" },
    { key: "codeLanguage", displayName: "Coding Language" },
    { key: "backend", displayName: "Backend" },
    { key: "customer", displayName: "Customer" },
    { key: "partner", displayName: "Relevant Partner Tag" }
];

const getForKeys = (items: IConfigurationItem[]): string[] => {
    const unique = Array.from(new Set([
        ...knownConfigTypes.map(type => type.key),
        ...items.map(i => (i.isFor ?? "").trim()).filter(Boolean)
    ]));
    unique.sort((a, b) => a.localeCompare(b));
    return unique;
};

export const ConfigManagement: React.FC<IConfigManagementProps> = ({ onChanged }) => {
    const [allItems, setAllItems] = React.useState<IConfigurationItem[]>([]);
    const [forKeys, setForKeys] = React.useState<string[]>([]);
    const [selectedFor, setSelectedFor] = React.useState<string>("");

    const [dialogMode, setDialogMode] = React.useState<DialogMode>(undefined);
    const [dialogTitle, setDialogTitle] = React.useState<string>("");
    const [dialogMessage, setDialogMessage] = React.useState<string>("");

    const [busy, setBusy] = React.useState<boolean>(false);

    // working item in dialog
    const [working, setWorking] = React.useState<IConfigurationItem | undefined>(undefined);

    const getConfig = async (): Promise<void> => {
        setBusy(true);
        try {
            const cItems = await DataSource.getConfig();
            setAllItems([...cItems ?? []].sort(
                (a, b) => a.isFor.localeCompare(b.isFor)
            ));

            const keys = getForKeys(cItems ?? []);
            setForKeys(keys);
            setSelectedFor(prev => (prev && keys.includes(prev) ? prev : (keys[0] ?? "")));
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
        // get config
        getConfig().catch((error) => {
            console.error("Unhandled promise rejection:", error);
        });
    }, []);

    const forDisplayMap = React.useMemo((): Map<string, string> => {
        const map = new Map<string, string>();

        allItems.forEach((item) => {
            const key = (item.isFor ?? "").trim();
            const displayName = (item.isForDisplayName ?? "").trim();

            if (key && !map.has(key)) {
                map.set(key, displayName || key);
            }
        });

        return map;
    }, [allItems]);

    const getForDisplayName = React.useCallback((key: string): string => {
        const normalizedKey = (key ?? "").trim();
        return forDisplayMap.get(normalizedKey)
            ?? knownConfigTypes.find(type => type.key === normalizedKey)?.displayName
            ?? key;
    }, [forDisplayMap]);

    const listForSelected = React.useMemo(() => {
        const key = (selectedFor ?? "").trim();
        return allItems
            .filter(i => (i.isFor ?? "").trim() === key)
            .sort((a, b) => (a.Title ?? "").localeCompare(b.Title ?? ""));
    }, [allItems, selectedFor]);

    const refreshLocal = (next: IConfigurationItem[]): void => {
        setAllItems(next);
        const keys = getForKeys(next);
        setForKeys(keys);
        setSelectedFor(prev => (prev && keys.includes(prev) ? prev : (keys[0] ?? "")));
        onChanged?.(next);
    };

    const openAdd = (): void => {
        if (!selectedFor) return;

        setWorking({
            Id: 0,
            Title: "",
            isFor: selectedFor,
            isForDisplayName: getForDisplayName(selectedFor),
            isActive: true,
            infoText: ""
        });

        setDialogMode("add");
        setDialogTitle(`Add: ${getForDisplayName(selectedFor)}`);
        setDialogMessage("");
    };

    const openEdit = (item: IConfigurationItem): void => {
        setWorking({ ...item });
        setDialogMode("edit");
        setDialogTitle(`Edit: ${item.Title}`);
        setDialogMessage("");
    };

    const openDelete = (item: IConfigurationItem): void => {
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

    const onSave = async (): Promise<void> => {
        if (!working) return;

        try {
            setBusy(true);

            if (!working.Title?.trim()) {
                setDialogMessage("Title is required.");
                setBusy(false);
                return;
            }

            if (dialogMode === "add") {
                const created = await ConfigService.create({
                    ...working,
                    Title: working.Title.trim(),
                    isFor: working.isFor.trim()
                });

                refreshLocal([...allItems, created]);
            }

            if (dialogMode === "edit") {
                const edited = await ConfigService.edit({
                    ...working,
                    Title: working.Title.trim(),
                    isFor: working.isFor.trim()
                });

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
            name: "Value",
            fieldName: "Title",
            minWidth: 80,
            maxWidth: 360,
            isResizable: true,
            onRender: (item: IConfigurationItem) => (
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
            key: "isActive",
            name: "Active",
            fieldName: "isActive",
            minWidth: 48,
            maxWidth: 80,
            isResizable: false,
            onRender: (item: IConfigurationItem) => <Text>{item.isActive ? "Yes" : "No"}</Text>
        },
        {
            key: "actions",
            name: "",
            minWidth: 64,
            maxWidth: 90,
            isResizable: false,
            onRender: (item: IConfigurationItem) => (
                <Stack horizontal tokens={{ childrenGap: 6 }}>
                    <IconButton
                        iconProps={{ iconName: "Edit" }}
                        title="Edit"
                        ariaLabel="Edit"
                        onClick={() => openEdit(item)}
                    />
                    <IconButton
                        iconProps={{ iconName: "Delete" }}
                        title="Delete"
                        ariaLabel="Delete"
                        onClick={() => openDelete(item)}
                    />
                </Stack>
            )
        }
    ];

    const navGroups: INavLinkGroup[] = [
        {
            links: forKeys.map((k) => ({
                key: k,
                name: `${getForDisplayName(k)} (${allItems.filter(i => i.isFor === k).length})`,
                url: "",
                onClick: () => setSelectedFor(k)
            }))
        }
    ];

    return (
        <Stack tokens={{ childrenGap: 16 }} styles={{ root: { width: "100%", marginTop: 24 } }}>

            <Stack horizontalAlign="space-between" verticalAlign="center" >
                <Text variant="xLarge">Configuration Management</Text>
                <Text variant="medium" >Adjust dropdown values within the app</Text>
            </Stack>

            {/* Left: vertical selector */}
            <Stack horizontal wrap tokens={{ childrenGap: 20 }}>

                {/* LEFT NAV */}
                <Stack styles={{ root: { flex: "0 1 220px", minWidth: 160 } }}>
                    <Nav
                        selectedKey={selectedFor}
                        groups={navGroups}
                        styles={{
                            root: {
                                borderRight: "1px solid #e1dfdd"
                            }
                        }}
                    />
                </Stack>

                {/* RIGHT CONTENT */}
                <Stack grow tokens={{ childrenGap: 12 }} styles={{ root: { minWidth: 0 } }}>
                    <Stack horizontal horizontalAlign="space-between" verticalAlign="center" wrap tokens={{ childrenGap: 12 }}>
                        <Text variant="xLarge">{getForDisplayName(selectedFor)}</Text>

                        <PrimaryButton
                            text="Add Value"
                            iconProps={{ iconName: "Add" }}
                            onClick={openAdd}
                        />
                    </Stack>

                    <Stack>
                        <DetailsList
                            items={listForSelected}
                            columns={columns}
                            selectionMode={SelectionMode.none}
                            layoutMode={DetailsListLayoutMode.fixedColumns}
                            compact
                            onRenderRow={(props?: IDetailsRowProps) => props ? (
                                <DetailsRow
                                    {...props}
                                    styles={{
                                        root: { minHeight: 64 },
                                        cell: { minHeight: 64, height: "auto" }
                                    }}
                                />
                            ) : null}
                            onItemInvoked={(item) => openEdit(item as IConfigurationItem)}
                        />
                    </Stack>

                </Stack>

            </Stack>

            {/* Dialog: Add/Edit */}
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
                            label="Value (Title)"
                            value={working.Title ?? ""}
                            onChange={(_, v) => setWorking({ ...working, Title: v ?? "" })}
                            required
                        />

                        <TextField
                            label="Info Text"
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

                        {/* show isFor but don’t let them change it (you can enable if desired) */}
                        <TextField
                            label="Category (isFor)"
                            value={working.isFor ?? ""}
                            disabled
                        />

                        <TextField
                            label="Category (Display Name)"
                            value={getForDisplayName(working.isFor ?? "")}
                            disabled
                        />
                    </Stack>
                )}

                <DialogFooter>
                    <PrimaryButton text={dialogMode === "add" ? "Create" : "Save"} onClick={onSave} disabled={busy} />
                    <DefaultButton text="Cancel" onClick={closeDialog} disabled={busy} />
                </DialogFooter>
            </Dialog>

            {/* Dialog: Delete */}
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

            {/* Dialog: ERROR */}
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

            {/* Loading Spinner Dialog */}
            <Dialog
                hidden={!busy}
                onDismiss={() => { setBusy(false) }}
                dialogContentProps={{
                    type: DialogType.normal,
                    title: "Loading...",
                    closeButtonAriaLabel: 'Close',
                }}
            >
                <Spinner size={SpinnerSize.large} label={"Working, please wait..."} />
            </Dialog>
        </Stack>
    );
};
