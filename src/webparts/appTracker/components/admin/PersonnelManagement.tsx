import * as React from "react";
import { useState, useEffect } from "react";
import {
    Stack, Text, DetailsList, type IColumn, SelectionMode, CommandBar, type ICommandBarItemProps,
    Pivot, PivotItem, IDropdownOption, DefaultButton, Dialog, DialogFooter, DialogType,
    Dropdown, PrimaryButton, Spinner, SpinnerSize,
    IconButton
} from "@fluentui/react";
import type { WebPartContext } from "@microsoft/sp-webpart-base";
import { Types } from "gd-sprest-bs";
import { Security } from "../services/Security";
import { formatError } from "../utils";
import Strings from "../../strings";
import styles from "../AppTracker.module.scss";
import { PersonnelService } from "../services/PersonnelService";
import { IPeoplePickerContext } from "@pnp/spfx-controls-react/lib/controls/peoplepicker/IPeoplePickerContext";
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { IPeoplePicker } from "../data/props";

interface IPersonnelManagementProps {
    context: WebPartContext;
}

export const PersonnelManagement: React.FunctionComponent<IPersonnelManagementProps> = ({ context }) => {
    const [visitors, setVisitors] = useState<Types.SP.GroupOData>();
    const [admins, setAdmins] = useState<Types.SP.GroupOData>();
    const [contributors, setContributors] = useState<Types.SP.GroupOData>();
    const [showSpinner, setShowSpinner] = useState(false);
    const [spinnerMessage, setSpinnerMessage] = useState<string>("Loading...");
    const [showAddUser, setShowAddUser] = useState<boolean>(false);
    const [appRole, setAppRole] = useState<string | undefined>(undefined);
    const [userId, setUserId] = useState<number | undefined>(undefined);
    const [showDialog, setShowDialog] = useState(false);
    const [dialogTitle, setDialogTitle] = useState<string>("");
    const [dialogMessage, setDialogMessage] = useState<string>("");
    const [selectedKey, setSelectedKey] = useState<string>(Strings.Groups.Visitors);
    const [selectedUser, setSelectedUser] = React.useState<Types.SP.User | undefined>(undefined);
    const [showRemoveUserDialog, setShowRemoveUserDialog] = useState(false);


    const setDialogProps = (title: string, message: string): void => {
        setShowDialog(true);
        setDialogTitle(title);
        setDialogMessage(message);
    };

    const loadPersonnel = (): void => {
        try {
            setVisitors(Security.Visitors);
            setAdmins(Security.Admins);
            setContributors(Security.Contributors);
            setShowSpinner(false);
        } catch (err) {
            console.error("Error loading personnel:", err);
            setDialogProps("Error loading personnel:", formatError(err));
        }
    };

    const refreshPersonnel = async (): Promise<void> => {
        setSpinnerMessage("Refreshing Personnel");
        setShowSpinner(true);
        Security.init().then(() =>
            loadPersonnel()
        ).catch(error => {
            setShowSpinner(false);
            setDialogProps("Error refreshing Personnel", formatError(error))
        })
    }

    useEffect(() => {
        loadPersonnel()
    }, []);

    const peoplePickerContext: IPeoplePickerContext = {
        absoluteUrl: context.pageContext.web.absoluteUrl,
        msGraphClientFactory: context.msGraphClientFactory,
        spHttpClient: context.spHttpClient
    };

    const columns: IColumn[] = [
        {
            key: "actions",
            name: "Actions",
            fieldName: "actions",
            minWidth: 75,
            maxWidth: 90,
            isResizable: false,
            onRender: (item: Types.SP.User) => (
                <IconButton
                    iconProps={{ iconName: "Delete" }}
                    title="Delete"
                    ariaLabel="Delete"
                    style={{ color: Strings.PillStyles.RedColor }}
                    onClick={() => {
                        setSelectedUser(item);
                        setShowRemoveUserDialog(true);
                    }}
                />
            )
        },
        {
            key: "title",
            name: "Name",
            fieldName: "Title",
            minWidth: 200,
            maxWidth: 300,
            isResizable: true,
            onRender: (item: Types.SP.User) => (
                <div style={{ display: "flex", alignItems: "center", height: "100%" }}>
                    <span>{item.Title}</span>
                </div>
            )
        },
        {
            key: "email",
            name: "Email",
            fieldName: "Email",
            minWidth: 250,
            maxWidth: 350,
            isResizable: true,
            onRender: (item: Types.SP.User) => (
                <div style={{ display: "flex", alignItems: "center", height: "100%" }}>
                    <span>{item.Email}</span>
                </div>
            )
        }
    ];

    const commandBarItems: ICommandBarItemProps[] = [
        {
            key: "new",
            text: "Add User",
            iconProps: { iconName: "Add" },
            onClick: () => setShowAddUser(true),
        },
        {
            key: "refresh",
            text: "Refresh",
            iconProps: { iconName: "Refresh" },
            onClick: () => refreshPersonnel()
        }
    ];

    // Decide which group to display based on selected pivot
    const getCurrentUsers = (): Types.SP.User[] => {
        switch (selectedKey) {
            case Strings.Groups.Visitors:
                return visitors?.Users?.results ?? [];
            case Strings.Groups.Admins:
                return admins?.Users?.results ?? [];
            case Strings.Groups.Contributors:
                return contributors?.Users?.results ?? [];
            default:
                return [];
        }
    };

    const groupOptions: IDropdownOption[] = [
        { key: Strings.Groups.Visitors, text: "Visitor" },
        { key: Strings.Groups.Contributors, text: "Contributor" },
        { key: Strings.Groups.Admins, text: "Owner/Admin" }
    ]

    const userExistsInGroup = (groupName: string, userId: number): boolean => {
        let users: Types.SP.User[] = [];

        switch (groupName) {
            case Strings.Groups.Visitors:
                users = visitors?.Users?.results ?? [];
                break;
            case Strings.Groups.Contributors:
                users = contributors?.Users?.results ?? [];
                break;
            case Strings.Groups.Admins:
                users = admins?.Users?.results ?? [];
                break;
        }

        return users.some(u => u.Id === userId);
    };

    const handleAddUser = async (groupName: string, userId: number): Promise<void> => {
        if (!groupName || !userId) {
            setDialogProps("Missing Information", "Please select both a user and a role.");
            return;
        }

        if (userExistsInGroup(groupName, userId)) {
            setDialogProps("Duplicate User", "This user already exists in the selected group.");
            return;
        }

        setSpinnerMessage("Adding user...");
        setShowSpinner(true);

        try {
            await PersonnelService.addUser(groupName, userId);
            await refreshPersonnel();
            setUserId(undefined);
            setAppRole(undefined);
        } catch (err) {
            const errorMessage = formatError(err);
            setDialogProps("Error Adding User", errorMessage);
            console.error("Error adding user", errorMessage);
        } finally {
            setShowSpinner(false);
            setShowAddUser(false);
        }
    };

    const handlePeoplePicker = (items: IPeoplePicker[]): void => {
        if (items && items.length > 0) {
            setUserId(parseInt(items[0].id!, 10));
        } else {
            setUserId(undefined);
        }
    };

    const handleDelete = async (): Promise<void> => {
        if (!selectedUser) return;

        setSpinnerMessage("Removing User...");
        setShowSpinner(true);

        try {
            await PersonnelService.delete(selectedKey, selectedUser.Id!);
            await refreshPersonnel();
            setSelectedUser(undefined);
        } catch (err) {
            const errorMessage = formatError(err);
            setDialogProps("Error Removing User", errorMessage);
            console.error("Error deleting user", errorMessage);
        } finally {
            setShowSpinner(false);
            setShowRemoveUserDialog(false);
        }
    };


    return (
        <Stack tokens={{ childrenGap: 20 }} style={{ marginTop: 20 }}>
            
            <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                <Text variant="xLarge">Personnel Management</Text>
            </Stack>

            <Stack horizontal horizontalAlign="space-between" verticalAlign="center" wrap tokens={{ childrenGap: 12 }}>
                <Pivot
                    selectedKey={selectedKey}
                    onLinkClick={(item) => setSelectedKey(item?.props.itemKey || Strings.Groups.Visitors)}
                >
                    <PivotItem headerText="Visitors" title="View visitors" ariaLabel="View visitors" itemKey={Strings.Groups.Visitors} />
                    <PivotItem headerText="Contributors" title="View contributors" ariaLabel="View contributors" itemKey={Strings.Groups.Contributors} />
                    <PivotItem headerText="Admins" title="View admins" ariaLabel="View admins" itemKey={Strings.Groups.Admins} />
                </Pivot>
                <CommandBar items={commandBarItems} styles={{ root: { padding: 0, background: "transparent" } }} />
            </Stack>

            <DetailsList
                items={getCurrentUsers()}
                columns={columns}
                selectionMode={SelectionMode.none}
                layoutMode={1}
                isHeaderVisible={true}
            />

            {/* Add User Dialog */}
            <Dialog
                hidden={!showAddUser}
                onDismiss={() => setShowAddUser(false)}
                dialogContentProps={{
                    type: DialogType.normal,
                    title: "Add a new User",
                    closeButtonAriaLabel: 'Close'
                }}
            >
                <Stack tokens={{ childrenGap: 16}}>
                <PeoplePicker
                    context={peoplePickerContext}
                    peoplePickerWPclassName={styles.formControl}
                    defaultSelectedUsers={[]}
                    titleText="User"
                    personSelectionLimit={1}
                    ensureUser
                    showtooltip={true}
                    required={true}
                    onChange={handlePeoplePicker}
                    principalTypes={[PrincipalType.User]}
                    resolveDelay={1000}
                    styles={{ root: { width: 300 } }}
                />
                <Dropdown
                    label="App Role"
                    required
                    options={groupOptions}
                    onChange={(event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption) => setAppRole(option?.key as string)}
                    style={{ marginBottom: 20 }}
                    styles={{ root: { width: 300 } }}
                />
                </Stack>
                <div style={{ padding: "5px 0" }}>&nbsp;</div>
                <DialogFooter>
                    <PrimaryButton text="Add User" iconProps={{ iconName: "Add" }} disabled={!appRole}
                        title="Add a new Document"
                        onClick={() => handleAddUser(appRole!, userId!)}
                    />
                    <DefaultButton onClick={() => setShowAddUser(false)} text="Cancel" title="Close Dialog Box" />
                </DialogFooter>
            </Dialog>

            {/* error message Dialog */}
            <Dialog
                hidden={!showDialog}
                onDismiss={() => setShowDialog(false)}
                dialogContentProps={{
                    type: DialogType.normal,
                    title: dialogTitle,
                    closeButtonAriaLabel: 'Close',
                    subText: dialogMessage
                }}
            >
                <DialogFooter>
                    <DefaultButton onClick={() => setShowDialog(false)} text="Close" title="Close Dialog Box" />
                </DialogFooter>
            </Dialog>


            {/* Loading Spinner Dialog */}
            <Dialog
                hidden={!showSpinner}
                onDismiss={() => setShowSpinner(false)}
                dialogContentProps={{
                    type: DialogType.normal,
                    title: "Loading...",
                    closeButtonAriaLabel: 'Close',
                }}
            >
                <Spinner size={SpinnerSize.large} label={spinnerMessage} />
            </Dialog>

            {/* Confirmation of user delete */}
            <Dialog
                hidden={!showRemoveUserDialog}
                onDismiss={() => setShowRemoveUserDialog(false)}
                dialogContentProps={{
                    type: DialogType.normal,
                    title: "Remove User",
                    closeButtonAriaLabel: 'Cancel',
                    subText: `Are you sure you want to remove ${selectedUser?.Title} from ${selectedKey}?`
                }}
            >
                <DialogFooter>
                    <PrimaryButton text="Remove"
                        className={styles.deleteButton}
                        title="Remove this User"
                        iconProps={{ iconName: "Delete" }}
                        onClick={handleDelete}
                    />
                    <DefaultButton onClick={() => setShowRemoveUserDialog(false)} text="Cancel" title="Close Dialog Box" />
                </DialogFooter>
            </Dialog>

        </Stack>
    );
};
