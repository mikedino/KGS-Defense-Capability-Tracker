import * as React from "react";
import {
    Stack, Text, TextField, Dropdown, IDropdownOption, PrimaryButton, DefaultButton, DatePicker,
    DayOfWeek, Dialog, DialogFooter, DialogType, ITextFieldProps,
    IPersonaProps
} from "@fluentui/react";
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { IPeoplePickerContext } from "@pnp/spfx-controls-react/lib/controls/peoplepicker/IPeoplePickerContext";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IContractItem } from "../data/props";
import { DataSource } from "../data/ds";
import styles from "../AppTracker.module.scss";
import { cardStackStyles } from "../ui/ComponentStyles";
import { Security } from "../services/Security";

const onFormatDate = (date?: Date): string => {
    if (!date) return "";
    return `${(date.getMonth() + 1).toString().padStart(2, "0")}/` +
        `${date.getDate().toString().padStart(2, "0")}/` +
        `${date.getFullYear().toString().slice(-2)}`;
};

export interface IContractFormProps {
    item?: IContractItem; // edit
    context: WebPartContext;
    onSave: (item: IContractItem) => void;
    onDelete?: (itemId: number) => void;
    onCancel: () => void;
}

type ContractTeamType = IContractItem["contractTeam"];

//type PeopleFields = "cor" | "ko" | "primaryPoc" | "secondaryPoc";
//const defaultBlankPerson = { Id: 0, EMail: "", Title: "" };

const onRenderCurrencyNumberInput: ITextFieldProps["onRenderInput"] = (inputProps, defaultRender) => {
    if (!defaultRender) return null;

    // Only applies when the underlying element is an <input>
    return defaultRender({
        ...inputProps,
        step: "0.01",
        min: "0"
    } as React.InputHTMLAttributes<HTMLInputElement>);
};

export const ContractForm: React.FC<IContractFormProps> = ({ item, context, onSave, onDelete, onCancel }) => {

    const [formData, setFormData] = React.useState<IContractItem>({
        Id: item?.Id || 0,
        Title: item?.Title || "",
        contractTeam: item?.contractTeam || "",
        start: item?.start || "",
        end: item?.end || "",
        contractValue: item?.contractValue || 0,
        cor: item?.cor?.Id ? item.cor : undefined,
        acor: item?.acor?.Id ? item.acor : undefined,
        ko: item?.ko?.Id ? item.ko : undefined,
        primaryPoc: item?.primaryPoc?.Id ? item.primaryPoc : undefined,
        secondaryPoc: item?.secondaryPoc?.Id ? item.secondaryPoc : undefined
    });

    const isFormValid = true;

    // ---- dropdown options (config-driven) ----
    const contractTeamOptions: IDropdownOption[] = React.useMemo(() => {
        const values = DataSource.getConfigValues("contractTeam");
        return (values ?? []).map((v) => ({ key: v, text: v }));
    }, []);

    // ---- people picker context ----
    const peoplePickerContext: IPeoplePickerContext = {
        absoluteUrl: context.pageContext.web.absoluteUrl,
        msGraphClientFactory: context.msGraphClientFactory,
        spHttpClient: context.spHttpClient
    };

    // ---- generic change helper ----
    const handleChange = <K extends keyof IContractItem>(field: K, value: IContractItem[K]): void => {
        setFormData((prev) => ({ ...prev, [field]: value }));
    };

    // ---- people pickers ----
    const handlePerson = (items: IPersonaProps[], field: keyof IContractItem): void => {
        if (!items.length) {
            handleChange(field, undefined); // allow undefined if clearing field
            return;
        }

        handleChange(field, {
            Id: parseInt(items[0].id!, 10),
            EMail: items[0].secondaryText!,
            Title: items[0].text!
        });
    };

    // ---- delete ----
    const [showDeleteConfirmation, setShowDeleteConfirmation] = React.useState(false);
    const handleDelete = (): void => onDelete?.(formData.Id);

    const isPoc = formData.cor?.Id === Security.currentUserID || formData.acor?.Id === Security.currentUserID || formData.ko?.Id === Security.currentUserID;
    const canEdit = Security.IsAdmin || isPoc;

    return (
        <Stack tokens={{ childrenGap: 20 }}>
            {/* ========================= BASIC INFORMATION  ========================= */}
            <Stack styles={{ root: [cardStackStyles.root, { paddingBottom: 24 }] }} tokens={{ childrenGap: 10 }}>
                <Stack>
                    <Text variant="large">Basic Information</Text>
                    <Text variant="small">Core details about the contract.  **Only the COR, Contracting Officer, or Admin can edit Contract information**</Text>
                </Stack>

                {/* Contract Name */}
                <TextField
                    label="Contract Name/Title"
                    className={styles.formControl}
                    value={formData.Title}
                    disabled={!canEdit}
                    onChange={(_, val) => handleChange("Title", val ?? "")}
                    required
                    maxLength={255}
                />

                <Stack horizontal wrap tokens={{ childrenGap: 12 }} verticalAlign="end">
                    <DatePicker
                        label="Start Date"
                        className={styles.formControl}
                        firstDayOfWeek={DayOfWeek.Sunday}
                        disabled={!canEdit}
                        value={formData.start ? new Date(formData.start) : undefined}
                        onSelectDate={(date) => handleChange("start", date ? date.toISOString() : "")}
                        allowTextInput
                        formatDate={onFormatDate}
                        styles={{ root: { width: 220, flexGrow: 0 } }}
                    />

                    <DatePicker
                        label="End Date"
                        className={styles.formControl}
                        firstDayOfWeek={DayOfWeek.Sunday}
                        disabled={!canEdit}
                        value={formData.end ? new Date(formData.end) : undefined}
                        onSelectDate={(date) => handleChange("end", date ? date.toISOString() : "")}
                        allowTextInput
                        formatDate={onFormatDate}
                        styles={{ root: { width: 220, flexGrow: 0 } }}
                    />
                </Stack>


                <Stack horizontal wrap tokens={{ childrenGap: 12 }} verticalAlign="end">

                    {/* Contract Lead - dropdown from config list */}
                    <Dropdown
                        label="Contract Lead"
                        disabled={!canEdit}
                        className={styles.formControl}
                        selectedKey={formData.contractTeam || undefined}
                        options={contractTeamOptions}
                        onChange={(_, option) => {
                            if (option) handleChange("contractTeam", option.key as ContractTeamType);
                        }}
                        styles={{ root: { width: 300, flexGrow: 0 } }}
                    />

                    <TextField
                        label="Contract Value"
                        disabled={!canEdit}
                        type="number"
                        prefix="$"
                        className={styles.formControl}
                        value={Number.isFinite(formData.contractValue) ? String(formData.contractValue) : ""}
                        onChange={(_, val) => {
                            const n = val === "" || val === null ? 0 : Number(val);
                            handleChange("contractValue", Number.isFinite(n) ? n : 0);
                        }}
                        onRenderInput={onRenderCurrencyNumberInput}
                        styles={{ root: { width: 200, flexGrow: 0 } }}
                    />
                </Stack>

                <Stack horizontal wrap tokens={{ childrenGap: 12 }} verticalAlign="end">
                    
                    <PeoplePicker
                        context={peoplePickerContext}
                        disabled={!canEdit}
                        peoplePickerWPclassName={styles.formControl}
                        defaultSelectedUsers={formData.cor?.EMail ? [formData.cor.EMail] : []}
                        titleText="COR"
                        personSelectionLimit={1}
                        ensureUser
                        showtooltip
                        onChange={(items) => handlePerson(items, "cor")}
                        principalTypes={[PrincipalType.User]}
                        resolveDelay={1000}
                        styles={{ root: { minWidth: 260, flexGrow: 1 } }}
                    />

                    <PeoplePicker
                        context={peoplePickerContext}
                        disabled={!canEdit}
                        peoplePickerWPclassName={styles.formControl}
                        defaultSelectedUsers={formData.acor?.EMail ? [formData.acor.EMail] : []}
                        titleText="ACOR"
                        personSelectionLimit={1}
                        ensureUser
                        showtooltip
                        onChange={(items) => handlePerson(items, "acor")}
                        principalTypes={[PrincipalType.User]}
                        resolveDelay={1000}
                        styles={{ root: { minWidth: 260, flexGrow: 1 } }}
                    />

                    <PeoplePicker
                        context={peoplePickerContext}
                        disabled={!canEdit}
                        peoplePickerWPclassName={styles.formControl}
                        defaultSelectedUsers={formData.ko?.EMail ? [formData.ko.EMail] : []}
                        titleText="Contracting Officer"
                        personSelectionLimit={1}
                        ensureUser
                        showtooltip
                        onChange={(items) => handlePerson(items, "ko")}
                        principalTypes={[PrincipalType.User]}
                        resolveDelay={1000}
                        styles={{ root: { minWidth: 260, flexGrow: 1 } }}
                    />

                </Stack>

                <Stack horizontal wrap tokens={{ childrenGap: 12 }} verticalAlign="end">
                    <PeoplePicker
                        context={peoplePickerContext}
                        disabled={!canEdit}
                        peoplePickerWPclassName={styles.formControl}
                        defaultSelectedUsers={formData.primaryPoc?.EMail ? [formData.primaryPoc.EMail] : []}
                        titleText="Contract PM"
                        personSelectionLimit={1}
                        ensureUser
                        showtooltip
                        onChange={(items) => handlePerson(items, "primaryPoc")}
                        principalTypes={[PrincipalType.User]}
                        resolveDelay={1000}
                        styles={{ root: { minWidth: 320, flexGrow: 1 } }}
                    />
                    <PeoplePicker
                        context={peoplePickerContext}
                        disabled={!canEdit}
                        peoplePickerWPclassName={styles.formControl}
                        defaultSelectedUsers={formData.secondaryPoc?.EMail ? [formData.secondaryPoc.EMail] : []}
                        titleText="Assistant Contract PM"
                        personSelectionLimit={1}
                        ensureUser
                        showtooltip
                        onChange={(items) => handlePerson(items, "secondaryPoc")}
                        principalTypes={[PrincipalType.User]}
                        resolveDelay={1000}
                        styles={{ root: { minWidth: 320, flexGrow: 1 } }}
                    />
                </Stack>

            </Stack>

            {/* ACTIONS */}
            <Stack horizontal horizontalAlign="space-between" tokens={{ childrenGap: 10 }} styles={{ root: { paddingTop: 20 } }}>
                {item ? (
                    <PrimaryButton
                        text="Delete"
                        className={styles.deleteButton}
                        style={{ width: 150 }}
                        onClick={() => setShowDeleteConfirmation(true)}
                        disabled={!canEdit}
                        title="Delete Record"
                    />
                ) : (
                    <span style={{ width: 150 }} />
                )}

                <Stack horizontal horizontalAlign="end" tokens={{ childrenGap: 10 }}>
                    <PrimaryButton
                        text="Save"
                        style={{ width: 150 }}
                        onClick={() => onSave(formData)}
                        disabled={!isFormValid || !canEdit}
                        title="Save Record"
                    />
                    <DefaultButton text="Cancel" style={{ width: 150 }} onClick={onCancel} title="Close Dialog Box" />
                </Stack>
            </Stack>

            <Dialog
                hidden={!showDeleteConfirmation}
                onDismiss={() => setShowDeleteConfirmation(false)}
                dialogContentProps={{
                    type: DialogType.normal,
                    title: "Delete Contract Entry",
                    subText: "Are you sure you want to delete this Contract Entry?",
                    closeButtonAriaLabel: "Cancel"
                }}
            >
                <DialogFooter>
                    <PrimaryButton text="Delete" className={styles.deleteButton} onClick={handleDelete} title="Delete Record" />
                    <DefaultButton text="Cancel" onClick={() => setShowDeleteConfirmation(false)} title="Close Dialog Box" />
                </DialogFooter>
            </Dialog>

        </Stack>
    );
};
