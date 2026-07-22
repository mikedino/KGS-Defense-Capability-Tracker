import * as React from "react";
import {
    DatePicker,
    DayOfWeek,
    DefaultButton,
    Dialog,
    DialogFooter,
    DialogType,
    Dropdown,
    IDropdownOption,
    IPersonaProps,
    PrimaryButton,
    Stack,
    Text,
    TextField
} from "@fluentui/react";
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { IPeoplePickerContext } from "@pnp/spfx-controls-react/lib/controls/peoplepicker/IPeoplePickerContext";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IContractItem } from "../common/props";
import { DataSource } from "../data/ds";
import styles from "../Dct.module.scss";
import { cardStackStyles } from "../ui/ComponentStyles";
import { Security } from "../services/Security";

const buildOptions = (values: string[]): IDropdownOption[] =>
    (values ?? []).map((v) => ({ key: v, text: v }));

const onFormatDate = (date?: Date): string => {
    if (!date) return "";
    return `${(date.getMonth() + 1).toString().padStart(2, "0")}/` +
        `${date.getDate().toString().padStart(2, "0")}/` +
        `${date.getFullYear().toString().slice(-2)}`;
};

export interface IContractFormProps {
    item?: IContractItem;
    context: WebPartContext;
    onSave: (item: IContractItem) => void;
    onDelete?: (itemId: number) => void;
    onCancel: () => void;
}

export const ContractForm: React.FC<IContractFormProps> = ({ item, context, onSave, onDelete, onCancel }) => {
    const [formData, setFormData] = React.useState<IContractItem>({
        Id: item?.Id || 0,
        Title: item?.Title || "",
        contractId: item?.contractId || "",
        invoice: item?.invoice || "",
        customerContractCode: item?.customerContractCode || "",
        customer: item?.customer || "",
        popStart: item?.popStart || "",
        popEnd: item?.popEnd || "",
        contractPm: item?.contractPm?.Id ? item.contractPm : undefined,
        primaryPoc: item?.primaryPoc?.Id ? item.primaryPoc : undefined,
        stakeholders: { results: item?.stakeholders?.results ?? [] },
        partner: item?.partner || "",
        infoLink: item?.infoLink || ""
    });
    const [showDeleteConfirmation, setShowDeleteConfirmation] = React.useState(false);

    type CustomerType = IContractItem["customer"];
    type PartnerType = IContractItem["partner"];

    const customerOptions = React.useMemo(() => buildOptions(DataSource.getConfigValues("customer")), []);
    const partnerOptions = React.useMemo(() => buildOptions(DataSource.getConfigValues("partner")), []);

    const peoplePickerContext: IPeoplePickerContext = {
        absoluteUrl: context.pageContext.web.absoluteUrl,
        msGraphClientFactory: context.msGraphClientFactory,
        spHttpClient: context.spHttpClient
    };

    const handleChange = <K extends keyof IContractItem>(field: K, value: IContractItem[K]): void => {
        setFormData((prev) => ({ ...prev, [field]: value }));
    };

    const handlePerson = (items: IPersonaProps[], field: "contractPm" | "primaryPoc"): void => {
        if (!items.length) {
            handleChange(field, undefined);
            return;
        }

        handleChange(field, {
            Id: parseInt(items[0].id!, 10),
            EMail: items[0].secondaryText!,
            Title: items[0].text!,
            JobTitle: items[0].tertiaryText
        });
    };

    const handleStakeholders = (items: IPersonaProps[]): void => {
        handleChange("stakeholders", {
            results: (items ?? []).map((p) => ({
                Id: parseInt(p.id!, 10),
                EMail: p.secondaryText!,
                Title: p.text!,
                JobTitle: p.tertiaryText
            }))
        });
    };

    const handleDelete = (): void => onDelete?.(formData.Id);
    const canEdit = Security.IsAdmin;

    return (
        <Stack tokens={{ childrenGap: 20 }}>
            <Stack styles={{ root: [cardStackStyles.root, { paddingBottom: 24 }] }} tokens={{ childrenGap: 12 }}>
                <Stack>
                    <Text variant="large">Contract Info</Text>
                    <Text variant="small">Contract reference fields. Lookup behavior can be wired in later.</Text>
                </Stack>

                <Stack horizontal wrap tokens={{ childrenGap: 12 }}>
                    <TextField
                        label="Contract ID"
                        className={styles.formControl}
                        value={formData.contractId ?? ""}
                        disabled={!canEdit}
                        onChange={(_, val) => handleChange("contractId", val ?? "")}
                        styles={{ root: { minWidth: 220, flexGrow: 1 } }}
                    />

                    <TextField
                        label="Task Order/Invoice ID"
                        className={styles.formControl}
                        value={formData.invoice ?? ""}
                        disabled={!canEdit}
                        onChange={(_, val) => handleChange("invoice", val ?? "")}
                        styles={{ root: { minWidth: 220, flexGrow: 1 } }}
                    />
                </Stack>

                <TextField
                    label="Contract Title"
                    className={styles.formControl}
                    value={formData.Title}
                    disabled={!canEdit}
                    onChange={(_, val) => handleChange("Title", val ?? "")}
                    required
                    maxLength={255}
                />

                <Stack horizontal wrap tokens={{ childrenGap: 12 }}>
                    <TextField
                        label="Customer Contract Code"
                        className={styles.formControl}
                        value={formData.customerContractCode ?? ""}
                        disabled={!canEdit}
                        onChange={(_, val) => handleChange("customerContractCode", val ?? "")}
                        styles={{ root: { minWidth: 240, flexGrow: 1 } }}
                    />

                    <Dropdown
                        label="Customer"
                        className={styles.formControl}
                        selectedKey={formData.customer || undefined}
                        options={customerOptions}
                        disabled={!canEdit}
                        onChange={(_, option) => {
                            if (option) handleChange("customer", option.key as CustomerType);
                        }}
                        styles={{ root: { minWidth: 220, flexGrow: 1 } }}
                    />
                </Stack>

                <Stack horizontal wrap tokens={{ childrenGap: 12 }}>
                    <DatePicker
                        label="PoP Start Date"
                        className={styles.formControl}
                        firstDayOfWeek={DayOfWeek.Sunday}
                        disabled={!canEdit}
                        value={formData.popStart ? new Date(formData.popStart) : undefined}
                        onSelectDate={(date) => handleChange("popStart", date ? date.toISOString() : "")}
                        allowTextInput
                        formatDate={onFormatDate}
                        styles={{ root: { width: 220 } }}
                    />

                    <DatePicker
                        label="PoP End Date"
                        className={styles.formControl}
                        firstDayOfWeek={DayOfWeek.Sunday}
                        disabled={!canEdit}
                        value={formData.popEnd ? new Date(formData.popEnd) : undefined}
                        onSelectDate={(date) => handleChange("popEnd", date ? date.toISOString() : "")}
                        allowTextInput
                        formatDate={onFormatDate}
                        styles={{ root: { width: 220 } }}
                    />
                </Stack>

                <Stack horizontal wrap tokens={{ childrenGap: 12 }}>
                    <PeoplePicker
                        context={peoplePickerContext}
                        disabled={!canEdit}
                        peoplePickerWPclassName={styles.formControl}
                        defaultSelectedUsers={formData.contractPm?.EMail ? [formData.contractPm.EMail] : []}
                        titleText="KGS Contract Project Manager"
                        personSelectionLimit={1}
                        ensureUser
                        showtooltip
                        onChange={(items) => handlePerson(items, "contractPm")}
                        principalTypes={[PrincipalType.User]}
                        resolveDelay={1000}
                        styles={{ root: { minWidth: 320, flexGrow: 1 } }}
                    />

                    <PeoplePicker
                        context={peoplePickerContext}
                        disabled={!canEdit}
                        peoplePickerWPclassName={styles.formControl}
                        defaultSelectedUsers={formData.primaryPoc?.EMail ? [formData.primaryPoc.EMail] : []}
                        titleText="Capability Primary POC"
                        personSelectionLimit={1}
                        ensureUser
                        showtooltip
                        onChange={(items) => handlePerson(items, "primaryPoc")}
                        principalTypes={[PrincipalType.User]}
                        resolveDelay={1000}
                        styles={{ root: { minWidth: 320, flexGrow: 1 } }}
                    />
                </Stack>

                <PeoplePicker
                    context={peoplePickerContext}
                    disabled={!canEdit}
                    peoplePickerWPclassName={styles.formControl}
                    defaultSelectedUsers={
                        formData.stakeholders?.results?.length
                            ? formData.stakeholders.results.map((p) => p.EMail)
                            : []
                    }
                    titleText="KGS Stakeholders"
                    personSelectionLimit={10}
                    ensureUser
                    showtooltip
                    onChange={handleStakeholders}
                    principalTypes={[PrincipalType.User]}
                    resolveDelay={1000}
                    styles={{ root: { minWidth: 420 } }}
                />

                <Stack horizontal wrap tokens={{ childrenGap: 12 }}>
                    <Dropdown
                        label="Relevant Partner Tag"
                        className={styles.formControl}
                        selectedKey={formData.partner || undefined}
                        options={partnerOptions}
                        disabled={!canEdit}
                        onChange={(_, option) => {
                            if (option) handleChange("partner", option.key as PartnerType);
                        }}
                        styles={{ root: { minWidth: 240, flexGrow: 1 } }}
                    />

                    <TextField
                        label="Contract Info Link/URL"
                        className={styles.formControl}
                        value={formData.infoLink ?? ""}
                        disabled={!canEdit}
                        onChange={(_, val) => handleChange("infoLink", val ?? "")}
                        styles={{ root: { minWidth: 280, flexGrow: 2 } }}
                    />
                </Stack>
            </Stack>

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
                        disabled={!canEdit}
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
