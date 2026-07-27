import * as React from "react";
import {
    ComboBox,
    DatePicker,
    DayOfWeek,
    DefaultButton,
    DirectionalHint,
    Dialog,
    DialogFooter,
    DialogType,
    Dropdown,
    IComboBox,
    IComboBoxOption,
    IDropdownOption,
    IPersonaProps,
    PrimaryButton,
    Stack,
    TextField
} from "@fluentui/react";
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IContractEndPointItem, IContractItem } from "../common/props";
import { DataSource } from "../data/ds";
import styles from "../Dct.module.scss";
import { Security } from "../services/Security";
import { onFormatDate, resolveUserByEmail } from "../common/utils";

export interface IContractFormProps {
    item?: IContractItem;
    context: WebPartContext;
    onSave: (item: IContractItem) => void;
    onDelete?: (itemId: number) => void;
    onCancel: () => void;
}

type JamisLookupField = "contractId" | "Title" | "customerContractCode";
const maxJamisResults = 20;

export const ContractForm: React.FC<IContractFormProps> = ({ item, context, onSave, onDelete, onCancel }) => {
    const [formData, setFormData] = React.useState<IContractItem>({
        Id: item?.Id || 0,
        capability: item?.capability,
        Title: item?.Title || "",
        contractId: item?.contractId || "",
        customerContractCode: item?.customerContractCode || "",
        customer: item?.customer || "",
        startDate: item?.startDate || "",
        endDate: item?.endDate || "",
        contractPm: item?.contractPm?.Id ? item.contractPm : undefined,
        partner: item?.partner || "",
        infoLink: item?.infoLink || "",
        ogTitle: item?.ogTitle || "",
        lobTitle: item?.lobTitle || ""
    });
    const [showDeleteConfirmation, setShowDeleteConfirmation] = React.useState(false);
    const [jamisSearchText, setJamisSearchText] = React.useState<Record<JamisLookupField, string>>({
        contractId: item?.contractId || "",
        Title: item?.Title || "",
        customerContractCode: item?.customerContractCode || ""
    });

    type CustomerType = IContractItem["customer"];
    type PartnerType = IContractItem["partner"];

    const customerOptions = React.useMemo<IDropdownOption[]>(() => DataSource.getConfigOptions("customer"), []);
    const partnerOptions = React.useMemo<IDropdownOption[]>(() => DataSource.getConfigOptions("partner"), []);
    const jamisContracts = React.useMemo<IContractEndPointItem[]>(() => DataSource.JamisContracts ?? [], []);

    const getOrgTitlesForJamisContract = (contract?: IContractEndPointItem): Pick<IContractItem, "ogTitle" | "lobTitle"> => {
        const ogTitle = (contract?.field_75 ?? "").trim();
        const ogItem = DataSource.OGs.find((og) => og.Title.toLowerCase() === ogTitle.toLowerCase());

        return {
            ogTitle,
            lobTitle: ogItem?.lob?.Title ?? ""
        };
    };

    const getJamisValue = (contract: IContractEndPointItem, field: JamisLookupField): string => {
        switch (field) {
            case "contractId":
                return contract.field_19 ?? "";
            case "Title":
                return contract.field_20 ?? "";
            case "customerContractCode":
                return contract.field_35 ?? "";
        }
    };

    const getJamisOptions = (field: JamisLookupField): IComboBoxOption[] => {
        const searchText = (jamisSearchText[field] ?? "").trim().toLowerCase();
        const matches = searchText
            ? jamisContracts.filter((contract) => getJamisValue(contract, field).toLowerCase().includes(searchText))
            : jamisContracts;

        return matches
            .map((contract): IComboBoxOption => ({
                key: contract.Id,
                text: getJamisValue(contract, field),
                data: contract
            }))
            .filter((option) => !!option.text)
            .sort((a, b) => a.text.localeCompare(b.text))
            .slice(0, maxJamisResults);
    };

    const contractIdOptions = React.useMemo<IComboBoxOption[]>(() => getJamisOptions("contractId"), [jamisContracts, jamisSearchText.contractId]);
    const contractTitleOptions = React.useMemo<IComboBoxOption[]>(() => getJamisOptions("Title"), [jamisContracts, jamisSearchText.Title]);
    const customerContractCodeOptions = React.useMemo<IComboBoxOption[]>(() => getJamisOptions("customerContractCode"), [jamisContracts, jamisSearchText.customerContractCode]);

    const handleChange = <K extends keyof IContractItem>(field: K, value: IContractItem[K]): void => {
        setFormData((prev) => ({ ...prev, [field]: value }));
    };

    const resolveProjectManager = async (email?: string, fallbackName?: string): Promise<void> => {
        const user = await resolveUserByEmail(email, fallbackName);
        if (!user) return;

        setFormData((prev) => ({
            ...prev,
            contractPm: user
        }));
    };

    const syncJamisContract = (contract: IContractEndPointItem): void => {
        const orgTitles = getOrgTitlesForJamisContract(contract);

        setFormData((prev) => ({
            ...prev,
            contractId: contract.field_19 ?? "",
            Title: contract.field_20 ?? "",
            customerContractCode: contract.field_35 ?? "",
            ...orgTitles
        }));
        setJamisSearchText({
            contractId: contract.field_19 ?? "",
            Title: contract.field_20 ?? "",
            customerContractCode: contract.field_35 ?? ""
        });

        resolveProjectManager(contract.field_21, contract.field_23).catch((error) =>
            console.warn("Unable to set Jamis Contract Project Manager", error)
        );
    };

    const handleJamisSelect = (option?: IComboBoxOption): void => {
        const selectedContract = option?.data as IContractEndPointItem | undefined;
        if (selectedContract) {
            syncJamisContract(selectedContract);
        }
    };

    const handleJamisInput = (field: JamisLookupField, value: string): void => {
        setJamisSearchText((prev) => ({ ...prev, [field]: value }));

        const normalizedValue = value.trim().toLowerCase();
        const matchingContract = normalizedValue
            ? jamisContracts.find((contract) => getJamisValue(contract, field).trim().toLowerCase() === normalizedValue)
            : undefined;

        if (matchingContract) {
            syncJamisContract(matchingContract);
            return;
        }

        setFormData((prev) => ({
            ...prev,
            [field]: value,
            ogTitle: "",
            lobTitle: ""
        }));
    };

    const renderJamisOption = (activeField: JamisLookupField): ((option?: IComboBoxOption) => JSX.Element) => (option?: IComboBoxOption): JSX.Element => {
        const contract = option?.data as IContractEndPointItem | undefined;
        const weight = (field: JamisLookupField): 400 | 600 => field === activeField ? 600 : 400;
        const empty = "-";

        return (
            <div style={{ display: "block", padding: "8px 10px", lineHeight: 1.25, minHeight: 50 }}>
                <div style={{ display: "block", fontSize: 12, fontWeight: weight("contractId") }}>
                    ID: {contract?.field_19 || empty}
                </div>
                <div style={{ display: "block", fontSize: 12, fontWeight: weight("Title"), marginTop: 2 }}>
                    Title: {contract?.field_20 || empty}
                </div>
                <div style={{ display: "block", fontSize: 12, fontWeight: weight("customerContractCode"), marginTop: 2 }}>
                    Code: {contract?.field_35 || empty}
                </div>
            </div>
        );
    };

    const jamisComboProps = {
        allowFreeform: true,
        allowFreeInput: true,
        autoComplete: "off" as const,
        openOnKeyboardFocus: true,
        dropdownWidth: 560,
        calloutProps: {
            directionalHint: DirectionalHint.bottomLeftEdge,
            gapSpace: 4,
            directionalHintFixed: true
        },
        comboBoxOptionStyles: {
            option: { height: 76, minHeight: 76, padding: 0 },
            optionText: { display: "block", whiteSpace: "normal" as const, height: "auto", overflow: "visible" as const }
        },
        styles: { root: { width: "100%" } }
    };

    const handlePerson = (items: IPersonaProps[]): void => {
        if (!items.length) {
            handleChange("contractPm", undefined);
            return;
        }

        handleChange("contractPm", {
            Id: parseInt(items[0].id!, 10),
            EMail: items[0].secondaryText!,
            Title: items[0].text!,
            JobTitle: items[0].tertiaryText
        });
    };

    const handleDelete = (): void => onDelete?.(formData.Id);
    const canEdit = Security.IsAdmin || Security.IsContributor;

    return (
        <div className={styles.capForm}>
            <section className={styles.formSection}>
                <div className={styles.formSectionHeader}>
                    <div>
                        <h3>Contract Info</h3>
                        <p>Use Contract ID, Contract Title, or Customer Contract Code to find and select a Jamis contract.</p>
                    </div>
                </div>

                <div className={`${styles.formGridThree} ${styles.contractLookupGrid}`}>
                    <ComboBox
                        label="Contract ID"
                        className={styles.formControl}
                        selectedKey={undefined}
                        text={formData.contractId ?? ""}
                        options={contractIdOptions}
                        disabled={!canEdit}
                        {...jamisComboProps}
                        onRenderOption={renderJamisOption("contractId")}
                        onInputValueChange={(value) => handleJamisInput("contractId", value)}
                        onChange={(
                            _event: React.FormEvent<IComboBox>,
                            option?: IComboBoxOption
                        ) => handleJamisSelect(option)}
                    />

                    <ComboBox
                        label="Contract Title"
                        className={styles.formControl}
                        selectedKey={undefined}
                        text={formData.Title}
                        options={contractTitleOptions}
                        disabled={!canEdit}
                        {...jamisComboProps}
                        onRenderOption={renderJamisOption("Title")}
                        onInputValueChange={(value) => handleJamisInput("Title", value)}
                        onChange={(
                            _event: React.FormEvent<IComboBox>,
                            option?: IComboBoxOption
                        ) => handleJamisSelect(option)}
                        required
                    />

                    <ComboBox
                        label="Customer Contract Code"
                        className={styles.formControl}
                        selectedKey={undefined}
                        text={formData.customerContractCode ?? ""}
                        options={customerContractCodeOptions}
                        disabled={!canEdit}
                        {...jamisComboProps}
                        onRenderOption={renderJamisOption("customerContractCode")}
                        onInputValueChange={(value) => handleJamisInput("customerContractCode", value)}
                        onChange={(
                            _event: React.FormEvent<IComboBox>,
                            option?: IComboBoxOption
                        ) => handleJamisSelect(option)}
                    />
                </div>

                <div className={styles.formGridThree}>
                    <TextField
                        label="OG"
                        className={styles.formControl}
                        value={formData.ogTitle ?? ""}
                        readOnly
                    />

                    <TextField
                        label="LOB"
                        className={styles.formControl}
                        value={formData.lobTitle ?? ""}
                        readOnly
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
                    />
                </div>

                <div className={styles.formGridTwo}>
                    <DatePicker
                        label="Capability Start Date"
                        className={styles.formControl}
                        firstDayOfWeek={DayOfWeek.Sunday}
                        disabled={!canEdit}
                        value={formData.startDate ? new Date(formData.startDate) : undefined}
                        onSelectDate={(date) => handleChange("startDate", date ? date.toISOString() : "")}
                        allowTextInput
                        formatDate={onFormatDate}
                    />

                    <DatePicker
                        label="Capability End Date"
                        className={styles.formControl}
                        firstDayOfWeek={DayOfWeek.Sunday}
                        disabled={!canEdit}
                        value={formData.endDate ? new Date(formData.endDate) : undefined}
                        onSelectDate={(date) => handleChange("endDate", date ? date.toISOString() : "")}
                        allowTextInput
                        formatDate={onFormatDate}
                    />
                </div>

                <div className={styles.formGridTwo}>
                    <PeoplePicker
                        key={`contractPm-${formData.contractPm?.EMail ?? "none"}`}
                        context={{
                            absoluteUrl: context.pageContext.web.absoluteUrl,
                            msGraphClientFactory: context.msGraphClientFactory,
                            spHttpClient: context.spHttpClient
                        }}
                        disabled={!canEdit}
                        peoplePickerWPclassName={styles.formControl}
                        defaultSelectedUsers={formData.contractPm?.EMail ? [formData.contractPm.EMail] : []}
                        titleText="KGS Contract Project Manager"
                        personSelectionLimit={1}
                        ensureUser
                        showtooltip
                        onChange={handlePerson}
                        principalTypes={[PrincipalType.User]}
                        resolveDelay={1000}
                    />
                </div>

                <div className={styles.formGridTwo}>
                    <Dropdown
                        label="Relevant Partner Tag"
                        className={styles.formControl}
                        selectedKey={formData.partner || undefined}
                        options={partnerOptions}
                        disabled={!canEdit}
                        onChange={(_, option) => {
                            if (option) handleChange("partner", option.key as PartnerType);
                        }}
                    />

                    <TextField
                        label="Contract Info Link/URL"
                        className={styles.formControl}
                        value={formData.infoLink ?? ""}
                        disabled={!canEdit}
                        onChange={(_, val) => handleChange("infoLink", val ?? "")}
                    />
                </div>
            </section>

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
        </div>
    );
};
