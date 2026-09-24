import * as React from "react";
import {
    ComboBox,
    Callout,
    Checkbox,
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
    MessageBar,
    MessageBarType,
    PrimaryButton,
    Stack,
    TextField
} from "@fluentui/react";
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IContractItem, IContractSourceItem } from "../common/props";
import { DataSource } from "../data/ds";
import styles from "../Dct.module.scss";
import { Security } from "../services/Security";
import { formatCurrencyInputValue, onFormatDate, parseCurrencyValue, resolveUserByEmail } from "../common/utils";
import { ContractService } from "../services/ContractService";

export interface IContractFormProps {
    item?: IContractItem;
    context: WebPartContext;
    onSave: (item: IContractItem) => void;
    onDelete?: () => void;
    onCancel: () => void;
    children?: React.ReactNode;
    allowExistingContractSave?: boolean;
}

type ContractLookupField = "contractId" | "Title" | "customerContractCode";
const maxContractSourceResults = 20;

interface IStateAutocompleteProps {
    value?: string;
    options: IComboBoxOption[];
    disabled?: boolean;
    onChange: (value: string) => void;
}

const StateAutocomplete: React.FC<IStateAutocompleteProps> = ({ value, options, disabled, onChange }) => {
    const targetRef = React.useRef<HTMLDivElement>(null);
    const selectedOption = options.find((option) => String(option.key) === (value ?? ""));
    const [inputValue, setInputValue] = React.useState<string>(selectedOption?.text ?? "");
    const [searchQuery, setSearchQuery] = React.useState<string>("");
    const [isOpen, setIsOpen] = React.useState<boolean>(false);
    const [activeIndex, setActiveIndex] = React.useState<number>(0);

    React.useEffect(() => {
        setInputValue(selectedOption?.text ?? "");
    }, [selectedOption?.key, selectedOption?.text]);

    const filteredOptions = React.useMemo(() => {
        const search = searchQuery.trim().toLowerCase();
        if (!search) return options;

        return options.filter((option) =>
            option.text.toLowerCase().includes(search) || String(option.key).toLowerCase().includes(search)
        );
    }, [searchQuery, options]);

    const selectOption = (option: IComboBoxOption): void => {
        onChange(String(option.key));
        setInputValue(option.text);
        setSearchQuery("");
        setIsOpen(false);
        setActiveIndex(0);
    };

    const resetInput = (): void => {
        setInputValue(selectedOption?.text ?? "");
        setSearchQuery("");
        setIsOpen(false);
        setActiveIndex(0);
    };

    return (
        <div ref={targetRef} className={`${styles.formControl} ${styles.stateAutocomplete}`}>
            <TextField
                label="State"
                value={inputValue}
                placeholder="Search by state name or abbreviation..."
                disabled={disabled}
                autoComplete="off"
                role="combobox"
                aria-expanded={isOpen}
                aria-autocomplete="list"
                onFocus={() => {
                    setSearchQuery("");
                    setIsOpen(true);
                }}
                onBlur={resetInput}
                onChange={(_, nextValue) => {
                    setInputValue(nextValue ?? "");
                    setSearchQuery(nextValue ?? "");
                    setIsOpen(true);
                    setActiveIndex(0);
                }}
                onKeyDown={(event) => {
                    if (event.key === "ArrowDown") {
                        event.preventDefault();
                        setIsOpen(true);
                        setActiveIndex((index) => Math.min(index + 1, Math.max(filteredOptions.length - 1, 0)));
                    } else if (event.key === "ArrowUp") {
                        event.preventDefault();
                        setActiveIndex((index) => Math.max(index - 1, 0));
                    } else if (event.key === "Enter" && isOpen && filteredOptions[activeIndex]) {
                        event.preventDefault();
                        selectOption(filteredOptions[activeIndex]);
                    } else if (event.key === "Escape") {
                        event.preventDefault();
                        resetInput();
                    }
                }}
            />

            {isOpen && !disabled && (
                <Callout
                    target={targetRef.current}
                    directionalHint={DirectionalHint.bottomLeftEdge}
                    gapSpace={2}
                    isBeakVisible={false}
                    setInitialFocus={false}
                    styles={{ root: { width: targetRef.current?.offsetWidth ?? 280 } }}
                >
                    <div className={styles.stateAutocompleteOptions} role="listbox" aria-label="State suggestions">
                        {filteredOptions.length ? filteredOptions.map((option, index) => (
                            <button
                                key={String(option.key) || "not-set"}
                                type="button"
                                role="option"
                                aria-selected={index === activeIndex}
                                className={`${styles.stateAutocompleteOption} ${index === activeIndex ? styles.stateAutocompleteOptionActive : ""}`}
                                onMouseDown={(event) => event.preventDefault()}
                                onMouseEnter={() => setActiveIndex(index)}
                                onClick={() => selectOption(option)}
                            >
                                {option.text}
                            </button>
                        )) : (
                            <div className={styles.stateAutocompleteEmpty}>No matching states</div>
                        )}
                    </div>
                </Callout>
            )}
        </div>
    );
};

export const ContractForm: React.FC<IContractFormProps> = ({
    item,
    context,
    onSave,
    onDelete,
    onCancel,
    children,
    allowExistingContractSave = false
}) => {
    const [formData, setFormData] = React.useState<IContractItem>({
        Id: item?.Id || 0,
        capability: { results: item?.capability?.results ?? [] },
        Title: item?.Title || "",
        synonyms: item?.synonyms || "",
        contractId: item?.contractId || "",
        contractType: item?.contractType || "",
        customerContractCode: item?.customerContractCode || "",
        isFlagged: item?.isFlagged ?? false,
        clearance: item?.clearance || "",
        customer: item?.customer || "",
        startDate: item?.startDate || "",
        endDate: item?.endDate || "",
        contractPm: item?.contractPm?.Id ? item.contractPm : undefined,
        partner: item?.partner || "",
        contractValue: item?.contractValue || 0,
        infoLink: item?.infoLink || "",
        ogTitle: item?.ogTitle || "",
        lobTitle: item?.lobTitle || "",
        city: item?.city || "",
        state: item?.state || "",
        country: item?.country || "US",
        location: item?.location
    });
    const [showDeleteConfirmation, setShowDeleteConfirmation] = React.useState(false);
    const [contractSourceSearchText, setContractSourceSearchText] = React.useState<Record<ContractLookupField, string>>({
        contractId: item?.contractId || "",
        Title: item?.Title || "",
        customerContractCode: item?.customerContractCode || ""
    });
    const [contractValueText, setContractValueText] = React.useState<string>(
        formatCurrencyInputValue(item?.contractValue ?? 0)
    );
    const [formMessage, setFormMessage] = React.useState<string | undefined>(undefined);

    type CustomerType = IContractItem["customer"];
    type PartnerType = IContractItem["partner"];
    type ContractType = IContractItem["contractType"];
    type ClearanceType = IContractItem["clearance"];

    const customerOptions = React.useMemo<IDropdownOption[]>(() => DataSource.getConfigOptions("customer"), []);
    const partnerOptions = React.useMemo<IDropdownOption[]>(() => DataSource.getConfigOptions("partner"), []);
    const contractTypeOptions = React.useMemo<IDropdownOption[]>(() => DataSource.getConfigOptions("contractType"), []);
    const clearanceOptions = React.useMemo<IDropdownOption[]>(() => DataSource.getConfigOptions("contractClearance"), []);
    const allStateOptions = React.useMemo<IComboBoxOption[]>(() => [
        { key: "", text: "Not set" },
        ...DataSource.getConfigOptions("state")
    ], []);
    const contractSources = React.useMemo<IContractSourceItem[]>(() => DataSource.ContractSources ?? [], []);

    // Read a normalized source field so all three lookup boxes can share filtering and rendering.
    const getContractSourceValue = (contract: IContractSourceItem, field: ContractLookupField): string => {
        switch (field) {
            case "contractId":
                return contract.contractId ?? "";
            case "Title":
                return contract.Title ?? "";
            case "customerContractCode":
                return contract.customerContractCode ?? "";
        }
    };

    // Create combo box options from the combined Jamis/CMS source collection.
    const getContractSourceOptions = (field: ContractLookupField): IComboBoxOption[] => {
        const searchText = (contractSourceSearchText[field] ?? "").trim().toLowerCase();
        const matches = searchText
            ? contractSources.filter((contract) => getContractSourceValue(contract, field).toLowerCase().includes(searchText))
            : contractSources;

        return matches
            .map((contract): IComboBoxOption => ({
                key: `${contract.source}-${contract.sourceId}`,
                text: getContractSourceValue(contract, field),
                data: contract
            }))
            .filter((option) => !!option.text)
            .sort((a, b) => a.text.localeCompare(b.text))
            .slice(0, maxContractSourceResults);
    };

    const contractIdOptions = React.useMemo<IComboBoxOption[]>(() => getContractSourceOptions("contractId"), [contractSources, contractSourceSearchText.contractId]);
    const contractTitleOptions = React.useMemo<IComboBoxOption[]>(() => getContractSourceOptions("Title"), [contractSources, contractSourceSearchText.Title]);
    const customerContractCodeOptions = React.useMemo<IComboBoxOption[]>(() => getContractSourceOptions("customerContractCode"), [contractSources, contractSourceSearchText.customerContractCode]);

    const handleChange = <K extends keyof IContractItem>(field: K, value: IContractItem[K]): void => {
        setFormMessage(undefined);
        setFormData((prev) => ({ ...prev, [field]: value }));
    };

    const handleContractValueChange = (value?: string): void => {
        const nextText = value ?? "";
        setContractValueText(nextText);
        handleChange("contractValue", parseCurrencyValue(nextText));
    };

    const handleContractValueBlur = (): void => {
        const contractValue = parseCurrencyValue(contractValueText);
        handleChange("contractValue", contractValue);
        setContractValueText(formatCurrencyInputValue(contractValue));
    };

    const applyContractData = (contract: IContractItem): void => {
        setFormData((prev) => ({
            ...prev,
            ...contract,
            capability: {
                results: contract.capability?.results ?? prev.capability?.results ?? []
            }
        }));
        setContractSourceSearchText({
            contractId: contract.contractId ?? "",
            Title: contract.Title ?? "",
            customerContractCode: contract.customerContractCode ?? ""
        });
        setContractValueText(formatCurrencyInputValue(contract.contractValue ?? 0));
    };

    const findLocalContractMatch = (
        contract: Pick<IContractItem, "Id" | "contractId" | "customerContractCode" | "Title">
    ): IContractItem | undefined =>
        ContractService.findMatchingContract(contract, DataSource.Contracts);

    const resolveProjectManager = async (email?: string, fallbackName?: string): Promise<void> => {
        const user = await resolveUserByEmail(email, fallbackName);
        if (!user) return;

        setFormData((prev) => ({
            ...prev,
            contractPm: user
        }));
    };

    // Apply the selected normalized source row to the DCT contract shape used by saves.
    const syncContractSource = (contract: IContractSourceItem): void => {
        const sourceContract: IContractItem = {
            Id: 0,
            contractId: contract.contractId ?? "",
            Title: contract.Title ?? "",
            customerContractCode: contract.customerContractCode ?? "",
            contractValue: 0,
            ogTitle: contract.ogTitle ?? "",
            lobTitle: contract.lobTitle ?? "",
            contractPm: contract.projectManager
        };
        const localContract = findLocalContractMatch(sourceContract);

        if (localContract) {
            applyContractData(localContract);
            return;
        }

        setFormData((prev) => ({
            ...prev,
            contractId: contract.contractId ?? "",
            Title: contract.Title ?? "",
            customerContractCode: contract.customerContractCode ?? "",
            ogTitle: contract.ogTitle ?? "",
            lobTitle: contract.lobTitle ?? "",
            contractPm: contract.projectManager ?? prev.contractPm
        }));
        setContractSourceSearchText({
            contractId: contract.contractId ?? "",
            Title: contract.Title ?? "",
            customerContractCode: contract.customerContractCode ?? ""
        });

        if (!contract.projectManager) {
            resolveProjectManager(contract.projectManagerEmail, contract.projectManagerName).catch((error) =>
                console.warn("Unable to set Contract Project Manager", error)
            );
        }
    };

    // Handle source dropdown selection for Jamis or CMS rows.
    const handleContractSourceSelect = (option?: IComboBoxOption): void => {
        const selectedContract = option?.data as IContractSourceItem | undefined;
        if (selectedContract) {
            syncContractSource(selectedContract);
        }
    };

    // Handle freeform source lookup text while still auto-applying exact Jamis/CMS matches.
    const handleContractSourceInput = (field: ContractLookupField, value: string): void => {
        setFormMessage(undefined);
        setContractSourceSearchText((prev) => ({ ...prev, [field]: value }));

        const normalizedValue = value.trim().toLowerCase();
        const matchingContract = normalizedValue
            ? contractSources.find((contract) => getContractSourceValue(contract, field).trim().toLowerCase() === normalizedValue)
            : undefined;

        if (matchingContract) {
            syncContractSource(matchingContract);
            return;
        }

        const localContract = findLocalContractMatch({
            Id: 0,
            contractId: field === "contractId" ? value : formData.contractId,
            Title: field === "Title" ? value : formData.Title,
            customerContractCode: field === "customerContractCode" ? value : formData.customerContractCode
        });

        if (localContract) {
            applyContractData(localContract);
            return;
        }

        setFormData((prev) => ({
            ...prev,
            [field]: value,
            ogTitle: "",
            lobTitle: ""
        }));
    };

    // Render source options with all identity fields plus the origin system to explain mixed Jamis/CMS results.
    const renderContractSourceOption = (activeField: ContractLookupField): ((option?: IComboBoxOption) => JSX.Element) => (option?: IComboBoxOption): JSX.Element => {
        const contract = option?.data as IContractSourceItem | undefined;
        const weight = (field: ContractLookupField): 400 | 600 => field === activeField ? 600 : 400;
        const empty = "-";

        return (
            <div style={{ display: "block", padding: "8px 10px", lineHeight: 1.25, minHeight: 50 }}>
                <div style={{ display: "block", fontSize: 12, fontWeight: weight("contractId") }}>
                    ID: {contract?.contractId || empty}
                </div>
                <div style={{ display: "block", fontSize: 12, fontWeight: weight("Title"), marginTop: 2 }}>
                    Title: {contract?.Title || empty}
                </div>
                <div style={{ display: "block", fontSize: 12, fontWeight: weight("customerContractCode"), marginTop: 2 }}>
                    Code: {contract?.customerContractCode || empty}
                </div>
                <div style={{ display: "block", fontSize: 11, marginTop: 2 }}>
                    Source: {contract?.sourceLabel || empty}
                </div>
            </div>
        );
    };

    // Shared combo behavior for each contract source lookup field.
    const contractSourceComboProps = {
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

    const findDuplicateContract = (): IContractItem | undefined => {
        const originalItemId = item?.Id ?? 0;

        return DataSource.Contracts.find((contract) => {
            if (originalItemId > 0 && contract.Id === originalItemId) return false;
            return ContractService.contractsMatch(formData, contract);
        });
    };

    const handleSave = (): void => {
        if (formData.country && formData.country.length !== 2) {
            setFormMessage("Country must be a two-letter ISO country code, such as US.");
            return;
        }

        const duplicateContract = findDuplicateContract();

        if (duplicateContract && !allowExistingContractSave) {
            setFormMessage(
                `A contract already exists with the same Contract Title, Customer Contract Code, or Contract ID: ${duplicateContract.Title || duplicateContract.contractId || "existing contract"}. Please use the existing contract instead.`
            );
            return;
        }

        onSave(formData);
    };

    const canEdit = Security.IsAdmin || Security.IsContributor;

    return (
        <div className={styles.capForm}>
            <section className={styles.formSection}>
                <div className={styles.formSectionHeader}>
                    <div>
                        <h3>Contract Info</h3>
                        <p>Use Contract ID, Contract Title, or Customer Contract Code to find and select a source contract.</p>
                    </div>
                </div>

                {formMessage && (
                    <MessageBar messageBarType={MessageBarType.warning} isMultiline>
                        {formMessage}
                    </MessageBar>
                )}

                <div className={styles.contractFormHeaderGrid}>
                    <ComboBox
                        label="Contract Title"
                        className={styles.formControl}
                        selectedKey={undefined}
                        text={formData.Title}
                        options={contractTitleOptions}
                        placeholder="Search..."
                        disabled={!canEdit}
                        {...contractSourceComboProps}
                        onRenderOption={renderContractSourceOption("Title")}
                        onInputValueChange={(value) => handleContractSourceInput("Title", value)}
                        onChange={(
                            _event: React.FormEvent<IComboBox>,
                            option?: IComboBoxOption
                        ) => handleContractSourceSelect(option)}
                        required
                    />

                    <Checkbox
                        label="Flag"
                        checked={formData.isFlagged ?? false}
                        disabled={!canEdit}
                        className={styles.contractFlagCheckbox}
                        title="Flag this contract to hide details from general users"
                        ariaLabel="Flag this contract to hide details from general users"
                        onChange={(_, checked) => handleChange("isFlagged", checked ?? false)}
                    />
                </div>

                <div className={styles.formGridThree}>
                    <ComboBox
                        label="Contract ID"
                        className={styles.formControl}
                        selectedKey={undefined}
                        text={formData.contractId ?? ""}
                        options={contractIdOptions}
                        placeholder="Search..."
                        disabled={!canEdit}
                        {...contractSourceComboProps}
                        onRenderOption={renderContractSourceOption("contractId")}
                        onInputValueChange={(value) => handleContractSourceInput("contractId", value)}
                        onChange={(
                            _event: React.FormEvent<IComboBox>,
                            option?: IComboBoxOption
                        ) => handleContractSourceSelect(option)}
                    />

                    <ComboBox
                        label="Customer Contract Code"
                        className={styles.formControl}
                        selectedKey={undefined}
                        text={formData.customerContractCode ?? ""}
                        options={customerContractCodeOptions}
                        placeholder="Search..."
                        disabled={!canEdit}
                        {...contractSourceComboProps}
                        onRenderOption={renderContractSourceOption("customerContractCode")}
                        onInputValueChange={(value) => handleContractSourceInput("customerContractCode", value)}
                        onChange={(
                            _event: React.FormEvent<IComboBox>,
                            option?: IComboBoxOption
                        ) => handleContractSourceSelect(option)}
                    />

                    <Dropdown
                        label="Contract Type"
                        className={styles.formControl}
                        selectedKey={formData.contractType || undefined}
                        options={contractTypeOptions}
                        disabled={!canEdit}
                        onChange={(_, option) => {
                            if (option) handleChange("contractType", option.key as ContractType);
                        }}
                    />

                    <TextField
                        label="Contract Value"
                        className={styles.formControl}
                        prefix="$"
                        value={contractValueText}
                        disabled={!canEdit}
                        onChange={(_, val) => handleContractValueChange(val)}
                        onBlur={handleContractValueBlur}
                    />

                    <Dropdown
                        label="Clearance Level"
                        className={styles.formControl}
                        selectedKey={formData.clearance || undefined}
                        options={clearanceOptions}
                        disabled={!canEdit}
                        onChange={(_, option) => {
                            if (option) handleChange("clearance", option.key as ClearanceType);
                        }}
                    />

                    <TextField
                        label="Synonyms"
                        className={styles.formControl}
                        value={formData.synonyms ?? ""}
                        maxLength={255}
                        disabled={!canEdit}
                        onChange={(_, val) => handleChange("synonyms", val ?? "")}
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

                <div className={styles.formGridThree}>
                    <Dropdown
                        label="Partner Tag"
                        className={styles.formControl}
                        selectedKey={formData.partner || undefined}
                        options={partnerOptions}
                        disabled={!canEdit}
                        onChange={(_, option) => {
                            if (option) handleChange("partner", option.key as PartnerType);
                        }}
                    />

                    <DatePicker
                        label="Start Date"
                        className={styles.formControl}
                        firstDayOfWeek={DayOfWeek.Sunday}
                        disabled={!canEdit}
                        value={formData.startDate ? new Date(formData.startDate) : undefined}
                        onSelectDate={(date) => handleChange("startDate", date ? date.toISOString() : "")}
                        allowTextInput
                        formatDate={onFormatDate}
                    />

                    <DatePicker
                        label="End Date"
                        className={styles.formControl}
                        firstDayOfWeek={DayOfWeek.Sunday}
                        disabled={!canEdit}
                        value={formData.endDate ? new Date(formData.endDate) : undefined}
                        onSelectDate={(date) => handleChange("endDate", date ? date.toISOString() : "")}
                        allowTextInput
                        formatDate={onFormatDate}
                    />
                </div>

                <div className={styles.formGridThree}>
                    <TextField
                        label="City"
                        className={styles.formControl}
                        value={formData.city ?? ""}
                        maxLength={255}
                        disabled={!canEdit}
                        onChange={(_, val) => handleChange("city", val ?? "")}
                    />

                    <StateAutocomplete
                        value={formData.state}
                        options={allStateOptions}
                        disabled={!canEdit}
                        onChange={(state) => handleChange("state", state)}
                    />

                    <TextField
                        label="Country"
                        description="Two-letter ISO country code"
                        className={styles.formControl}
                        value={formData.country ?? ""}
                        maxLength={2}
                        disabled={!canEdit}
                        onChange={(_, val) => handleChange(
                            "country",
                            (val ?? "").replace(/[^a-z]/gi, "").toUpperCase().slice(0, 2)
                        )}
                    />
                </div>

                <div className={styles.formGridWide}>
                    <TextField
                        label="Contract Info Link/URL"
                        className={styles.formControl}
                        value={formData.infoLink ?? ""}
                        disabled={!canEdit}
                        onChange={(_, val) => handleChange("infoLink", val ?? "")}
                    />

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
            </section>

            {children}

            <Stack horizontal horizontalAlign="space-between" tokens={{ childrenGap: 10 }} styles={{ root: { paddingTop: 20 } }}>
                {onDelete ? (
                    <PrimaryButton
                        text="Remove"
                        className={styles.deleteButton}
                        style={{ width: 150 }}
                        onClick={() => setShowDeleteConfirmation(true)}
                        disabled={!canEdit}
                        title="Remove Contract Relationship"
                    />
                ) : (
                    <span style={{ width: 150 }} />
                )}

                <Stack horizontal horizontalAlign="end" tokens={{ childrenGap: 10 }}>
                    <PrimaryButton
                        text="Save"
                        style={{ width: 150 }}
                        onClick={handleSave}
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
                    title: "Remove Contract Relationship",
                    subText: "Remove this contract from the capability?",
                    closeButtonAriaLabel: "Cancel"
                }}
            >
                <DialogFooter>
                    <PrimaryButton text="Remove" className={styles.deleteButton} onClick={onDelete} title="Remove Contract Relationship" />
                    <DefaultButton text="Cancel" onClick={() => setShowDeleteConfirmation(false)} title="Close Dialog Box" />
                </DialogFooter>
            </Dialog>
        </div>
    );
};
