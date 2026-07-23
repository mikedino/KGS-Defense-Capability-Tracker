import * as React from "react";
import {
    DefaultButton,
    Dialog,
    DialogFooter,
    DialogType,
    DetailsList,
    Dropdown,
    IColumn,
    IDropdownOption,
    IconButton,
    IPersonaProps,
    PrimaryButton,
    SelectionMode,
    Stack,
    TagPicker,
    Text,
    TextField
} from "@fluentui/react";
import { ITag } from "@fluentui/react/lib/Pickers";
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { IPeoplePickerContext } from "@pnp/spfx-controls-react/lib/controls/peoplepicker/IPeoplePickerContext";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ICapabilityContractDraft, ICapabilityItem, ICapFormSaveResult, IContractItem, IOpportunityItem, IOppNetTagValue, licenseReqdChoices } from "../common/props";
import { DataSource } from "../data/ds";
import { licenseReqdOptions } from "../ui/DropdownChoices";
import { cardStackStyles } from "../ui/ComponentStyles";
import styles from "../Dct.module.scss";
import { ContractForm } from "./ContractForm";
import { formatDate } from "../common/utils";
import Strings from "../common/strings";
import { buildOppNetItemUrl } from "../common/tagUtils";

export interface ICapFormProps {
    item?: ICapabilityItem;
    context: WebPartContext;
    onSave: (item: ICapFormSaveResult) => void;
    onDelete?: (itemId: number) => void;
    onCancel: () => void;
}

interface IOppNetTag extends IOppNetTagValue {
    key: string;
    name: string;
}

export const CapForm: React.FC<ICapFormProps> = ({ item, context, onSave, onDelete, onCancel }) => {
    const [submitted, setSubmitted] = React.useState<boolean>(false);
    const [titleError, setTitleError] = React.useState<string | undefined>(undefined);
    const [showDeleteConfirmation, setShowDeleteConfirmation] = React.useState(false);
    const [opportunities, setOpportunities] = React.useState<IOpportunityItem[]>(DataSource.Opportunities ?? []);
    const [oppNetLoading, setOppNetLoading] = React.useState<boolean>(false);
    const [capContracts, setCapContracts] = React.useState<ICapabilityContractDraft[]>(
        () => item?.Id
            ? DataSource.Contracts
                .filter((contract) => contract.capability?.Id === item.Id)
                .map((contract) => ({ ...contract }))
            : []
    );
    const [deletedContractIds, setDeletedContractIds] = React.useState<number[]>([]);
    const [showContractDialog, setShowContractDialog] = React.useState<boolean>(false);
    const [selectedContract, setSelectedContract] = React.useState<ICapabilityContractDraft | undefined>(undefined);

    type CapStatusType = ICapabilityItem["capStatus"];
    type PlatformType = ICapabilityItem["platform"];
    type HostingEnvType = ICapabilityItem["hostingEnv"];
    type ConnectivityType = ICapabilityItem["connectivity"];
    type ComplianceType = ICapabilityItem["compliance"];
    type CodeLanguageType = ICapabilityItem["codeLanguage"];
    type BackendType = ICapabilityItem["backend"];

    const [formData, setFormData] = React.useState<ICapabilityItem>({
        Id: item?.Id || 0,
        Title: item?.Title || "",
        description: item?.description || "",
        capabilities: item?.capabilities || "",
        link: item?.link || "",
        capStatus: item?.capStatus || "",
        primaryPoc: item?.primaryPoc?.Id ? item.primaryPoc : undefined,
        stakeholders: { results: item?.stakeholders?.results ?? [] },
        notes: item?.notes || "",
        platform: item?.platform || "",
        hostingEnv: item?.hostingEnv || "",
        connectivity: item?.connectivity || "",
        compliance: item?.compliance || "",
        licenseReqd: item?.licenseReqd ?? "No",
        licenseReqmts: item?.licenseReqmts || "",
        extensibility: item?.extensibility || "",
        serverReqmts: item?.serverReqmts || "",
        codeLanguage: item?.codeLanguage || "",
        backend: item?.backend || "",
        oppNetTagsJson: item?.oppNetTagsJson || "",
        oppNetTags: item?.oppNetTags ?? []
    });

    const capStatusOptions = React.useMemo<IDropdownOption[]>(() => DataSource.getConfigOptions("capabilityStatus"), []);
    const platformOptions = React.useMemo<IDropdownOption[]>(() => DataSource.getConfigOptions("platform"), []);
    const hostingEnvOptions = React.useMemo<IDropdownOption[]>(() => DataSource.getConfigOptions("hostingEnvironment"), []);
    const connectivityOptions = React.useMemo<IDropdownOption[]>(() => DataSource.getConfigOptions("connectivity"), []);
    const complianceOptions = React.useMemo<IDropdownOption[]>(() => DataSource.getConfigOptions("compliance"), []);
    const codeLanguageOptions = React.useMemo<IDropdownOption[]>(() => DataSource.getConfigOptions("codingLanguage"), []);
    const backendOptions = React.useMemo<IDropdownOption[]>(() => DataSource.getConfigOptions("backend"), []);
    const peoplePickerContext: IPeoplePickerContext = {
        absoluteUrl: context.pageContext.web.absoluteUrl,
        msGraphClientFactory: context.msGraphClientFactory,
        spHttpClient: context.spHttpClient
    };

    React.useEffect(() => {
        let mounted = true;

        if (DataSource.Opportunities.length) {
            setOpportunities(DataSource.Opportunities);
            return;
        }

        setOppNetLoading(true);
        DataSource.getOpportunities()
            .then((items) => {
                if (mounted) setOpportunities(items);
            })
            .catch((error) => console.error("Error loading OppNet opportunities:", error))
            .finally(() => {
                if (mounted) setOppNetLoading(false);
            });

        return () => {
            mounted = false;
        };
    }, []);

    const handleChange = <K extends keyof ICapabilityItem>(field: K, value: ICapabilityItem[K]): ICapabilityItem[K] => {
        setFormData((prev) => ({ ...prev, [field]: value }));
        return value;
    };

    const validateTitle = (value?: string): string | undefined =>
        !value?.trim() ? "Capability name is required" : undefined;

    const cleanText = (value?: string | null): string => (value ?? "").trim();

    const toOppNetTag = (item: IOpportunityItem | IOppNetTagValue): IOppNetTag => {
        const isOpportunity = "Id" in item;
        const id = isOpportunity ? item.Id : item.id;
        const title = cleanText(isOpportunity ? item.Title : item.title);
        const customer = cleanText(isOpportunity ? item.Customer : item.customer);
        const status = cleanText(isOpportunity ? item.Status : item.status);

        return {
            id,
            title,
            key: id.toString(),
            name: title,
            customer,
            status,
            url: "url" in item ? item.url : buildOppNetItemUrl(Strings.Sites.oppNet.url, id)
        };
    };

    const getSelectedOppNetTags = (): IOppNetTag[] =>
        (formData.oppNetTags ?? []).map(toOppNetTag);

    const resolveOppNetSuggestions = (filterText: string, selectedItems?: IOppNetTag[]): IOppNetTag[] => {
        const selectedIds = new Set((selectedItems ?? []).map(tag => tag.id));
        const search = filterText.trim().toLowerCase();

        return opportunities
            .filter(opp => !selectedIds.has(opp.Id))
            .filter(opp => !!cleanText(opp.Title))
            .filter(opp => {
                if (!search) return true;
                return [opp.Title, opp.Customer, opp.Status, opp.Id.toString()]
                    .some(value => cleanText(value).toLowerCase().includes(search));
            })
            .sort((a, b) => cleanText(a.Title).localeCompare(cleanText(b.Title)))
            .slice(0, 20)
            .map(toOppNetTag);
    };

    const handleOppNetTagsChange = (items?: IOppNetTag[]): void => {
        handleChange("oppNetTags", (items ?? []).map(tag => ({
                id: tag.id,
                title: tag.title,
                customer: tag.customer,
                status: tag.status,
                url: tag.url
        })));
    };

    const handlePrimaryPoc = (items: IPersonaProps[]): void => {
        if (!items.length) {
            handleChange("primaryPoc", undefined);
            return;
        }

        handleChange("primaryPoc", {
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

    const createTempContractId = (): string =>
        `new-${Date.now()}-${Math.random().toString(36).slice(2)}`;

    const getContractRowKey = (contract: ICapabilityContractDraft): string =>
        contract.Id > 0 ? `id-${contract.Id}` : contract.tempId ?? "";

    const handleAddContract = (): void => {
        setSelectedContract(undefined);
        setShowContractDialog(true);
    };

    const handleEditContract = (contract: ICapabilityContractDraft): void => {
        setSelectedContract(contract);
        setShowContractDialog(true);
    };

    const handleRemoveContract = (contract: ICapabilityContractDraft): void => {
        if (contract.Id > 0) {
            setDeletedContractIds((prev) => prev.indexOf(contract.Id) >= 0 ? prev : [...prev, contract.Id]);
        }

        const key = getContractRowKey(contract);
        setCapContracts((prev) => prev.filter((row) => getContractRowKey(row) !== key));
    };

    const handleSaveContract = (contract: IContractItem): void => {
        const nextContract: ICapabilityContractDraft = {
            ...contract,
            tempId: selectedContract?.tempId ?? createTempContractId()
        };

        if (selectedContract) {
            const selectedKey = getContractRowKey(selectedContract);
            setCapContracts((prev) =>
                prev.map((row) => getContractRowKey(row) === selectedKey ? nextContract : row)
            );
        } else {
            setCapContracts((prev) => [...prev, nextContract]);
        }

        setSelectedContract(undefined);
        setShowContractDialog(false);
    };

    const contractColumns = React.useMemo<IColumn[]>(() => [
        {
            key: "Title",
            name: "Contract Title",
            fieldName: "Title",
            minWidth: 220,
            isResizable: true
        },
        {
            key: "customer",
            name: "Customer",
            fieldName: "customer",
            minWidth: 120,
            maxWidth: 180,
            isResizable: true
        },
        {
            key: "endDate",
            name: "End Date",
            fieldName: "endDate",
            minWidth: 90,
            maxWidth: 120,
            isResizable: true,
            onRender: (contract: ICapabilityContractDraft) => contract.endDate ? formatDate(contract.endDate) : ""
        },
        {
            key: "contractPm",
            name: "PM",
            fieldName: "contractPm",
            minWidth: 140,
            maxWidth: 220,
            isResizable: true,
            onRender: (contract: ICapabilityContractDraft) => contract.contractPm?.Title ?? ""
        },
        {
            key: "actions",
            name: "",
            minWidth: 72,
            maxWidth: 72,
            onRender: (contract: ICapabilityContractDraft) => (
                <Stack horizontal tokens={{ childrenGap: 4 }}>
                    <IconButton
                        iconProps={{ iconName: "Edit" }}
                        title="Edit related contract"
                        ariaLabel="Edit related contract"
                        onClick={() => handleEditContract(contract)}
                    />
                    <IconButton
                        iconProps={{ iconName: "Delete" }}
                        title="Remove related contract"
                        ariaLabel="Remove related contract"
                        onClick={() => handleRemoveContract(contract)}
                    />
                </Stack>
            )
        }
    ], []);

    React.useEffect(() => {
        setTitleError(validateTitle(formData.Title));
    }, [formData.Title]);

    const handleSubmit = (ev: React.FormEvent<HTMLFormElement>): boolean => {
        ev.preventDefault();
        setSubmitted(true);

        const nextTitleError = validateTitle(formData.Title);
        setTitleError(nextTitleError);

        if (nextTitleError) {
            return false;
        }

        onSave({ capability: formData, contracts: capContracts, deletedContractIds });
        return true;
    };

    const handleDelete = (): void => onDelete?.(formData.Id);

    return (
        <>
        <form onSubmit={handleSubmit}>
            <Stack tokens={{ childrenGap: 20 }}>
                <Stack styles={cardStackStyles} tokens={{ childrenGap: 12 }}>
                    <Stack>
                        <Text variant="large">General Info</Text>
                        <Text variant="small">Core details about the capability</Text>
                    </Stack>

                    <TextField
                        label="Capability Name"
                        className={styles.formControl}
                        value={formData.Title}
                        onChange={(_, val) => handleChange("Title", val ?? "")}
                        required
                        maxLength={255}
                        errorMessage={submitted ? titleError : undefined}
                    />

                    <TextField
                        label="Capability Description"
                        className={styles.formControl}
                        multiline
                        autoAdjustHeight
                        value={formData.description ?? ""}
                        onChange={(_, val) => handleChange("description", val ?? "")}
                    />

                    <TextField
                        label="Technical Capabilities"
                        className={styles.formControl}
                        multiline
                        autoAdjustHeight
                        value={formData.capabilities ?? ""}
                        onChange={(_, val) => handleChange("capabilities", val ?? "")}
                    />

                    <TextField
                        label="Link/URL"
                        className={styles.formControl}
                        value={formData.link ?? ""}
                        maxLength={255}
                        onChange={(_, val) => handleChange("link", val ?? "")}
                    />

                    <Dropdown
                        label="Capability Status"
                        className={styles.formControl}
                        selectedKey={formData.capStatus || undefined}
                        options={capStatusOptions}
                        onChange={(_, option) => {
                            if (option) handleChange("capStatus", option.key as CapStatusType);
                        }}
                        styles={{ root: { width: 240, maxWidth: "100%" } }}
                    />

                    <TextField
                        label="Additional Notes"
                        className={styles.formControl}
                        multiline
                        autoAdjustHeight
                        value={formData.notes ?? ""}
                        onChange={(_, val) => handleChange("notes", val ?? "")}
                    />

                    <Stack horizontal wrap tokens={{ childrenGap: 12 }}>
                        <PeoplePicker
                            context={peoplePickerContext}
                            peoplePickerWPclassName={styles.formControl}
                            defaultSelectedUsers={formData.primaryPoc?.EMail ? [formData.primaryPoc.EMail] : []}
                            titleText="Capability Primary POC"
                            personSelectionLimit={1}
                            ensureUser
                            showtooltip
                            onChange={handlePrimaryPoc}
                            principalTypes={[PrincipalType.User]}
                            resolveDelay={1000}
                            styles={{ root: { minWidth: 320, flexGrow: 1 } }}
                        />

                        <PeoplePicker
                            context={peoplePickerContext}
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
                            styles={{ root: { minWidth: 320, flexGrow: 1 } }}
                        />
                    </Stack>
                </Stack>

                <Stack styles={cardStackStyles} tokens={{ childrenGap: 12 }}>
                    <Stack>
                        <Text variant="large">Technical Info</Text>
                        <Text variant="small">Platform, compliance, licensing, and implementation details</Text>
                    </Stack>

                    <Stack horizontal wrap tokens={{ childrenGap: 12 }}>
                        <Dropdown
                            label="Platform"
                            className={styles.formControl}
                            selectedKey={formData.platform || undefined}
                            options={platformOptions}
                            onChange={(_, option) => {
                                if (option) handleChange("platform", option.key as PlatformType);
                            }}
                            styles={{ root: { minWidth: 220, flexGrow: 1 } }}
                        />

                        <Dropdown
                            label="Hosting Environment"
                            className={styles.formControl}
                            selectedKey={formData.hostingEnv || undefined}
                            options={hostingEnvOptions}
                            onChange={(_, option) => {
                                if (option) handleChange("hostingEnv", option.key as HostingEnvType);
                            }}
                            styles={{ root: { minWidth: 220, flexGrow: 1 } }}
                        />
                    </Stack>

                    <Stack horizontal wrap tokens={{ childrenGap: 12 }}>
                        <Dropdown
                            label="Connectivity"
                            className={styles.formControl}
                            selectedKey={formData.connectivity || undefined}
                            options={connectivityOptions}
                            onChange={(_, option) => {
                                if (option) handleChange("connectivity", option.key as ConnectivityType);
                            }}
                            styles={{ root: { minWidth: 220, flexGrow: 1 } }}
                        />

                        <Dropdown
                            label="Compliance"
                            className={styles.formControl}
                            selectedKey={formData.compliance || undefined}
                            options={complianceOptions}
                            onChange={(_, option) => {
                                if (option) handleChange("compliance", option.key as ComplianceType);
                            }}
                            styles={{ root: { minWidth: 220, flexGrow: 1 } }}
                        />
                    </Stack>

                    <Stack horizontal wrap tokens={{ childrenGap: 12 }}>
                        <Dropdown
                            label="License Required?"
                            required
                            className={styles.formControl}
                            selectedKey={formData.licenseReqd || undefined}
                            options={licenseReqdOptions(false)}
                            onChange={(_, option) => {
                                if (option) handleChange("licenseReqd", option.key as licenseReqdChoices);
                            }}
                            styles={{ root: { width: 220 } }}
                        />

                        <Dropdown
                            label="Coding Language"
                            className={styles.formControl}
                            selectedKey={formData.codeLanguage || undefined}
                            options={codeLanguageOptions}
                            onChange={(_, option) => {
                                if (option) handleChange("codeLanguage", option.key as CodeLanguageType);
                            }}
                            styles={{ root: { minWidth: 220, flexGrow: 1 } }}
                        />

                        <Dropdown
                            label="Backend"
                            className={styles.formControl}
                            selectedKey={formData.backend || undefined}
                            options={backendOptions}
                            onChange={(_, option) => {
                                if (option) handleChange("backend", option.key as BackendType);
                            }}
                            styles={{ root: { minWidth: 220, flexGrow: 1 } }}
                        />
                    </Stack>

                    <TextField
                        label="Licensing Requirements"
                        className={styles.formControl}
                        multiline
                        autoAdjustHeight
                        value={formData.licenseReqmts ?? ""}
                        onChange={(_, val) => handleChange("licenseReqmts", val ?? "")}
                    />

                    <TextField
                        label="APIs/Extensibility"
                        className={styles.formControl}
                        multiline
                        autoAdjustHeight
                        value={formData.extensibility ?? ""}
                        onChange={(_, val) => handleChange("extensibility", val ?? "")}
                    />

                    <TextField
                        label="Server Requirements"
                        className={styles.formControl}
                        value={formData.serverReqmts ?? ""}
                        onChange={(_, val) => handleChange("serverReqmts", val ?? "")}
                    />
                </Stack>

                <Stack styles={cardStackStyles} tokens={{ childrenGap: 12 }}>
                    <Stack>
                        <Text variant="large">Contracts</Text>
                        <Text variant="small">Add one or more contract relationships for this capability</Text>
                    </Stack>

                    <Stack horizontal horizontalAlign="end">
                        <PrimaryButton
                            text="Add Contract"
                            iconProps={{ iconName: "Add" }}
                            onClick={handleAddContract}
                        />
                    </Stack>

                    {capContracts.length ? (
                        <DetailsList
                            items={capContracts}
                            columns={contractColumns}
                            selectionMode={SelectionMode.none}
                            getKey={(contract) => getContractRowKey(contract as ICapabilityContractDraft)}
                        />
                    ) : (
                        <Text styles={{ root: { color: "gray", fontStyle: "italic" } }}>
                            No contracts have been added yet.
                        </Text>
                    )}
                </Stack>

                <Stack styles={cardStackStyles} tokens={{ childrenGap: 12 }}>
                    <Stack>
                        <Text variant="large">Tagging</Text>
                        <Text variant="small">Apply relevant tags to help with search and data mining</Text>
                    </Stack>

                    <TagPicker
                        label="OppNet Tags"
                        className={styles.formControl}
                        selectedItems={getSelectedOppNetTags()}
                        onResolveSuggestions={(filter, selectedItems) =>
                            resolveOppNetSuggestions(filter, selectedItems as IOppNetTag[])
                        }
                        onEmptyResolveSuggestions={(selectedItems) =>
                            resolveOppNetSuggestions("", selectedItems as IOppNetTag[])
                        }
                        onChange={(items?: ITag[]) => handleOppNetTagsChange(items as IOppNetTag[])}
                        getTextFromItem={(tag) => tag.name}
                        pickerSuggestionsProps={{
                            suggestionsHeaderText: "Matching OppNet opportunities",
                            noResultsFoundText: oppNetLoading ? "Loading OppNet opportunities..." : "No OppNet opportunities found",
                            resultsMaximumNumber: 20
                        }}
                        inputProps={{
                            placeholder: oppNetLoading ? "Loading OppNet opportunities..." : "Search OppNet opportunities"
                        }}
                        onRenderSuggestionsItem={(tag) => {
                            const oppTag = tag as IOppNetTag;
                            return (
                            <div style={{ padding: "8px 10px", textAlign: "left", lineHeight: 1.25, whiteSpace: "normal" }}>
                                <Text block styles={{ root: { fontWeight: 600, whiteSpace: "normal" } }}>{oppTag.title}</Text>
                                <Text block variant="small" styles={{ root: { whiteSpace: "normal" } }}>
                                    {[oppTag.customer, oppTag.status].filter(Boolean).join(" | ")}
                                </Text>
                            </div>
                        )}}
                        styles={{ root: { minWidth: 280, maxWidth: 520 } }}
                    />
                </Stack>

                <Stack horizontal horizontalAlign="space-between" tokens={{ childrenGap: 10 }} styles={{ root: { paddingTop: 20 } }}>
                    {item ? (
                        <PrimaryButton
                            text="Delete"
                            className={styles.deleteButton}
                            style={{ width: 150 }}
                            onClick={() => setShowDeleteConfirmation(true)}
                            title="Delete Record"
                        />
                    ) : (
                        <span style={{ width: 150 }} />
                    )}

                    <Stack horizontal horizontalAlign="end" tokens={{ childrenGap: 10 }}>
                        <PrimaryButton type="submit" text="Save" style={{ width: 150 }} title="Save Record" />
                        <DefaultButton text="Cancel" style={{ width: 150 }} onClick={onCancel} title="Close Dialog Box" />
                    </Stack>
                </Stack>

                <Dialog
                    hidden={!showDeleteConfirmation}
                    onDismiss={() => setShowDeleteConfirmation(false)}
                    dialogContentProps={{
                        type: DialogType.normal,
                        title: "Delete Capability Entry",
                        subText: "Are you sure you want to delete this capability entry?",
                        closeButtonAriaLabel: "Cancel"
                    }}
                >
                    <DialogFooter>
                        <PrimaryButton text="Delete" className={styles.deleteButton} onClick={handleDelete} title="Delete Record" />
                        <DefaultButton text="Cancel" onClick={() => setShowDeleteConfirmation(false)} title="Close Dialog Box" />
                    </DialogFooter>
                </Dialog>
            </Stack>
        </form>
        <Dialog
            hidden={!showContractDialog}
            onDismiss={() => {
                setSelectedContract(undefined);
                setShowContractDialog(false);
            }}
            dialogContentProps={{
                type: DialogType.largeHeader,
                title: selectedContract ? "Edit Related Contract" : "Add Related Contract",
                showCloseButton: true
            }}
            modalProps={{
                isBlocking: true,
                styles: { main: { width: "90vw !important", maxWidth: "900px !important", minWidth: "675px !important" } }
            }}
        >
            <ContractForm
                item={selectedContract}
                context={context}
                onSave={handleSaveContract}
                onDelete={() => {
                    if (selectedContract) {
                        handleRemoveContract(selectedContract);
                    }
                    setSelectedContract(undefined);
                    setShowContractDialog(false);
                }}
                onCancel={() => {
                    setSelectedContract(undefined);
                    setShowContractDialog(false);
                }}
            />
        </Dialog>
        </>
    );
};
