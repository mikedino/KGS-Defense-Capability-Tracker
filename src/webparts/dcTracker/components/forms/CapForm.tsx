import * as React from "react";
import {
    DefaultButton,
    Dialog,
    DialogFooter,
    DialogType,
    DetailsList,
    DetailsRow,
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
import { ICapabilityContractDraft, ICapabilityItem, ICapFormSaveResult, IContractItem, IOpportunityItem, IPastPerformanceItem, IProposalItem, licenseReqdChoices } from "../common/props";
import { DataSource } from "../data/ds";
import { licenseReqdOptions } from "../ui/DropdownChoices";
import styles from "../Dct.module.scss";
import { ContractForm } from "./ContractForm";
import { formatDate } from "../common/utils";
import { buildOppNetItemUrl, buildPastPerformanceItemUrl, buildProposalItemUrl } from "../common/tagUtils";

export interface ICapFormProps {
    item?: ICapabilityItem;
    context: WebPartContext;
    onSave: (item: ICapFormSaveResult) => void;
    onDelete?: (itemId: number) => void;
    onCancel: () => void;
}

interface ISourceTag<T> extends ITag {
    key: string;
    name: string;
    item: T;
    title: string;
    subtitle?: string;
}

type IOppNetTag = ISourceTag<IOpportunityItem>;
type IPastPerformanceTag = ISourceTag<IPastPerformanceItem>;
type IProposalTag = ISourceTag<IProposalItem>;

export const CapForm: React.FC<ICapFormProps> = ({ item, context, onSave, onDelete, onCancel }) => {
    const [submitted, setSubmitted] = React.useState<boolean>(false);
    const [titleError, setTitleError] = React.useState<string | undefined>(undefined);
    const [showDeleteConfirmation, setShowDeleteConfirmation] = React.useState(false);
    const [opportunities, setOpportunities] = React.useState<IOpportunityItem[]>(DataSource.Opportunities ?? []);
    const [pastPerformance, setPastPerformance] = React.useState<IPastPerformanceItem[]>(DataSource.PastPerformance ?? []);
    const [proposals, setProposals] = React.useState<IProposalItem[]>(DataSource.Proposals ?? []);
    const [oppNetLoading, setOppNetLoading] = React.useState<boolean>(false);
    const [pastPerformanceLoading, setPastPerformanceLoading] = React.useState<boolean>(false);
    const [proposalsLoading, setProposalsLoading] = React.useState<boolean>(false);
    const opportunitiesPromiseRef = React.useRef<Promise<IOpportunityItem[]> | undefined>(undefined);
    const pastPerformancePromiseRef = React.useRef<Promise<IPastPerformanceItem[]> | undefined>(undefined);
    const proposalsPromiseRef = React.useRef<Promise<IProposalItem[]> | undefined>(undefined);
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
        oppNetTags: item?.oppNetTags ?? [],
        pastPerformanceTagsJson: item?.pastPerformanceTagsJson || "",
        pastPerformanceTags: item?.pastPerformanceTags ?? [],
        proposalTagsJson: item?.proposalTagsJson || "",
        proposalTags: item?.proposalTags ?? []
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

    const loadOpportunities = React.useCallback((): Promise<IOpportunityItem[]> => {
        if (DataSource.Opportunities.length) {
            setOpportunities(DataSource.Opportunities);
            return Promise.resolve(DataSource.Opportunities);
        }

        if (!opportunitiesPromiseRef.current) {
            setOppNetLoading(true);
            opportunitiesPromiseRef.current = DataSource.getOpportunities()
                .then((items) => {
                    setOpportunities(items);
                    return items;
                })
                .catch((error) => {
                    opportunitiesPromiseRef.current = undefined;
                    console.error("Error loading OppNet opportunities:", error);
                    return [];
                })
                .finally(() => {
                    setOppNetLoading(false);
                });
        }

        return opportunitiesPromiseRef.current;
    }, []);

    React.useEffect(() => {
        loadOpportunities().catch((error) => console.error("Error loading OppNet opportunities:", error));
    }, [loadOpportunities]);

    const loadPastPerformance = React.useCallback((): Promise<IPastPerformanceItem[]> => {
        if (DataSource.PastPerformance.length) {
            setPastPerformance(DataSource.PastPerformance);
            return Promise.resolve(DataSource.PastPerformance);
        }

        if (!pastPerformancePromiseRef.current) {
            setPastPerformanceLoading(true);
            pastPerformancePromiseRef.current = DataSource.getPastPerformance()
                .then((items) => {
                    setPastPerformance(items);
                    return items;
                })
                .catch((error) => {
                    pastPerformancePromiseRef.current = undefined;
                    console.error("Error loading Past Performance:", error);
                    return [];
                })
                .finally(() => {
                    setPastPerformanceLoading(false);
                });
        }

        return pastPerformancePromiseRef.current;
    }, []);

    React.useEffect(() => {
        loadPastPerformance().catch((error) => console.error("Error loading Past Performance:", error));
    }, [loadPastPerformance]);

    const loadProposals = React.useCallback((): Promise<IProposalItem[]> => {
        if (DataSource.Proposals.length) {
            setProposals(DataSource.Proposals);
            return Promise.resolve(DataSource.Proposals);
        }

        if (!proposalsPromiseRef.current) {
            setProposalsLoading(true);
            proposalsPromiseRef.current = DataSource.getProposals()
                .then((items) => {
                    setProposals(items);
                    return items;
                })
                .catch((error) => {
                    proposalsPromiseRef.current = undefined;
                    console.error("Error loading Proposals:", error);
                    return [];
                })
                .finally(() => {
                    setProposalsLoading(false);
                });
        }

        return proposalsPromiseRef.current;
    }, []);

    React.useEffect(() => {
        loadProposals().catch((error) => console.error("Error loading Proposals:", error));
    }, [loadProposals]);

    const handleChange = <K extends keyof ICapabilityItem>(field: K, value: ICapabilityItem[K]): ICapabilityItem[K] => {
        setFormData((prev) => ({ ...prev, [field]: value }));
        return value;
    };

    const validateTitle = (value?: string): string | undefined =>
        !value?.trim() ? "Capability name is required" : undefined;

    const cleanText = (value?: unknown): string => {
        if (value === null || value === undefined) return "";
        if (typeof value === "string") return value.trim();
        if (Array.isArray(value)) return value.map(cleanText).filter(Boolean).join(", ");
        if (typeof value === "object" && "results" in value) {
            return cleanText((value as { results?: unknown }).results);
        }

        return String(value).trim();
    };

    const getTagItem = <T,>(tag: ISourceTag<T> | T): T =>
        "item" in (tag as ISourceTag<T>) ? (tag as ISourceTag<T>).item : tag as T;

    const getSelectedIds = <T extends { Id: number }>(selectedItems?: Array<ISourceTag<T> | T>): Set<number> =>
        new Set((selectedItems ?? []).map(tag => getTagItem(tag)?.Id).filter((id): id is number => typeof id === "number"));

    const toOppNetTag = (item: IOpportunityItem): IOppNetTag => {
        const tagItem: IOpportunityItem = {
            Id: item.Id,
            Title: cleanText(item.Title),
            Customer: cleanText(item.Customer),
            Status: cleanText(item.Status),
            url: item.url ?? buildOppNetItemUrl(item.Id)
        };
        const title = cleanText(tagItem.Title);
        const subtitle = [title, tagItem.Customer].map(cleanText).filter(Boolean).join(" | ");
        const displayName = tagItem.Id.toString();

        return {
            key: `opp-${tagItem.Id}`,
            name: displayName,
            title: `ID: ${displayName}`,
            subtitle,
            item: tagItem
        };
    };

    const toPastPerformanceTag = (item: IPastPerformanceItem): IPastPerformanceTag => {
        const tagItem: IPastPerformanceItem = {
            ...item,
            Id: item.Id,
            Contract_x0023_: cleanText(item.Contract_x0023_),
            Customer_x0020_Agency: cleanText(item.Customer_x0020_Agency),
            url: item.url ?? buildPastPerformanceItemUrl(item.Id)
        };
        const title = cleanText(tagItem.Contract_x0023_);
        const subtitle = [tagItem.Customer_x0020_Agency, tagItem.Capability_x0020_Area].map(cleanText).filter(Boolean).join(" | ");
        const displayName = title || tagItem.Id.toString();

        return {
            key: `past-performance-${tagItem.Id}`,
            name: displayName,
            title: displayName,
            subtitle,
            item: tagItem
        };
    };

    const toProposalTag = (item: IProposalItem): IProposalTag => {
        const tagItem: IProposalItem = {
            ...item,
            Id: item.Id,
            Title: cleanText(item.Title),
            OpportunityStage: cleanText(item.OpportunityStage),
            TypeOfOpportunity: cleanText(item.TypeOfOpportunity),
            Entity: cleanText(item.Entity),
            url: item.url ?? buildProposalItemUrl(item.Id)
        };
        const title = cleanText(tagItem.Title);
        const displayName = title || tagItem.Id.toString();
        const subtitle = [tagItem.TypeOfOpportunity, tagItem.Entity].map(cleanText).filter(Boolean).join(" | ");

        return {
            key: `proposal-${tagItem.Id}`,
            name: displayName,
            title: displayName,
            subtitle,
            item: tagItem
        };
    };

    const getSelectedOppNetTags = (): IOppNetTag[] =>
        (formData.oppNetTags ?? []).filter((tag): tag is IOpportunityItem => !!tag).map(toOppNetTag);

    const getSelectedPastPerformanceTags = (): IPastPerformanceTag[] =>
        (formData.pastPerformanceTags ?? []).filter((tag): tag is IPastPerformanceItem => !!tag).map(toPastPerformanceTag);

    const getSelectedProposalTags = (): IProposalTag[] =>
        (formData.proposalTags ?? []).filter((tag): tag is IProposalItem => !!tag).map(toProposalTag);

    const filterOppNetSuggestions = (items: IOpportunityItem[], filterText: string, selectedItems?: Array<IOppNetTag | IOpportunityItem>): IOppNetTag[] => {
        const selectedIds = getSelectedIds(selectedItems);
        const search = filterText.trim().toLowerCase();

        return items
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

    const resolveOppNetSuggestions = async (filterText: string, selectedItems?: IOppNetTag[]): Promise<IOppNetTag[]> => {
        return filterOppNetSuggestions(opportunities.length ? opportunities : await loadOpportunities(), filterText, selectedItems);
    }

    const filterPastPerformanceSuggestions = (items: IPastPerformanceItem[], filterText: string, selectedItems?: Array<IPastPerformanceTag | IPastPerformanceItem>): IPastPerformanceTag[] => {
        const selectedIds = getSelectedIds(selectedItems);
        const search = filterText.trim().toLowerCase();

        return items
            .filter(item => !selectedIds.has(item.Id))
            .filter(item => !!cleanText(item.Contract_x0023_))
            .filter(item => {
                if (!search) return true;
                return [item.Contract_x0023_, item.Customer_x0020_Agency, item.Doc_x0020_Type, item.Capability_x0020_Area, item.Id.toString()]
                    .some(value => cleanText(value).toLowerCase().includes(search));
            })
            .sort((a, b) => cleanText(a.Contract_x0023_).localeCompare(cleanText(b.Contract_x0023_)))
            .slice(0, 20)
            .map(toPastPerformanceTag);
    };

    const resolvePastPerformanceSuggestions = async (filterText: string, selectedItems?: IPastPerformanceTag[]): Promise<IPastPerformanceTag[]> =>
        filterPastPerformanceSuggestions(pastPerformance.length ? pastPerformance : await loadPastPerformance(), filterText, selectedItems);

    const filterProposalSuggestions = (items: IProposalItem[], filterText: string, selectedItems?: Array<IProposalTag | IProposalItem>): IProposalTag[] => {
        const selectedIds = getSelectedIds(selectedItems);
        const search = filterText.trim().toLowerCase();

        return items
            .filter(proposal => !selectedIds.has(proposal.Id))
            .filter(proposal => !!cleanText(proposal.Title))
            .filter(proposal => {
                if (!search) return true;
                return [proposal.Title, proposal.OpportunityStage, proposal.TypeOfOpportunity, proposal.Entity, proposal.Id.toString()]
                    .some(value => cleanText(value).toLowerCase().includes(search));
            })
            .sort((a, b) => cleanText(a.Title).localeCompare(cleanText(b.Title)))
            .slice(0, 20)
            .map(toProposalTag);
    };

    const resolveProposalSuggestions = async (filterText: string, selectedItems?: IProposalTag[]): Promise<IProposalTag[]> =>
        filterProposalSuggestions(proposals.length ? proposals : await loadProposals(), filterText, selectedItems);

    const handleOppNetTagsChange = (items?: Array<IOppNetTag | IOpportunityItem>): void => {
        handleChange("oppNetTags", (items ?? []).map(getTagItem));
    };

    const handlePastPerformanceTagsChange = (items?: Array<IPastPerformanceTag | IPastPerformanceItem>): void => {
        handleChange("pastPerformanceTags", (items ?? []).map(getTagItem));
    };

    const handleProposalTagsChange = (items?: Array<IProposalTag | IProposalItem>): void => {
        handleChange("proposalTags", (items ?? []).map(getTagItem));
    };

    const renderOppNetSuggestion = (tag: IOppNetTag): JSX.Element => {
        const item = getTagItem(tag);

        return (
            <div
                style={{
                    boxSizing: "border-box",
                    color: "#323130",
                    display: "block",
                    lineHeight: 1.35,
                    minHeight: 48,
                    padding: "8px 10px",
                    textAlign: "left",
                    whiteSpace: "normal",
                    width: "100%"
                }}
            >
                <div style={{ color: "#323130", display: "block", fontSize: 14, fontWeight: 600, whiteSpace: "normal" }}>
                    {item.Id}
                </div>
                <div style={{ color: "#605e5c", display: "block", fontSize: 12, marginTop: 2, whiteSpace: "normal" }}>
                    {[item.Title, item.Customer].map(cleanText).filter(Boolean).join(" | ")}
                </div>
            </div>
        );
    };

    const renderPastPerformanceSuggestion = (tag: IPastPerformanceTag): JSX.Element => {
        const item = getTagItem(tag);

        return (
            <div
                style={{
                    boxSizing: "border-box",
                    color: "#323130",
                    display: "block",
                    lineHeight: 1.35,
                    minHeight: 48,
                    padding: "8px 10px",
                    textAlign: "left",
                    whiteSpace: "normal",
                    width: "100%"
                }}
            >
                <div style={{ color: "#323130", display: "block", fontSize: 14, fontWeight: 600, whiteSpace: "normal" }}>
                    {[item.Contract_x0023_, item.Doc_x0020_Type ? `(${item.Doc_x0020_Type})` : ""].map(cleanText).filter(Boolean).join(" ")}
                </div>
                <div style={{ color: "#605e5c", display: "block", fontSize: 12, marginTop: 2, whiteSpace: "normal" }}>
                    {[item.Customer_x0020_Agency, item.Capability_x0020_Area].map(cleanText).filter(Boolean).join(" | ")}
                </div>
            </div>
        );
    };

    const renderProposalSuggestion = (tag: IProposalTag): JSX.Element => {
        const item = getTagItem(tag);

        return (
            <div
                style={{
                    boxSizing: "border-box",
                    color: "#323130",
                    display: "block",
                    lineHeight: 1.35,
                    minHeight: 40,
                    padding: "8px 10px",
                    textAlign: "left",
                    whiteSpace: "normal",
                    width: "100%"
                }}
            >
                <div style={{ color: "#323130", display: "block", fontSize: 14, fontWeight: 600, whiteSpace: "normal" }}>
                    {cleanText(item.Title)}
                </div>
                <div style={{ color: "#605e5c", display: "block", fontSize: 12, marginTop: 2, whiteSpace: "normal" }}>
                    {[item.TypeOfOpportunity, item.Entity].map(cleanText).filter(Boolean).join(" | ")}
                </div>
            </div>
        );
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
                <Stack horizontal tokens={{ childrenGap: 4 }} verticalAlign="center">
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
                <div className={styles.capForm}>
                    <section className={styles.formSection}>
                        <div className={styles.formSectionHeader}>
                            <div>
                                <h3>General Info</h3>
                                <p>Core details about the capability</p>
                            </div>
                        </div>

                        <div className={styles.formGridTwo}>
                            <div className={styles.formFieldFull}>
                                <TextField
                                    label="Capability Name"
                                    className={styles.formControl}
                                    value={formData.Title}
                                    onChange={(_, val) => handleChange("Title", val ?? "")}
                                    required
                                    maxLength={255}
                                    errorMessage={submitted ? titleError : undefined}
                                />
                            </div>

                            <div className={styles.formFieldFull}>
                                <TextField
                                    label="Capability Description"
                                    className={styles.formControl}
                                    multiline
                                    autoAdjustHeight
                                    value={formData.description ?? ""}
                                    onChange={(_, val) => handleChange("description", val ?? "")}
                                />
                            </div>

                            <div className={styles.formFieldFull}>
                                <TextField
                                    label="Technical Capabilities"
                                    className={styles.formControl}
                                    multiline
                                    autoAdjustHeight
                                    value={formData.capabilities ?? ""}
                                    onChange={(_, val) => handleChange("capabilities", val ?? "")}
                                />
                            </div>

                            <Dropdown
                                label="Capability Status"
                                className={styles.formControl}
                                selectedKey={formData.capStatus || undefined}
                                options={capStatusOptions}
                                onChange={(_, option) => {
                                    if (option) handleChange("capStatus", option.key as CapStatusType);
                                }}
                            />

                            <TextField
                                label="Link/URL"
                                className={styles.formControl}
                                value={formData.link ?? ""}
                                maxLength={255}
                                onChange={(_, val) => handleChange("link", val ?? "")}
                            />

                            <div className={styles.formFieldFull}>
                                <TextField
                                    label="Additional Notes"
                                    className={styles.formControl}
                                    multiline
                                    autoAdjustHeight
                                    value={formData.notes ?? ""}
                                    onChange={(_, val) => handleChange("notes", val ?? "")}
                                />
                            </div>

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
                            />

                            <div className={styles.formFieldFull}>
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
                                />
                            </div>
                        </div>
                    </section>

                    <section className={styles.formSection}>
                        <div className={styles.formSectionHeader}>
                            <div>
                                <h3>Technical Info</h3>
                                <p>Platform, compliance, licensing, and implementation details</p>
                            </div>
                        </div>

                        <div className={styles.formGridTwo}>
                            <Dropdown
                                label="Platform"
                                className={styles.formControl}
                                selectedKey={formData.platform || undefined}
                                options={platformOptions}
                                onChange={(_, option) => {
                                    if (option) handleChange("platform", option.key as PlatformType);
                                }}
                            />

                            <Dropdown
                                label="Hosting Environment"
                                className={styles.formControl}
                                selectedKey={formData.hostingEnv || undefined}
                                options={hostingEnvOptions}
                                onChange={(_, option) => {
                                    if (option) handleChange("hostingEnv", option.key as HostingEnvType);
                                }}
                            />

                            <Dropdown
                                label="Connectivity"
                                className={styles.formControl}
                                selectedKey={formData.connectivity || undefined}
                                options={connectivityOptions}
                                onChange={(_, option) => {
                                    if (option) handleChange("connectivity", option.key as ConnectivityType);
                                }}
                            />

                            <Dropdown
                                label="Compliance"
                                className={styles.formControl}
                                selectedKey={formData.compliance || undefined}
                                options={complianceOptions}
                                onChange={(_, option) => {
                                    if (option) handleChange("compliance", option.key as ComplianceType);
                                }}
                            />
                        </div>

                        <div className={styles.formGridThree}>
                            <Dropdown
                                label="License Required?"
                                required
                                className={styles.formControl}
                                selectedKey={formData.licenseReqd || undefined}
                                options={licenseReqdOptions(false)}
                                onChange={(_, option) => {
                                    if (option) handleChange("licenseReqd", option.key as licenseReqdChoices);
                                }}
                            />

                            <Dropdown
                                label="Coding Language"
                                className={styles.formControl}
                                selectedKey={formData.codeLanguage || undefined}
                                options={codeLanguageOptions}
                                onChange={(_, option) => {
                                    if (option) handleChange("codeLanguage", option.key as CodeLanguageType);
                                }}
                            />

                            <Dropdown
                                label="Backend"
                                className={styles.formControl}
                                selectedKey={formData.backend || undefined}
                                options={backendOptions}
                                onChange={(_, option) => {
                                    if (option) handleChange("backend", option.key as BackendType);
                                }}
                            />
                        </div>

                        <div className={styles.formGrid}>
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
                        </div>
                    </section>

                    <section className={styles.formSection}>
                        <div className={styles.formSectionHeader}>
                            <div>
                                <h3>Contracts</h3>
                                <p>Add one or more contract relationships for this capability</p>
                            </div>
                            <PrimaryButton
                                text="Add Contract"
                                iconProps={{ iconName: "Add" }}
                                onClick={handleAddContract}
                            />
                        </div>

                        {capContracts.length ? (
                            <div className={styles.contractListFrame}>
                                <DetailsList
                                    items={capContracts}
                                    columns={contractColumns}
                                    selectionMode={SelectionMode.none}
                                    getKey={(contract) => getContractRowKey(contract as ICapabilityContractDraft)}
                                    onRenderRow={(props) => props ? (
                                        <DetailsRow
                                            {...props}
                                            styles={{
                                                root: { minHeight: 40 },
                                                cell: { alignItems: "center", display: "flex", minHeight: 40, paddingBottom: 4, paddingTop: 4 }
                                            }}
                                        />
                                    ) : null}
                                />
                            </div>
                        ) : (
                            <Text styles={{ root: { color: "gray", fontStyle: "italic" } }}>
                                No contracts have been added yet.
                            </Text>
                        )}
                    </section>

                    <section className={styles.formSection}>
                        <div className={styles.formSectionHeader}>
                            <div>
                                <h3>Tagging</h3>
                                <p>Apply relevant tags to help with search and data mining</p>
                            </div>
                        </div>

                        <div className={styles.formGrid}>
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
                                onRenderSuggestionsItem={(tag) => renderOppNetSuggestion(tag as IOppNetTag)}
                            />

                            <TagPicker
                                label="Past Performance Tags"
                                className={styles.formControl}
                                selectedItems={getSelectedPastPerformanceTags()}
                                onResolveSuggestions={(filter, selectedItems) =>
                                    resolvePastPerformanceSuggestions(filter, selectedItems as IPastPerformanceTag[])
                                }
                                onEmptyResolveSuggestions={(selectedItems) =>
                                    resolvePastPerformanceSuggestions("", selectedItems as IPastPerformanceTag[])
                                }
                                onChange={(items?: ITag[]) => handlePastPerformanceTagsChange(items as IPastPerformanceTag[])}
                                getTextFromItem={(tag) => tag.name}
                                pickerSuggestionsProps={{
                                    suggestionsHeaderText: "Matching Past Performance",
                                    noResultsFoundText: pastPerformanceLoading ? "Loading Past Performance..." : "No Past Performance found",
                                    resultsMaximumNumber: 20
                                }}
                                inputProps={{
                                    placeholder: pastPerformanceLoading ? "Loading Past Performance..." : "Search Past Performance"
                                }}
                                onRenderSuggestionsItem={(tag) => renderPastPerformanceSuggestion(tag as IPastPerformanceTag)}
                            />

                            <TagPicker
                                label="Proposal Tags"
                                className={styles.formControl}
                                selectedItems={getSelectedProposalTags()}
                                onResolveSuggestions={(filter, selectedItems) =>
                                    resolveProposalSuggestions(filter, selectedItems as IProposalTag[])
                                }
                                onEmptyResolveSuggestions={(selectedItems) =>
                                    resolveProposalSuggestions("", selectedItems as IProposalTag[])
                                }
                                onChange={(items?: ITag[]) => handleProposalTagsChange(items as IProposalTag[])}
                                getTextFromItem={(tag) => tag.name}
                                pickerSuggestionsProps={{
                                    suggestionsHeaderText: "Matching Proposals",
                                    noResultsFoundText: proposalsLoading ? "Loading Proposals..." : "No Proposals found",
                                    resultsMaximumNumber: 20
                                }}
                                inputProps={{
                                    placeholder: proposalsLoading ? "Loading Proposals..." : "Search Proposals"
                                }}
                                onRenderSuggestionsItem={(tag) => renderProposalSuggestion(tag as IProposalTag)}
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
                </div>
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
