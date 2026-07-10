import * as React from "react";
import {
    Stack, Text, TextField, Dropdown, IDropdownOption, PrimaryButton, DefaultButton, DatePicker,
    DayOfWeek, Dialog, DialogFooter, DialogType,
    IPersonaProps
} from "@fluentui/react";
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { IPeoplePickerContext } from "@pnp/spfx-controls-react/lib/controls/peoplepicker/IPeoplePickerContext";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { licenseReqdOptions } from "../ui/DropdownChoices";
import { getPlainTextLength, RichTextField } from "../ui/RichTextField";
import { IApplicationItem, licenseReqdChoices } from "../data/props";
import { DataSource } from "../data/ds";
import styles from "../AppTracker.module.scss";
import { cardStackStyles } from "../ui/ComponentStyles";

// If you have your own helper for "All", keep it; for form dropdowns we generally don't include All.
const buildOptions = (values: string[]): IDropdownOption[] =>
    (values ?? []).map((v) => ({ key: v, text: v }));

const onFormatDate = (date?: Date): string => {
    if (!date) return "";
    return `${(date.getMonth() + 1).toString().padStart(2, "0")}/` +
        `${date.getDate().toString().padStart(2, "0")}/` +
        `${date.getFullYear().toString().slice(-2)}`;
};

export interface IAppFormProps {
    item?: IApplicationItem; // edit
    context: WebPartContext;
    onSave: (item: IApplicationItem) => void;
    onDelete?: (itemId: number) => void;
    onCancel: () => void;
}

export const AppForm: React.FC<IAppFormProps> = ({ item, context, onSave, onDelete, onCancel }) => {

    const [submitted, setSubmitted] = React.useState<boolean>(false);
    const [appLinkError, setAppLinkError] = React.useState<string | undefined>(undefined);

    // ---- validation ----
    const [titleError, setTitleError] = React.useState<string | undefined>(undefined);
    const [relatedInfoError, setRelatedInfoError] = React.useState<string | undefined>(undefined);
    //const [highlightsError, setHighlightsError] = React.useState<string | undefined>(undefined);


    type SupportTeamType = IApplicationItem["supportTeam"];
    type ConnectivityType = IApplicationItem["connectivity"];
    type PlatformType = IApplicationItem["platform"];
    type EnvironmentType = IApplicationItem["environment"];
    type ManagingGroupType = IApplicationItem["managingGroup"];
    type AppStatusType = IApplicationItem["appStatus"];

    const [formData, setFormData] = React.useState<IApplicationItem>({
        Id: item?.Id || 0,
        Title: item?.Title || "",
        description: item?.description || "",
        //highlights: item?.highlights || "",
        appUrl: item?.appUrl || "",
        relatedInfo: item?.relatedInfo || "",
        primaryPoc: item?.primaryPoc || undefined,
        stakeholders: { results: item?.stakeholders?.results ?? [] },
        systemOwner: item?.systemOwner || undefined,
        managingGroup: item?.managingGroup || "",
        supportTeam: item?.supportTeam || "",
        licenseReqd: item?.licenseReqd ?? "No",
        connectivity: item?.connectivity || "",
        integrationsInput: item?.integrationsInput || "",
        integrationsOutput: item?.integrationsOutput || "",
        synonyms: item?.synonyms || "",
        platform: item?.platform || "",
        environment: item?.environment || "",
        contract: item?.contract?.Id ? item.contract : undefined,
        appStatus: item?.appStatus || "",
        appLaunchDate: item?.appLaunchDate || "",
        userCount: item?.userCount || 0
    });

    // ---- dropdown options (config-driven) ----
    const supportTeamOptions = React.useMemo(() => buildOptions(DataSource.getConfigValues("supportTeam")), []);
    const connectivityOptions = React.useMemo(() => buildOptions(DataSource.getConfigValues("connectivity")), []);
    const platformOptions = React.useMemo(() => buildOptions(DataSource.getConfigValues("platform")), []);
    const environmentOptions = React.useMemo(() => buildOptions(DataSource.getConfigValues("environment")), []);
    const managingGroupOptions = React.useMemo(() => buildOptions(DataSource.getConfigValues("managingGroup")), []);
    const appStatusOptions = React.useMemo(() => buildOptions(DataSource.getConfigValues("appStatus")), []);
    const contractOptions = React.useMemo(() => DataSource.Contracts.map(c => ({ key: c.Id.toString(), text: c.Title })), []);

    // ---- people picker context ----
    const peoplePickerContext: IPeoplePickerContext = {
        absoluteUrl: context.pageContext.web.absoluteUrl,
        msGraphClientFactory: context.msGraphClientFactory,
        spHttpClient: context.spHttpClient
    };

    // ---- generic change helper ----
    const handleChange = <K extends keyof IApplicationItem>(field: K, value: IApplicationItem[K]): IApplicationItem[K] => {
        setFormData((prev) => ({ ...prev, [field]: value }));
        return value;
    };

    // ---- people pickers ----
    const handlePerson = (items: IPersonaProps[], field: keyof IApplicationItem): void => {
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

    const handleStakeholders = (items: IPersonaProps[]): void => {
        const owners = (items ?? []).map((p) => ({
            Id: parseInt(p.id!, 10),
            EMail: p.secondaryText!,
            Title: p.text!
        }));

        handleChange("stakeholders", { results: owners });
    };

    const handleContractChange = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
        if (option) {
            handleChange("contract", { Id: parseInt(option.key as string), Title: option.text });
        }
    };

    // ---------- Validation helpers ----------------
    const validateTitle = (value?: string): string | undefined => {
        return !value?.trim() ? "Application name is required" : undefined;
    };

    const validateRelatedInfo = (value: string): string | undefined => {
        const plainTextLength: number = getPlainTextLength(value);

        if (plainTextLength > 500) {
            return "Please keep to 500 characters or fewer.";
        }

        return undefined;
    };

    // const validateHighlights = (value: string): string | undefined => {
    //     const plainTextLength: number = getPlainTextLength(value);

    //     if (plainTextLength > 500) {
    //         return "Please keep to 500 characters or fewer.";
    //     }

    //     return undefined;
    // };

    const validateForm = (): boolean => {
        const nextTitleError: string | undefined = validateTitle(formData.Title);
        const nextRelatedInfoError: string | undefined = validateRelatedInfo(formData.relatedInfo ?? "");
        //const nextHighlightsError: string | undefined = validateHighlights(formData.highlights ?? "");

        setTitleError(nextTitleError);
        setRelatedInfoError(nextRelatedInfoError);
        //setHighlightsError(nextHighlightsError);

        return !nextTitleError
            && !nextRelatedInfoError
            //&& !nextHighlightsError
            && !appLinkError;
    };

    // ---------- Validation effects (real-time) ----------------
    React.useEffect(() => {
        setTitleError(validateTitle(formData.Title));
    }, [formData.Title]);

    React.useEffect(() => {
        setRelatedInfoError(validateRelatedInfo(formData.relatedInfo ?? ""));
    }, [formData.relatedInfo]);

    // React.useEffect(() => {
    //     setHighlightsError(validateHighlights(formData.highlights ?? ""));
    // }, [formData.highlights]);

    React.useEffect(() => {
        if ((formData.appUrl?.length ?? 0) >= 255) {
            setAppLinkError("The App URL is limited to 255 characters. The URL may be cut off and not work correctly so please verify before saving.");
        } else {
            setAppLinkError(undefined);
        }
    }, [formData.appUrl]);


    // ---------- FORM SUBMIT - DOUBLE-CHECK VALIDATION ----------------
    const handleSubmit = (ev: React.FormEvent<HTMLFormElement>): boolean => {
        ev.preventDefault();
        setSubmitted(true);

        const isValid: boolean = validateForm();

        if (!isValid) {
            return false;
        }

        onSave(formData);
        return true;
    };

    // ---- delete ----
    const [showDeleteConfirmation, setShowDeleteConfirmation] = React.useState(false);
    const handleDelete = (): void => onDelete?.(formData.Id);

    return (

        <form onSubmit={handleSubmit}>
            <Stack tokens={{ childrenGap: 20 }}>
                {/* ========================= BASIC INFORMATION  ========================= */}
                <Stack styles={cardStackStyles} tokens={{ childrenGap: 10 }}>
                    <Stack>
                        <Text variant="large">Basic Information</Text>
                        <Text variant="small">Core details about the application</Text>
                    </Stack>

                    {/* Application Name */}
                    <TextField
                        label="Application Name"
                        className={styles.formControl}
                        value={formData.Title}
                        onChange={(_, val) => handleChange("Title", val ?? "")}
                        required
                        maxLength={255}
                        errorMessage={submitted ? titleError : undefined}
                    />

                    {/* Description (500 char) - multiline on its own line */}
                    <TextField
                        label="Description"
                        description="Please keep to 1000 characters."
                        className={styles.formControl}
                        multiline
                        autoAdjustHeight
                        value={formData.description ?? ""}
                        onChange={(_, val) => handleChange("description", val ?? "")}
                        maxLength={1000}
                    />

                    {/* Related Info - multiline on its own line */}
                    {/* renamed 7/6/26 -MDL */}
                    <RichTextField
                        label="Core Capabilities & Relevant Software"
                        className={styles.formControl}
                        description="Please keep to 500 characters & use bullets"
                        value={formData.relatedInfo ?? ""}
                        maxLength={500}
                        errorMessage={submitted ? relatedInfoError : undefined}
                        onChange={(value: string): string => handleChange("relatedInfo", value)}
                    />

                    {/* Hide for now Highlights & Features - multiline on its own line */}
                    {/* <RichTextField
                        label="Highlights & Features"
                        className={styles.formControl}
                        description="Please keep to 500 characters & use bullets"
                        value={formData.highlights ?? ""}
                        maxLength={500}
                        errorMessage={submitted ? highlightsError : undefined}
                        onChange={(value: string): string => handleChange("highlights", value)}
                    /> */}

                    {/* URL */}
                    <TextField
                        label="URL"
                        className={styles.formControl}
                        value={formData.appUrl ?? ""}
                        maxLength={255}
                        onChange={(_, val) => handleChange("appUrl", val ?? "")}
                        errorMessage={appLinkError}
                    />

                    {/* Row: App Status / App Launch Date */}
                    <Stack horizontal wrap tokens={{ childrenGap: 12 }} style={{ margin: 6 }}>
                        <Dropdown
                            label="App Status"
                            className={styles.formControl}
                            selectedKey={formData.appStatus || undefined}
                            options={appStatusOptions}
                            onChange={(_, option) => {
                                if (option) handleChange("appStatus", option.key as AppStatusType);
                            }}
                            styles={{ root: { width: 200, flexGrow: 0, alignSelf: "anchor-center" } }}
                        />

                        <DatePicker
                            label="App Launch Date"
                            className={styles.formControl}
                            firstDayOfWeek={DayOfWeek.Sunday}
                            value={formData.appLaunchDate ? new Date(formData.appLaunchDate) : undefined}
                            onSelectDate={(date) => handleChange("appLaunchDate", date ? date.toISOString() : "")}
                            allowTextInput
                            formatDate={onFormatDate}
                            styles={{ root: { width: 220, flexGrow: 0, alignSelf: "anchor-center" } }}
                        />
                    </Stack>

                    {/* Synonyms keep it full-width line */}
                    <TextField
                        label="Synonyms"
                        className={styles.formControl}
                        value={formData.synonyms ?? ""}
                        onChange={(_, val) => handleChange("synonyms", val ?? "")}
                        styles={{ root: { minWidth: 260, flexGrow: 1 } }}
                    />

                    <Stack horizontal wrap tokens={{ childrenGap: 12 }} style={{ margin: 6 }}>
                        <Dropdown
                            label="License Required?"
                            required
                            className={styles.formControl}
                            selectedKey={formData.licenseReqd || undefined}
                            options={licenseReqdOptions(false)}
                            onChange={(_, option) => {
                                if (option) handleChange("licenseReqd", option.key as licenseReqdChoices);
                            }}
                            styles={{ root: { width: 220, flexGrow: 0, alignSelf: "anchor-center" } }}
                        />

                        <TextField
                            label="User Count"
                            type="number"
                            className={styles.formControl}
                            value={Number.isFinite(formData.userCount) ? formData.userCount.toString() : ""}
                            onChange={(_, val) => handleChange("userCount", Number(val))}
                            styles={{ root: { width: 150, flexGrow: 0, alignSelf: "anchor-center" } }}
                        />
                    </Stack>
                </Stack>

                {/* =========================  OWNERSHIP & SUPPORT  ========================= */}
                <Stack styles={cardStackStyles} tokens={{ childrenGap: 10 }}>
                    <Stack>
                        <Text variant="large">Ownership &amp; Support</Text>
                        <Text variant="small">Points of contact and support structure</Text>
                    </Stack>

                    {/* Row: Primary POC + Stakeholders */}

                    <PeoplePicker
                        context={peoplePickerContext}
                        peoplePickerWPclassName={styles.formControl}
                        defaultSelectedUsers={formData.primaryPoc?.EMail ? [formData.primaryPoc.EMail] : []}
                        titleText="Primary POC"
                        personSelectionLimit={1}
                        ensureUser
                        showtooltip
                        onChange={(items) => handlePerson(items, "primaryPoc")}
                        principalTypes={[PrincipalType.User]}
                        resolveDelay={1000}
                        styles={{ root: { width: 350, flexGrow: 0 } }}
                    />

                    <PeoplePicker
                        context={peoplePickerContext}
                        peoplePickerWPclassName={styles.formControl}
                        defaultSelectedUsers={
                            formData.stakeholders?.results?.length
                                ? formData.stakeholders.results.map((p) => p.EMail)
                                : []
                        }
                        titleText="Stakeholders"
                        personSelectionLimit={10}
                        ensureUser
                        showtooltip
                        onChange={(items) => handleStakeholders(items)}
                        principalTypes={[PrincipalType.User]}
                        resolveDelay={1000}
                        styles={{ root: { minWidth: 420 } }}
                    />

                    {/* Row: System Owner / Managing Group */}
                    <Stack horizontal wrap tokens={{ childrenGap: 12 }}>
                        <PeoplePicker
                            context={peoplePickerContext}
                            peoplePickerWPclassName={styles.formControl}
                            defaultSelectedUsers={formData.systemOwner?.EMail ? [formData.systemOwner.EMail] : []}
                            titleText="System Owner"
                            personSelectionLimit={1}
                            ensureUser
                            showtooltip
                            onChange={(items) => handlePerson(items, "systemOwner")}
                            principalTypes={[PrincipalType.User]}
                            resolveDelay={1000}
                            styles={{ root: { width: 350, flexGrow: 0 } }}
                        />

                        <Dropdown
                            label="Managing Group"
                            required
                            className={styles.formControl}
                            selectedKey={formData.managingGroup || undefined}
                            options={managingGroupOptions}
                            onChange={(_, option) => {
                                if (option) handleChange("managingGroup", option.key as ManagingGroupType);
                            }}
                            styles={{ root: { minWidth: 260, flexGrow: 0, alignSelf: "anchor-center" } }}
                        />
                    </Stack>

                    {/* Row: Support Team / Contract */}
                    <Stack horizontal wrap tokens={{ childrenGap: 12 }}>
                        <Dropdown
                            label="Contractor Name"
                            className={styles.formControl}
                            selectedKey={formData.supportTeam || undefined}
                            options={supportTeamOptions}
                            onChange={(_, option) => {
                                if (option) handleChange("supportTeam", option.key as SupportTeamType);
                            }}
                            styles={{ root: { minWidth: 200, flexGrow: 0, alignSelf: "anchor-center" } }}
                        />

                        <Dropdown
                            label="Contract"
                            className={styles.formControl}
                            selectedKey={formData.contract?.Id?.toString()}
                            options={contractOptions}
                            onChange={handleContractChange}
                            styles={{ root: { minWidth: 250, flexGrow: 1, alignSelf: "anchor-center" } }}
                        />
                    </Stack>
                </Stack>

                {/* ========================= PLATFORM, DATA, & INTEGRATIONS  ========================= */}
                <Stack styles={cardStackStyles} tokens={{ childrenGap: 10 }}>
                    <Stack>
                        <Text variant="large">Platform, Data, &amp; Integrations</Text>
                        <Text variant="small">Where it runs and how it connects</Text>
                    </Stack>

                    {/* Integrations (Input) - multiline on its own line */}
                    <TextField
                        label="Integrations (Input)"
                        className={styles.formControl}
                        multiline
                        autoAdjustHeight
                        description="Please keep to 500 characters."
                        maxLength={500}
                        value={formData.integrationsInput ?? ""}
                        onChange={(_, val) => handleChange("integrationsInput", val ?? "")}
                    />

                    {/* Integrations (Output) - multiline on its own line */}
                    <TextField
                        label="Integrations (Output)"
                        className={styles.formControl}
                        multiline
                        autoAdjustHeight
                        description="Please keep to 500 characters."
                        maxLength={500}
                        value={formData.integrationsOutput ?? ""}
                        onChange={(_, val) => handleChange("integrationsOutput", val ?? "")}
                    />

                    {/* Row: Platform / Hosting Environment / Connectivity */}
                    <Stack horizontal wrap tokens={{ childrenGap: 12 }} style={{ margin: 6 }}>
                        <Dropdown
                            label="Platform"
                            className={styles.formControl}
                            selectedKey={formData.platform || undefined}
                            options={platformOptions}
                            onChange={(_, option) => {
                                if (option) handleChange("platform", option.key as PlatformType);
                            }}
                            styles={{ root: { minWidth: 200, flexGrow: 1, alignSelf: "anchor-center" } }}
                        />

                        <Dropdown
                            label="Hosting Environment"
                            className={styles.formControl}
                            selectedKey={formData.environment || undefined}
                            options={environmentOptions}
                            onChange={(_, option) => {
                                if (option) handleChange("environment", option.key as EnvironmentType);
                            }}
                            styles={{ root: { minWidth: 240, flexGrow: 1, alignSelf: "anchor-center" } }}
                        />

                        <Dropdown
                            label="Connectivity"
                            className={styles.formControl}
                            selectedKey={formData.connectivity || undefined}
                            options={connectivityOptions}
                            onChange={(_, option) => {
                                if (option) handleChange("connectivity", option.key as ConnectivityType);
                            }}
                            styles={{ root: { minWidth: 240, flexGrow: 1, alignSelf: "anchor-center" } }}
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
                            title="Delete Record"
                        />
                    ) : (
                        <span style={{ width: 150 }} />
                    )}

                    <Stack horizontal horizontalAlign="end" tokens={{ childrenGap: 10 }}>
                        <PrimaryButton
                            type="submit"
                            text="Save"
                            style={{ width: 150 }}
                            //onClick={(e) => handleSubmit(e)}
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
                        title: "Delete Application Entry",
                        subText: "Are you sure you want to delete this App Entry?",
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

    );
};
