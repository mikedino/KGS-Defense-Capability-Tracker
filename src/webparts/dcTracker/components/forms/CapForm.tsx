import * as React from "react";
import {
    DefaultButton,
    Dialog,
    DialogFooter,
    DialogType,
    Dropdown,
    IDropdownOption,
    PrimaryButton,
    Stack,
    Text,
    TextField
} from "@fluentui/react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ICapabilityItem, licenseReqdChoices } from "../common/props";
import { DataSource } from "../data/ds";
import { licenseReqdOptions } from "../ui/DropdownChoices";
import { cardStackStyles } from "../ui/ComponentStyles";
import styles from "../Dct.module.scss";

const buildOptions = (values: string[]): IDropdownOption[] =>
    (values ?? []).map((v) => ({ key: v, text: v }));

export interface ICapFormProps {
    item?: ICapabilityItem;
    context: WebPartContext;
    onSave: (item: ICapabilityItem) => void;
    onDelete?: (itemId: number) => void;
    onCancel: () => void;
}

export const CapForm: React.FC<ICapFormProps> = ({ item, onSave, onDelete, onCancel }) => {
    const [submitted, setSubmitted] = React.useState<boolean>(false);
    const [titleError, setTitleError] = React.useState<string | undefined>(undefined);
    const [showDeleteConfirmation, setShowDeleteConfirmation] = React.useState(false);

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
        contract: item?.contract?.Id ? item.contract : undefined
    });

    const capStatusOptions = React.useMemo(() => buildOptions(DataSource.getConfigValues("capStatus")), []);
    const platformOptions = React.useMemo(() => buildOptions(DataSource.getConfigValues("platform")), []);
    const hostingEnvOptions = React.useMemo(() => buildOptions(DataSource.getConfigValues("hostingEnv")), []);
    const connectivityOptions = React.useMemo(() => buildOptions(DataSource.getConfigValues("connectivity")), []);
    const complianceOptions = React.useMemo(() => buildOptions(DataSource.getConfigValues("compliance")), []);
    const codeLanguageOptions = React.useMemo(() => buildOptions(DataSource.getConfigValues("codeLanguage")), []);
    const backendOptions = React.useMemo(() => buildOptions(DataSource.getConfigValues("backend")), []);
    const contractOptions = React.useMemo(
        () => DataSource.Contracts.map(c => ({ key: c.Id.toString(), text: c.Title })),
        []
    );

    const handleChange = <K extends keyof ICapabilityItem>(field: K, value: ICapabilityItem[K]): ICapabilityItem[K] => {
        setFormData((prev) => ({ ...prev, [field]: value }));
        return value;
    };

    const validateTitle = (value?: string): string | undefined =>
        !value?.trim() ? "Capability name is required" : undefined;

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

        onSave(formData);
        return true;
    };

    const handleDelete = (): void => onDelete?.(formData.Id);

    return (
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
                        <Text variant="large">Contract Info</Text>
                        <Text variant="small">Contract association placeholder</Text>
                    </Stack>

                    <Dropdown
                        label="Contract"
                        className={styles.formControl}
                        selectedKey={formData.contract?.Id?.toString()}
                        options={contractOptions}
                        onChange={(_, option) => {
                            if (option) {
                                handleChange("contract", { Id: parseInt(option.key as string, 10), Title: option.text });
                            }
                        }}
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
    );
};
