import * as React from "react";
import { DefaultButton, Dropdown, PrimaryButton, Stack, TextField } from "@fluentui/react";
import { IContractDocumentItem } from "../common/props";
import { DataSource } from "../data/ds";
import styles from "../Dct.module.scss";

interface IContractDocumentFormProps {
    item: IContractDocumentItem;
    onSave: (updatedItem: IContractDocumentItem) => void;
    onCancel: () => void;
    onDelete: (item: IContractDocumentItem) => void;
}

export const ContractDocumentForm: React.FC<IContractDocumentFormProps> = ({ item, onSave, onCancel, onDelete }) => {
    const [formData, setFormData] = React.useState<IContractDocumentItem>({ ...item });
    const cdocTypeOptions = DataSource.getConfigOptions("cdocType");

    return (
        <Stack tokens={{ childrenGap: 12 }}>
            <TextField label="Filename" value={formData.FileLeafRef} disabled />
            <TextField
                label="Title"
                value={formData.Title ?? ""}
                onChange={(_, newValue) => setFormData({ ...formData, Title: newValue || "" })}
                maxLength={255}
            />
            <Dropdown
                label="Contract Document Type"
                selectedKey={formData.cdocType || undefined}
                options={cdocTypeOptions}
                onChange={(_, option) => setFormData({ ...formData, cdocType: (option?.key as string) ?? "" })}
            />
            <Stack horizontal horizontalAlign="space-between" tokens={{ childrenGap: 10 }} style={{ marginTop: 24 }}>
                <PrimaryButton text="Save" onClick={() => onSave(formData)} title="Save Record" />
                <DefaultButton text="Cancel" onClick={onCancel} title="Close Dialog Box" />
                <DefaultButton text="Delete" className={styles.deleteButton} onClick={() => onDelete(item)} styles={{ root: { marginLeft: "auto" } }} title="Delete Record" />
            </Stack>
        </Stack>
    );
};
