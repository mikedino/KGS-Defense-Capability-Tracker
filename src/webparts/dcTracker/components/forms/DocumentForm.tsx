import * as React from "react";
import { TextField, PrimaryButton, DefaultButton, Stack } from "@fluentui/react";
import { IDocumentItem } from "../common/props";
import styles from "../Dct.module.scss";

interface IDocumentFormProps {
  item: IDocumentItem;
  onSave: (updatedItem: IDocumentItem) => void;
  onCancel: () => void;
  onDelete: (item: IDocumentItem) => void;
}

export const DocumentForm: React.FC<IDocumentFormProps> = ({ item, onSave, onCancel, onDelete }) => {
  const [formData, setFormData] = React.useState<IDocumentItem>({ ...item })
  //const isSaveDisabled = !formData.docType;

  return (
    <Stack tokens={{ childrenGap: 12 }}>
      <TextField label="Filename" value={formData.FileLeafRef} disabled />
      <TextField
        label="Title"
        value={formData.Title}
        onChange={(_, newValue) => setFormData({ ...formData, Title: newValue || "" })}
        required
        maxLength={255}
      />
    <Stack horizontal horizontalAlign="space-between" tokens={{ childrenGap: 10 }} style={{ marginTop: 40 }}> 
        <PrimaryButton text="Save" onClick={() => onSave(formData)} title="Save Record"/>
        <DefaultButton text="Cancel" onClick={onCancel} title="Close Dialog Box"/>
        <DefaultButton text="Delete" className={styles.deleteButton} onClick={() => onDelete(item)} styles={{ root: { marginLeft: "auto" } }} title="Delete Record"/>
      </Stack>
    
    </Stack>
  );
};
