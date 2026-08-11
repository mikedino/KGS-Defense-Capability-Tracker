import * as React from "react";
import { Icon, IconButton, Link, PrimaryButton, Text } from "@fluentui/react";
import { IContractDocumentItem } from "../../common/props";
import styles from "../../Dct.module.scss";

export interface IContractDocumentsPanelProps {
    documents: IContractDocumentItem[];
    canEdit: boolean;
    onAdd?: () => void;
    onEdit?: (document: IContractDocumentItem) => void;
    onDelete?: (document: IContractDocumentItem) => void;
    showAddButton?: boolean;
}

const getFileIcon = (fileType?: string): string => {
    switch ((fileType ?? "").toLowerCase()) {
        case "doc":
        case "docx":
            return "WordDocument";
        case "xls":
        case "xlsx":
            return "ExcelDocument";
        case "ppt":
        case "pptx":
            return "PowerPointDocument";
        case "pdf":
            return "PDF";
        case "jpg":
        case "jpeg":
        case "png":
            return "FileImage";
        default:
            return "Page";
    }
};

export const ContractDocumentsPanel: React.FC<IContractDocumentsPanelProps> = ({ documents, canEdit, onAdd, onEdit, onDelete, showAddButton = true }) => {
    return (
        <div className={styles.contractDocumentsPanel}>
            <div className={styles.contractDocumentsHeader}>
                <span className={styles.contractDocumentsTitle}>Contract Documents ({documents.length})</span>
                {canEdit && showAddButton && onAdd && <PrimaryButton text="Add Document" iconProps={{ iconName: "Add" }} onClick={onAdd} />}
            </div>

            {documents.length ? (
                <div className={styles.contractDocumentLinks}>
                    {documents.map((documentItem) => (
                        <div key={documentItem.Id} className={styles.contractDocumentRow}>
                            {canEdit && (
                                <div className={styles.contractDocumentActions}>
                                    <IconButton iconProps={{ iconName: "Edit" }} title="Edit document metadata" ariaLabel="Edit document metadata" onClick={() => onEdit?.(documentItem)} />
                                    <IconButton iconProps={{ iconName: "Delete" }} title="Delete document" ariaLabel="Delete document" onClick={() => onDelete?.(documentItem)} />
                                </div>
                            )}
                            <Icon className={styles.contractDocumentIcon} iconName={getFileIcon(documentItem.File_x0020_Type)} />
                            <Link className={styles.contractDocumentLink} href={documentItem.ServerRedirectedEmbedUrl || documentItem.EncodedAbsUrl} target="_blank" rel="noopener noreferrer">
                                {documentItem.FileLeafRef}
                            </Link>
                            {documentItem.cdocType && <span className={styles.contractDocumentBadge}>{documentItem.cdocType}</span>}
                        </div>
                    ))}
                </div>
            ) : (
                <Text styles={{ root: { color: "gray", fontStyle: "italic" } }}>
                    No documents have been added for this contract.
                </Text>
            )}
        </div>
    );
};
