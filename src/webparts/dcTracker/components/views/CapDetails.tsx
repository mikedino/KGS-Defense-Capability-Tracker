import * as React from 'react';
import { useState, useEffect } from 'react';
import {
    Stack, CommandBarButton, Pivot, PivotItem, Spinner, Dialog, DialogType,
    DialogFooter, DefaultButton, SpinnerSize, PrimaryButton, Text, 
    mergeStyleSets, DetailsList, IDetailsRowProps, DetailsRow,
    SelectionMode, DetailsListLayoutMode, IconButton, IColumn, Icon,
    Dropdown, IDropdownOption
} from "@fluentui/react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { DataSource } from '../data/ds';
import { ContextInfo, Helper } from 'gd-sprest-bs';
import { Web } from 'gd-sprest-bs';
import { ICapabilityItem, ICapFormSaveResult, IContractItem, IDocumentItem } from "../common/props";
import styles from '../Dct.module.scss';
import Strings from '../common/strings';
import { customPivotStyles } from '../ui/ComponentStyles';
import { formatError } from '../common/utils';
import { Pill } from '../ui/Pill';
import { getAtoStatusFill } from '../ui/StatusColors';
import { CapabilityService } from '../services/CapabilityService';
import { DocumentService } from '../services/DocumentService';
import { ContractService } from '../services/ContractService';
import { CapForm } from '../forms/CapForm';
import { DocumentForm } from '../forms/DocumentForm';
import { CapabilityOverview } from './capDetails/CapOverview';
import { SupportingInfo } from './capDetails/CapSupportingInfo';
import { ContractInfo } from './capDetails/CapContract';
import { TaggingInfo } from './capDetails/CapTagging';
import { Security } from '../services/Security';
import { fetchFileDataUrlFromDocumentItem } from '../export/ExportPdfUtils';
import { exportCapabilityPdf } from '../export/ExportPdfWrapper';
import { ScreenshotCarousel } from '../ui/ScreenshotCarousel';
import { AppHeader } from '../ui/AppHeader';
import { CapRouteTab } from '../routing/routes';

interface CapDetailsProps {
    capability: ICapabilityItem;
    context: WebPartContext;
    onBack: () => void;
    activeTab?: CapRouteTab;
    onTabChange?: (tab: CapRouteTab) => void;
    onNewCapability?: () => void;
}

interface CustomFile extends File {
    data: ArrayBuffer;
}

// Capability form style override
const dialogStyles = mergeStyleSets({
    mainOverride: {
        width: '90vw !important',
        maxWidth: '900px !important',
        minWidth: '675px !important',
        height: 'auto'
    },
});

export type DocFolderStatus = "unknown" | "ready" | "missing" | "error";

const isCapabilityTab = (v: unknown): v is CapRouteTab =>
    v === "overview" || v === "supporting" || v === "tagging" || v === "contract" || v === "documentation";

export const CapDetails: React.FC<CapDetailsProps> = ({ capability, context, onBack, activeTab = "overview", onTabChange, onNewCapability }) => {

    const [capState, setCapState] = useState<ICapabilityItem>(capability);
    const [showCapabilityForm, setShowCapabilityForm] = useState<boolean>(false);
    const [documents, setDocuments] = React.useState<IDocumentItem[]>([]);
    const [selectedDocument, setSelectedDocument] = React.useState<IDocumentItem | undefined>(undefined);
    const [documentType, setDocumentType] = useState<string | undefined>(undefined);
    const [docFolderStatus, setDocFolderStatus] = useState<DocFolderStatus>("unknown");

    const [showDialog, setShowDialog] = useState(false);
    const [showSpinner, setShowSpinner] = useState(false);
    const [spinnerMessage, setSpinnerMessage] = useState<string>("Loading...");
    const [showDocUploadDialog, setShowDocUploadDialog] = useState(false);
    const [showDocDeleteDialog, setShowDocDeleteDialog] = useState(false);
    const [showDocEditDialog, setShowDocEditDialog] = useState(false);
    const [dialogTitle, setDialogTitle] = useState<string>("");
    const [dialogMessage, setDialogMessage] = useState<string>("");

    const setDialogProps = (title: string, message: string): void => {
        setShowDialog(true);
        setDialogTitle(title);
        setDialogMessage(message);
    };

    const setSpinnerProps = (message: string): void => {
        setShowSpinner(true);
        setSpinnerMessage(message);
    };

    const hideDialog = (): void => {
        setShowDialog(false);
    };

    const getDocs = async (appId: number): Promise<boolean> => {
        try {
            const docs = await DataSource.getDocumentsByCapability(appId);

            if (docs?.length) {
                setDocFolderStatus("ready");
            }

            setDocuments(
                [...docs].sort(
                    (a, b) => new Date(b.Modified).getTime() - new Date(a.Modified).getTime()
                )
            );

            return true;
        } catch (error) {
            const errorMessage = formatError(error);
            console.error("Error getting Capability Documents:", errorMessage);
            setDialogProps("Error getting Capability Documents:", errorMessage);
            return false;
        }
    };

    useEffect(() => {
        setCapState(capability);
    }, [capability]);

    const contracts: IContractItem[] = DataSource.Contracts
        .filter((contract) => contract.capability?.Id === capState.Id)
        .sort((a, b) => (a.Title ?? "").localeCompare(b.Title ?? ""));

    const primaryContract: IContractItem | undefined = contracts[0];

    useEffect(() => {
        getDocs(capState.Id).catch((error) => {
            console.error("Unhandled promise rejection:", error);
        });
    }, [capState.Id]);

    const statusProps = getAtoStatusFill(capState.capStatus);
    // const getUserPhotoUrl = (email: string): string => `/_layouts/15/userphoto.aspx?size=M&accountname=${encodeURIComponent(email)}`;
    const screenshots = documents.filter(doc => doc.docType === "Screenshot");

    /*************************************************************************
    **** adding new documents all done here in this component due to the amount of information needed before/during/after upload
    ****************************************************************************/

    // reset document type on dialog open
    useEffect(() => {
        if (showDocUploadDialog) {
            setDocumentType(undefined);
        }
    }, [showDocUploadDialog]);

    const docTypeOptions = React.useMemo<IDropdownOption[]>(() => {
        return DataSource.getConfigOptions("documentType");
    }, [showDocUploadDialog]);

    const uploadFile = (fileName: string, file: CustomFile): Promise<IDocumentItem> => {
        return new Promise((resolve, reject) => {
            setSpinnerProps("Uploading file...");

            const targetFolderUrl = `${ContextInfo.webServerRelativeUrl}/${Strings.Sites.main.lists.Documents}/${capability.Id}`;

            const list = Web().Lists(Strings.Sites.main.lists.Documents);

            // Upload the file
            // list.RootFolder().Files().add(fileName, true, file.data).execute(
            Web().getFolderByServerRelativeUrl(targetFolderUrl).Files().add(fileName, true, file.data).execute(
                (uploadedFile) => {
                    if (uploadedFile) {
                        // Retrieve the ListItem ID
                        uploadedFile.ListItemAllFields().execute((listItem) => {
                            const newItemId = listItem?.Id;
                            setSpinnerMessage("Updating File Properties");
                            // Update item properties
                            list.Items().getById(newItemId).update({
                                capabilityId: capability.Id,
                                docType: documentType
                            }).execute(
                                () => {
                                    //on successful update, get the updated item again from the list
                                    setSpinnerMessage("Updating File Properties");
                                    list.Items().query({
                                        Select: ["File_x0020_Type", "UniqueId", "Id", "FileLeafRef", "capability/Id", "ServerRedirectedEmbedUrl"],
                                        Expand: ["capability"],
                                        Filter: `Id eq ${newItemId}`
                                    }).execute(
                                        items => {
                                            if (items && items.results && Array.isArray(items.results)) {
                                                const docItem = items.results[0] as unknown as IDocumentItem;
                                                setShowSpinner(false);
                                                resolve(docItem);
                                            } else {
                                                setShowSpinner(false);
                                                setDialogProps("Error", "No capability documents found or unexpected data structure");
                                                reject(new Error("No capability documents found or unexpected data structure"));
                                            }

                                        },
                                        error => {
                                            setShowSpinner(false);
                                            setDialogProps("Error", "Error updating document properties");
                                            reject(error);
                                        }
                                    )
                                },
                                error => {
                                    setShowSpinner(false);
                                    setDialogProps("Error", "Error retrieving updated document");
                                    reject(error);
                                }
                            );
                        });
                    } else {
                        setShowSpinner(false);
                        setDialogProps("Error", "File uploaded, but no file metadata returned.");
                        reject(new Error("File uploaded, but no file metadata returned."));
                    }
                },
                uploadError => {
                    setShowSpinner(false);
                    setDialogProps("Error Uploading", formatError(uploadError));
                    reject(uploadError);
                }
            );
        });
    };

    const handleOnAddDocument = async (): Promise<void> => {
        try {
            const file = await Helper.ListForm.showFileDialog();

            if (!file || !file.src) {

                throw new Error("No file selected or invalid file structure.");
            }

            const buffer: ArrayBuffer = await file.src.arrayBuffer();
            const customFile: CustomFile = Object.assign(file.src, { data: buffer });

            // Step 1: verify doc folder status
            if (docFolderStatus !== "ready") {
                // if folder status is not ready, check for one using helper
                setSpinnerProps("Checking for document folder...");
                const isFolderPresent = await DocumentService.ensureCapDocumentFolder(capState.Id);
                if (isFolderPresent) {
                    setDocFolderStatus("ready");
                } else {
                    try {
                        await DocumentService.createCapabilityFolder(capState.Id);
                    } catch (error) {
                        setDocFolderStatus("error")
                        setDialogProps("Error Creating Document Folder", formatError(error))
                        console.log("Error creating document folder", error)
                        return;
                    }
                }
            }

            // Step 2: Upload the file and get the new associated item
            const newDocument = await uploadFile(file.name, customFile);

            if (!newDocument) {
                throw new Error("Failed to upload file and retrieve document.");
            }

            await new Promise((resolve) => setTimeout(resolve, 100)); // Yield control for React to update

            // Step 3: refresh doc lib
            getDocs(capState.Id).catch((error) => {
                console.error("Unhandled promise rejection:", error);
            });

        } catch (error) {
            console.error("Error during upload or edit:", error);
            setDialogProps("Error Handling Document", formatError(error));
        }
    };

    const handleDownloadDoc = async (fileUrl: string, webUrl: string = ContextInfo.webServerRelativeUrl): Promise<void> => {
        // Downloads the document
        window.open(webUrl + "/_layouts/15/download.aspx?SourceUrl=" + fileUrl, "_blank");
    }

    const getFileIcon = (fileType: string): string => {
        switch (fileType.toLowerCase()) {
            // Word Documents
            case 'doc':
            case 'docx':
            case 'dotx':
            case 'dotm':
            case 'docm':
                return 'WordDocument';

            // Excel Files
            case 'xls':
            case 'xlsx':
            case 'xlsm':
            case 'xlsb':
            case 'xltx':
            case 'xltm':
                return 'ExcelDocument';

            // PowerPoint Files
            case 'ppt':
            case 'pptx':
            case 'pptm':
            case 'ppsx':
            case 'ppsm':
            case 'potx':
            case 'potm':
                return 'PowerPointDocument';

            // PDF Files
            case 'pdf':
                return 'PDF';

            // Image Files
            case 'jpg':
            case 'jpeg':
            case 'png':
            case 'gif':
            case 'bmp':
            case 'tiff':
                return 'FileImage';

            // Text Files
            case 'txt':
                return 'TextDocument';

            // Video Files
            case 'mp4':
            case 'avi':
            case 'mov':
            case 'wmv':
            case 'flv':
            case 'mkv':
                return 'Video';

            // Code Files
            case 'html':
            case 'css':
            case 'js':
            case 'ts':
            case 'json':
            case 'xml':
            case 'py':
            case 'java':
            case 'c':
            case 'cpp':
            case 'cs':
            case 'php':
            case 'rb':
            case 'swift':
            case 'go':
                return 'FileCode';

            // Default / Other Files
            default:
                return 'Page';
        }
    };

    const handleEditDocument = async (documentItem: IDocumentItem): Promise<void> => {
        if (!documentItem) return;

        setSpinnerProps("Updating document properties...");
        setShowSpinner(true);

        try {
            const modDocument = await DocumentService.edit(documentItem);

            if (modDocument) {
                // Update the existing item
                setDocuments((prevDocs) =>
                    prevDocs.map((pd) => (pd.Id === modDocument.Id ? modDocument : pd))
                );
            } else {
                // Remove the deleted item
                setDocuments((prevDocs) =>
                    prevDocs.filter((pd) => pd.Id !== documentItem.Id)
                );
            }
        } catch (err) {
            const errorMessage = formatError(err);
            setDialogProps("Error Editing Document", errorMessage);
            console.error("Error editing document", err);
        } finally {
            setShowSpinner(false);
            setShowDocEditDialog(false);
        }
    };

    // handle inline document deletion
    const handleDeleteInline = async (documentItem: IDocumentItem): Promise<void> => {
        if (!documentItem) return;

        setSpinnerProps("Deleting document...");
        setShowSpinner(true);

        try {
            await DocumentService.delete(documentItem.Id, () => {
                setDocuments((prevDocs) =>
                    prevDocs.filter((pd) => pd.Id !== documentItem.Id)
                );
            });
        } catch (err) {
            const errorMessage = formatError(err);
            setDialogProps("Error Deleting Document", errorMessage);
            console.error("Error deleting document", errorMessage);
        } finally {
            setShowSpinner(false);
            setShowDocDeleteDialog(false);
            setShowDocEditDialog(false);
        }
    };

    const documentColumns: IColumn[] = [
        {
            key: "icon",
            name: "",
            fieldName: "FileLeafRef",
            minWidth: 24,
            maxWidth: 24,
            isResizable: false,
            onRender: (item: IDocumentItem) => {
                const icon = getFileIcon(item.File_x0020_Type);
                return <Icon iconName={icon} />
            },
        },
        {
            key: "filename",
            name: "Filename",
            fieldName: "FileLeafRef",
            minWidth: 300,
            isResizable: true,
            onRender: (item: IDocumentItem) => {

                const officeExtensions = [
                    "docx", "docm", "dotx", "dotm", "doc",
                    "xlsx", "xlsm", "xlsb", "xltx", "xltm", "xls",
                    "pptx", "pptm", "ppsx", "ppsm", "potx", "potm", "ppt",
                    "one", "onetoc2",
                    "vsdx", "vsdm", "vssx", "vssm", "vstx", "vstm"
                ];
                const isOfficeDoc = officeExtensions.includes(item.File_x0020_Type?.toLowerCase() ?? "");

                // Use WOPI for Office docs, direct URL for others
                const openUrl = isOfficeDoc
                    ? `${ContextInfo.webAbsoluteUrl}/_layouts/15/WopiFrame.aspx?sourcedoc=${item.UniqueId}&file=${encodeURIComponent(item.FileLeafRef)}&action=default`
                    : item.ServerRedirectedEmbedUrl || item.EncodedAbsUrl; //for text and images, fall back to AbsUrl

                return (
                    <a
                        href="#"
                        onClick={(e) => {
                            e.preventDefault(); // Prevent the default anchor behavior
                            window.open(openUrl, '_blank'); // Open the URL in a new tab
                        }}
                        rel="noopener noreferrer"
                        style={{ textDecoration: "none" }}
                    >
                        {item.FileLeafRef}
                    </a>
                );
            },
        },
        {
            key: "docType",
            name: "Doc Type",
            fieldName: "docType",
            minWidth: 120,
            isResizable: true,
        },
        {
            key: "modified",
            name: "Modified",
            fieldName: "Modified",
            minWidth: 100,
            maxWidth: 120,
            isResizable: true,
            isSorted: true,
            isSortedDescending: true,
            onRender: (item: IDocumentItem) =>
                new Date(item.Modified).toLocaleDateString(),
        },
        {
            key: "editor",
            name: "Modified By",
            fieldName: "Editor",
            minWidth: 120,
            isResizable: true,
            onRender: (item: IDocumentItem) => item.Editor?.Title,
        },
        {
            key: "actions",
            name: "Actions",
            fieldName: "actions",
            minWidth: 100,
            maxWidth: 120,
            isResizable: false,
            onRender: (item: IDocumentItem) => (
                <div style={{ display: "flex", gap: 6 }}>
                    <IconButton
                        iconProps={{ iconName: "Download" }}
                        title="Download"
                        ariaLabel="Download"
                        onClick={() =>
                            handleDownloadDoc(`${ContextInfo.webAbsoluteUrl}/${Strings.Sites.main.lists.Documents}/${item.FileLeafRef}`)
                                .catch(error => {
                                    const errorMessage = formatError(error);
                                    setDialogProps(`Error Downloading`, errorMessage);
                                    console.error('Error downloading inline', errorMessage);
                                })
                        }
                    />
                    <IconButton
                        iconProps={{ iconName: "Edit" }}
                        title="Edit"
                        ariaLabel="Edit"
                        onClick={() => {
                            setSelectedDocument(item);
                            setShowDocEditDialog(true);
                        }}
                    />
                    <IconButton
                        iconProps={{ iconName: "Delete" }}
                        title="Delete"
                        ariaLabel="Delete"
                        style={{ color: Strings.PillStyles.RedColor }}
                        onClick={() => {
                            setSelectedDocument(item);
                            setShowDocDeleteDialog(true);
                        }}
                    />
                </div>
            ),
        }
    ];

    const isPoc = 
        capState.primaryPoc?.Id === Security.currentUserID ||
        contracts.some((contract) => contract.contractPm?.Id === Security.currentUserID) || 
        capState.Author?.Id === Security.currentUserID;
    const canViewContract = Security.IsAdmin || isPoc;
    const canEdit = Security.IsAdmin || isPoc;

    useEffect(() => {
        if (activeTab === "contract" && !canViewContract) {
            onTabChange?.("overview");
        }
    }, [activeTab, canViewContract, onTabChange]);

    return (
        <div className={styles.dcTracker}>
            <AppHeader onNewCapability={onNewCapability} />

            {/* Page details */}
            <Stack tokens={{ childrenGap: 16 }} className={styles.pageContent}>
                {/* Back Button & Header */}
                <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                    <CommandBarButton text="Back" iconProps={{ iconName: "NavigateBack" }} onClick={onBack} />
                </Stack>
                <Stack className={styles.recordShell}>
                    <Stack horizontal horizontalAlign="space-between" verticalAlign="start" tokens={{ childrenGap: 16 }} className={styles.recordHeader}>
                        <Stack verticalAlign="center" tokens={{ childrenGap: 4 }}>
                            <Stack horizontal wrap verticalAlign="center" tokens={{ childrenGap: 10 }}>
                                <Text variant="xxLarge" styles={{ root: { fontWeight: 600 } }}>{capState.Title}</Text>
                                <Pill
                                    text={capState.capStatus}
                                    backgroundColor={statusProps.backgroundColor}
                                    textColor={statusProps.textColor}
                                    style={{ height: 16, whiteSpace: 'nowrap' }}
                                />
                            </Stack>
                        </Stack>
                        <Stack horizontal tokens={{ childrenGap: 4 }}>
                            <PrimaryButton
                                text="Edit"
                                title="Edit Capability Details"
                                iconProps={{ iconName: "Edit" }}
                                disabled={!canEdit}
                                onClick={() => setShowCapabilityForm(true)}
                                style={{ minWidth: 125 }}
                            />
                            <DefaultButton
                                text="Export"
                                title="Export Capability to a single-page PDF"
                                iconProps={{ iconName: "Export" }}
                                onClick={async (): Promise<void> => {

                                    setSpinnerProps("Preparing Export...");
                                    setShowSpinner(true);

                                    // get first screenshot (only 1)
                                    const screenshot = documents.find(docs => docs.docType === "Screenshot")
                                    let screenshotBinary: string | undefined;

                                    if (screenshot) {
                                        try {
                                            //get screenshot image binary for PDF export
                                            screenshotBinary = await fetchFileDataUrlFromDocumentItem(context, screenshot);
                                        } catch (e) {
                                            console.error("Error getting screenshot binary: ", e)
                                            screenshotBinary = undefined; // fallback to "No screenshot available"
                                        }
                                    }

                                    await exportCapabilityPdf({
                                        capability: capState,
                                        contract: primaryContract,
                                        kgsLogoDataUrl: Strings.Logo,
                                        fileName: capState.Title,
                                        screenshotBinary
                                    });

                                    setShowSpinner(false);
                                }}
                                style={{ minWidth: 125 }}
                            />
                        </Stack>
                    </Stack>

                    {/********************************** Tabs ***********************************/}
                    <div className={styles.recordTabs}>
                        <Pivot
                            selectedKey={activeTab}
                            styles={customPivotStyles}
                            linkFormat="tabs"
                            onLinkClick={(item) => {
                                const key = item?.props.itemKey;
                                onTabChange?.(isCapabilityTab(key) ? key : "overview");
                            }}
                        >
                            <PivotItem itemKey="overview" headerText="Overview" />
                            <PivotItem itemKey="supporting" headerText="Supporting Info" />
                            <PivotItem itemKey="tagging" headerText="Tagging" />
                            {canViewContract && <PivotItem itemKey="contract" headerText="Contract" />}
                            <PivotItem itemKey="documentation" headerText="Documentation" />
                        </Pivot>
                    </div>

                    {/******************************* TAB CONTENT *****************************/}
                    <Stack className={styles.recordContent}>
                        {activeTab === "overview" && (
                            <CapabilityOverview
                                capState={capState}
                                rightContent={
                                    screenshots && screenshots.length > 0 ? (
                                        <ScreenshotCarousel
                                            screenshots={screenshots}
                                            solutionTitle={capState.Title}
                                            webUrl={ContextInfo.webAbsoluteUrl}
                                        />
                                    ) : null
                                }
                            />
                        )}

                        {activeTab === "supporting" && (
                            <SupportingInfo capState={capState} />
                        )}

                        {activeTab === "tagging" && (
                            <TaggingInfo capState={capState} />
                        )}

                        {activeTab === "contract" && canViewContract && (
                            <ContractInfo contracts={contracts} />
                        )}

                        {activeTab === "documentation" && (
                            <Stack tokens={{ childrenGap: 12 }}>
                                <Stack horizontal horizontalAlign="end">
                                    {/* {Security.IsMember && */}
                                    <PrimaryButton
                                        text="Add Document"
                                        title="Add a new Document"
                                        disabled={!canEdit}
                                        iconProps={{ iconName: "Add" }}
                                        onClick={() => setShowDocUploadDialog(true)}
                                    />
                                    {/* } */}
                                </Stack>
                                <Stack tokens={{ childrenGap: 20 }} className={styles.detailCard} style={{ width: "100%" }}>

                                    {documents.length === 0 ? (
                                        <div style={{ padding: "20px 0" }}>
                                            No documents are posted for this solution
                                        </div>
                                    ) : (
                                        <DetailsList
                                            items={documents}
                                            columns={documentColumns}
                                            selectionMode={SelectionMode.none}
                                            layoutMode={DetailsListLayoutMode.fixedColumns}
                                            onRenderRow={(props?: IDetailsRowProps) => {
                                                if (!props) return null;
                                                return (
                                                    <DetailsRow
                                                        {...props}
                                                        styles={{
                                                            root: { fontSize: 14, display: 'flex', alignItems: 'center' },
                                                            cell: { fontSize: 14, display: 'flex', alignItems: 'center' },
                                                        }}
                                                    />
                                                );
                                            }}
                                        />
                                    )}

                                </Stack>
                            </Stack>
                        )}
                    </Stack>
                </Stack>

            </Stack>

            {/* Capability Form Dialog */}
            <Dialog
                hidden={!showCapabilityForm}
                onDismiss={() => setShowCapabilityForm(false)}
                dialogContentProps={{
                    type: DialogType.largeHeader,
                    title: `Edit Capability - ${capState.Title}`,
                    showCloseButton: true
                }}
                modalProps={{
                    isBlocking: false,
                    styles: { main: dialogStyles.mainOverride }
                }}
            >
                <CapForm
                    key={[
                        capState.Id,
                        capState.Modified ?? "",
                        capState.oppNetTagsJson ?? "",
                        capState.pastPerformanceTagsJson ?? "",
                        capState.proposalTagsJson ?? ""
                    ].join("|")}
                    item={capState}
                    context={context}
                    onCancel={() => setShowCapabilityForm(false)}
                    onSave={async (result: ICapFormSaveResult) => {
                        try {
                            setSpinnerProps("Editing Capability...");
                            const updatedCapability = await CapabilityService.edit(result.capability);
                            await ContractService.saveForCapability(updatedCapability.Id, result.contracts, result.deletedContractIds);
                            await DataSource.init(true, context);
                            setCapState(updatedCapability);
                            setShowCapabilityForm(false);
                            setShowSpinner(false);
                        } catch (err) {
                            setShowSpinner(false);
                            console.error(`Error editing Capability: ${formatError(err)}`);
                            setDialogProps("Error editing Capability", formatError(err));
                        }
                    }}
                    onDelete={async () => {
                        setSpinnerMessage("Deleting Capability...");
                        setShowSpinner(true);
                        try {
                            await CapabilityService.delete(capState.Id);
                            //setCapState(undefined);
                            setShowCapabilityForm(false);
                            setShowSpinner(false);
                            onBack();
                        } catch (error) {
                            console.error(`Error deleting Capability: ${formatError(error)}`);
                            setShowSpinner(false);
                            setDialogProps("Error deleting Capability", formatError(error));
                        }
                    }}

                />
            </Dialog>

            {/* Confirmation of document delete */}
            <Dialog
                hidden={!showDocDeleteDialog}
                onDismiss={() => setShowDocDeleteDialog(false)}
                dialogContentProps={{
                    type: DialogType.normal,
                    title: "Delete Document",
                    closeButtonAriaLabel: 'Cancel',
                    subText: "Are you sure you want to delete this document?"
                }}
            >
                <DialogFooter>
                    <PrimaryButton text="Delete"
                        className={styles.deleteButton}
                        title="Delete this Document"
                        iconProps={{ iconName: "Delete" }}
                        onClick={() => handleDeleteInline(selectedDocument!)}
                    />
                    <DefaultButton onClick={() => setShowDocDeleteDialog(false)} text="Cancel" title="Close Dialog Box" />
                </DialogFooter>
            </Dialog>

            {/* Upload/Add document Dialog */}
            <Dialog
                hidden={!showDocUploadDialog}
                onDismiss={() => setShowDocUploadDialog(false)}
                dialogContentProps={{
                    type: DialogType.normal,
                    title: "Upload new Document",
                    closeButtonAriaLabel: 'Close',
                    subText: "You must select a document type before uploading"
                }}
            >
                <Dropdown
                    label="Document Type"
                    required
                    options={docTypeOptions}
                    onChange={(_, option?: IDropdownOption) => setDocumentType((option?.key as string) ?? undefined)}
                    style={{ marginBottom: 20 }}
                />
                <div style={{ padding: "10px 0" }}>&nbsp;</div>
                <DialogFooter>
                    <PrimaryButton
                        text="Upload Document"
                        iconProps={{ iconName: "Add" }}
                        disabled={!documentType}
                        title="Upload a new Document"
                        onClick={async () => {
                            try {
                                await handleOnAddDocument();
                            } catch (error) {
                                setShowDocUploadDialog(false);
                                const errorMessage = formatError(error);
                                setDialogProps(`Error adding`, errorMessage);
                                console.error("Error adding doc", errorMessage);
                            } finally {
                                setShowSpinner(false);
                                setShowDocUploadDialog(false);
                            }
                        }}
                    />
                    <DefaultButton onClick={() => setShowDocUploadDialog(false)} text="Cancel" title="Close Dialog Box" />
                </DialogFooter>
            </Dialog>

            {/* Edit document props Dialog */}
            <Dialog
                hidden={!showDocEditDialog}
                onDismiss={() => setShowDocEditDialog(false)}
                dialogContentProps={{
                    type: DialogType.normal,
                    title: "Edit Document Properties",
                    closeButtonAriaLabel: 'Close'
                }}
            >
                <DocumentForm
                    item={selectedDocument!}
                    onSave={handleEditDocument}
                    onCancel={() => setShowDocEditDialog(false)}
                    onDelete={(doc) => {
                        setSelectedDocument(doc);
                        setShowDocDeleteDialog(true);
                    }}
                />

            </Dialog>


            {/* Loading Spinner Dialog */}
            <Dialog
                hidden={!showSpinner}
                onDismiss={() => setShowSpinner(false)}
                dialogContentProps={{
                    type: DialogType.normal,
                    title: "Loading...",
                    closeButtonAriaLabel: 'Close',
                }}
            >
                <Spinner size={SpinnerSize.large} label={spinnerMessage} />
            </Dialog>

            {/* error message Dialog */}
            <Dialog
                hidden={!showDialog}
                onDismiss={hideDialog}
                dialogContentProps={{
                    type: DialogType.normal,
                    title: dialogTitle,
                    closeButtonAriaLabel: 'Close',
                    subText: dialogMessage
                }}
            >
                <DialogFooter>
                    <DefaultButton onClick={hideDialog} text="Close" title="Close Dialog Box" />
                </DialogFooter>
            </Dialog>

        </div>
    );
}
