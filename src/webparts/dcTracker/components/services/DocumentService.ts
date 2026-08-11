import { IContractDocumentItem, IDocumentItem } from "../common/props";
import { Web } from "gd-sprest-bs";
import Strings from "../common/strings";
import { formatError } from "../common/utils";
import { ContextInfo } from "gd-sprest-bs";

export class DocumentService {

    //upon creation of a new capability > create a folder in the documents library to store docs
    static createCapabilityFolder(capabilityId: number): Promise<void> {
        return new Promise<void>((resolve, reject) => {
            const list = Web().Lists(Strings.Sites.main.lists.Documents);

            try {
                list.RootFolder().Folders()
                    .add(capabilityId.toString())
                    .execute(
                        () => resolve(),
                        error => reject(error)
                    );

            } catch (error) {
                console.error("Error creating document folder for the new Capability", error);
                reject(error);
            }
        });
    }

    // check to verify a folder exists before uploading
    static async ensureCapDocumentFolder(capId: number): Promise<boolean> {
        const folderUrl = `${ContextInfo.webServerRelativeUrl}/${Strings.Sites.main.lists.Documents}/${capId}`;

        try {
            const resp = await Web().getFolderByServerRelativeUrl(folderUrl).executeAndWait();
            if(resp.existsFl || resp.Exists) { 
                return true;
            } else return false;
        } catch {
            return false;

        }

    }

    // Create the contract document folder named after the DCTContracts item ID.
    static createContractFolder(contractId: number): Promise<void> {
        return new Promise<void>((resolve, reject) => {
            const list = Web().Lists(Strings.Sites.main.lists.ContractDocuments);

            try {
                list.RootFolder().Folders()
                    .add(contractId.toString())
                    .execute(
                        () => resolve(),
                        error => reject(error)
                    );

            } catch (error) {
                console.error("Error creating document folder for the new Contract", error);
                reject(error);
            }
        });
    }

    // Check whether the contract document folder exists before upload.
    static async ensureContractDocumentFolder(contractId: number): Promise<boolean> {
        const folderUrl = `${ContextInfo.webServerRelativeUrl}/${Strings.Sites.main.lists.ContractDocuments}/${contractId}`;

        try {
            const resp = await Web().getFolderByServerRelativeUrl(folderUrl).executeAndWait();
            return !!(resp.existsFl || resp.Exists);
        } catch {
            return false;
        }
    }

    // Upload a document into the contract-specific folder and stamp its contract metadata.
    static uploadContractDocument(contractId: number, fileName: string, data: ArrayBuffer, cdocType: string): Promise<IContractDocumentItem> {
        return new Promise<IContractDocumentItem>((resolve, reject) => {
            const targetFolderUrl = `${ContextInfo.webServerRelativeUrl}/${Strings.Sites.main.lists.ContractDocuments}/${contractId}`;
            const list = Web().Lists(Strings.Sites.main.lists.ContractDocuments);

            Web().getFolderByServerRelativeUrl(targetFolderUrl).Files().add(fileName, true, data).execute(
                (uploadedFile) => {
                    if (!uploadedFile) {
                        reject(new Error("File uploaded, but no file metadata returned."));
                        return;
                    }

                    uploadedFile.ListItemAllFields().execute((listItem) => {
                        const newItemId = listItem?.Id;
                        list.Items().getById(newItemId).update({
                            contractId,
                            cdocType
                        }).execute(
                            () => {
                                list.Items().query({
                                    Select: ["File_x0020_Type", "UniqueId", "Id", "ServerRedirectedEmbedUrl", "EncodedAbsUrl", "FileLeafRef", "contract/Id",
                                        "Modified", "Editor/Id", "Editor/EMail", "Editor/Title", "cdocType", "Title"],
                                    Filter: `Id eq ${newItemId}`,
                                    Expand: ["contract", "Editor"]
                                }).execute(
                                    (items) => {
                                        if (items?.results?.length) {
                                            resolve(items.results[0] as unknown as IContractDocumentItem);
                                        } else {
                                            reject(new Error("Contract document uploaded, but the updated item could not be retrieved."));
                                        }
                                    },
                                    error => reject(error)
                                );
                            },
                            error => reject(error)
                        );
                    });
                },
                error => reject(error)
            );
        });
    }

    static edit(doc: IDocumentItem): Promise<IDocumentItem> {
        return new Promise<IDocumentItem>((resolve, reject) => {
            Web().Lists(Strings.Sites.main.lists.Documents).Items(doc.Id).update({
                Title: doc.Title
            }).execute(
                //success
                item => {
                    console.info(`Updated document ${item?.Id} !`)
                    resolve(item as unknown as IDocumentItem);
                },
                //error
                (error) => {
                    const err = formatError(error);
                    console.error(`Error updating document ${err}`);
                    reject(error);
                }
            )
        })
    }

    // Update editable metadata for a contract document item.
    static editContractDocument(doc: IContractDocumentItem): Promise<IContractDocumentItem> {
        return new Promise<IContractDocumentItem>((resolve, reject) => {
            Web().Lists(Strings.Sites.main.lists.ContractDocuments).Items(doc.Id).update({
                Title: doc.Title,
                cdocType: doc.cdocType
            }).execute(
                item => {
                    console.info(`Updated contract document ${item?.Id} !`)
                    resolve({ ...doc } as IContractDocumentItem);
                },
                (error) => {
                    const err = formatError(error);
                    console.error(`Error updating contract document ${err}`);
                    reject(error);
                }
            )
        })
    }

    // DELETE DOCUMENT BY SOLUTION
    static delete(documentId: number, onUpdated: () => void): Promise<void> {

        return new Promise<void>((resolve, reject) => {

            // Delete the item
            Web().Lists(Strings.Sites.main.lists.Documents).Items(documentId).delete().execute(
                // Success
                () => {
                    // Resolve the promise
                    resolve();

                    // Call the update event
                    onUpdated();
                },
                // Error
                (error) => {
                    const errorMessage = formatError(error);
                    console.error("Error Deleting Document", errorMessage);
                    reject(errorMessage);
                }
            );
        });

    }

    // Delete a contract document item from the contract documents library.
    static deleteContractDocument(documentId: number): Promise<void> {

        return new Promise<void>((resolve, reject) => {

            Web().Lists(Strings.Sites.main.lists.ContractDocuments).Items(documentId).delete().execute(
                () => resolve(),
                (error) => {
                    const errorMessage = formatError(error);
                    console.error("Error Deleting Contract Document", errorMessage);
                    reject(errorMessage);
                }
            );
        });

    }

}
