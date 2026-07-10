import { IDocumentItem } from "../data/props";
import { Web } from "gd-sprest-bs";
import Strings from "../../strings";
import { formatError } from "../utils";
import { ContextInfo } from "gd-sprest-bs";

export class DocumentService {

    //upon creation of a new solution > create a folder in the documents library to store docs
    static createApplicationFolder(applicationId: number): Promise<void> {
        return new Promise<void>((resolve, reject) => {
            const list = Web().Lists(Strings.Lists.Documents);

            try {
                list.RootFolder().Folders()
                    .add(applicationId.toString())
                    .execute(
                        () => resolve(),
                        error => reject(error)
                    );

            } catch (error) {
                console.error("Error creating document folder for the new Application", error);
                reject(error);
            }
        });
    }

    // check to verify a folder exists before uploading
    // check to verify a folder exists before uploading
    static async ensureAppDocumentFolder(appId: number): Promise<boolean> {
        const folderUrl = `${ContextInfo.webServerRelativeUrl}/${Strings.Lists.Documents}/${appId}`;

        try {
            const resp = await Web().getFolderByServerRelativeUrl(folderUrl).executeAndWait();
            if(resp.existsFl || resp.Exists) { 
                return true;
            } else return false;
        } catch {
            return false;

        }

    }


    static edit(doc: IDocumentItem): Promise<IDocumentItem> {
        return new Promise<IDocumentItem>((resolve, reject) => {
            Web().Lists(Strings.Lists.Documents).Items(doc.Id).update({
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

    // DELETE DOCUMENT BY SOLUTION
    static delete(documentId: number, onUpdated: () => void): Promise<void> {

        return new Promise<void>((resolve, reject) => {

            // Delete the item
            Web().Lists(Strings.Lists.Documents).Items(documentId).delete().execute(
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

}