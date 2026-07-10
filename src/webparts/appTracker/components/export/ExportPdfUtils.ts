import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClientResponse, SPHttpClient } from "@microsoft/sp-http";
import { IApplicationItem, IContractItem, IDocumentItem } from "../data/props";
import { IApplicationBookItem } from "./PDFBook";

const getImageMimeTypeFromFileType = (
    fileType?: string
): "image/png" | "image/jpeg" => {

    switch (fileType?.toLowerCase()) {
        case "jpg":
        case "jpeg":
            return "image/jpeg";

        case "png":
        default:
            return "image/png";
    }
};

/**
 * Helper to get the binary for the screenshot for the PDF export.
 * @param context - The web part context
 * @param doc - document object
 * @returns encoded picture base64
 */
export const fetchFileDataUrlFromDocumentItem = async (
    context: WebPartContext,
    doc: IDocumentItem
): Promise<string> => {

    const mimeType: "image/png" | "image/jpeg" = getImageMimeTypeFromFileType(doc.File_x0020_Type);

    const url: string =
        `${context.pageContext.web.absoluteUrl}` +
        `/_api/web/getfilebyid(guid'${doc.UniqueId}')/$value`;

    const response: SPHttpClientResponse =
        await context.spHttpClient.get(
            url,
            SPHttpClient.configurations.v1,
            { headers: { accept: "application/octet-stream" } }
        );

    if (!response.ok) {
        throw new Error(`Failed to fetch file bytes. HTTP ${response.status}`);
    }

    const buffer: ArrayBuffer = await response.arrayBuffer();

    // Convert ArrayBuffer → base64 → data URL
    let binary: string = "";
    const bytes: Uint8Array = new Uint8Array(buffer);

    for (let i = 0; i < bytes.byteLength; i++) {
        binary += String.fromCharCode(bytes[i]);
    }

    const base64: string = btoa(binary);
    return `data:${mimeType};base64,${base64}`;
};

/**
 * Helper to group iterations/docs by application Id
 * @param item - iterations or docs
 * @returns map of grouped items by application id
 */
export const groupByApplicationId = <T extends { application: { Id: number } }>(
    items: T[]
): Map<number, T[]> => {
    const map: Map<number, T[]> = new Map<number, T[]>();

    for (const item of items) {
        const sid: number = item.application.Id;
        const arr: T[] | undefined = map.get(sid);

        if (arr) {
            arr.push(item);
        } else {
            map.set(sid, [item]);
        }
    }

    return map;
};


/**
 * Compile the PDF Book, by applications
 * @param context - web part context
 * @param applications - filtered list of applications
 * @param contracts - all contracts
 * @param screenshotDocs - screenshots for app array
 * @returns map of grouped items by application ID
 */
export const buildPdfBookItems = async (
    context: WebPartContext,
    applications: IApplicationItem[],
    contracts: IContractItem[],
    screenshotDocs: IDocumentItem[]
): Promise<IApplicationBookItem[]> => {

    // contracts keyed by contract list item Id
    const contractById: Map<number, IContractItem> = new Map<number, IContractItem>();
    for (const contract of contracts) {
        contractById.set(contract.Id, contract);
    }

    // pick 1 screenshot per application (first one)
    const screenshotByApplication: Map<number, IDocumentItem> = new Map<number, IDocumentItem>();
    for (const doc of screenshotDocs) {
        const applicationId: number = doc.application.Id;

        if (!screenshotByApplication.has(applicationId)) {
            screenshotByApplication.set(applicationId, doc);
        }
    }

    const items: IApplicationBookItem[] = [];

    for (const app of applications) {
        const contractId: number | undefined = app.contract?.Id;
        const contract: IContractItem | undefined = contractId !== undefined
            ? contractById.get(contractId)
            : undefined;

        const screenshotDoc: IDocumentItem | undefined = screenshotByApplication.get(app.Id);
        let screenshotBinary: string | undefined;

        if (screenshotDoc) {
            try {
                screenshotBinary = await fetchFileDataUrlFromDocumentItem(context, screenshotDoc);
            } catch {
                screenshotBinary = undefined;
            }
        }

        items.push({
            application: app,
            contract,
            screenshotBinary
        });
    }

    return items;
};