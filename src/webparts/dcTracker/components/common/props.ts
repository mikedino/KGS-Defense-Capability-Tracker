import { IPersonaProps } from "@fluentui/react/lib/Persona";
import type { WebPartContext } from "@microsoft/sp-webpart-base"

export interface IDCTrackerProps {
  appDescription: string;
  context: WebPartContext;
}

interface ILookupItem {
    readonly Id: number;
    Title: string;
    contractId?: string;
}

export interface IPeoplePicker extends IPersonaProps {
    EMail: string;
    Id: number;
    Title: string;
}

export interface IPeoplePickerExtended extends IPeoplePicker {
    JobTitle?: string;
    Department?: string;
}

export type DocumentType = "Screenshot" | "Other";
export type licenseReqdChoices = "Yes" | "No";
//export type CapabilityStatus = "Active" | "In Dev" | "Pending"

export interface ICapabilityItem {
    readonly Id: number;
    readonly Author?: IPeoplePicker;
    readonly Modified?: string;

    /* OVERVIEW */
    Title: string;
    description?: string;  //multi-line
    capabilities?: string; //multi-line
    link?: string;
    capStatus: string; //config
    notes?: string; //multi-line

    /* TECHNICAL INFO */
    platform: string; //config
    hostingEnv?: string; //config
    connectivity: string; //config
    compliance?: string; //config
    licenseReqd: licenseReqdChoices;
    licenseReqmts?: string; //multi-line
    extensibility?: string; //multi-line
    serverReqmts?: string;
    codeLanguage?: string;
    backend?: string;

    /* CONTRACT INFO */
    contract?: ILookupItem;
}

export interface IContractItem {
    readonly Id: number;
    Title: string; //Contract Title
    contractId?: string;
    invoice?: string;
    customerContractCode?: string;
    customer?: string;
    popStart?: string; //date
    popEnd?: string; //date
    contractPm?: IPeoplePickerExtended; //KGS Contract Project Manager
    primaryPoc?: IPeoplePickerExtended; //Capability Primary POC
    stakeholders?: { results: IPeoplePickerExtended[] }
    partner?: string; //config 
    infoLink?: string;
}

export interface IContractEndPointItem {
    readonly Id: number;
    Title: string;
    field_19: string; // Contract ID (e.g. 100158)
    field_20: string; // Contract Title
    field_35: string; // Customer Contract Code (e.g. 19AQMM21D0121)
    field_21: string; // Manager 1 Email (Project Manager)
    field_23: string; // Manager 1 Name (Project Manager)
    //field_73: string; // NAICS Code (e.g. 541519)
    //field_75: string; // OG
    //field_16?: string; // Completion Date (NOT USED - EMPTY)
}

export interface IDocumentItem {
    readonly Id: number;
    Title?: string;
    capability: ILookupItem;
    docType: string; //config
    readonly UniqueId: string;  //GUID
    readonly FileLeafRef: string; //filename
    readonly EncodedAbsUrl: string; //direct file path
    readonly ServerRedirectedEmbedUrl: string;
    readonly Modified: string;
    readonly File_x0020_Type: string;
    readonly Editor: IPeoplePicker;
}


export interface IConfigurationItem {
    readonly Id: number;
    Title: string;
    isFor: string;
    isForDisplayName: string;
    isActive: boolean;
    infoText: string;
}
