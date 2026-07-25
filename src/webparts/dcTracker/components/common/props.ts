import { IPersonaProps } from "@fluentui/react/lib/Persona";
import type { WebPartContext } from "@microsoft/sp-webpart-base"

export interface IDCTrackerProps {
    appDescription: string;
    context: WebPartContext;
}

export interface ILookupItem {
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
export type ConfigType =
    | "backend"
    | "capabilityStatus"
    | "codingLanguage"
    | "compliance"
    | "connectivity"
    | "customer"
    | "partner"
    | "hostingEnvironment"
    | "platform";
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
    primaryPoc?: IPeoplePickerExtended; //Capability Primary POC
    stakeholders?: { results: IPeoplePickerExtended[] }
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

    /* TAGS */
    oppNetTagsJson?: string;
    oppNetTags?: IOpportunityItem[];
    pastPerformanceTagsJson?: string;
    pastPerformanceTags?: IPastPerformanceItem[];
    proposalTagsJson?: string;
    proposalTags?: IProposalItem[];
}

export interface IContractItem {
    readonly Id: number;
    capability?: ILookupItem;
    Title: string; //Contract Title
    contractId?: string;
    customerContractCode?: string;
    customer?: string;
    startDate?: string; //date
    endDate?: string; //date
    contractPm?: IPeoplePickerExtended; //KGS Contract Project Manager
    partner?: string; //config 
    infoLink?: string;
}

export interface ICapabilityContractDraft extends IContractItem {
    tempId?: string;
}

export interface ICapFormSaveResult {
    capability: ICapabilityItem;
    contracts: ICapabilityContractDraft[];
    deletedContractIds: number[];
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


export interface IConfigItem {
    readonly Id: number;
    Title: string; // Display Text
    configType: ConfigType | string;
    configValue: string; // compared value/saved value
    sortOrder?: number;
    isActive: boolean;
    infoText?: string;
}

export interface IOpportunityItem {
    readonly Id: number;
    Title?: string;
    Customer?: string;
    Status?: string;
    url?: string;
}

export interface IPastPerformanceItem {
    readonly Id: number;
    Contract_x0023_?: string; //Contract #
    Customer_x0020_Agency?: string; //Customer Agency
    Doc_x0020_Type?: string;
    Capability_x0020_Area?: { results: string[] }; //multi-choice field
    url?: string;
}

export interface IProposalItem {
    readonly Id: number;
    Title: string;
    OpportunityStage?: string;
    TypeOfOpportunity?: string;
    Entity?: string;
    url?: string;
}
