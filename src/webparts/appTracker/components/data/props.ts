import { IPersonaProps } from "@fluentui/react/lib/Persona";

interface ILookupItem {
    readonly Id: number;
    Title: string;
    contractTeam?: string;
}

export interface IPeoplePicker extends IPersonaProps {
    EMail: string;
    Id: number;
    Title: string;
}

export interface IPeoplePickerExtended extends IPeoplePicker{
    JobTitle?: string;
    Department?: string;
}

export type licenseReqdChoices = "Yes" | "No" | "Depends";
export type DocumentType = "Screenshot" | "Other";

export interface IApplicationItem {
    readonly Id: number;
    readonly Author?: IPeoplePicker;
    readonly Modified?: string;
    
    /* OVERVIEW */
    Title: string;
    description?: string;  //multi-line
    highlights?: string; //multi-line
    appUrl: string;
    relatedInfo: string; //multi-line - now "Core Capabilities & Relevant Software"
    primaryPoc?: IPeoplePickerExtended;
    stakeholders?: { results: IPeoplePicker[] }; 
    systemOwner?: IPeoplePickerExtended;
    managingGroup: string; //config
    supportTeam: string; //config
    userCount: number;
    licenseReqd: licenseReqdChoices;

    /* SUPPORTING INFO */
    connectivity: string; //config
    integrationsInput: string; //multi-line
    integrationsOutput: string; //multi-line
    synonyms: string;
    platform: string; //config
    environment: string; //config
    appStatus: string; //config
    appLaunchDate: string;
    
    /* CONTRACT INFO */
    contract?: ILookupItem;
}
// explicit form type leaving out readonly fields
//export type IApplicationForm = Omit<IApplicationItem, "Created" | "Modified">;

export interface IConfigurationItem {
    readonly Id: number;
    Title: string;
    isFor: string;
    isForDisplayName: string;
    isActive: boolean;
    infoText: string;
}

export interface IContractItem {
    readonly Id: number;
    Title: string;
    contractTeam: string; //config
    start: string; //date
    end: string; //date
    contractValue: number;
    cor?: IPeoplePickerExtended;
    acor?: IPeoplePickerExtended;
    ko?: IPeoplePickerExtended;
    primaryPoc?: IPeoplePickerExtended;
    secondaryPoc?: IPeoplePickerExtended;
}

export interface IDocumentItem {
    readonly Id: number;
    Title?: string;
    application: ILookupItem;
    docType: DocumentType;
    readonly UniqueId: string;  //GUID
    readonly FileLeafRef: string; //filename
    readonly EncodedAbsUrl: string; //direct file path
    readonly ServerRedirectedEmbedUrl: string;
    readonly Modified: string;
    readonly File_x0020_Type: string;
    readonly Editor: IPeoplePicker;
}
