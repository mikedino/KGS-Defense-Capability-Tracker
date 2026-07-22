import { IDropdownOption } from "@fluentui/react";

/****************************
 * 
 * LICENSE COST OPTIONS
 *  
 *****************************/
export const licenseReqdOptions = (includeAll: boolean = true): IDropdownOption[] => {
    const baseOptions: IDropdownOption[] = [
        { key: "Yes", text: "Yes" },
        { key: "No", text: "No" },
        { key: "Depends", text: "Depends" }
    ];
    return includeAll
        ? [{ key: "all", text: "All Locations" }, ...baseOptions]
        : baseOptions;
}
