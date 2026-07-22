import { ListSecurity } from "dattatable";
import { ContextInfo, SPTypes, Types } from "gd-sprest-bs";
import Strings from "../common/strings";

/**
 * Security
 * Code related to the security groups the user belongs to.
 */
export class Security {

    // Role display for user
    static RoleDisplay: string;

    // UserID for use throughout
    static currentUserID: number;

    // initialize object
    private static _listSecurity: ListSecurity | undefined;

    // Admin
    private static _isAdmin = false;
    static get IsAdmin(): boolean { return this._isAdmin; }
    private static _admins: Types.SP.GroupOData | undefined;
    static get Admins(): Types.SP.GroupOData | undefined { return this._admins; }
    private static _adminGroupInfo: Types.SP.GroupCreationInformation = {
        AllowMembersEditMembership: true,
        AutoAcceptRequestToJoinLeave: false,
        Description: "Admins/Owners for the Defense Capabilities Tracker app",
        Title: Strings.Groups.Admins,
        OnlyAllowMembersViewMembership: false
    };

    // Contributors
    private static _isContributor = false;
    static get IsContributor(): boolean { return this._isContributor; }
    private static _contributors: Types.SP.GroupOData | undefined;
    static get Contributors(): Types.SP.GroupOData | undefined { return this._contributors; }
    private static _contributorGroupInfo: Types.SP.GroupCreationInformation = {
        AllowMembersEditMembership: true,
        AutoAcceptRequestToJoinLeave: false,
        Description: "All Contributors for the Defense Capabilities Tracker app",
        Title: Strings.Groups.Contributors,
        OnlyAllowMembersViewMembership: false
    };

    // Visitors
    private static _isVisitor = false;
    static get IsVisitor(): boolean { return this._isVisitor; }
    private static _visitors: Types.SP.GroupOData | undefined;
    static get Visitors(): Types.SP.GroupOData | undefined { return this._visitors; }
    private static _visitorGroupInfo: Types.SP.GroupCreationInformation = {
        AllowMembersEditMembership: false,
        AutoAcceptRequestToJoinLeave: false,
        Description: "Visitors for the Defense Capabilities Tracker app",
        Title: Strings.Groups.Visitors,
        OnlyAllowMembersViewMembership: false
    };

    // Initializes the class
    static init(): Promise<void> {
        return new Promise((resolve, reject) => {

            // set the user ID
            this.currentUserID = ContextInfo.userId;

            this._listSecurity = new ListSecurity({
                webUrl: Strings.SiteUrl,
                groups: [
                    this._adminGroupInfo, this._contributorGroupInfo, this._visitorGroupInfo
                ],
                listItems: [
                    //Capabilities list
                    {
                        listName: Strings.Lists.Capabilities,
                        groupName: Strings.Groups.Admins,
                        permission: SPTypes.RoleType.Administrator
                    },
                    {
                        listName: Strings.Lists.Capabilities,
                        groupName: Strings.Groups.Contributors,
                        permission: SPTypes.RoleType.Contributor
                    },
                    {
                        listName: Strings.Lists.Capabilities,
                        groupName: Strings.Groups.Visitors,
                        permission: SPTypes.RoleType.Reader
                    },
                    //Documents list
                    {
                        listName: Strings.Lists.Documents,
                        groupName: Strings.Groups.Admins,
                        permission: SPTypes.RoleType.Administrator
                    },
                    {
                        listName: Strings.Lists.Documents,
                        groupName: Strings.Groups.Contributors,
                        permission: SPTypes.RoleType.Contributor
                    },
                    {
                        listName: Strings.Lists.Documents,
                        groupName: Strings.Groups.Visitors,
                        permission: SPTypes.RoleType.Reader
                    },
                    //Configuration list
                    {
                        listName: Strings.Lists.Configuration,
                        groupName: Strings.Groups.Admins,
                        permission: SPTypes.RoleType.Administrator
                    },
                    {
                        listName: Strings.Lists.Configuration,
                        groupName: Strings.Groups.Contributors,
                        permission: SPTypes.RoleType.Reader
                    },
                    {
                        listName: Strings.Lists.Configuration,
                        groupName: Strings.Groups.Visitors,
                        permission: SPTypes.RoleType.Reader
                    },
                    //Contracts list
                    {
                        listName: Strings.Lists.Contracts,
                        groupName: Strings.Groups.Admins,
                        permission: SPTypes.RoleType.Administrator
                    },
                    {
                        listName: Strings.Lists.Contracts,
                        groupName: Strings.Groups.Contributors,
                        permission: SPTypes.RoleType.Contributor
                    },
                    {
                        listName: Strings.Lists.Contracts,
                        groupName: Strings.Groups.Visitors,
                        permission: SPTypes.RoleType.Reader
                    }
                ],
                onGroupsLoaded: () => {
                    // Set the groups
                    this._admins = this._listSecurity?.getGroup(Strings.Groups.Admins);
                    this._contributors = this._listSecurity?.getGroup(Strings.Groups.Contributors);                    
                    this._visitors = this._listSecurity?.getGroup(Strings.Groups.Visitors);

                    // Set the user flags
                    if (this._listSecurity) {
                        if (this._listSecurity.isInGroup(ContextInfo.userId, Strings.Groups.Admins)) {
                            this._isAdmin = true;
                            this._isContributor = true;
                            this.RoleDisplay = "Administrator";
                        } else if (this._listSecurity.isInGroup(ContextInfo.userId, Strings.Groups.Contributors)) {
                            this._isContributor = true;
                            this.RoleDisplay = "Contributor";
                        } else if (this._listSecurity.isInGroup(ContextInfo.userId, Strings.Groups.Visitors)) {
                            this._isVisitor = true;
                            this.RoleDisplay = "Visitor";
                        } else {
                            this.RoleDisplay = "NoRole";
                        }
                    }

                    // Ensure the groups exist
                    if (this._admins && this._contributors && this._visitors) {
                        resolve();
                    } else {
                        reject();
                    }
                }
            });
        });
    }

    // See if the user has permissions
    public static hasPermissions(): Promise<boolean> {
        if (!this._listSecurity) {
            return Promise.resolve(false);
        }
        // checkUserPermissions only checks if you have ManangeWeb permissions !!!
        //return Promise.resolve(this._listSecurity.checkUserPermissions());
        return Promise.resolve(this._isAdmin);
    }

    // Displays the security group configuration
    static show(onComplete: () => void): void {
        if (!this._listSecurity) {
            throw new Error("ListSecurity object has not been initialized.");
        } else this._listSecurity.show(true, onComplete);
    }

}
