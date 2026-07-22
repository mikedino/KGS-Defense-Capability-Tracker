import { InstallationRequired } from "dattatable";
import { Components } from "gd-sprest-bs";
import { Configuration } from "../data/cfg";
import { Security } from "./Security";
import Strings from "../common/strings";

/**
 * Installation Modal
 */
export class InstallationModal {
    /**
     * Shows the installation modal if install is required or if showFl is true.
     * @param showFl - Forces the modal to appear, e.g., if security groups are missing
     * @param onComplete - Optional callback to execute after setup (e.g., render web part)
     */
    static show(showFl: boolean = false, onComplete?: () => void): Promise<void> {

        //return a promise
        return new Promise<void>((resolve, reject) => {

            InstallationRequired.requiresInstall({ cfg: Configuration }).then(installFl => {

                const errors: Components.IListGroupItem[] = [];
                const groupsMissing = !Security.Admins || !Security.Contributors || !Security.Visitors;

                if (groupsMissing) {
                    errors.push({ content: "The security groups have not been initialized." });
                }

                if (errors.length > 0) {
                    showFl = true;
                }

                if (installFl || showFl) {
                    InstallationRequired.showDialog({
                        errors,
                        onBodyRendered: el => {
                            const elem: HTMLDivElement = document.createElement("div");
                            elem.innerHTML = "<p>IMPORTANT: Install lists before the security groups! (if applicable)</p>";
                            elem.style.color = "red"; // Set font color to red
                            elem.style.paddingTop = "20px";
                            el.append(elem);
                        },
                        onFooterRendered: el => {
                            if (groupsMissing) {
                                Components.Tooltip({
                                    el,
                                    content: "Manage the security groups",
                                    type: Components.ButtonTypes.OutlinePrimary,
                                    btnProps: {
                                        text: "Install Security",
                                        //isDisabled: !InstallationRequired.ListsExist,
                                        onClick: () => {
                                            Security.show(() => {
                                                onComplete?.();
                                                resolve(); // done
                                            });
                                            // Security.createGroups().then(() => {
                                            //     resolve();
                                            //     window.location.reload();
                                            // }, error => {
                                            //     reject(error);
                                            // });
                                        }
                                    }
                                });
                            } else {
                                onComplete?.();
                                resolve(); // done
                            }
                        }
                    });
                } else {
                    console.warn(`[${Strings.ProjectName}] No installation or security setup required.`);
                    onComplete?.();
                    resolve(); // done
                }
            }, (err) => {
                console.error("Installation check failed:", err);
                reject(err); // allow caller to handle error
            });
        });

    }

}