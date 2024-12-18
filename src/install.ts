import { InstallationRequired } from "dattatable";
import { Configuration } from "./cfg";
import Strings from "./strings";

/**
 * Install Dialog
 */
export class InstallDialog {
    static show(force: boolean = false) {
        // See if an installation is required
        InstallationRequired.requiresInstall({ cfg: Configuration }).then(installFl => {
            // See if an install is required
            if (installFl || force) {
                // Show the dialog
                InstallationRequired.showDialog();
            } else {
                // Log
                console.error("[" + Strings.ProjectName + "] Error initializing the solution.");
            }
        });
    }
}