import { Components, Helper } from "gd-sprest-bs";
import { Configuration } from "./cfg";
import { Message } from "./message";
import Strings from "./strings";

/**
 * SharePoint Banner
 */
const GlobalVariable = (el: HTMLElement) => {
    // Ensure the element exists
    if (el == null) {
        // Log
        console.error("[SP Banner] The 'el' property was not passed.");
        return;
    }

    // See if this is a popup dialog for a classic page, and do nothing
    if (document.location.search.indexOf("IsDlg") > 0) { return; }

    // Get the message
    Message().then(message => {
        // Create the alert
        Components.Alert({
            el,
            className: "m-0",
            isDismissible: false,
            content: message,
            type: Components.AlertTypes.Info
        });
    });
}

// Set the configuration
GlobalVariable["Configuration"] = Configuration;

// Set the global variable
window[Strings.GlobalVariable] = GlobalVariable;

// Notify other scripts that this library is loaded
Helper.SP.SOD.notifyScriptLoadedAndExecuteWaitingJobs("sp-banner");