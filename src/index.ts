import { Components, Helper } from "gd-sprest-bs";
import { Configuration } from "./cfg";
import { Message } from "./message";
import Strings, { updateConfigWeb } from "./strings";

interface ISPBanner {
    el: HTMLElement;
    context?: any;
    webUrl?: string;
}

/**
 * SharePoint Banner
 */
const GlobalVariable = (props: ISPBanner) => {
    // Ensure the element exists
    if (props.el == null) {
        // Log
        console.error("[SP Banner] The 'el' property was not passed.");
        return;
    }

    // Clear the element
    while (props.el.firstChild) { props.el.removeChild(props.el.firstChild); }

    // See if this is a popup dialog for a classic page, and do nothing
    if (document.location.search.indexOf("IsDlg") > 0) { return; }

    // See if the web url exists
    if (props.webUrl) {
        // Update the config
        updateConfigWeb(props.webUrl);
    }

    // Get the message
    Message().then(message => {
        // Create the alert
        Components.Alert({
            el: props.el,
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