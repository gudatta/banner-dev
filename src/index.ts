import { Components, Helper } from "gd-sprest-bs";
import { Configuration } from "./cfg";
import { Message } from "./message";
import Strings, { setContext, updateConfigWeb } from "./strings";
import { UserInformation } from "./userInfo";

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

    // See if this is a popup dialog for a classic page, and do nothing
    if (document.location.search.indexOf("IsDlg") > 0) { return; }

    // See if the context was provided
    if (props.context) {
        setContext(props.context, props.webUrl);
    }
    // Else, see if the web url was provided
    else if (props.webUrl) {
        // Update the config
        updateConfigWeb(props.webUrl);
    }

    // Clear the element
    while (props.el.firstChild) { props.el.removeChild(props.el.firstChild); }

    // Get the message
    Message().then(message => {
        // Get the user information
        UserInformation().then(userInfo => {
            // See if the user information exists
            if(userInfo) {
                // Append the user to the message
                message += `<span>User: ${userInfo.displayName}</span>`;
            }

            // Create the alert
            Components.Alert({
                el: props.el,
                className: "m-0",
                isDismissible: false,
                content: message,
                type: Components.AlertTypes.Info
            });
        });
    });
}

// Set the configuration
GlobalVariable["Configuration"] = Configuration;

// Set the global variable
window[Strings.GlobalVariable] = GlobalVariable;

// Notify other scripts that this library is loaded
Helper.SP.SOD.notifyScriptLoadedAndExecuteWaitingJobs("sp-banner");