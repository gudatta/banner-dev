import { ContextInfo } from "gd-sprest-bs";

// Sets the context information
// This is for SPFx or Teams solutions
export const setContext = (context, sourceUrl?: string) => {
    // Set the context
    ContextInfo.setPageContext(context.pageContext);

    // Update the source url
    updateConfigWeb(sourceUrl || ContextInfo.webServerRelativeUrl);
}

// Updates the config web target
export const updateConfigWeb = (webUrl:string) => {
    // Set the url
    Strings.SourceUrl = webUrl;
    Strings.ConfigUrl = Strings.SourceUrl + "/SiteAssets/sp-banner/banner.html";
}

/**
 * Global Constants
 */
const Strings = {
    AppElementId: "sp-banner",
    ConfigUrl: ContextInfo.webServerRelativeUrl + "/SiteAssets/sp-banner/banner.html",
    GlobalVariable: "SPBanner",
    ProjectName: "SP Banner",
    ProjectDescription: "Displays a banner on classic pages.",
    SourceUrl: ContextInfo.webServerRelativeUrl,
    Version: "0.1"
};
export default Strings;