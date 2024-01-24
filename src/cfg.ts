import { Helper } from "gd-sprest-bs";

/**
 * Configuration
 */
export const Configuration = (webUrl: string) => {
    return Helper.SPConfig({
        CustomActionCfg: {
            Site: [
                {
                    Name: "SPBanner",
                    Title: "SharePoint Banner",
                    Description: "Displays a banner in IE browsers.",
                    Location: "ScriptLink",
                    Scope: 100000,
                    ScriptBlock: `
                    var s = document.createElement("script");
                    s.src = ${webUrl}/SiteAssets/sp-banner/sp-banner.js";
                    document.head.appendChild(s); SP.SOD.executeOrDelayUntilScriptLoaded(function() {
                        var el=document.createElement("div");
                        document.body.insertBefore(el, document.body.firstChild);
                        SPBanner(el);
                    }, "sp-banner");`.trim()
                }
            ]
        }
    });
}