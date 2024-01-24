import { Graph, SPTypes } from "gd-sprest-bs";

export interface IUserInformation {
    businessPhones: string[];
    displayName: string;
    givenName: string;
    id: string;
    jobTitle: string;
    mail: string;
    mobilePhone: string;
    officeLocation: string;
    preferredLanguage: string;
    suname: string;
    userPrincipalName: string;
}

/**
 * User Infomration
 */
export const UserInformation = (): PromiseLike<IUserInformation> => {
    // Return a promise
    return new Promise(resolve => {
        // Set the cloud env
        Graph.Cloud = SPTypes.CloudEnvironment.Default;

        // Get the access token
        Graph.getAccessToken().execute(token => {
            // Set the token
            Graph.Token = token.access_token;

            // Get the current user
            Graph({
                url: "me"
            }).execute(currentUser => {
                // Resolve the request
                resolve(currentUser as IUserInformation);
            }, resolve);
        });
    });
}