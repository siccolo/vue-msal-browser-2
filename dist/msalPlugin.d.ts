import * as msal from "@azure/msal-browser";
import Vue from "vue";
declare let msalInstance: msal.PublicClientApplication | null;
interface ExtendedBrowserAuthOptions extends msal.BrowserAuthOptions {
    scopes?: Array<string>;
}
interface ExtendedConfiguration extends msal.Configuration {
    graph?: Response | {};
    mode?: "redirect" | "popup";
    auth: ExtendedBrowserAuthOptions;
}
export default class msalPlugin extends msal.PublicClientApplication {
    static install(vue: typeof Vue, msalConfig: ExtendedConfiguration): void;
    extendedConfiguration: ExtendedConfiguration;
    loginRequest: {
        scopes: Array<string>;
    };
    constructor(options: ExtendedConfiguration);
    callMSGraph(endpoint: string, accessToken: string): Promise<Response | void>;
    callMSGraphWithCallback(endpoint: string, accessToken: string, callback: any): void;
    callMSGraphAsPromise(endpoint: string, accessToken: string, callback: any): Promise<any>;
    postMSGraph(endpoint: string, accessToken: string, data: any, callback: any): void;
    getSilentToken(account: msal.AccountInfo, scopes?: string[]): Promise<msal.AuthenticationResult | void>;
    authenticate(): Promise<msal.AccountInfo[] | msal.AuthenticationResult>;
    authenticateRedirect(): Promise<msal.AccountInfo[]>;
    authenticatePopup(): Promise<msal.AuthenticationResult>;
}
export { msalInstance, ExtendedConfiguration, ExtendedBrowserAuthOptions };
