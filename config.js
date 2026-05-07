/**
 * Zentrale Konfiguration für die OrderCheck App.
 *
 * Alle sensiblen IDs und Endpoints sind hier an einer Stelle definiert.
 * Bei Änderungen müssen die Auth-Files (auth-start.html, auth-end.html)
 * und script.js nicht separat angepasst werden.
 */
export const CONFIG = {
    /** Microsoft Authentication Library (MSAL) Konfiguration */
    msal: {
        /** Azure AD App Registration Client ID */
        clientId: "9603fe2f-fad8-4a18-9950-14f1eb195745",
        /** Azure AD Tenant ID */
        tenantId: "2b71e8a3-7c31-40db-84b0-9c7e50b9ea71",
        /** Benötigte Microsoft Graph API Scopes */
        scopes: ['User.Read', 'Files.ReadWrite', 'Sites.ReadWrite.All']
    },
    /** SharePoint Online Konfiguration */
    sharePoint: {
        /** SharePoint Site ID (Format: domain,siteId,webId) */
        siteId: "edox.sharepoint.com,e68093ae-074a-4851-88dd-31284b07b284,e94104cd-b270-4e09-b29a-1932cc934dbf",
        /** SharePoint List ID für Prüfvorgänge und Fehlerpositionen */
        listId: "ecac4b28-01fb-489a-ad32-f5ee4fd4bd83"
    }
};
