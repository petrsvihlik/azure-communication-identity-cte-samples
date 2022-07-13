---
page_type: sample
languages:
- javascript
products:
- azure
- azure-communication-services
---

# Create and manage Communication access tokens for Teams users in a single-page application (SPA)

This code sample walks you through the process of acquiring a Communication Token Credential by exchanging an Azure AD token of a user with a Teams license for a valid Communication access token.

The client part of this sample utilizes the [MSAL.js v2.0](https://github.com/AzureAD/microsoft-authentication-library-for-js/tree/dev/lib/msal-browser) (`msal-browser`) package for authentication against the Azure AD and acquisition of a token with delegated permissions.
The initialization of a Communication credential object that can be used for Calling is achieved by the `@azure/communication-common` package.

The server part of the sample is based on [Express.js](https://expressjs.com/) and relies on widely used libraries such as `express-jwt` and `jwks-rsa` for Azure AD token validation. The token exchange itself is then facilitated by the `@azure/communication-identity` package.

## Prerequisites

- An Azure account with an active subscription. Create an account for free.
- Node.js [Active LTS version](https://nodejs.org/en/about/releases/)
- An active Communication Services resource and connection string. Create a Communication Services resource.
- Azure Active Directory tenant with users that have a Teams license.

## Before running sample code

1. Complete the [Administrator actions](https://docs.microsoft.com/azure/communication-services/quickstarts/manage-teams-identity?pivots=programming-language-javascript#administrator-actions) from the [Manage access tokens for Teams users quickstart](https://docs.microsoft.com/azure/communication-services/quickstarts/manage-teams-identity).
   - Take a not of Fabrikam's Azure AD Tenant ID and Contoso's Azure AD App Client ID. You'll need the values in the following steps.
1. On the Authentication pane of your Azure AD App, add a new platform of the SPA (single-page application) type with the Redirect URI of `http://localhost:3000/spa`.
1. Open an instance of PowerShell, Windows Terminal, Command Prompt or equivalent and navigate to the directory that you'd like to clone the sample to.
1. `git clone https://github.com/Azure-Samples/communication-services-javascript-quickstarts.git`
1. With the Communication Services procured in pre-requisites and Azure AD Tenant and App Registration procured as part of the Administrator actions, you can now add the connection string, tenant ID and app client ID to the environment variables using the commands below.

    ```powershell
    setx COMMUNICATION_SERVICES_CONNECTION_STRING <YOUR_COMMUNICATION_SERVICES_CONNECTION_STRING>
    setx AAD_TENANT_ID <FABRIKAM_AZURE_AD_TENANT_ID>
    setx AAD_CLIENT_ID <CONTOSO_AZURE_AD_APP_CLIENT_ID>
    ```

   - *Alternatively, you can add these values to a `.env` file in the root of the sample directory.*
1. Edit the `./App/authConfig.js` and set the `msalConfig.auth.clientId` to Contoso's Azure AD App Client ID.

    ```js
    msalConfig = {
        auth: {
          clientId: "<CONTOSO_AZURE_AD_APP_CLIENT_ID>"
        }
    }
    ```

## Run the code

From a console prompt, navigate to the directory containing the server.js file, then execute the following node command to run the app.

`npm start`
