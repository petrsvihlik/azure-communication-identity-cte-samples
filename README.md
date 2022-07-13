---
page_type: sample
languages:
- javascript
products:
- azure
- azure-communication-services
---

# Create and manage access tokens

TODO

## Prerequisites

- An Azure account with an active subscription. Create an account for free.
- Node.js [Active LTS version](https://nodejs.org/en/about/releases/)
- An active Communication Services resource and connection string. Create a Communication Services resource.
- Azure Active Directory tenant with users that have a Teams license.

## Before running sample code

1. Complete the [Administrator actions](https://docs.microsoft.com/azure/communication-services/quickstarts/manage-teams-identity?pivots=programming-language-javascript#administrator-actions) from the [Manage access tokens for Teams users quickstart](https://docs.microsoft.com/azure/communication-services/quickstarts/manage-teams-identity).
  1. For the next steps, you will need Fabrikam's Azure AD tenant ID and Contoso's Azure AD App Client ID.
1. On the Authentication pane of your Azure AD App, add a new platform of the SPA type with the Redirect URI of `http://localhost:3000/spa`.
1. Open an instance of PowerShell, Windows Terminal, Command Prompt or equivalent and navigate to the directory that you'd like to clone the sample to.
1. `git clone https://github.com/Azure-Samples/communication-services-javascript-quickstarts.git`
1. With the Communication Services procured in pre-requisites, add connection string to environment variable using below command

setx COMMUNICATION_SERVICES_CONNECTION_STRING <YOUR_COMMUNICATION_SERVICES_CONNECTION_STRING>
setx AAD_TENANT_ID <FABRIKAM_AZURE_AD_TENANT_ID>
setx AAD_CLIENT_ID <CONTOSO_AZURE_AD_APP_CLIENT_ID>

1. Edit the `./App/authConfig.js` and set the `msalConfig.auth.clientId` to Contoso's Azure AD App Client ID.
1. 

## Run the code

From a console prompt, navigate to the directory containing the server.js file, then execute the following node command to run the app.

`npm start`
