/**
 * Configuration object to be passed to MSAL instance on creation. 
 * For a full list of MSAL.js configuration parameters, visit:
 * https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-browser/docs/configuration.md
 * For more details on using MSAL.js with Azure AD B2C, visit:
 * https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-browser/docs/working-with-b2c.md 
 */

const msalConfig = {
  auth: {
    clientId: "1875691f-131f-4802-95a5-4511bde1408e", // Multi-tenant
    //clientId: "834c8592-72f5-4890-ba10-fc04d1cb392e", // Single-tenant
    redirectUri: "http://localhost", // You must register this URI on Azure Portal/App Registration. Defaults to "window.location.href".
  },
  cache: {
    cacheLocation: "sessionStorage", // Configures cache location. "sessionStorage" is more secure, but "localStorage" gives you SSO between tabs.
    storeAuthStateInCookie: false, // If you wish to store cache items in cookies as well as browser cache, set this to "true".
  },
  system: {
    loggerOptions: {
      loggerCallback: (level, message, containsPii) => {
        if (containsPii) {
          return;
        }
        switch (level) {
          case msal.LogLevel.Error:
            console.error(message);
            return;
          case msal.LogLevel.Info:
            console.info(message);
            return;
          case msal.LogLevel.Verbose:
            console.debug(message);
            return;
          case msal.LogLevel.Warning:
            console.warn(message);
            return;
        }
      }
    }
  }
};

// Create the main myMSALObj instance
// configuration parameters are located at authConfig.js
const myMSALObj = new msal.PublicClientApplication(msalConfig);

let accountId = "";
let username = "";

function setAccount(account) {
  accountId = account.homeAccountId;
  username = account.username;
  welcomeUser(username);
}

function selectAccount() {
  /**
   * See here for more info on account retrieval: 
   * https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-common/docs/Accounts.md
   */

  const currentAccounts = myMSALObj.getAllAccounts();

  if (currentAccounts.length < 1) {
    return;
  } else if (currentAccounts.length > 1) {

    /**
     * Due to the way MSAL caches account objects, the auth response from initiating a user-flow
     * is cached as a new account, which results in more than one account in the cache. Here we make
     * sure we are selecting the account with homeAccountId that contains the sign-up/sign-in user-flow, 
     * as this is the default flow the user initially signed-in with.
     */
    const accounts = currentAccounts.filter(account =>

      account.idTokenClaims.aud === msalConfig.auth.clientId
    );

    if (accounts.length > 1) {
      // localAccountId identifies the entity for which the token asserts information.
      if (accounts.every(account => account.localAccountId === accounts[0].localAccountId)) {
        // All accounts belong to the same user
        setAccount(accounts[0]);
      } else {
        // Multiple users detected. Logout all to be safe.
        signOut();
      };
    } else if (accounts.length === 1) {
      setAccount(accounts[0]);
    }

  } else if (currentAccounts.length === 1) {
    setAccount(currentAccounts[0]);
  }
}

// in case of page refresh
selectAccount();

function handleResponse(response) {
  /**
   * To see the full list of response object properties, visit:
   * https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-browser/docs/request-response-object.md#response
   */

  if (response !== null) {
    setAccount(response.account);
  } else {
    selectAccount();
  }
}

function signIn() {

  /**
   * You can pass a custom request object below. This will override the initial configuration. For more information, visit:
   * https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-browser/docs/request-response-object.md#request
   */

  myMSALObj.loginPopup({
    scopes: ["openid"], // By default, MSAL.js will add OIDC scopes (openid, profile, email) to any login request.
  })
    .then(handleResponse)
    .catch(error => {
      console.log(error);
    });
}

function signOut() {

  /**
   * You can pass a custom request object below. This will override the initial configuration. For more information, visit:
   * https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-browser/docs/request-response-object.md#request
   */

  const logoutRequest = {
    postLogoutRedirectUri: msalConfig.auth.redirectUri,
    mainWindowRedirectUri: msalConfig.auth.redirectUri
  };

  myMSALObj.logoutPopup(logoutRequest);
}

function getTokenPopup(request) {

  /**
  * See here for more information on account retrieval: 
  * https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-common/docs/Accounts.md
  */
  request.account = myMSALObj.getAccountByHomeId(accountId);
  request.forceRefresh = true; // just for testing purposes

  /**
   * 
   */
  return myMSALObj.acquireTokenSilent(request)
    .then((response) => {
      // In case the response from B2C server has an empty accessToken field
      // throw an error to initiate token acquisition
      if (!response.accessToken || response.accessToken === "") {
        throw new msal.InteractionRequiredAuthError;
      }
      return response;
    })
    .catch(error => {
      console.log("Silent token acquisition fails. Acquiring token using popup. \n", error);
      if (error instanceof msal.InteractionRequiredAuthError) {
        // fallback to interaction when silent call fails
        return myMSALObj.acquireTokenPopup(request)
          .then(response => {
            console.log(response);
            return response;
          }).catch(error => {
            console.log(error);
          });
      } else {
        console.log(error);
      }
    });
}



function passTokenToCteApi() {
  getTokenPopup({
    scopes: ["api://1875691f-131f-4802-95a5-4511bde1408e/CTE.Exchange"]
  })
    .then(response => {
      if (response) {
        console.log("access_token acquired at: " + new Date().toString());
        try {
          let apiAccessToken = response.accessToken;
          callCte(apiAccessToken);
        } catch (error) {
          console.log(error);
        }
      }
    })
    .catch(function (error) {
      console.log(error);
    });
}

function callCte(apiAccessToken) {


  const manageCallsTokenRequest = { scopes: ["https://auth.msft.communication.azure.com/Teams.ManageCalls"] };

  manageCallsTokenRequest.account = myMSALObj.getAccountByHomeId(accountId);

  myMSALObj.acquireTokenSilent(manageCallsTokenRequest).then(function (accessTokenResponse) {
    // Acquire token silent success
    let teamsUserAccessToken = accessTokenResponse.accessToken;
    // Call your API with token
    callExchange(apiAccessToken, teamsUserAccessToken);
  }).catch(function (error) {
    //Acquire token silent failure, and send an interactive request
    if (error instanceof msal.InteractionRequiredAuthError) {
      myMSALObj.acquireTokenPopup(manageCallsTokenRequest).then(function (accessTokenResponse) {
        // Acquire token interactive success
        let teamsUserAccessToken = accessTokenResponse.accessToken;
        // Call your API with token
        callExchange(apiAccessToken, teamsUserAccessToken);
      }).catch(function (error) {
        // Acquire token interactive failure
        console.log(error);
      });
    }
    console.log(error);
  });

}

function callExchange(apiAccessToken, teamsUserAccessToken) {

  const headers = new Headers();
  const bearer = `Bearer ${apiAccessToken}`;

  headers.append("Authorization", bearer);
  headers.append("Content-Type", "application/json");

  fetch("/exchange", {
    method: "POST",
    headers: headers,
    body: JSON.stringify({ accessToken: teamsUserAccessToken })
  })
    .then(response => response.json())
    .then(response => {
      if (response) {
        logMessage('Token: ' + JSON.stringify(response));
      }
    })
    .catch(error => {
      console.log(error);
    });
}
