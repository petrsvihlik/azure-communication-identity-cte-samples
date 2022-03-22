const express = require('express');
const morgan = require('morgan');
const path = require('path');
const dotenv = require('dotenv');
const { CommunicationIdentityClient } = require('@azure/communication-identity');
const { PublicClientApplication, CryptoProvider } = require('@azure/msal-node');
const jwt_decode = require('jwt-decode');

dotenv.config();

const HOSTNAME = process.env.HOST || 'localhost';
const PORT = process.env.PORT || 80;
const HOST_URI = `http://${HOSTNAME}:${PORT}`;
const COMMUNICATION_SERVICES_CONNECTION_STRING = process.env.COMMUNICATION_SERVICES_CONNECTION_STRING;

// initialize express.
const app = express();
app.use(express.json());
app.use(express.urlencoded());

// Initialize variables.
let port = PORT;

// Configure morgan module to log all requests.
app.use(morgan('dev'));

// Setup app folders.
app.use(express.static('App'));

/*
const msalConfig = {
    auth: {
        clientId: process.env.AAD_CLIENT_ID,
        authority: process.env.AAD_AUTHORITY,
    }
};
const pca = new PublicClientApplication(msalConfig);
const provider = new CryptoProvider();
let pkceVerifier = "";
//TODO
app.get('/cte',
    async (req, res) => {


        const { verifier, challenge } = await provider.generatePkceCodes();
        pkceVerifier = verifier;
        // Get the auth code
        pca.getAuthCodeUrl({
            scopes: ["https://auth.msft.communication.azure.com/Teams.ManageCalls"],
            redirectUri: `${HOST_URI}/redirect`,
            codeChallenge: challenge,
            codeChallengeMethod: "S256"
        }).then((response) => {
            res.redirect(response);
        }).catch((error) => {
            console.log(JSON.stringify(error));
        });
    });


app.get('/redirect', async (req, res) => {
    // Acquire a token with the Teams.ManageCalls permission 
    pca.acquireTokenByCode({
        code: req.query.code,
        scopes: ["https://auth.msft.communication.azure.com/Teams.ManageCalls"],
        redirectUri: `${HOST_URI}/redirect`,
        codeVerifier: pkceVerifier,
    }).then(async (response) => {
        res.status(200).send(response.accessToken);
    }).catch((error) => {
        console.log(error);
        res.status(500).send(error);
    });
});
*/
app.post('/exchange', /* !! SOME AUTHORIZATION HERE!!*/ async (req, res, next) => {

    try {
        // Get Azure AD App client id
        const appId = process.env.AAD_CLIENT_ID;
        // Get user's oid
        const userId = jwt_decode(req.headers.authorization).oid;

        const identityClient = new CommunicationIdentityClient(COMMUNICATION_SERVICES_CONNECTION_STRING);

        // Pass the Client ID and oid
        let communicationIdentityToken = await identityClient.getTokenForTeamsUser(req.body.accessToken, appId, userId);

        res.status(200).send(communicationIdentityToken);
    }
    catch (err) {
        next(err);
    }
});


// Set up a route for index.html
app.get('*', (req, res) => {
    res.sendFile(path.join(__dirname + '/index.html'));
});

// Start the server.
app.listen(port);
console.log(`Listening on port ${port}...`);