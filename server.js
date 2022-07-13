const express = require('express');
const morgan = require('morgan');
const path = require('path');
const dotenv = require('dotenv');
const { CommunicationIdentityClient } = require('@azure/communication-identity');
const { expressjwt: jwt } = require("express-jwt");
const jwksClient = require('jwks-rsa');


// Initialize variables
dotenv.config();
const PORT = process.env.PORT || 3000;
const COMMUNICATION_SERVICES_CONNECTION_STRING = process.env.COMMUNICATION_SERVICES_CONNECTION_STRING;

// Initialize express
const app = express();
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Configure morgan module to log all requests
app.use(morgan('dev'));

// Setup app folders
app.use(express.static('App'));

const checkJwt = jwt({
    secret: jwksClient.expressJwtSecret({
        cache: true,
        rateLimit: true,
        jwksRequestsPerMinute: 5,
        jwksUri: `https://login.microsoftonline.com/${process.env.AAD_TENANT_ID}/discovery/keys?appid=${process.env.AAD_CLIENT_ID}`
    }),
    requestProperty: 'user',
    algorithms: ['RS256'],
});

app.post('/exchange',
    checkJwt,
    async (req, res, next) => {
        try {
            // Get Azure AD App client id
            const appId = process.env.AAD_CLIENT_ID;

            // Get user's oid
            const userId = req.user.oid;

            // The Teams user token to be exchanged for a Communication token
            const teamsUserAadToken = req.body.accessToken;

            // Create a new CommunicationIdentityClient
            const identityClient = new CommunicationIdentityClient(COMMUNICATION_SERVICES_CONNECTION_STRING);

            // Pass the Teams Azure AD token, Azure App's Client ID and Teams user's oid
            const communicationIdentityToken = await identityClient.getTokenForTeamsUser(teamsUserAadToken, appId, userId);

            res.status(200).send(communicationIdentityToken);
        }
        catch (err) {
            next(err);
        }
    });

app.get('/spa', function (req, res) {
    // A dedicated Redirect URI path meet the URI restrictions and to prevent the identity platform from choosing an arbitrary URI
    // More about the Redirect/Reply URI restrictions https://docs.microsoft.com/azure/active-directory/develop/reply-url
    res.redirect('/');
});

// Set up a route for index.html
app.get('*', (req, res) => {
    res.sendFile(path.join(__dirname + '/index.html'));
});

// Start the server
app.listen(PORT);
console.log(`Listening on port ${PORT}...`);