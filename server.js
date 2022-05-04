const express = require('express');
const morgan = require('morgan');
const path = require('path');
const dotenv = require('dotenv');
const { CommunicationIdentityClient } = require('@azure/communication-identity');
var jwt = require("express-azure-jwt");
const jwtScope = require('express-jwt-scope');

// Initialize variables
dotenv.config();
const HOSTNAME = process.env.HOST || 'localhost';
const PORT = process.env.PORT || 80;
const HOST_URI = `http://${HOSTNAME}:${PORT}`;
const COMMUNICATION_SERVICES_CONNECTION_STRING = process.env.COMMUNICATION_SERVICES_CONNECTION_STRING;

// Initialize express
const app = express();
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Configure morgan module to log all requests
app.use(morgan('dev'));

// Setup app folders
app.use(express.static('App'));

app.post('/exchange',
    jwt({ aadIssuerUrlTemplate: 'https://login.microsoftonline.com/{tenantId}/v2.0' }), // Verify JWT integrity
    jwtScope('CTE.Exchange', { scopeKey: 'scp' }), // Verify required scopes
    async (req, res, next) => {

        try {
            // Get Azure AD App client id
            const appId = process.env.AAD_CLIENT_ID;

            // Get user's oid
            const userId = req.user.oid;

            // Create a new CommunicationIdentityClient
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

// Start the server
app.listen(PORT);
console.log(`Listening on port ${PORT}...`);