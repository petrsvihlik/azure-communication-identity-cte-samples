
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