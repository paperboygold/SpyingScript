const axios = require('axios');
const jwt = require('jsonwebtoken');
const fs = require('fs');

/**
 * Request an access token using client credentials flow with a JWT bearer assertion.
 * 
 * @param {Function} callback - The callback that handles the response.
 */
function requestToken(callback) {
    // Configuration: Replace these with your specific Azure AD details.
    const clientId = 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx';
    const tenantId = 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx';
    const certificatePrivateKey = fs.readFileSync('./keys/key.pem'); // Path to your private key file.
    const tokenUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
    
    // Thumbprint: Convert hex to base64 after uploading the key.pem as a secret in Azure.
    const thumbprintHex = "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX";

    const thumbprintBase64 = Buffer.from(thumbprintHex, 'hex').toString('base64');

    // JWT payload creation
    const payload = {
        aud: tokenUrl,
        iss: clientId,
        sub: clientId,
        jti: Math.random().toString(),
        nbf: Math.floor(Date.now() / 1000),
        exp: Math.floor(Date.now() / 1000) + (60 * 60)  // Token expiration set to 1 hour
    };

    // JWT signing
    const clientAssertion = jwt.sign(payload, certificatePrivateKey, {
        algorithm: 'RS256',
        header: {
            x5t: thumbprintBase64
        }
    });

    // Token request
    axios({
        method: 'post',
        url: tokenUrl,
        headers: {
            'Content-Type': 'application/x-www-form-urlencoded'
        },
        data: {
            grant_type: 'client_credentials',
            client_id: clientId,
            client_assertion_type: 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer',
            client_assertion: clientAssertion,
            scope: 'https://graph.microsoft.com/.default'
        },
        transformRequest: (data) => {
            return Object.entries(data).map(([key, value]) => `${encodeURIComponent(key)}=${encodeURIComponent(value)}`).join('&');
        }
    })
    .then(function(response) {
        callback(response.data.access_token); 
    })
    .catch(function(error) {
        console.log(error);
    });
}

module.exports = requestToken;