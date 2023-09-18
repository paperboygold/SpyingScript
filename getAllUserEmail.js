const axios = require('axios');
const requestToken = require('./token.js');

function getAllUserMail(userId, subjectFilter) {
    // get access token
    requestToken(async function(accessToken) {
        // make the API request
        try {
            const result = await axios({
                method: 'get',
                url: `https://graph.microsoft.com/v1.0/users/${userId}/mailFolders/Inbox/messages?$top=10`,
                headers: {
                    Authorization: `Bearer ${accessToken}`
                }
            });

            // print the messages
            console.log(result.data);

            // count the number of filtered messages and print it
            console.log(`The number of messages is: ${result.data.value.length}`);
        } catch (error) {
            console.log(error);
        }
    });
}

getAllUserMail('lucas.townsend@movember.com'); // replace with user's email and sender's email