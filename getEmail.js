const axios = require('axios');
const requestToken = require('./token.js');

/**
 * Strips HTML tags and converts certain HTML entities from the input string.
 * 
 * @param {string} html - The input HTML string.
 * @returns {string} - Cleaned string.
 */
function stripHTML(html) {
    let text = html;
    text = text.replace(/<!--[\s\S]*?-->/g, '');
    text = text.replace(/<style[\s\S]*?>[\s\S]*?<\/style>/g, '');
    text = text.replace(/<br\s*\/?>/g, '\n');
    text = text.replace(/<\/p>/g, '\n');
    text = text.replace(/<\/div>/g, '\n');
    text = text.replace(/<\/h[1-6]>/g, '\n');
    text = text.replace(/<\/li>/g, '\n');
    text = text.replace(/&nbsp;/g, ' ');
    text = text.replace(/<[\s\S]*?>/g, '');
    return text.trim();
}

/**
 * Retrieves the specified folder for a given user.
 * 
 * @param {string} accessToken - OAuth2 access token.
 * @param {string} userId - The user's email or ID.
 * @param {string} folderName - Name of the folder to be retrieved.
 * @returns {Object} - Folder details.
 */
async function getFolder(accessToken, userId, folderName) {
    const result = await axios({
        method: 'get',
        url: `https://graph.microsoft.com/v1.0/users/${userId}/mailFolders`,
        headers: {
            Authorization: `Bearer ${accessToken}`,
        },
    });

    const folder = result.data.value.find(folder => folder.displayName === folderName);
    return folder;
}

/**
 * Retrieves emails from a specified folder based on subject and body filters.
 * 
 * @param {string} accessToken - OAuth2 access token.
 * @param {string} userId - The user's email or ID.
 * @param {Object} folder - Folder details.
 * @param {string} subjectFilter - Filter for email subjects.
 * @param {string} [bodyFilter] - Optional filter for email body.
 */
async function getEmailsFromFolder(accessToken, userId, folder, subjectFilter, bodyFilter = null) {
    let subjectFilterEncoded = encodeURIComponent(`contains(subject, '${subjectFilter}')`);
    let filter;
    
    if (bodyFilter) {
        let bodyFilterEncoded = encodeURIComponent(`contains(body/content, '${bodyFilter}')`);
        filter = `$filter=${subjectFilterEncoded} and ${bodyFilterEncoded}`;
    } else {
        filter = `$filter=${subjectFilterEncoded}`;
    }

    let nextLink = `https://graph.microsoft.com/v1.0/users/${userId}/mailFolders/${folder.id}/messages?$top=100&${filter}`;
    do {
        const result = await axios({
            method: 'get',
            url: nextLink,
            headers: {
                Authorization: `Bearer ${accessToken}`,
            },
        });

        const filteredMessages = result.data.value;

        filteredMessages.forEach((message) => {
            console.log('Subject:', message.subject);
            console.log('Received Date:', message.receivedDateTime);
            console.log('Message ID:', message.id);

            let formattedBody = stripHTML(message.body.content);
            if (formattedBody.startsWith(message.subject)) {
                formattedBody = formattedBody.slice(message.subject.length).trim();
            }

            console.log('Body:');
            console.log(formattedBody);
            console.log('------------------------------------------------------\n');
        });

        if (result.data["@odata.nextLink"]) {
            nextLink = result.data["@odata.nextLink"];
            let url = new URL(nextLink);
            let params = new URLSearchParams(url.search);
            if (!params.has("$filter")) {
                params.append("$filter", filter);
                url.search = params.toString();
                nextLink = url.toString();
            }
        } else {
            nextLink = null;
        }

    } while (nextLink);
}

/**
 * Iterates over a list of folder names and retrieves emails based on filters.
 * 
 * @param {string} accessToken - OAuth2 access token.
 * @param {string} userId - The user's email or ID.
 * @param {Array} folderNames - List of folder names.
 * @param {string} subjectFilter - Filter for email subjects.
 * @param {string} [bodyFilter] - Optional filter for email body.
 */
async function getEmailsFromFolders(accessToken, userId, folderNames, subjectFilter, bodyFilter = null) {
    for (const folderName of folderNames) {
        const folder = await getFolder(accessToken, userId, folderName);
        if (folder) {
            await getEmailsFromFolder(accessToken, userId, folder, subjectFilter, bodyFilter);
        }
    }
}

/**
 * Initiates the process to retrieve emails by subject and optional body filter.
 * 
 * @param {string} userId - The user's email or ID.
 * @param {string} subjectFilter - Filter for email subjects.
 * @param {string} [bodyFilter] - Optional filter for email body.
 */
function getEmail(userId, subjectFilter, bodyFilter = null) {
    let sanitizedSubjectFilter = subjectFilter.replace(/'/g, "''");
    let sanitizedBodyFilter = bodyFilter ? bodyFilter.replace(/'/g, "''") : null;

    requestToken(async function (accessToken) {
        await getEmailsFromFolders(accessToken, userId, ['Inbox', 'Deleted Items', 'Sent Items', 'Archive'], sanitizedSubjectFilter, sanitizedBodyFilter);
    });
}

// *****************************************************************************
// This is where you customize the input parameters:
// 1. The user's email or ID.
// 2. The subject filter to search for in emails.
// 3. [Optional] A filter for the email body.
// *****************************************************************************
getEmail('example@domain.com', 'RE: Example');