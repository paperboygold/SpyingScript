const axios = require('axios');
const requestToken = require('./token.js');

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

async function getEmailsFromFolders(accessToken, userId, folderNames, subjectFilter, bodyFilter = null) {
    for (const folderName of folderNames) {
        const folder = await getFolder(accessToken, userId, folderName);
        if (folder) {
            await getEmailsFromFolder(accessToken, userId, folder, subjectFilter, bodyFilter);
        }
    }
}

function getEmailBySubject(userId, subjectFilter, bodyFilter = null) {
    // Sanitize the subject and body filters to escape the single quotes
    let sanitizedSubjectFilter = subjectFilter.replace(/'/g, "''");
    let sanitizedBodyFilter = bodyFilter ? bodyFilter.replace(/'/g, "''") : null;

    requestToken(async function (accessToken) {
        await getEmailsFromFolders(accessToken, userId, ['Inbox', 'Deleted Items', 'Sent Items', 'Archive'], sanitizedSubjectFilter, sanitizedBodyFilter);
    });
}

getEmailBySubject('lucas.townsend@movember.com', 'Thanks for your interest in ChatGPT Enterprise');