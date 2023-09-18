const axios = require('axios');
const requestToken = require('./token.js');

async function getMailFolders(accessToken, userId) {
  try {
    const result = await axios({
      method: 'get',
      url: `https://graph.microsoft.com/v1.0/users/${userId}/mailFolders/`,
      headers: {
        Authorization: `Bearer ${accessToken}`,
      },
    });

    return result.data;
    
  } catch (error) {
    console.log(error);
  }
}

function fetchMailFolders(userId) {
  requestToken(async function (accessToken) {
    const folders = await getMailFolders(accessToken, userId);
    
    // display folders
    if (folders && folders.value) {
      folders.value.forEach((folder) => {
        console.log('Folder ID:', folder.id);
        console.log('Folder Name:', folder.displayName);
        console.log('------------------------------------------------------\n');
      });
    }
  });
}

fetchMailFolders('lucas.townsend@movember.com');
