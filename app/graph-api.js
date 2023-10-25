import { getTokenPopup } from './authPopup.js';
import { callMSGraph, callMSGraphPaginated } from './graph.js';

// ms-graph endpoint configuration
const graphConfig = {
    graphMeEndpoint: "https://graph.microsoft.com/v1.0/me",
    graphContactsEndpoint: "https://graph.microsoft.com/v1.0/me/contacts",
    graphContactFoldersEndpoint: "https://graph.microsoft.com/v1.0/me/contactFolders",
};

function readAllContacts() {
    getTokenPopup(tokenRequest)
        .then(response => {
            callMSGraphPaginated(graphConfig.graphContactsEndpoint, response.accessToken, updateUI);
        }).catch(error => {
            console.error(error);
        });
}

