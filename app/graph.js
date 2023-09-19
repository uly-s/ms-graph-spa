/** 
 * Helper function to call MS Graph API endpoint
 * using the authorization bearer token scheme
*/
function callMSGraph(endpoint, token, callback) {
    const headers = new Headers();
    const bearer = `Bearer ${token}`;

    headers.append("Authorization", bearer);

    const options = {
        method: "GET",
        headers: headers
    };

    console.log('request made to Graph API at: ' + new Date().toString());
    console.log(endpoint);
    const wtf = endpoint == graphConfig.graphMeEndpoint ? endpoint : "https://graph.microsoft.com/v1.0/me/contacts";
    fetch(wtf, options)
        .then(response => endpoint == graphConfig.graphMeEndpoint ? response.json() : response.json())
        .then(response => callback(response, wtf))
        .catch(error => console.log(error));
}

