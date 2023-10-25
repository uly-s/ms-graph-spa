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
    headers: headers,
  };

  console.log("request made to Graph API at: " + new Date().toString());
  console.log(endpoint);

  fetch(endpoint, options)
    .then((response) => response.json())
    .then((response) => callback(response, endpoint))
    .catch((error) => console.log(error));
}

var contacts = [];

async function callMSGraphPaginated(endpoint, token, callback) {
  const headers = new Headers();
  const bearer = `Bearer ${token}`;

  headers.append("Authorization", bearer);

  let nextLink = endpoint;
  let allResults = [];

  while (nextLink) {
    console.log("request made to Graph API at: " + new Date().toString());
    console.log(nextLink);

    const options = {
      method: "GET",
      headers: headers,
    };

    const response = await fetch(nextLink, options);
    const data = await response.json();
    allResults = allResults.concat(data.value);

    nextLink = data["@odata.nextLink"];
  }
  
  callback(allResults, endpoint);
  contacts = allResults;
  extendUI();
}

function deleteAllContacts(_token) {
    let token = _token;

  contacts.forEach((contact) => {
    const interval = setInterval(() => {
        let headers = new Headers();
        let token = getToken();
        const bearer = `Bearer ${token}`;
        headers.append("Authorization", bearer);
        const endpoint = `https://graph.microsoft.com/v1.0/me/contacts/${contact.id}`;
        const options = {
          method: "DELETE",
          headers: headers,
        };
    
        console.log("Delete request made for: " + contact.displayName + " at: " + new Date().toString());
    
        fetch(endpoint, options)
          .then((response) => response.json())
          .then((response) => console.log(response))
          .catch((error) => token = getToken());
    }, 500);

  });
}

function deleteContacts(_endpoint, token, callback) {
  const headers = new Headers();
  const bearer = `Bearer ${token}`;
  const endpoint = "https://graph.microsoft.com/v1.0/me/contacts/{id}";
  headers.append("Authorization", bearer);

  const options = {
    method: "DELETE",
    headers: headers,
  };

  console.log("request made to Graph API at: " + new Date().toString());
  console.log(endpoint);

  fetch(endpoint, options)
    .then((response) => response.json())
    .then((response) => callback(response, endpoint))
    .catch((error) => console.log(error));
}

function extendUI(){
// Create a Blob object from the JSON data
  const blob = new Blob([JSON.stringify(contacts, null, 2)], {
    type: "application/json",
  });
  // Create a URL for the Blob object
  const url = URL.createObjectURL(blob);

  // Create a link element for the download
  const link = document.createElement("a");
  link.href = url;
  link.download = "data.json";
  link.innerText = "Download Backup";
  link.classList = "btn btn-primary";
  link.addEventListener("click", () => {    
    const deleteButton = document.createElement("button");
    deleteButton.innerText = "Delete Contacts";
    deleteButton.classList = "btn btn-primary";
    deleteButton.addEventListener("click", () => {
        deleteAllContacts(getToken());
    });
    const div = document.getElementById("delete");
    div.appendChild(deleteButton);
  });
  let div = document.getElementById("downloads");
  div.appendChild(link);
}
