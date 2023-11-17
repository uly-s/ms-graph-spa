"use strict";
import express from 'express';
import morgan from 'morgan';
import { join } from 'path';
import fetch from 'node-fetch';

const DEFAULT_PORT = process.env.PORT || 3000;

// initialize express.
const app = express();

// Initialize variables.
let port = DEFAULT_PORT;

// Configure morgan module to log all requests.
app.use(morgan('dev'));

// Setup app folders.
app.use(express.static('app'));

// Set up a route for index.html
app.get('*', (req, res) => {
    res.sendFile(join(__dirname + '/index.html'));
});


// Setup app folders.
app.use(express.static('app'));

// Set up a route for index.html
app.get('*', (req, res) => {
    res.sendFile(join(__dirname + '/index.html'));
});

import { ConfidentialClientApplication } from '@azure/msal-node';
import multer from 'multer';

const config = {
    auth: {
        clientId: "6932e19a-340d-4abe-a5a6-2a1d841ebae2",
        authority: 'https://login.microsoftonline.com/4ffb5de3-16b3-4329-8a70-58b1807ab447',
        clientSecret: 'Ttj8Q~OdZAJlErZe0fg_Z~b_ny7P2_dvYlERPbTi'
    }
};

const cca = new ConfidentialClientApplication(config);

// Set up a route to initiate backend code against the msgraph api.
app.post('/api/msgraph', async (req, res) => {
    const graphEndpoint = 'https://graph.microsoft.com/v1.0/me';

    try {
        const authResult = await cca.acquireTokenByClientCredential({
            scopes: ['https://graph.microsoft.com/.default']
        });

        for (const contact of toremove) {
            const endpoint = `https://graph.microsoft.com/v1.0/me/contacts/${contact.id}`;
            const response = await fetch(endpoint, {
                method: 'DELETE',
                headers: {
                    'Authorization': `Bearer ${authResult.accessToken}`
                }
            });
            if (response.status === 204) {
                console.log(`Contact ${contact.id} deleted successfully`);
            } else {
                console.log(`Error deleting contact ${contact.id}: ${response}`);
            }
        }

        const response = await fetch(graphEndpoint, {
            headers: {
                'Authorization': `Bearer ${authResult.accessToken}`
            }
        });

        const data = await response.json();
        res.send(data);
    } catch (error) {
        console.error(error);
        res.status(500).send(error);
    }
});


var contacts = [];

var deduped = [];

var toremove = [];



// Set up multer storage engine
const storage = multer.memoryStorage();
const upload = multer({ storage: storage });

// Set up a route to handle file uploads
app.post('/api/upload', upload.single('file'), (req, res) => {
    try {
        const fileContents = req.file.buffer.toString();
        const jsonData = JSON.parse(fileContents);
        contacts = jsonData;
        dedupe();
        console.log(contacts.length);
        console.log(deduped.length);
        console.log(toremove.length);
        res.status(200).send('File uploaded successfully');
        
    } catch (error) {
        console.error(error);
        res.status(500).send(error);
    }
});

function dedupe() {
    const seen = new Set();
    deduped = contacts.filter(contact => {
        const displayName = `${contact.givenName} ${contact.surname}`;
        const email = contact.emailAddresses && contact.emailAddresses.length > 0 && contact.emailAddresses[0].address;
        if (!email) {
            return true;
        }
        const key = `${displayName}:${email}`;
        if (seen.has(key)) {
            toremove.push(contact);
            return false;
        } else {
            seen.add(key);
            return true;
        }
    });
}







// Start the server.
app.listen(port);
console.log(`Listening on port ${port}...`);



