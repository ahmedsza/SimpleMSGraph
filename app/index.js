// index.js
// index.js


const tableBody = document.querySelector('#users-table tbody');
//import * as MicrosoftGraph from '@microsoft/microsoft-graph-client';
// Rest of the code
//import { Client } from 'https://cdn.jsdelivr.net/npm/@microsoft/microsoft-graph-client/lib/graph-js-sdk.js'


const client = MicrosoftGraph.Client.init({
    authProvider: (done) => {
        // Set the access token for the request
        const accessToken = myauthresult;
        done(null, accessToken);
    }
});

const msalConfig = {
    auth: {
      clientId: 'REPLACE-WITH-YOUR-APP-ID',
      authority: 'https://login.microsoftonline.com/PUTYOURTENANTIDHERE'
    }
  };
  


  const msalApplication = new msal.PublicClientApplication(msalConfig);
  var myauthresult;
  const login = async () => {
    try {
      const authResult = await msalApplication.loginPopup({
        scopes: ['user.read','User.Read.All']
      });
      console.log('Access token:', authResult.accessToken);
      myauthresult = authResult.accessToken;
      getUsers(authResult.accessToken);
      console.log("agetr users");
    } catch (error) {
      console.log('Error:', error);
    }
  };

const getUsers = async (accessToken) => {
    console.log("in getusers start");
    let users = [];
    //client.accessToken = accessToken;
    //client.
    let usersPage = await client.api('/users').top(5).get();
    console.log("in getysers");
    while (usersPage) {
        users = users.concat(usersPage.value);
        if (usersPage['@odata.nextLink']) {
            console.log("in getusers nextlink");
            usersPage = await client.api(usersPage['@odata.nextLink']).get();
        } else {
            usersPage = null;
        }
    }

    // Process the retrieved data
    const tableBody = document.querySelector('#users-table tbody');
    for (const user of users) {
        console.log(user);
        const row = document.createElement('tr');
        const nameCell = document.createElement('td');
        nameCell.textContent = user.displayName;
        const emailCell = document.createElement('td');
        emailCell.textContent = user.mail;
        const jobTitleCell = document.createElement('td');
        jobTitleCell.textContent = user.jobTitle;
        row.appendChild(nameCell);
        row.appendChild(emailCell);
        row.appendChild(jobTitleCell);
        tableBody.appendChild(row);
    }
};
const loginButton = document.querySelector('#login-button');
loginButton.addEventListener('click', login);
