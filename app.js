const express = require('express');
const path = require('path');
const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
require('dotenv').config();

const app = express();

// Serve static files from the public directory
app.use(express.static(path.join(__dirname, 'public')));

async function testAuthentication() {
  try {
    const credential = new ClientSecretCredential(
      process.env.TENANT_ID,
      process.env.CLIENT_ID,
      process.env.CLIENT_SECRET
    );

    console.log('Obtaining access token');
    const tokenResponse = await credential.getToken(['https://graph.microsoft.com/.default']);
    console.log('Access token obtained:', tokenResponse.token);
  } catch (error) {
    console.error('Error obtaining access token:', error);
  }
}

async function getDocumentContent(sharingUrl) {
  try {
    console.log('Inside getDocumentContent function');
    console.log('Sharing URL:', sharingUrl);
    console.log('Tenant ID:', process.env.TENANT_ID);
    console.log('Client ID:', process.env.CLIENT_ID);
    console.log('Client Secret:', process.env.CLIENT_SECRET);

    console.log('Creating ClientSecretCredential');
    const credential = new ClientSecretCredential(
      process.env.TENANT_ID,
      process.env.CLIENT_ID,
      process.env.CLIENT_SECRET
    );
    console.log('ClientSecretCredential created');
    console.log('Credential:', credential);

    console.log('Creating Microsoft Graph client');
    const client = Client.initWithMiddleware({
      authProvider: {
        getAccessToken: async () => {
          console.log('Obtaining access token');
          const tokenResponse = await credential.getToken('https://graph.microsoft.com/.default');
          console.log('Access token obtained');
          return tokenResponse.token;
        }
      }
    });
    console.log('Microsoft Graph client created');

    const encodedUrl = encodeURIComponent(sharingUrl);
    console.log('Encoded URL:', encodedUrl);

    console.log('Making API call to get item metadata');
    const itemResponse = await client.api(`/shares/${encodedUrl}/driveItem`).get();
    console.log('Item Response:', itemResponse);

    console.log('Making API call to get item content');
    const contentResponse = await client.api(`/drives/${itemResponse.parentReference.driveId}/items/${itemResponse.id}/content`).get();
    console.log('Content Response:', contentResponse);

    return contentResponse;
  } catch (error) {
    console.error('Error retrieving document:', error);
    if (error.statusCode === 401) {
      console.error('Unauthorized. Check your Azure AD application permissions.');
    } else if (error.statusCode === 404) {
      console.error('Document not found. Check the sharing URL.');
    } else {
      console.error('Unexpected error:', error.statusCode, error.message);
    }
    return null;
  }
}

function formatContent(content) {
  const formattedContent = content.replace(/\n\n/g, '</p><p>').replace(/\n/g, '<br>');
  return `<p>${formattedContent}</p>`;
}

async function getFormattedContent(sharingUrl) {
  console.log('Inside getFormattedContent function');
  const content = await getDocumentContent(sharingUrl);
  console.log('Retrieved content:', content);
  if (content) {
    return formatContent(content);
  } else {
    console.log('Document content is empty or null');
    return '<p>Document content not available.</p>';
  }
}

app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

app.get('/display1', async (req, res) => {
  console.log('Inside /display1 route handler');
  try {
    const sharingUrl = 'https://dehartmhk.sharepoint.com/:w:/s/Team/ER_lRUDzbgZOoWg_uyrpL0oBqdKLKGl_eZNN-3yCPOwKRQ?e=zKCS8A';
    const formattedContent = await getFormattedContent(sharingUrl);
    res.send(`
      <html>
        <head>
          <link rel="stylesheet" href="/styles.css">
        </head>
        <body>
          <div id="content">${formattedContent}</div>
          <button onclick="location.href='/'">Back to Selection</button>
        </body>
      </html>
    `);
  } catch (error) {
    console.error('Error in /display1 route:', error);
    res.status(500).send(`
      <html>
        <head>
          <link rel="stylesheet" href="/styles.css">
        </head>
        <body>
          <div id="content">Error retrieving document content. Please try again later.</div>
        </body>
      </html>
    `);
  }
});

// ... Repeat for displays 2-6 ...

// Call the test function
testAuthentication();

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});