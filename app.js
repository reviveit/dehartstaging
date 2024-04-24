const express = require('express');
const path = require('path');
const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
require('dotenv').config();

const app = express();
app.use(express.static(path.join(__dirname, 'public')));

// Convert sharing URL to a format suitable for Microsoft Graph API
function convertToGraphApiSharingUrl(sharingUrl) {
  const buffer = Buffer.from(sharingUrl, "utf8");
  const base64Url = buffer.toString('base64');
  return `u!${base64Url.replace(/=/g, '').replace(/\+/g, '-').replace(/\//g, '_')}`;
}

// Function to test the authentication setup
async function testAuthentication() {
  const credential = new ClientSecretCredential(
    process.env.TENANT_ID,
    process.env.CLIENT_ID,
    process.env.CLIENT_SECRET
  );

  console.log('Testing authentication...');
  try {
    const tokenResponse = await credential.getToken(['https://graph.microsoft.com/.default']);
    console.log('Authentication test successful. Access token obtained:', tokenResponse.token);
  } catch (error) {
    console.error('Authentication test failed:', error);
  }
}

// Function to retrieve document content from a sharing URL
async function getDocumentContent(sharingUrl) {
  const encodedUrl = convertToGraphApiSharingUrl(sharingUrl);
  console.log('Encoded URL for Graph API:', encodedUrl);
  
  const credential = new ClientSecretCredential(
    process.env.TENANT_ID,
    process.env.CLIENT_ID,
    process.env.CLIENT_SECRET
  );
  
  const client = Client.initWithMiddleware({
    authProvider: {
      getAccessToken: async () => {
        const tokenResponse = await credential.getToken(['https://graph.microsoft.com/.default']);
        return tokenResponse.token;
      }
    }
  });

  try {
    const itemResponse = await client.api(`/shares/${encodedUrl}/driveItem`).get();
    const contentResponse = await client.api(`/drives/${itemResponse.parentReference.driveId}/items/${itemResponse.id}/content`)
      .responseType('stream')  // Set responseType to 'stream' for handling binary data
      .get();

    // Convert the stream to text (assuming it's a text document)
    return new Promise((resolve, reject) => {
      let data = '';
      contentResponse.on('data', (chunk) => { data += chunk; });
      contentResponse.on('end', () => { resolve(data.toString()); });
      contentResponse.on('error', reject);
    });
  } catch (error) {
    console.error('Error retrieving document:', error);
    throw error;
  }
}

// Function to format the retrieved content (assumes content is text)
function formatContent(content) {
  return `<p>${content.replace(/\n/g, '<br>')}</p>`;
}

// Route to display formatted content from a specific URL
app.get('/display1', async (req, res) => {
  const sharingUrl = 'https://dehartmhk.sharepoint.com/:w:/s/Team/ER_lRUDzbgZOoWg_uyrpL0oBqdKLKGl_eZNN-3yCPOwKRQ?e=zKCS8A';
  try {
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

// Start the server and run a test of authentication
testAuthentication();
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
