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
    const contentResponse = await client.api(`/drives/${itemResponse.parentReference.driveId}/items/${itemResponse.id}/content`).get();
    return contentResponse;
  } catch (error) {
    console.error('Error retrieving document:', error);
    throw error;
  }
}

async function getFormattedContent(sharingUrl) {
  try {
    const content = await getDocumentContent(sharingUrl);
    return formatContent(content);
  } catch (error) {
    console.error('Failed to retrieve or format content:', error);
    return '<p>Document content not available.</p>';
  }
}

function formatContent(content) {
  return `<p>${content.replace(/\n/g, '<br>')}</p>`;
}

app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

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

// Initialize testing and server
testAuthentication();
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
