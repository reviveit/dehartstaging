const express = require('express');
const path = require('path');
const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
require('dotenv').config();

const app = express();

// Serve static files from the public directory
app.use(express.static(path.join(__dirname, 'public')));

const credential = new ClientSecretCredential(
  process.env.TENANT_ID,
  process.env.CLIENT_ID,
  process.env.CLIENT_SECRET
);

const client = Client.initWithMiddleware({
  authProvider: {
    getAccessToken: async () => (await credential.getToken()).token
  }
});

async function getDocumentContent(sharingUrl) {
  try {
    const encodedUrl = encodeURIComponent(sharingUrl);
    const itemResponse = await client.api(`/shares/${encodedUrl}/driveItem`).get();
    const contentResponse = await client.api(`/drives/${itemResponse.parentReference.driveId}/items/${itemResponse.id}/content`).get();
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
  console.log('Sharing URL:', sharingUrl);
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
  try {
    const sharingUrl = 'https://dehartmhk.sharepoint.com/:w:/s/Team/ER_lRUDzbgZOoWg_uyrpL0oBqdKLKGl_eZNN-3yCPOwKRQ?e=qBLtp7';
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

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});