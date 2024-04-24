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

async function getFormattedContent(driveId, itemId) {
  const content = await getDocumentContent(driveId, itemId);
  if (content) {
    return formatContent(content);
  }
  return '<p>Document content not available.</p>';
}

app.get('/display1', async (req, res) => {
  const sharingUrl = 'https://dehartmhk.sharepoint.com/:w:/s/Team/ER_lRUDzbgZOoWg_uyrpL0oBfeIPXJ8_zc9HjheAXDfjug?e=8Oa9rN';
  const formattedContent = await getFormattedContent(sharingUrl);
  res.send(`
    <html>
      <head>
        <link rel="stylesheet" href="/styles.css">
      </head>
      <body>
        <div id="content">${formattedContent}</div>
      </body>
    </html>
  `);
});

// ... Repeat for displays 2-8 ...

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});