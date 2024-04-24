const express = require('express');
const path = require('path');
const axios = require('axios');
const mammoth = require('mammoth');
const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
require('dotenv').config();

const app = express();
app.use(express.static(path.join(__dirname, 'public')));

function convertToGraphApiSharingUrl(sharingUrl) {
  const buffer = Buffer.from(sharingUrl, "utf8");
  const base64Url = buffer.toString('base64');
  return `u!${base64Url.replace(/=/g, '').replace(/\+/g, '-').replace(/\//g, '_')}`;
}

async function testAuthentication() {
  const credential = new ClientSecretCredential(
    process.env.TENANT_ID,
    process.env.CLIENT_ID,
    process.env.CLIENT_SECRET
  );
  try {
    const tokenResponse = await credential.getToken(['https://graph.microsoft.com/.default']);
    console.log('Access token obtained:', tokenResponse.token);
  } catch (error) {
    console.error('Error obtaining access token:', error);
  }
}

async function getDocumentContent(sharingUrl) {
  const encodedUrl = convertToGraphApiSharingUrl(sharingUrl);
  const credential = new ClientSecretCredential(
    process.env.TENANT_ID,
    process.env.CLIENT_ID,
    process.env.CLIENT_SECRET
  );
  try {
    const tokenResponse = await credential.getToken('https://graph.microsoft.com/.default');
    const accessToken = tokenResponse.token;
    const itemResponse = await axios.get(`https://graph.microsoft.com/v1.0/shares/${encodedUrl}/driveItem`, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
      },
    });
    const itemData = itemResponse.data;
    const contentResponse = await axios.get(`https://graph.microsoft.com/v1.0/drives/${itemData.parentReference.driveId}/items/${itemData.id}/content`, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
      },
    });
    const content = contentResponse.data;
    return content;
  } catch (error) {
    console.error('Error retrieving document:', error);
    throw error;
  }
}

async function convertToHtml(content) {
  const result = await mammoth.convertToHtml({ buffer: content });
  return result.value;
}

async function getFormattedContent(sharingUrl) {
  try {
    const content = await getDocumentContent(sharingUrl);
    const html = await convertToHtml(content);
    return html;
  } catch (error) {
    console.error('Failed to retrieve or format content:', error);
    return '<p>Sorry, the content could not be retrieved at the moment. Please try again later.</p>';
  }
}

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
    res.status(500).send('Error retrieving document content. Please try again later.');
  }
});

testAuthentication();

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});