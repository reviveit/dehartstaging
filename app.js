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

    const contentResponse = await axios({
      url: `https://graph.microsoft.com/v1.0/drives/${itemData.parentReference.driveId}/items/${itemData.id}/content`,
      method: 'GET',
      responseType: 'arraybuffer',  
      headers: {
        Authorization: `Bearer ${accessToken}`,
      },
    });
    return Buffer.from(contentResponse.data);  
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
    if (error.message.includes('Corrupted zip')) {
      return '<p>Sorry, the document appears to be corrupted. Please try again with a different document.</p>';
    } else {
      return '<p>Sorry, an error occurred while retrieving the document content. Please try again later.</p>';
    }
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
        <div class="page-transition">
          <div class="header">
            <img src="/logo.png" alt="Logo" class="logo">
            <h1>Display 1</h1>
          </div>
          <div class="document-container">
            <div id="content">${formattedContent}</div>
            <button id="fullScreenBtn">Full Screen</button>
          </div>
          <div class="button-container">
            <button onclick="navigateToDisplay('/')">Back to Selection</button>
          </div>
        </div>
        <div class="loading-spinner"></div>
        <script>
          function navigateToDisplay(url) {
            const pageTransition = document.querySelector('.page-transition');
            const loadingSpinner = document.querySelector('.loading-spinner');
            pageTransition.classList.add('fade-out');
            loadingSpinner.style.display = 'block';
            setTimeout(() => {
              location.href = url;
            }, 500);
          }

          const fullScreenBtn = document.getElementById('fullScreenBtn');
          const content = document.getElementById('content');

          fullScreenBtn.addEventListener('click', () => {
            content.classList.add('full-screen');
            document.body.style.overflow = 'hidden';
          });

          content.addEventListener('click', () => {
            content.classList.remove('full-screen');
            document.body.style.overflow = 'auto';
          });
        </script>
      </body>
    </html>
  `);
  } catch (error) {
    console.error('Error in /display1 route:', error);
    res.status(500).send('Error retrieving document content. Please try again later.');
  }
});

app.get('/display2', async (req, res) => {
  const sharingUrl = 'https://dehartmhk.sharepoint.com/:w:/s/Team/EeAUJ4jWFFdKlWJUodu4Eh4BObOVrhI-ahp37WuWkbw2uw?e=9VVWZ6';
  try {
    const formattedContent = await getFormattedContent(sharingUrl);
    res.send(`
      <html>
        <head>
          <link rel="stylesheet" href="/styles.css">
        </head>
        <body>
          <div class="document-container">
            <div id="content">${formattedContent}</div>
          </div>
          <div class="button-container">
            <button onclick="location.href='/'">Back to Selection</button>
          </div>
        </body>
      </html>
    `);
  } catch (error) {
    console.error('Error in /display2 route:', error);
    res.status(500).send('Error retrieving document content. Please try again later.');
  }
});

app.get('/display3', async (req, res) => {
  const sharingUrl = 'https://dehartmhk.sharepoint.com/:w:/s/Team/EdKznNWC8-FBgmXlzBOuyTYBAzTail-2aB-MtxKrKFtGog?e=iaCFse';
  try {
    const formattedContent = await getFormattedContent(sharingUrl);
    res.send(`
      <html>
        <head>
          <link rel="stylesheet" href="/styles.css">
        </head>
        <body>
          <div class="document-container">
            <div id="content">${formattedContent}</div>
          </div>
          <div class="button-container">
            <button onclick="location.href='/'">Back to Selection</button>
          </div>
        </body>
      </html>
    `);
  } catch (error) {
    console.error('Error in /display3 route:', error);
    res.status(500).send('Error retrieving document content. Please try again later.');
  }
});

app.get('/display4', async (req, res) => {
  const sharingUrl = 'https://dehartmhk.sharepoint.com/:w:/s/Team/Ec00txZbL_tPoh5XJ8Qh3gMBKdcSqCH7qaoL7KhtstF0wA?e=kWtmnI';
  try {
    const formattedContent = await getFormattedContent(sharingUrl);
    res.send(`
      <html>
        <head>
          <link rel="stylesheet" href="/styles.css">
        </head>
        <body>
          <div class="document-container">
            <div id="content">${formattedContent}</div>
          </div>
          <div class="button-container">
            <button onclick="location.href='/'">Back to Selection</button>
          </div>
        </body>
      </html>
    `);
  } catch (error) {
    console.error('Error in /display4 route:', error);
    res.status(500).send('Error retrieving document content. Please try again later.');
  }
});

app.get('/display5', async (req, res) => {
  const sharingUrl = 'https://dehartmhk.sharepoint.com/:w:/s/Team/EYOEVTc6bN1DvM2ifudIc1IB_QZrKgR5CosXNJb9smYycw?e=QSjkW6';
  try {
    const formattedContent = await getFormattedContent(sharingUrl);
    res.send(`
      <html>
        <head>
          <link rel="stylesheet" href="/styles.css">
        </head>
        <body>
          <div class="document-container">
            <div id="content">${formattedContent}</div>
          </div>
          <div class="button-container">
            <button onclick="location.href='/'">Back to Selection</button>
          </div>
        </body>
      </html>
    `);
  } catch (error) {
    console.error('Error in /display5 route:', error);
    res.status(500).send('Error retrieving document content. Please try again later.');
  }
});

app.get('/display6', async (req, res) => {
  const sharingUrl = 'https://dehartmhk.sharepoint.com/:w:/s/Team/EZ1RmV7OFQdJvlw9_VGwyTQB8KIiQ7tpaIzZZ2iiNGYj7w?e=gePgwq';
  try {
    const formattedContent = await getFormattedContent(sharingUrl);
    res.send(`
      <html>
        <head>
          <link rel="stylesheet" href="/styles.css">
        </head>
        <body>
          <div class="document-container">
            <div id="content">${formattedContent}</div>
          </div>
          <div class="button-container">
            <button onclick="location.href='/'">Back to Selection</button>
          </div>
        </body>
      </html>
    `);
  } catch (error) {
    console.error('Error in /display6 route:', error);
    res.status(500).send('Error retrieving document content. Please try again later.');
  }
});

testAuthentication();

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});