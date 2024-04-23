// Import necessary libraries
const express = require('express');
const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
const mammoth = require('mammoth');


// Initialize express app
const app = express();

// Azure AD credentials
const credential = new ClientSecretCredential(
  process.env.TENANT_ID,
  process.env.CLIENT_ID,
  process.env.CLIENT_SECRET
);

// Microsoft Graph client setup
const client = Client.init({
  authProvider: (done) => {
    credential.getToken().then((token) => {
      done(null, token.token); // Provide the token to Microsoft Graph API
    }).catch((err) => {
      done(err, null); // Handle error in getting token
    });
  }
});

// Endpoint to get text from a Word document
app.get('/document-text/:docId', async (req, res) => {
    const docPath = `path_to_your_documents/${req.params.docId}.docx`;

    try {
        const result = await mammoth.extractRawText({path: docPath});
        const text = result.value; // The raw text
        res.send(text);
    } catch (error) {
        console.error('Error reading document:', error);
        res.status(500).send('Failed to extract text');
    }
});

document.addEventListener('DOMContentLoaded', function() {
    fetch('/document-text/document-id-here')
        .then(response => response.text())
        .then(text => {
            document.getElementById('documentContent').textContent = text;
        })
        .catch(error => console.error('Error loading the document:', error));
});


// Start server
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
