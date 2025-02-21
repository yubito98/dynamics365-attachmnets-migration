const axios = require('axios');
const fs = require('fs');
const path = require('path');
require('dotenv').config()


async function getAuthToken() {
    const tokenEndpoint = `https://login.microsoftonline.com/${process.env.TENANT_ID}/oauth2/token`;
  
    const params = new URLSearchParams({
      grant_type: "client_credentials",
      client_id: process.env.CLIENT_ID,
      client_secret: process.env.CLIENT_SECRET,
      resource: process.env.RESOURCE,
    });
  
    try {
      const response = await axios.post(tokenEndpoint, params, {
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
      });
  
      console.log("Access Token:", response.data.access_token);
      return response.data.access_token;
    } catch (error) {
      console.error("Error fetching token:", error.response?.data || error.message);
    }
  }

  async function getAttachments() {
    const token = await getAuthToken();
  
    // Dynamics Web API endpoint for Annotations (notes) with attachments.
    // This URL queries the annotations entity for records that have a non-null documentbody.
    const apiUrl = 'https://credentialcheck.crm.dynamics.com/api/data/v9.1/annotations?' +
                   '$select=annotationid,filename,mimetype,documentbody,objectid_account&' +
                   '$filter=documentbody ne null';
  
    try {
      const response = await axios.get(apiUrl, {
        headers: {
          Authorization: `Bearer ${token}`,
          Accept: 'application/json',
        },
      });
  
      const annotations = response.data.value;
      console.log(`Found ${annotations.length} attachments`);
  
      // Ensure the ./files directory exists
      const filesDir = path.join(__dirname, 'files');
      if (!fs.existsSync(filesDir)) {
        fs.mkdirSync(filesDir, { recursive: true });
      }
  
      // Loop through each annotation (file attachment)
      for (const note of annotations) {
        // Destructure needed fields. 'objectid' represents the Dynamics record the note is linked to.
        const { annotationid, filename, mimetype, documentbody, objectid_account } = note;
  
        // Decode the base64 encoded file content provided in the Dynamics annotation
        const fileBuffer = Buffer.from(documentbody, 'base64');
  
        // Define the local file path and write the file
        const filePath = path.join(filesDir, filename);
        fs.writeFileSync(filePath, fileBuffer);
        console.log(`Downloaded file: ${filename}`);
  
        // Create metadata JSON for the file
        const metadata = {
          dynamicsRecordId: objectid_account,  // The Dynamics record ID associated with the attachment
          fileName: filename,
          fileType: mimetype,
        };
  
        // Save the metadata JSON file (e.g., "document.pdf.json")
        const jsonFilePath = path.join(filesDir, `${filename}.json`);
        fs.writeFileSync(jsonFilePath, JSON.stringify(metadata, null, 2));
        console.log(`Created metadata JSON for: ${filename}`);
      }
    } catch (error) {
      console.error('Error fetching attachments:', error.response?.data || error.message);
    }
  }
  
  getAttachments();