const axios = require('axios');
const fs = require('fs');
const path = require('path');
require('dotenv').config();

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
  if (!token) {
    console.error("No token available, exiting.");
    return;
  }

  // Dynamics Web API endpoint for Annotations (notes) with attachments.
  const apiUrl = 'https://credentialcheck.crm.dynamics.com/api/data/v9.1/annotations?' +
                 '$select=annotationid,filename,mimetype,documentbody&' +
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

    // Ensure the ./files directory exists for storing downloaded files
    const filesDir = path.join(__dirname, 'annotation-attachments');
    if (!fs.existsSync(filesDir)) {
      fs.mkdirSync(filesDir, { recursive: true });
    }

    // Array to hold aggregated metadata for each file
    const aggregatedMetadata = [];

    // Loop through each annotation (file attachment)
    for (const note of annotations) {
      // Destructure needed fields; removed objectid_account
      const { annotationid, filename, mimetype, documentbody } = note;

      // Decode the base64 encoded file content provided in the Dynamics annotation
      const fileBuffer = Buffer.from(documentbody, 'base64');

      // Define the local file path and write the file in the files folder
      const filePath = path.join(filesDir, filename);
      fs.writeFileSync(filePath, fileBuffer);
      console.log(`Downloaded file: ${filename}`);

      // Add the file metadata to the aggregated array
      aggregatedMetadata.push({
        fileName: filename,
        fileType: mimetype,
      });
    }

    // Save the aggregated metadata JSON file outside of the files folder
    const metadataJSONPath = path.join(__dirname, 'annotationAttachments.json');
    fs.writeFileSync(metadataJSONPath, JSON.stringify(aggregatedMetadata, null, 2));
    console.log(`Created aggregated metadata JSON: ${metadataJSONPath}`);

    // Create CSV content from the aggregated metadata
    const csvHeader = "FileName,FileType\n";
    const csvRows = aggregatedMetadata.map(item =>
      `${item.fileName},${item.fileType}`
    );
    const csvContent = csvHeader + csvRows.join("\n");

    // Save the CSV file outside of the files folder
    const metadataCSVPath = path.join(__dirname, 'annotationAttachments.csv');
    fs.writeFileSync(metadataCSVPath, csvContent);
    console.log(`Created metadata CSV file: ${metadataCSVPath}`);

  } catch (error) {
    console.error('Error fetching attachments:', error.response?.data || error.message);
  }
}

getAttachments();
