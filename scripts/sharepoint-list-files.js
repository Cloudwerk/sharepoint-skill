#!/usr/bin/env node
/**
 * SharePoint List Files
 * Lists files in a SharePoint folder via Microsoft Graph API
 * 
 * Usage:
 *   node sharepoint-list-files.js <site-name> <folder-path>
 *   node sharepoint-list-files.js TeamSite "Shared Documents/General"
 *   node sharepoint-list-files.js TeamSite ""  # Root folder
 * 
 * Output: JSON array of files
 */

const https = require('https');
const { execSync } = require('child_process');
const path = require('path');

const args = process.argv.slice(2);

if (args.length < 1) {
  console.error('Usage: node sharepoint-list-files.js <site-name> [folder-path]');
  console.error('');
  console.error('Examples:');
  console.error('  node sharepoint-list-files.js TeamSite');
  console.error('  node sharepoint-list-files.js TeamSite "Shared Documents/General"');
  process.exit(1);
}

const siteName = args[0];
const folderPath = args[1] || '';

// Get access token
const authScript = path.join(__dirname, 'sharepoint-auth.js');
let token;

try {
  token = execSync(`node ${authScript}`, { encoding: 'utf8' }).trim();
} catch (err) {
  console.error('ERROR: Failed to get access token');
  console.error(err.message);
  process.exit(1);
}

// Get tenant
const getTenantScript = path.join(__dirname, 'get-tenant.js');
let tenant;

try {
  tenant = execSync(`node ${getTenantScript}`, { encoding: 'utf8' }).trim();
} catch (err) {
  console.error('ERROR: Failed to get tenant');
  console.error(err.message);
  process.exit(1);
}

// Graph API request
const hostname = 'graph.microsoft.com';
const siteId = `${tenant}.sharepoint.com:/sites/${siteName}`;

// First, get the site ID
getSiteId(siteId, token, (actualSiteId) => {
  // Then get drive ID
  getDriveId(actualSiteId, token, (driveId) => {
    // Finally, list files
    listFiles(driveId, folderPath, token);
  });
});

function getSiteId(site, token, callback) {
  const apiPath = `/v1.0/sites/${site}`;
  
  makeRequest(hostname, apiPath, token, (data) => {
    if (data.id) {
      callback(data.id);
    } else {
      console.error('ERROR: Site not found');
      console.error(JSON.stringify(data, null, 2));
      process.exit(1);
    }
  });
}

function getDriveId(siteId, token, callback) {
  const apiPath = `/v1.0/sites/${siteId}/drive`;
  
  makeRequest(hostname, apiPath, token, (data) => {
    if (data.id) {
      callback(data.id);
    } else {
      console.error('ERROR: Drive not found');
      console.error(JSON.stringify(data, null, 2));
      process.exit(1);
    }
  });
}

function listFiles(driveId, folderPath, token) {
  let apiPath;
  
  if (folderPath) {
    // List specific folder
    const encodedPath = encodeURIComponent(folderPath);
    apiPath = `/v1.0/drives/${driveId}/root:/${encodedPath}:/children`;
  } else {
    // List root
    apiPath = `/v1.0/drives/${driveId}/root/children`;
  }
  
  makeRequest(hostname, apiPath, token, (data) => {
    if (data.value) {
      const files = data.value.map(item => ({
        name: item.name,
        size: item.size,
        webUrl: item.webUrl,
        downloadUrl: item['@microsoft.graph.downloadUrl'],
        lastModified: item.lastModifiedDateTime,
        isFolder: !!item.folder,
        type: item.folder ? 'folder' : (item.file?.mimeType || 'file')
      }));
      
      console.log(JSON.stringify(files, null, 2));
    } else {
      console.error('ERROR: Unexpected response');
      console.error(JSON.stringify(data, null, 2));
      process.exit(1);
    }
  });
}

function makeRequest(hostname, path, token, callback) {
  const options = {
    hostname,
    port: 443,
    path,
    method: 'GET',
    headers: {
      'Authorization': `Bearer ${token}`,
      'Accept': 'application/json'
    }
  };

  const req = https.request(options, (res) => {
    let data = '';

    res.on('data', (chunk) => {
      data += chunk;
    });

    res.on('end', () => {
      try {
        const parsed = JSON.parse(data);
        callback(parsed);
      } catch (err) {
        console.error('ERROR: Invalid JSON response');
        console.error(data);
        process.exit(1);
      }
    });
  });

  req.on('error', (e) => {
    console.error(`ERROR: ${e.message}`);
    process.exit(1);
  });

  req.end();
}
