#!/usr/bin/env node
/**
 * SharePoint Upload File
 * Uploads a file to SharePoint via Microsoft Graph API
 * 
 * Usage:
 *   node sharepoint-upload.js <site-name> <local-file> <remote-path>
 *   node sharepoint-upload.js TeamSite ./document.pdf "Shared Documents/General/document.pdf"
 *   node sharepoint-upload.js TeamSite ./file.txt "file.txt"  # Upload to root
 * 
 * Output: Upload confirmation with file URL
 */

const https = require('https');
const fs = require('fs');
const { execSync } = require('child_process');
const path = require('path');

const args = process.argv.slice(2);

if (args.length < 3) {
  console.error('Usage: node sharepoint-upload.js <site-name> <local-file> <remote-path>');
  console.error('');
  console.error('Examples:');
  console.error('  node sharepoint-upload.js TeamSite ./document.pdf "document.pdf"');
  console.error('  node sharepoint-upload.js TeamSite ./file.txt "General/file.txt"');
  process.exit(1);
}

const siteName = args[0];
const localFile = args[1];
const remotePath = args[2];

if (!fs.existsSync(localFile)) {
  console.error(`ERROR: Local file not found: ${localFile}`);
  process.exit(1);
}

const fileSize = fs.statSync(localFile).size;
const fileName = path.basename(remotePath);

console.error(`Uploading: ${localFile} → ${remotePath}`);
console.error(`Size: ${fileSize} bytes`);

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

// Get site ID -> drive ID -> upload file
getSiteId(siteId, token, (actualSiteId) => {
  getDriveId(actualSiteId, token, (driveId) => {
    if (fileSize < 4 * 1024 * 1024) {
      // Small file: simple upload
      uploadSmallFile(driveId, remotePath, localFile, token);
    } else {
      // Large file: resumable upload session
      uploadLargeFile(driveId, remotePath, localFile, token);
    }
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

function uploadSmallFile(driveId, remotePath, localFile, token) {
  const encodedPath = encodeURIComponent(remotePath);
  const apiPath = `/v1.0/drives/${driveId}/root:/${encodedPath}:/content`;
  
  const fileContent = fs.readFileSync(localFile);
  
  const options = {
    hostname: 'graph.microsoft.com',
    port: 443,
    path: apiPath,
    method: 'PUT',
    headers: {
      'Authorization': `Bearer ${token}`,
      'Content-Type': 'application/octet-stream',
      'Content-Length': fileContent.length
    }
  };

  const req = https.request(options, (res) => {
    let data = '';

    res.on('data', (chunk) => {
      data += chunk;
    });

    res.on('end', () => {
      if (res.statusCode === 200 || res.statusCode === 201) {
        try {
          const result = JSON.parse(data);
          console.log(`✓ Uploaded: ${result.name}`);
          console.log(`  URL: ${result.webUrl}`);
          console.log(`  Size: ${result.size} bytes`);
        } catch (err) {
          console.log('✓ Upload successful');
        }
      } else {
        console.error('ERROR: Upload failed');
        console.error(`Status: ${res.statusCode}`);
        console.error(data);
        process.exit(1);
      }
    });
  });

  req.on('error', (e) => {
    console.error(`ERROR: ${e.message}`);
    process.exit(1);
  });

  req.write(fileContent);
  req.end();
}

function uploadLargeFile(driveId, remotePath, localFile, token) {
  console.error('Large file upload (>4MB) not yet implemented');
  console.error('Use simple upload by splitting file or implement resumable upload session');
  process.exit(1);
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
