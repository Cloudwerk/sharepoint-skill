#!/usr/bin/env node
/**
 * SharePoint Download File
 * Downloads a file from SharePoint via Microsoft Graph API
 * 
 * Usage:
 *   node sharepoint-download.js <site-name> <file-path> [output-path]
 *   node sharepoint-download.js TeamSite "Shared Documents/General/document.docx" ./report.docx
 *   node sharepoint-download.js TeamSite "document.pdf"  # Downloads to ./document.pdf
 * 
 * Output: Downloaded file
 */

const https = require('https');
const fs = require('fs');
const { execSync } = require('child_process');
const path = require('path');

const args = process.argv.slice(2);

if (args.length < 2) {
  console.error('Usage: node sharepoint-download.js <site-name> <file-path> [output-path]');
  console.error('');
  console.error('Examples:');
  console.error('  node sharepoint-download.js TeamSite "document.docx"');
  console.error('  node sharepoint-download.js TeamSite "General/document.docx" ./downloads/report.docx');
  process.exit(1);
}

const siteName = args[0];
const filePath = args[1];
const outputPath = args[2] || `./${path.basename(filePath)}`;

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

// Get site ID -> drive ID -> download file
getSiteId(siteId, token, (actualSiteId) => {
  getDriveId(actualSiteId, token, (driveId) => {
    downloadFile(driveId, filePath, token, outputPath);
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

function downloadFile(driveId, filePath, token, outputPath) {
  const encodedPath = encodeURIComponent(filePath);
  const apiPath = `/v1.0/drives/${driveId}/root:/${encodedPath}:/content`;
  
  const options = {
    hostname: 'graph.microsoft.com',
    port: 443,
    path: apiPath,
    method: 'GET',
    headers: {
      'Authorization': `Bearer ${token}`
    }
  };

  const file = fs.createWriteStream(outputPath);

  const req = https.request(options, (res) => {
    if (res.statusCode === 200) {
      res.pipe(file);
      
      file.on('finish', () => {
        file.close();
        console.log(`✓ Downloaded: ${outputPath}`);
        console.log(`  Size: ${fs.statSync(outputPath).size} bytes`);
      });
    } else if (res.statusCode === 302 || res.statusCode === 301) {
      // Follow redirect
      const redirectUrl = new URL(res.headers.location);
      downloadFromUrl(redirectUrl.href, outputPath);
    } else {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        console.error('ERROR: Download failed');
        console.error(`Status: ${res.statusCode}`);
        console.error(data);
        process.exit(1);
      });
    }
  });

  req.on('error', (e) => {
    console.error(`ERROR: ${e.message}`);
    process.exit(1);
  });

  req.end();
}

function downloadFromUrl(url, outputPath) {
  const file = fs.createWriteStream(outputPath);
  
  https.get(url, (res) => {
    res.pipe(file);
    
    file.on('finish', () => {
      file.close();
      console.log(`✓ Downloaded: ${outputPath}`);
      console.log(`  Size: ${fs.statSync(outputPath).size} bytes`);
    });
  }).on('error', (e) => {
    fs.unlink(outputPath, () => {});
    console.error(`ERROR: ${e.message}`);
    process.exit(1);
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
