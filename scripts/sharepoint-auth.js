#!/usr/bin/env node
/**
 * SharePoint Authentication Helper
 * Generates OAuth2 access token for Microsoft Graph API / SharePoint REST API
 * 
 * Usage:
 *   node sharepoint-auth.js
 *   
 * Returns: Access token (stdout)
 * 
 * Credentials from: ~/.config/bobby/sharepoint.env
 */

const https = require('https');
const fs = require('fs');
const path = require('path');

// Load credentials
const envPath = path.join(process.env.HOME, '.config/bobby/sharepoint.env');

if (!fs.existsSync(envPath)) {
  console.error('ERROR: SharePoint credentials not found');
  console.error(`Expected: ${envPath}`);
  process.exit(1);
}

const envContent = fs.readFileSync(envPath, 'utf8');
const credentials = {};

envContent.split('\n').forEach(line => {
  const match = line.match(/^([^=]+)=(.*)$/);
  if (match) {
    credentials[match[1].trim()] = match[2].trim().replace(/^['"]|['"]$/g, '');
  }
});

const TENANT = credentials.TEAMS_TENANT_ID || credentials.TENANT;
const CLIENT_ID = credentials.TEAMS_CLIENT_ID || credentials.CLIENT_ID;
const CLIENT_SECRET = credentials.TEAMS_CLIENT_SECRET || credentials.CLIENT_SECRET;

if (!TENANT) {
  console.error('ERROR: Missing TEAMS_TENANT_ID or TENANT in sharepoint.env');
  console.error('Run setup.js to configure credentials');
  process.exit(1);
}

if (!CLIENT_ID || !CLIENT_SECRET) {
  console.error('ERROR: Missing TEAMS_CLIENT_ID/CLIENT_ID or TEAMS_CLIENT_SECRET/CLIENT_SECRET in sharepoint.env');
  console.error('Run setup.js to configure credentials');
  process.exit(1);
}

// Request token
const tokenUrl = `login.microsoftonline.com`;
const tokenPath = `/${TENANT}/oauth2/v2.0/token`;

const postData = new URLSearchParams({
  grant_type: 'client_credentials',
  client_id: CLIENT_ID,
  client_secret: CLIENT_SECRET,
  scope: 'https://graph.microsoft.com/.default'
}).toString();

const options = {
  hostname: tokenUrl,
  port: 443,
  path: tokenPath,
  method: 'POST',
  headers: {
    'Content-Type': 'application/x-www-form-urlencoded',
    'Content-Length': Buffer.byteLength(postData)
  }
};

const req = https.request(options, (res) => {
  let data = '';

  res.on('data', (chunk) => {
    data += chunk;
  });

  res.on('end', () => {
    if (res.statusCode === 200) {
      const token = JSON.parse(data);
      console.log(token.access_token);
    } else {
      console.error('ERROR: Token request failed');
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

req.write(postData);
req.end();
