#!/usr/bin/env node
/**
 * Get Tenant Helper
 * Reads tenant from sharepoint.env config
 * 
 * Returns: Tenant domain (e.g., "contoso")
 */

const fs = require('fs');
const path = require('path');

const envPath = path.join(process.env.HOME, '.config/bobby/sharepoint.env');

if (!fs.existsSync(envPath)) {
  console.error('ERROR: SharePoint credentials not found. Run setup.js first.');
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

const tenantId = credentials.TEAMS_TENANT_ID || credentials.TENANT || credentials.SHAREPOINT_TENANT;

if (!tenantId) {
  console.error('ERROR: No tenant configured. Run setup.js first.');
  process.exit(1);
}

// Extract tenant name from domain (e.g., "contoso.onmicrosoft.com" -> "contoso")
const tenant = tenantId.replace('.onmicrosoft.com', '');

console.log(tenant);
