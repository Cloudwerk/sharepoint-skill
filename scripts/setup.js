#!/usr/bin/env node
/**
 * SharePoint Skill Setup
 * Interactive setup wizard for SharePoint credentials and permissions validation
 * 
 * Usage:
 *   node setup.js
 * 
 * This script will:
 * 1. Check for existing credentials
 * 2. Test App Registration permissions
 * 3. Validate SharePoint access
 * 4. Guide through any missing configuration
 */

const readline = require('readline');
const fs = require('fs');
const path = require('path');
const https = require('https');

const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout
});

const CONFIG_PATH = path.join(process.env.HOME, '.config/bobby/sharepoint.env');
const SKILL_CONFIG_PATH = path.join(__dirname, '..', '.sharepoint-config.json');

let config = {};

console.log('');
console.log('════════════════════════════════════════════════════════════');
console.log('  SharePoint Skill Setup');
console.log('════════════════════════════════════════════════════════════');
console.log('');

async function main() {
  // Step 1: Check existing credentials
  console.log('Step 1: Checking existing credentials...\n');
  
  if (fs.existsSync(CONFIG_PATH)) {
    console.log(`✓ Found credentials: ${CONFIG_PATH}`);
    loadCredentials();
  } else {
    console.log(`✗ No credentials found at: ${CONFIG_PATH}`);
    console.log('  Creating new configuration...\n');
  }

  // Step 2: Prompt for missing values
  await promptForCredentials();

  // Step 3: Test authentication
  console.log('\nStep 2: Testing authentication...\n');
  const token = await testAuthentication();
  
  if (!token) {
    console.log('✗ Authentication failed. Please check your credentials.\n');
    rl.close();
    process.exit(1);
  }
  
  console.log('✓ Authentication successful!\n');

  // Step 4: Test permissions
  console.log('Step 3: Testing SharePoint permissions...\n');
  const permissions = await testPermissions(token);
  
  displayPermissionResults(permissions);

  // Step 5: Test site access
  console.log('\nStep 4: Testing site access...\n');
  const siteAccess = await testSiteAccess(token, config.testSite || 'TeamSite');
  
  if (siteAccess.success) {
    console.log(`✓ Successfully accessed site: ${config.testSite || 'TeamSite'}`);
  } else {
    console.log(`✗ Could not access site: ${config.testSite || 'TeamSite'}`);
    console.log(`  Error: ${siteAccess.error}`);
  }

  // Step 6: Save configuration
  console.log('\nStep 5: Saving configuration...\n');
  saveConfiguration();

  console.log('');
  console.log('════════════════════════════════════════════════════════════');
  console.log('  Setup Complete!');
  console.log('════════════════════════════════════════════════════════════');
  console.log('');
  console.log('Your SharePoint skill is ready to use.');
  console.log('');
  console.log('Try these commands:');
  console.log('  node scripts/sharepoint-list-files.js TeamSite');
  console.log('  node scripts/sharepoint-download.js TeamSite "file.docx"');
  console.log('');

  rl.close();
}

function loadCredentials() {
  const envContent = fs.readFileSync(CONFIG_PATH, 'utf8');
  
  envContent.split('\n').forEach(line => {
    const match = line.match(/^([^=]+)=(.*)$/);
    if (match) {
      const key = match[1].trim();
      const value = match[2].trim().replace(/^['"]|['"]$/g, '');
      config[key] = value;
    }
  });

  // Map to standard names
  config.clientId = config.TEAMS_CLIENT_ID || config.CLIENT_ID;
  config.clientSecret = config.TEAMS_CLIENT_SECRET || config.CLIENT_SECRET;
  config.tenantId = config.TEAMS_TENANT_ID || config.TENANT || config.SHAREPOINT_TENANT;
}

async function promptForCredentials() {
  if (!config.tenantId) {
    config.tenantId = await question('Enter your tenant ID or domain (e.g., contoso.onmicrosoft.com): ');
  } else {
    console.log(`  Tenant: ${config.tenantId}`);
  }

  if (!config.clientId) {
    config.clientId = await question('Enter your Azure App Registration Client ID: ');
  } else {
    console.log(`  Client ID: ${config.clientId}`);
  }

  if (!config.clientSecret) {
    config.clientSecret = await question('Enter your Client Secret: ');
  } else {
    console.log(`  Client Secret: ****${config.clientSecret.slice(-4)}`);
  }

  if (!config.testSite) {
    config.testSite = await question('Enter a SharePoint site to test (e.g., TeamSite): ');
  }
}

function question(prompt) {
  return new Promise((resolve) => {
    rl.question(prompt, (answer) => {
      resolve(answer.trim());
    });
  });
}

async function testAuthentication() {
  return new Promise((resolve) => {
    const tokenUrl = 'login.microsoftonline.com';
    const tokenPath = `/${config.tenantId}/oauth2/v2.0/token`;

    const postData = new URLSearchParams({
      grant_type: 'client_credentials',
      client_id: config.clientId,
      client_secret: config.clientSecret,
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
          resolve(token.access_token);
        } else {
          console.error(`  Authentication failed: ${res.statusCode}`);
          console.error(`  ${data}`);
          resolve(null);
        }
      });
    });

    req.on('error', (e) => {
      console.error(`  Error: ${e.message}`);
      resolve(null);
    });

    req.write(postData);
    req.end();
  });
}

async function testPermissions(token) {
  // Decode token to check permissions
  const parts = token.split('.');
  if (parts.length !== 3) {
    return { valid: false, permissions: [] };
  }

  try {
    const payload = JSON.parse(Buffer.from(parts[1], 'base64').toString());
    const roles = payload.roles || [];
    
    const requiredPermissions = [
      'Sites.Read.All',
      'Sites.ReadWrite.All',
      'Files.Read.All',
      'Files.ReadWrite.All'
    ];

    const results = {};
    requiredPermissions.forEach(perm => {
      results[perm] = roles.includes(perm);
    });

    return { valid: true, permissions: results, roles };
  } catch (err) {
    return { valid: false, error: err.message };
  }
}

function displayPermissionResults(permissions) {
  if (!permissions.valid) {
    console.log('✗ Could not validate permissions');
    if (permissions.error) {
      console.log(`  Error: ${permissions.error}`);
    }
    return;
  }

  console.log('Checking required permissions:\n');

  const checks = [
    { name: 'Sites.Read.All', required: true },
    { name: 'Sites.ReadWrite.All', required: false },
    { name: 'Files.Read.All', required: true },
    { name: 'Files.ReadWrite.All', required: false }
  ];

  checks.forEach(check => {
    const hasPermission = permissions.permissions[check.name];
    const symbol = hasPermission ? '✓' : '✗';
    const label = check.required ? '(required)' : '(optional)';
    
    console.log(`  ${symbol} ${check.name.padEnd(25)} ${label}`);
  });

  console.log('');

  const hasRequired = permissions.permissions['Sites.Read.All'] && permissions.permissions['Files.Read.All'];
  
  if (!hasRequired) {
    console.log('⚠ WARNING: Missing required permissions!');
    console.log('');
    console.log('To add permissions:');
    console.log('1. Go to Azure Portal → App Registrations');
    console.log(`2. Find app: ${config.clientId}`);
    console.log('3. API Permissions → Add permission → Microsoft Graph → Application permissions');
    console.log('4. Add: Sites.Read.All, Files.Read.All (minimum)');
    console.log('5. Grant admin consent');
    console.log('');
  }
}

async function testSiteAccess(token, siteName) {
  return new Promise((resolve) => {
    const hostname = 'graph.microsoft.com';
    
    // Extract tenant from config
    let tenant = config.tenantId.replace('.onmicrosoft.com', '');
    const siteId = `${tenant}.sharepoint.com:/sites/${siteName}`;
    const apiPath = `/v1.0/sites/${siteId}`;

    const options = {
      hostname,
      port: 443,
      path: apiPath,
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
        if (res.statusCode === 200) {
          const site = JSON.parse(data);
          resolve({ success: true, site });
        } else {
          try {
            const error = JSON.parse(data);
            resolve({ success: false, error: error.error?.message || 'Unknown error' });
          } catch {
            resolve({ success: false, error: data });
          }
        }
      });
    });

    req.on('error', (e) => {
      resolve({ success: false, error: e.message });
    });

    req.end();
  });
}

function saveConfiguration() {
  // Save to ~/.config/bobby/sharepoint.env
  const envDir = path.dirname(CONFIG_PATH);
  if (!fs.existsSync(envDir)) {
    fs.mkdirSync(envDir, { recursive: true });
  }

  const envContent = `# SharePoint Skill Configuration
# Generated by setup.js on ${new Date().toISOString()}

TEAMS_CLIENT_ID=${config.clientId}
TEAMS_CLIENT_SECRET=${config.clientSecret}
TEAMS_TENANT_ID=${config.tenantId}

# Legacy compatibility
SHAREPOINT_TENANT=${config.tenantId.replace('.onmicrosoft.com', '')}
SHAREPOINT_BASE_URL=https://${config.tenantId.replace('.onmicrosoft.com', '')}.sharepoint.com
`;

  fs.writeFileSync(CONFIG_PATH, envContent);
  console.log(`✓ Saved credentials: ${CONFIG_PATH}`);

  // Save skill-specific config
  const skillConfig = {
    tenant: config.tenantId,
    testSite: config.testSite,
    setupCompleted: new Date().toISOString()
  };

  fs.writeFileSync(SKILL_CONFIG_PATH, JSON.stringify(skillConfig, null, 2));
  console.log(`✓ Saved skill config: ${SKILL_CONFIG_PATH}`);
}

main().catch(err => {
  console.error('Setup failed:', err);
  rl.close();
  process.exit(1);
});
