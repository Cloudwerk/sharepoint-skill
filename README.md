# SharePoint Skill

**SharePoint file operations via Microsoft Graph API**

List, download, and upload files from SharePoint sites programmatically using Microsoft Graph API. Designed for AI agents and automation workflows.

## Features

- 📂 **List files** in SharePoint folders and document libraries
- ⬇️ **Download files** from SharePoint to local filesystem
- ⬆️ **Upload files** to SharePoint (files <4MB)
- 🔐 **OAuth2 authentication** via Azure App Registration
- 🎯 **Interactive setup wizard** - guided configuration with validation
- 🌐 **Multi-tenant support** - works with any Microsoft 365 tenant
- ✅ **Permission validation** - checks required Graph API permissions

## Quick Start

### 1. Run Setup

**No dependencies required!** Scripts use Node.js built-in `https` module.

```bash
node scripts/setup.js
```

The setup wizard will:
- Prompt for Azure App Registration credentials (Tenant ID, Client ID, Client Secret)
- Test authentication
- Validate SharePoint permissions
- Test site access
- Save configuration

### 3. Use the Scripts

```bash
# List files in a SharePoint site
node scripts/sharepoint-list-files.js TeamSite

# Download a file
node scripts/sharepoint-download.js TeamSite "document.docx" ./document.docx

# Upload a file
node scripts/sharepoint-upload.js TeamSite ./report.pdf "Shared Documents/report.pdf"
```

## Azure App Registration Setup

### Required Permissions

Your Azure App Registration needs these Microsoft Graph **Application permissions**:

- `Sites.Read.All` ✅ Required
- `Files.Read.All` ✅ Required
- `Sites.ReadWrite.All` (optional, for upload)
- `Files.ReadWrite.All` (optional, for upload)

### Setup Steps

1. Go to [Azure Portal](https://portal.azure.com) → **App Registrations**
2. Create new registration or use existing
3. **API Permissions** → Add permission → **Microsoft Graph** → **Application permissions**
4. Add the required permissions listed above
5. **Grant admin consent** (important!)
6. Copy **Application (client) ID** and **Directory (tenant) ID**
7. **Certificates & secrets** → New client secret → Copy the secret value

Run `node scripts/setup.js` and enter these values when prompted.

## Scripts

### setup.js

Interactive setup wizard for first-time configuration.

```bash
node scripts/setup.js
```

### sharepoint-list-files.js

List files and folders in a SharePoint site.

```bash
node scripts/sharepoint-list-files.js <site-name> [folder-path]

# Examples
node scripts/sharepoint-list-files.js TeamSite
node scripts/sharepoint-list-files.js TeamSite "Shared Documents/General"
```

**Output:** JSON array with file details (name, size, URLs, last modified, type)

### sharepoint-download.js

Download a file from SharePoint.

```bash
node scripts/sharepoint-download.js <site-name> <file-path> [output-path]

# Examples
node scripts/sharepoint-download.js TeamSite "document.docx"
node scripts/sharepoint-download.js TeamSite "General/report.docx" ./downloads/report.docx
```

### sharepoint-upload.js

Upload a file to SharePoint (max 4MB).

```bash
node scripts/sharepoint-upload.js <site-name> <local-file> <remote-path>

# Examples
node scripts/sharepoint-upload.js TeamSite ./document.pdf "document.pdf"
node scripts/sharepoint-upload.js TeamSite ./file.txt "General/file.txt"
```

**Note:** Large file uploads (>4MB) require resumable upload session (not yet implemented).

## Configuration

Credentials are stored in `~/.config/bobby/sharepoint.env`:

```env
TEAMS_CLIENT_ID=your-client-id
TEAMS_CLIENT_SECRET=your-client-secret
TEAMS_TENANT_ID=contoso.onmicrosoft.com
```

The setup wizard creates this file automatically, but you can also edit it manually.

## Site Names

SharePoint site names are the last part of the site URL.

**Example:** For `https://contoso.sharepoint.com/sites/TeamSite`, use `TeamSite` as the site name.

Site names are case-sensitive - use the exact name from the URL.

## Tips

- 🔍 **Explore first:** Use `sharepoint-list-files.js` to see available files before downloading
- 📝 **Quote paths:** Always quote paths with spaces: `"Shared Documents/General"`
- 🌍 **Library names:** Default library is often `Shared Documents` (English) or localized in your tenant language
- 📂 **Root folder:** Use empty string `""` or omit the folder parameter for root
- ⚠️ **Case sensitive:** File and folder names are case-sensitive

## Troubleshooting

### Access Denied

**Error:** `"code": "accessDenied"`

**Solution:** Your Azure App Registration is missing SharePoint permissions. Add `Sites.Read.All` and `Files.Read.All`, then grant admin consent.

### Site Not Found

**Error:** `"error": "Site not found"`

**Possible causes:**
- Site name is case-sensitive (check exact name in URL)
- Site does not exist
- Site is not accessible to the app

### Authentication Failed

**Error:** Token request failed

**Solution:**
- Verify Client ID and Client Secret are correct
- Check Tenant ID format (e.g., `contoso.onmicrosoft.com`)
- Ensure client secret hasn't expired in Azure Portal

## Use Cases

- 🤖 **AI Agent workflows** - Download documents for processing, upload generated reports
- 📊 **Automation** - Sync files between systems, backup documents
- 🔄 **Integration** - Connect SharePoint with other tools and platforms
- 📝 **Document management** - Programmatic access to SharePoint libraries

## Architecture

```
SharePoint Skill
├── scripts/
│   ├── setup.js              # Interactive setup wizard
│   ├── sharepoint-auth.js    # OAuth2 token generation
│   ├── sharepoint-list-files.js
│   ├── sharepoint-download.js
│   ├── sharepoint-upload.js
│   └── get-tenant.js         # Tenant helper (loads from config)
└── SKILL.md                  # Full documentation
```

All scripts use Node.js with built-in `https` module - no external HTTP libraries needed.

## Contributing

Contributions welcome! Open issues or pull requests on GitHub.

## License

MIT License - See [LICENSE](LICENSE) file for details.

## Credits

Created by [CloudWerk GmbH](https://www.cloudwerk.com) for AI agent workflows.
