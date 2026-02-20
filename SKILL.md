---
name: sharepoint
description: "List, download, and upload files from SharePoint sites via Microsoft Graph API. Use when working with SharePoint files, document libraries, or team sites. Specific triggers: listing files in SharePoint folders, downloading documents from SharePoint sites, uploading files to SharePoint, or accessing SharePoint document libraries programmatically."
---

# SharePoint

Access SharePoint sites programmatically via Microsoft Graph API.

## First-Time Setup

Before using this skill, run the interactive setup wizard:

```bash
node scripts/setup.js
```

The setup will:
1. Check for existing credentials
2. Prompt for missing values (Tenant ID, Client ID, Client Secret)
3. Test authentication
4. Validate SharePoint permissions
5. Test access to a SharePoint site
6. Save configuration to `~/.config/bobby/sharepoint.env`

**What you'll need:**
- Azure App Registration Client ID
- Client Secret
- Tenant ID or domain (e.g., `contoso.onmicrosoft.com`)
- A SharePoint site name to test (e.g., `TeamSite`)

**Required Azure permissions:**
- `Sites.Read.All` (required for listing/downloading)
- `Files.Read.All` (required for file access)
- `Sites.ReadWrite.All` (optional, needed for upload)
- `Files.ReadWrite.All` (optional, needed for upload)

After setup, the skill is ready to use across all agents.

## Quick Start

```bash
# List files in a SharePoint site
node scripts/sharepoint-list-files.js TeamSite

# Download a file
node scripts/sharepoint-download.js TeamSite "document.docx" ./document.docx

# Upload a file
node scripts/sharepoint-upload.js TeamSite ./report.pdf "General/report.pdf"
```

All scripts automatically handle authentication using credentials from `~/.config/bobby/sharepoint.env`.

## Core Operations

### 1. List Files

**Script:** `scripts/sharepoint-list-files.js`

**Usage:**
```bash
node scripts/sharepoint-list-files.js <site-name> [folder-path]
```

**Examples:**
```bash
# List root folder
node scripts/sharepoint-list-files.js TeamSite

# List specific folder
node scripts/sharepoint-list-files.js TeamSite "Shared Documents/General"

# List another site
node scripts/sharepoint-list-files.js ProjectSite
```

**Output:** JSON array with file details:
```json
[
  {
    "name": "document.docx",
    "size": 12345,
    "webUrl": "https://...",
    "downloadUrl": "https://...",
    "lastModified": "2026-02-20T10:00:00Z",
    "isFolder": false,
    "type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
  }
]
```

### 2. Download File

**Script:** `scripts/sharepoint-download.js`

**Usage:**
```bash
node scripts/sharepoint-download.js <site-name> <file-path> [output-path]
```

**Examples:**
```bash
# Download to current directory
node scripts/sharepoint-download.js TeamSite "document.docx"

# Download to specific location
node scripts/sharepoint-download.js TeamSite "General/report.docx" ./downloads/report.docx

# Download from subfolder
node scripts/sharepoint-download.js TeamSite "Shared Documents/Marketing/slides.pptx" ./slides.pptx
```

**Output:** Downloaded file + confirmation message

### 3. Upload File

**Script:** `scripts/sharepoint-upload.js`

**Usage:**
```bash
node scripts/sharepoint-upload.js <site-name> <local-file> <remote-path>
```

**Examples:**
```bash
# Upload to root
node scripts/sharepoint-upload.js TeamSite ./document.pdf "document.pdf"

# Upload to subfolder
node scripts/sharepoint-upload.js TeamSite ./report.docx "General/report.docx"

# Upload to specific library
node scripts/sharepoint-upload.js TeamSite ./updated-file.docx "Shared Documents/General/file.docx"
```

**Output:** Upload confirmation with file URL

**Note:** Currently supports files <4MB. For larger files, split or implement resumable upload session.

## Site Names

SharePoint site names are used as parameters in all commands.

**Format:** The site name is the last part of the SharePoint URL.

Example: For `https://contoso.sharepoint.com/sites/TeamSite`, use `TeamSite` as the site name.

**Tips:**
- Site names are case-sensitive
- Use the exact name as shown in SharePoint URL
- If unsure, visit the site in browser and copy the name from URL

## Authentication

All scripts use OAuth2 Client Credentials flow via Microsoft Graph API.

**Setup:** Run `node scripts/setup.js` once to configure credentials interactively.

**Credentials location:** `~/.config/bobby/sharepoint.env`

**Required fields:**
- `TEAMS_CLIENT_ID` - Azure App Registration Client ID
- `TEAMS_CLIENT_SECRET` - Client Secret
- `TEAMS_TENANT_ID` - Tenant ID (e.g., `contoso.onmicrosoft.com`)

**Token generation:** Automatically handled by `scripts/sharepoint-auth.js`

**Manual setup:** If you prefer not to use the setup wizard, create `~/.config/bobby/sharepoint.env` manually:

```bash
TEAMS_CLIENT_ID=your-client-id
TEAMS_CLIENT_SECRET=your-client-secret
TEAMS_TENANT_ID=contoso.onmicrosoft.com
```

## Error Handling

### Access Denied

**Symptom:** `"code": "accessDenied"`

**Cause:** Azure App Registration lacks SharePoint/Sites permissions

**Fix:** Add these permissions in Azure Portal:
- `Sites.Read.All` - Read all site collections
- `Sites.ReadWrite.All` - Read/write all site collections (if upload needed)
- `Files.Read.All` - Read all files
- `Files.ReadWrite.All` - Read/write all files (if upload needed)

After adding permissions, grant admin consent.

### Site Not Found

**Symptom:** `"error": "Site not found"`

**Possible causes:**
- Site name is case-sensitive
- Site does not exist
- Site is not accessible to the app

**Solution:** Check `references/SITES.md` for correct site names

### File Not Found

**Symptom:** `404` or `"itemNotFound"`

**Possible causes:**
- File path is incorrect (case-sensitive)
- File is in a different library
- Missing `Freigegebene Dokumente/` prefix

**Solution:** 
1. List the folder first to see available files
2. Use exact path including library name

## Tips

- **Folder paths:** Default document library is often `Shared Documents` (English) or localized name in your tenant language
- **Quotes:** Always quote paths with spaces: `"Shared Documents/General"`
- **Root folder:** Use empty string `""` or omit parameter for root
- **Case sensitivity:** File and folder names are case-sensitive in some Graph API endpoints
- **Testing:** Use `sharepoint-list-files.js` first to explore site structure before downloading/uploading
- **Library names:** Use `sharepoint-list-files.js` to discover the exact document library name in your tenant

## Integration Example

**Workflow: Update a document**

```bash
# 1. Download current version
node scripts/sharepoint-download.js TeamSite "Shared Documents/General/report.docx" ./report-current.docx

# 2. Edit file (using docx skill or manual editing)
# ... modifications ...

# 3. Upload updated version
node scripts/sharepoint-upload.js TeamSite ./report-updated.docx "Shared Documents/General/report.docx"
```

## Limitations

- Large file uploads (>4MB) not yet implemented (requires resumable upload session)
- No folder creation support (add if needed)
- No file deletion support (add if needed)
- Read-only access until admin grants write permissions

## Resources

- **SITES.md** - Detailed site information and folder structures
- **sharepoint-auth.js** - Authentication helper (used by all scripts)
- **sharepoint-list-files.js** - List files and folders
- **sharepoint-download.js** - Download files
- **sharepoint-upload.js** - Upload files (<4MB)
