[![MseeP.ai Security Assessment Badge](https://mseep.net/pr/ryaker-outlook-mcp-badge.png)](https://mseep.ai/app/ryaker-outlook-mcp)

# M365 Assistant MCP Server

A comprehensive MCP (Model Context Protocol) server that connects Claude with Microsoft 365 services through the Microsoft Graph API and Power Automate API.

## Supported Services

- **Outlook** - Email, calendar, folders, and rules
- **OneDrive** - Files, folders, search, and sharing
- **Power Automate** - Flows, environments, and run history

## Directory Structure

```
├── index.js                 # Main entry point
├── config.js                # Configuration settings
├── auth/                    # Authentication modules
│   ├── index.js             # Authentication exports
│   ├── token-manager.js     # Token storage and refresh (Graph + Flow)
│   └── tools.js             # Auth-related tools
├── calendar/                # Calendar functionality
│   ├── index.js             # Calendar exports
│   ├── list.js              # List events
│   ├── create.js            # Create event
│   ├── delete.js            # Delete event
│   ├── cancel.js            # Cancel event
│   ├── accept.js            # Accept event
│   └── decline.js           # Decline event
├── email/                   # Email functionality
│   ├── index.js             # Email exports
│   ├── list.js              # List emails
│   ├── search.js            # Search emails
│   ├── read.js              # Read email
│   ├── send.js              # Send email
│   └── mark-as-read.js      # Mark email read/unread
├── folder/                  # Folder functionality
│   ├── index.js             # Folder exports
│   ├── list.js              # List folders
│   ├── create.js            # Create folder
│   └── move.js              # Move emails
├── rules/                   # Email rules functionality
│   ├── index.js             # Rules exports
│   ├── list.js              # List rules
│   └── create.js            # Create rule
├── onedrive/                # OneDrive functionality
│   ├── index.js             # OneDrive exports
│   ├── list.js              # List files/folders
│   ├── search.js            # Search files
│   ├── download.js          # Get download URL
│   ├── upload.js            # Simple upload (<4MB)
│   ├── upload-large.js      # Chunked upload (>4MB)
│   ├── share.js             # Create sharing link
│   └── folder.js            # Create/delete folders
├── power-automate/          # Power Automate functionality
│   ├── index.js             # Power Automate exports
│   ├── flow-api.js          # Flow API client
│   ├── list-environments.js # List environments
│   ├── list-flows.js        # List flows
│   ├── run-flow.js          # Trigger flow
│   ├── list-runs.js         # Run history
│   └── toggle-flow.js       # Enable/disable flow
└── utils/                   # Utility functions
    ├── graph-api.js         # Microsoft Graph API helper
    ├── odata-helpers.js     # OData query building
    └── mock-data.js         # Test mode data
```

## Features

- **Authentication**: OAuth 2.0 authentication with Microsoft Graph API (+ Flow API for Power Automate)
- **Email Management**: List, search, read, send, and organize emails
- **Calendar Management**: List, create, accept, decline, and delete calendar events
- **OneDrive Integration**: List, search, upload, download, and share files
- **Power Automate**: List environments/flows, trigger flows, view run history
- **Modular Structure**: Clean separation of concerns for maintainability
- **Test Mode**: Simulated responses for testing without real API calls

## Available Tools

### Outlook (Email & Calendar)
| Tool | Description |
|------|-------------|
| `list-emails` | List emails in a folder with preview and unread/flag/attachment markers; filter by date range, unread, flag status; oldest-first; up to 500 |
| `search-emails` | Search emails with filters (same date/flag filters; `strict` disables broadened fallbacks) |
| `read-email` | Read email content (sanitized; does not mark as read) |
| `send-email` | Send a new email |
| `draft-email` | Create a new draft |
| `draft-reply` | Create a reply / reply-all draft in the thread (never sends) |
| `send-draft` | Send an existing draft |
| `update-draft` | Edit a draft's subject, body, recipients or importance; on replies, keeps the quoted thread by default (never sends) |
| `add-attachment` | Attach a local file or a OneDrive file to a draft, up to 150 MB (never sends) |
| `mark-as-read` | Mark email as read/unread |
| `flag-email` | Flag, unflag or complete one or more emails (shows in Microsoft To Do) |
| `delete-email` | Move to Deleted Items, or permanently delete |
| `list-events` | List calendar events |
| `create-event` | Create calendar event |
| `accept-event` | Accept event invitation |
| `decline-event` | Decline event invitation |
| `delete-event` | Delete calendar event |
| `list-folders` | List mail folders |
| `create-folder` | Create mail folder |
| `move-emails` | Move emails between folders |
| `list-rules` | List inbox rules |
| `create-rule` | Create inbox rule |

### OneDrive
| Tool | Description |
|------|-------------|
| `onedrive-list` | List files in a path |
| `onedrive-search` | Search files by query |
| `onedrive-download` | Get download URL |
| `onedrive-upload` | Upload small file (<4MB) |
| `onedrive-upload-large` | Chunked upload (>4MB) |
| `onedrive-share` | Create sharing link |
| `onedrive-create-folder` | Create folder |
| `onedrive-delete` | Delete file or folder |

### Power Automate
| Tool | Description |
|------|-------------|
| `flow-list-environments` | List Power Platform environments |
| `flow-list` | List flows in environment |
| `flow-run` | Trigger a manual flow |
| `flow-list-runs` | Get flow run history |
| `flow-toggle` | Enable/disable a flow |

## Quick Start

1. **Install dependencies**: `npm install`
2. **Azure setup**: Register app in Azure Portal (see detailed steps below)
3. **Configure environment**: Copy `.env.example` to `.env` and add your Azure credentials
4. **Configure Claude**: Update your Claude Desktop config with the server path
5. **Start auth server**: `npm run auth-server`
6. **Authenticate**: Use the authenticate tool in Claude to get the OAuth URL
7. **Start using**: Access your M365 data through Claude!

## Installation

### Prerequisites
- Node.js 14.0.0 or higher
- npm or yarn package manager
- Azure account for app registration

### Install Dependencies

```bash
npm install
```

## Azure App Registration & Configuration

### App Registration

1. Open [Azure Portal](https://portal.azure.com/)
2. Search for "App registrations"
3. Click "New registration"
4. Name: "M365 MCP Server"
5. Account type: "Accounts in any organizational directory and personal Microsoft accounts"
6. Redirect URI: Web → `http://localhost:3333/auth/callback`
7. Click "Register"
8. Copy the "Application (client) ID" for your `.env` file

### App Permissions

1. Go to "API permissions" under Manage
2. Click "Add a permission" → "Microsoft Graph" → "Delegated permissions"
3. Add these permissions:
   - `offline_access`
   - `User.Read`
   - `Mail.Read`, `Mail.ReadWrite`, `Mail.Send`
   - `Calendars.Read`, `Calendars.ReadWrite`
   - `Contacts.Read`
   - `Files.Read`, `Files.ReadWrite`
4. Click "Add permissions"

The server requests exactly this list at login **and on every token refresh**
(see `DEFAULT_GRAPH_SCOPES` in `config.js`). If your app registration grants a
different set, set `MS_SCOPES` (space-separated) in `.env` so login and refresh
stay in sync — a refresh that asks for fewer scopes silently downgrades the token.

**For Power Automate** (optional):
- Requires additional Azure AD configuration with Flow API scope
- See Power Automate section below for details

### Client Secret

1. Go to "Certificates & secrets" → "Client secrets"
2. Click "New client secret"
3. Add description and select expiration
4. **Copy the VALUE** (not the Secret ID)

## Configuration

### 1. Environment Variables

```bash
cp .env.example .env
```

Edit `.env`:
```bash
# Get these values from Azure Portal > App Registrations > Your App
MS_CLIENT_ID=your-application-client-id-here
MS_CLIENT_SECRET=your-client-secret-VALUE-here
MS_TENANT_ID=your-tenant-id-here
USE_TEST_MODE=false
```

**Important Notes:**
- Use `MS_CLIENT_ID` and `MS_CLIENT_SECRET` in the `.env` file. Both `index.js` and the auth server load the `.env` next to them, wherever they are started from
- Personal accounts (outlook.com / hotmail / live): set `MS_TENANT_ID=consumers` (or `common`) and register the app for personal accounts
- Set `MS_TENANT_ID` to your tenant GUID for single-tenant work/school apps to avoid `/common` endpoint errors
- `OUTLOOK_MCP_DEBUG=true` logs full Graph URLs (including search terms) to stderr; off by default
- `OUTLOOK_MCP_ATTACHMENT_DIRS` limits the local folders `add-attachment` may read from (separated by `:`; `;` on Windows). Hidden files and hidden folders are always refused
- For Claude Desktop config, you'll use `OUTLOOK_CLIENT_ID` and `OUTLOOK_CLIENT_SECRET`
- Always use the client secret **VALUE**, never the Secret ID

### 2. Claude Desktop Configuration

Add to your Claude Desktop config:

```json
{
  "mcpServers": {
    "m365-assistant": {
      "command": "node",
      "args": ["/path/to/outlook-mcp/index.js"],
      "env": {
        "USE_TEST_MODE": "false",
        "OUTLOOK_CLIENT_ID": "your-client-id",
        "OUTLOOK_CLIENT_SECRET": "your-client-secret"
      }
    }
  }
}
```

## Authentication

### Graph API (Outlook + OneDrive)

1. Start auth server: `npm run auth-server` (listens on 127.0.0.1:3333 only)
2. Use the `authenticate` tool in Claude
3. Visit the provided URL and sign in
4. Tokens saved to `~/.outlook-mcp-tokens.json` and refreshed automatically

On a remote/headless machine, open an SSH tunnel from the machine with the
browser first: `ssh -L 3333:localhost:3333 <host>`, then visit the URL locally.

### Power Automate (Optional)

Power Automate requires a separate token with the Flow API scope. Configure additional Azure AD permissions for `https://service.flow.microsoft.com//.default` scope.

**Limitations:**
- Only solution-aware flows are accessible
- Only manual trigger flows can be run via API
- Requires environment ID for most operations

## Troubleshooting

### Common Issues

**"Cannot find module"**
```bash
npm install
```

**"Port 3333 in use"**
```bash
npx kill-port 3333
npm run auth-server
```

**"Invalid client secret" (AADSTS7000215)**
- Use the secret **VALUE**, not the Secret ID

**"Authentication required"**
- Delete `~/.outlook-mcp-tokens.json` and re-authenticate

## Testing

```bash
# Run with MCP Inspector
npm run inspect

# Run in test mode (mock data)
npm run test-mode

# Run Jest tests
npm test
```

## Extending the Server

1. Create new module directory
2. Implement tool handlers in separate files
3. Export tool definitions from module index
4. Import and add to `TOOLS` array in `index.js`
