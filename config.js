/**
 * Configuration for Outlook MCP Server
 */
const path = require('path');
const os = require('os');

// Ensure we have a home directory path even if process.env.HOME is undefined
const homeDir = process.env.HOME || process.env.USERPROFILE || os.homedir() || '/tmp';

// Graph scopes requested at login AND on every token refresh. They must match:
// a refresh that asks for fewer scopes gets back a token with fewer permissions
// (e.g. Mail.Read only), which silently breaks moving/flagging mail.
// Override with MS_SCOPES (space-separated) if your app registration differs.
const DEFAULT_GRAPH_SCOPES = [
  'offline_access',
  'User.Read',
  'Mail.Read',
  'Mail.ReadWrite',
  'Mail.Send',
  'Calendars.Read',
  'Calendars.ReadWrite',
  'Contacts.Read',
  'Files.Read',
  'Files.ReadWrite'
];

function getGraphScopes(envValue = process.env.MS_SCOPES) {
  const fromEnv = (envValue || '').split(/\s+/).filter(Boolean);
  return fromEnv.length > 0 ? fromEnv : [...DEFAULT_GRAPH_SCOPES];
}

module.exports = {
  DEFAULT_GRAPH_SCOPES,
  getGraphScopes,

  // Verbose request logging (full URLs incl. search terms) - off by default
  DEBUG: process.env.OUTLOOK_MCP_DEBUG === 'true',

  // Server information
  SERVER_NAME: "m365-assistant",
  SERVER_VERSION: "2.0.0",
  
  // Test mode setting
  USE_TEST_MODE: process.env.USE_TEST_MODE === 'true',
  
  // Authentication configuration
  AUTH_CONFIG: {
    clientId: process.env.OUTLOOK_CLIENT_ID || process.env.MS_CLIENT_ID || '',
    clientSecret: process.env.OUTLOOK_CLIENT_SECRET || process.env.MS_CLIENT_SECRET || '',
    redirectUri: 'http://localhost:3333/auth/callback',
    scopes: getGraphScopes(),
    tokenStorePath: path.join(homeDir, '.outlook-mcp-tokens.json'),
    authServerUrl: 'http://localhost:3333'
  },
  
  // Microsoft Graph API
  GRAPH_API_ENDPOINT: 'https://graph.microsoft.com/v1.0/',
  
  // Email constants
  EMAIL_SELECT_FIELDS: 'id,subject,from,toRecipients,ccRecipients,receivedDateTime,bodyPreview,hasAttachments,importance,isRead,flag,categories,conversationId',
  EMAIL_DETAIL_FIELDS: 'id,subject,from,toRecipients,ccRecipients,bccRecipients,receivedDateTime,bodyPreview,body,hasAttachments,importance,isRead,flag,categories,conversationId,internetMessageHeaders',
  
  // Calendar constants
  CALENDAR_SELECT_FIELDS: 'id,subject,bodyPreview,start,end,location,organizer,attendees,isAllDay,isCancelled',
  
  // Pagination
  DEFAULT_PAGE_SIZE: 25,
  MAX_RESULT_COUNT: 50,
  // Upper bound for list/search when paging through a large folder (e.g. a backlog)
  MAX_LIST_COUNT: 500,

  // Timezone
  DEFAULT_TIMEZONE: "Central European Standard Time",

  // OneDrive constants
  ONEDRIVE_SELECT_FIELDS: 'id,name,size,lastModifiedDateTime,webUrl,folder,file,parentReference',
  ONEDRIVE_UPLOAD_THRESHOLD: 4 * 1024 * 1024, // 4MB - files larger than this need chunked upload

  // Power Automate / Flow constants
  FLOW_API_ENDPOINT: 'https://api.flow.microsoft.com',
  FLOW_SCOPE: 'https://service.flow.microsoft.com/.default',
};
