/**
 * List emails functionality
 */
const config = require('../config');
const { callGraphAPIPaginated } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');
const { resolveFolderPath } = require('./folder-utils');
const { clampCount, buildMessageFilter, formatEmailList, UNTRUSTED_NOTE } = require('./format');

/**
 * List emails handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleListEmails(args = {}) {
  const folder = args.folder || "inbox";
  const requestedCount = clampCount(args.count, 10);

  try {
    // Validate filters before touching the API
    const filter = buildMessageFilter({
      receivedAfter: args.receivedAfter,
      receivedBefore: args.receivedBefore,
      unreadOnly: args.unreadOnly,
      hasAttachments: args.hasAttachments,
      flagStatus: args.flagStatus
    });

    // Get access token
    const accessToken = await ensureAuthenticated();

    // Resolve the folder path (throws if the folder does not exist)
    const endpoint = await resolveFolderPath(accessToken, folder);

    const queryParams = {
      $top: Math.min(50, requestedCount), // 50 per page; pagination fetches the rest
      $orderby: args.oldestFirst === true ? 'receivedDateTime asc' : 'receivedDateTime desc',
      $select: config.EMAIL_SELECT_FIELDS
    };
    if (filter) {
      queryParams.$filter = filter;
    }

    // Make API call with pagination support
    const response = await callGraphAPIPaginated(accessToken, 'GET', endpoint, queryParams, requestedCount);

    if (!response.value || response.value.length === 0) {
      return {
        content: [{
          type: "text",
          text: `No emails found in ${folder}.`
        }]
      };
    }

    const emailList = formatEmailList(response.value, { includePreview: args.includePreview !== false });

    return {
      content: [{
        type: "text",
        text: `Found ${response.value.length} emails in ${folder}:\n${UNTRUSTED_NOTE}\n\n${emailList}`
      }]
    };
  } catch (error) {
    if (error.message === 'Authentication required') {
      return {
        content: [{
          type: "text",
          text: "Authentication required. Please use the 'authenticate' tool first."
        }]
      };
    }

    return {
      content: [{
        type: "text",
        text: `Error listing emails: ${error.message}`
      }]
    };
  }
}

module.exports = handleListEmails;
