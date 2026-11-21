/**
 * Flag / unflag / complete emails.
 * Flagged messages also show up in Microsoft To Do ("Flagged email").
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');
const { FLAG_STATUSES } = require('./format');

/**
 * Flag email handler
 * @param {object} args
 * @param {string} args.ids - One email ID or a comma-separated list
 * @param {string} [args.status='flagged'] - flagged | notFlagged | complete
 * @returns {object} - MCP response
 */
async function handleFlagEmail(args = {}) {
  const ids = String(args.ids || args.id || '')
    .split(',')
    .map(id => id.trim())
    .filter(Boolean);
  const status = args.status || 'flagged';

  if (ids.length === 0) {
    return { content: [{ type: "text", text: "At least one email ID is required (ids)." }] };
  }
  if (!FLAG_STATUSES.includes(status)) {
    return { content: [{ type: "text", text: `status must be one of: ${FLAG_STATUSES.join(', ')}` }] };
  }

  try {
    const accessToken = await ensureAuthenticated();
    const failed = [];

    for (const id of ids) {
      try {
        await callGraphAPI(accessToken, 'PATCH', `me/messages/${encodeURIComponent(id)}`, {
          flag: { flagStatus: status }
        });
      } catch (error) {
        console.error(`Error setting flag on ${id}: ${error.message}`);
        failed.push({ id, error: error.message });
      }
    }

    const done = ids.length - failed.length;
    const label = { flagged: 'flagged', notFlagged: 'unflagged', complete: 'marked flag complete' }[status];
    let text = done > 0 ? `${done} email(s) ${label}.` : '';
    if (failed.length > 0) {
      if (text) text += '\n\n';
      text += `Failed for ${failed.length} email(s):\n` +
        failed.slice(0, 5).map(f => `- ${f.id}: ${f.error}`).join('\n');
    }
    return { content: [{ type: "text", text }] };
  } catch (error) {
    if (error.message === 'Authentication required') {
      return { content: [{ type: "text", text: "Authentication required. Please use the 'authenticate' tool first." }] };
    }
    return { content: [{ type: "text", text: `Error flagging email: ${error.message}` }] };
  }
}

module.exports = handleFlagEmail;
