/**
 * Reply drafts. draft-reply creates a reply / reply-all draft inside the
 * original thread and never sends it; send-draft sends an existing draft.
 * Keeping these separate lets a client allow drafting but gate sending.
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

function formatRecipients(list) {
  if (!list || list.length === 0) return 'None';
  return list.map(r => `${r.emailAddress.name || ''} <${r.emailAddress.address}>`.trim()).join(', ');
}

function authError(error) {
  return error.message === 'Authentication required' || error.message === 'UNAUTHORIZED';
}

/**
 * @param {object} args
 * @param {string} args.id - ID of the message to reply to
 * @param {string} [args.body] - Plain-text reply, placed above the quoted thread
 * @param {boolean} [args.replyAll=true] - Reply to all recipients
 */
async function handleDraftReply(args = {}) {
  if (!args.id) {
    return { content: [{ type: "text", text: "Email ID is required (id)." }] };
  }
  const replyAll = args.replyAll !== false;
  const action = replyAll ? 'createReplyAll' : 'createReply';

  try {
    const accessToken = await ensureAuthenticated();
    const draft = await callGraphAPI(
      accessToken,
      'POST',
      `me/messages/${encodeURIComponent(args.id)}/${action}`,
      { comment: args.body || '' }
    );

    return {
      content: [{
        type: "text",
        text: `${replyAll ? 'Reply-all' : 'Reply'} draft created (not sent).\n\n` +
          `Draft ID: ${draft.id}\n` +
          `Subject: ${draft.subject}\n` +
          `To: ${formatRecipients(draft.toRecipients)}\n` +
          `Cc: ${formatRecipients(draft.ccRecipients)}\n\n` +
          `Check the recipients, then send it with send-draft or from Outlook.`
      }]
    };
  } catch (error) {
    if (authError(error)) {
      return { content: [{ type: "text", text: "Authentication required. Please use the 'authenticate' tool first." }] };
    }
    return { content: [{ type: "text", text: `Error creating reply draft: ${error.message}` }] };
  }
}

/**
 * @param {object} args
 * @param {string} args.id - ID of the draft to send
 */
async function handleSendDraft(args = {}) {
  if (!args.id) {
    return { content: [{ type: "text", text: "Draft ID is required (id)." }] };
  }

  try {
    const accessToken = await ensureAuthenticated();
    const draft = await callGraphAPI(accessToken, 'GET', `me/messages/${encodeURIComponent(args.id)}`, null, {
      $select: 'id,isDraft,subject,toRecipients,ccRecipients,bccRecipients'
    });
    if (draft.isDraft !== true) {
      return { content: [{ type: "text", text: "That message is not a draft; nothing was sent." }] };
    }

    await callGraphAPI(accessToken, 'POST', `me/messages/${encodeURIComponent(args.id)}/send`);

    return {
      content: [{
        type: "text",
        text: `Draft sent.\n\nSubject: ${draft.subject}\nTo: ${formatRecipients(draft.toRecipients)}\n` +
          `Cc: ${formatRecipients(draft.ccRecipients)}\nBcc: ${formatRecipients(draft.bccRecipients)}`
      }]
    };
  } catch (error) {
    if (authError(error)) {
      return { content: [{ type: "text", text: "Authentication required. Please use the 'authenticate' tool first." }] };
    }
    return { content: [{ type: "text", text: `Error sending draft: ${error.message}` }] };
  }
}

module.exports = { handleDraftReply, handleSendDraft };
