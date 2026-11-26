/**
 * update-draft: edits an existing draft (subject, body, recipients,
 * importance). Refuses anything that is not a draft and never sends.
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');
const { parseRecipients, mergeRecipients, formatRecipients, buildBody } = require('./draft-utils');

const BODY_MODES = ['auto', 'replace', 'prepend'];
const IMPORTANCE = ['low', 'normal', 'high'];

function text(message) {
  return { content: [{ type: "text", text: message }] };
}

/**
 * @param {object} args
 * @param {string} args.id - Draft ID
 * @param {string} [args.subject]
 * @param {string} [args.body] - New body text (plain or HTML)
 * @param {string} [args.bodyMode] - auto | replace | prepend
 * @param {string} [args.to] - Replaces the To list ("" clears it)
 * @param {string} [args.cc] - Replaces the Cc list
 * @param {string} [args.bcc] - Replaces the Bcc list
 * @param {string} [args.addTo] - Adds to the To list
 * @param {string} [args.addCc] - Adds to the Cc list
 * @param {string} [args.addBcc] - Adds to the Bcc list
 * @param {string} [args.importance] - low | normal | high
 */
async function handleUpdateDraft(args = {}) {
  if (!args.id) return text("Draft ID is required (id).");

  const bodyMode = args.bodyMode || 'auto';
  if (!BODY_MODES.includes(bodyMode)) {
    return text(`bodyMode must be one of: ${BODY_MODES.join(', ')}.`);
  }
  if (args.importance !== undefined && !IMPORTANCE.includes(args.importance)) {
    return text(`importance must be one of: ${IMPORTANCE.join(', ')}.`);
  }

  const fields = ['subject', 'body', 'to', 'cc', 'bcc', 'addTo', 'addCc', 'addBcc', 'importance'];
  if (!fields.some(f => args[f] !== undefined)) {
    return text(`Nothing to change. Give at least one of: ${fields.join(', ')}.`);
  }

  try {
    const accessToken = await ensureAuthenticated();
    const endpoint = `me/messages/${encodeURIComponent(args.id)}`;
    const draft = await callGraphAPI(accessToken, 'GET', endpoint, null, {
      $select: 'id,isDraft,subject,body,toRecipients,ccRecipients,bccRecipients,importance'
    });
    if (draft.isDraft !== true) {
      return text("That message is not a draft; nothing was changed.");
    }

    const patch = {};
    const notes = [];

    if (args.subject !== undefined) patch.subject = args.subject;
    if (args.importance !== undefined) patch.importance = args.importance;

    for (const [field, setArg, addArg] of [
      ['toRecipients', 'to', 'addTo'],
      ['ccRecipients', 'cc', 'addCc'],
      ['bccRecipients', 'bcc', 'addBcc']
    ]) {
      if (args[setArg] === undefined && args[addArg] === undefined) continue;
      let list = args[setArg] !== undefined ? parseRecipients(args[setArg]) : (draft[field] || []);
      if (args[addArg] !== undefined) list = mergeRecipients(list, parseRecipients(args[addArg]));
      patch[field] = list;
    }

    if (args.body !== undefined) {
      const { body, note } = buildBody(draft.body, args.body, bodyMode);
      patch.body = body;
      notes.push(note);
    }

    const updated = await callGraphAPI(accessToken, 'PATCH', endpoint, patch);
    const result = { ...draft, ...patch, ...(updated && updated.id ? updated : {}) };

    return text(
      `Draft updated (not sent).\n\n` +
      `Draft ID: ${result.id || args.id}\n` +
      `Subject: ${result.subject || '(no subject)'}\n` +
      `To: ${formatRecipients(result.toRecipients)}\n` +
      `Cc: ${formatRecipients(result.ccRecipients)}\n` +
      `Bcc: ${formatRecipients(result.bccRecipients)}\n` +
      `Importance: ${result.importance || 'normal'}` +
      (notes.length ? `\n${notes.join('\n')}` : '')
    );
  } catch (error) {
    if (error.message === 'Authentication required' || error.message === 'UNAUTHORIZED') {
      return text("Authentication required. Please use the 'authenticate' tool first.");
    }
    return text(`Error updating draft: ${error.message}`);
  }
}

module.exports = handleUpdateDraft;
