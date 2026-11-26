/**
 * Helpers shared by the draft tools: recipient lists and body editing.
 */

/**
 * Turns "a@x.com, b@y.com" into Graph recipients. An empty string gives [].
 */
function parseRecipients(value) {
  if (value === undefined || value === null) return [];
  return String(value)
    .split(',')
    .map(address => address.trim())
    .filter(Boolean)
    .map(address => ({ emailAddress: { address } }));
}

/**
 * Appends recipients that are not already in the list (case-insensitive).
 */
function mergeRecipients(existing = [], added = []) {
  const seen = new Set(existing.map(r => (r.emailAddress?.address || '').toLowerCase()));
  const merged = [...existing];
  for (const r of added) {
    const key = r.emailAddress.address.toLowerCase();
    if (!seen.has(key)) {
      seen.add(key);
      merged.push(r);
    }
  }
  return merged;
}

function formatRecipients(list) {
  if (!list || list.length === 0) return 'None';
  return list
    .map(r => {
      const { name, address } = r.emailAddress;
      return name && name !== address ? `${name} <${address}>` : address;
    })
    .join(', ');
}

const HTML_TAG = /<\/?(html|body|p|div|br|span|a|b|i|u|strong|em|ul|ol|li|table|h[1-6])\b[^>]*>/i;

function looksLikeHtml(text) {
  return typeof text === 'string' && HTML_TAG.test(text);
}

function escapeHtml(text) {
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/**
 * Plain text to an HTML fragment that keeps line breaks.
 */
function textToHtml(text) {
  return escapeHtml(text).replace(/\r\n|\r|\n/g, '<br>');
}

/**
 * HTML for a piece of new body text: kept as-is when it already is HTML,
 * otherwise escaped with line breaks preserved.
 */
function toHtmlFragment(text) {
  if (!looksLikeHtml(text)) return textToHtml(text);
  // Drop document wrappers so the fragment can sit inside an existing body
  return String(text)
    .replace(/<!DOCTYPE[^>]*>/i, '')
    .replace(/<head[\s\S]*?<\/head>/i, '')
    .replace(/<\/?(html|body)\b[^>]*>/gi, '')
    .trim();
}

/**
 * Splits an HTML body into the part up to and including <body>, the
 * author's text, and the quoted thread Outlook adds to replies and
 * forwards (an <hr> followed by <div id="divRplyFwdMsg">). quote is null
 * when there is no quoted thread.
 */
function splitHtmlBody(html) {
  const content = html || '';
  const bodyTag = content.match(/<body\b[^>]*>/i);
  const start = bodyTag ? bodyTag.index + bodyTag[0].length : 0;
  const head = content.slice(0, start);
  const rest = content.slice(start);

  const marker = rest.search(/<div\b[^>]*\bid=["']?(divRplyFwdMsg|appendonsend)\b/i);
  if (marker === -1) {
    return { head, own: rest, quote: null };
  }
  // Include the separator line Outlook puts just before the header block
  let cut = marker;
  const hr = rest.lastIndexOf('<hr', marker);
  if (hr !== -1 && marker - hr < 300) cut = hr;
  return { head, own: rest.slice(0, cut), quote: rest.slice(cut) };
}

/**
 * Builds a new draft body.
 * @param {{contentType: string, content: string}} current - The draft's body
 * @param {string} text - New text (plain or HTML)
 * @param {string} mode - 'auto' (replace the text above a quoted thread, or
 *   the whole body when there is none), 'replace' (whole body, quote
 *   included) or 'prepend' (insert above the existing body)
 * @returns {{body: {contentType: string, content: string}, note: string}}
 */
function buildBody(current, text, mode = 'auto') {
  const contentType = current?.contentType === 'html' ? 'html' : 'text';
  const content = current?.content || '';

  if (contentType === 'text' && !looksLikeHtml(text)) {
    if (mode === 'prepend' && content) {
      return { body: { contentType: 'text', content: `${text}\n\n${content}` }, note: 'New text added above the existing body.' };
    }
    return { body: { contentType: 'text', content: text }, note: 'Body replaced.' };
  }

  // HTML draft, or HTML text going into a plain-text draft
  const fragment = toHtmlFragment(text);
  const existingHtml = contentType === 'html' ? content : textToHtml(content);

  if (mode === 'replace') {
    return { body: { contentType: 'html', content: fragment }, note: 'Whole body replaced (any quoted thread removed).' };
  }

  const { head, own, quote } = splitHtmlBody(existingHtml);
  if (mode === 'prepend') {
    const separator = own.trim() || quote ? '<br><br>' : '';
    return {
      body: { contentType: 'html', content: `${head}${fragment}${separator}${own}${quote || ''}` },
      note: 'New text added above the existing body.'
    };
  }

  // auto
  if (quote) {
    return {
      body: { contentType: 'html', content: `${head}${fragment}<br>${quote}` },
      note: 'Text above the quoted thread replaced; quoted thread kept.'
    };
  }
  return { body: { contentType: 'html', content: fragment }, note: 'Body replaced.' };
}

module.exports = {
  parseRecipients,
  mergeRecipients,
  formatRecipients,
  looksLikeHtml,
  escapeHtml,
  textToHtml,
  toHtmlFragment,
  splitHtmlBody,
  buildBody
};
