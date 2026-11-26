/**
 * add-attachment: attaches a file to an existing draft. The file comes
 * from a local path on the machine running this server or from OneDrive.
 * Files under 3 MB are sent in one request; larger ones (up to 150 MB)
 * through an upload session.
 */
const fs = require('fs');
const path = require('path');
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');
const { downloadToBuffer, putChunk } = require('../utils/http-binary');

const SMALL_LIMIT = 3 * 1024 * 1024;      // single POST (base64 must stay under 4 MB)
const MAX_SIZE = 150 * 1024 * 1024;       // Outlook's attachment limit
const CHUNK_SIZE = 320 * 1024 * 10;       // 3.125 MB per PUT (must be < 4 MB)

const MIME_TYPES = {
  '.pdf': 'application/pdf',
  '.txt': 'text/plain',
  '.csv': 'text/csv',
  '.html': 'text/html',
  '.htm': 'text/html',
  '.json': 'application/json',
  '.xml': 'application/xml',
  '.zip': 'application/zip',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.gif': 'image/gif',
  '.webp': 'image/webp',
  '.heic': 'image/heic',
  '.svg': 'image/svg+xml',
  '.doc': 'application/msword',
  '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  '.xls': 'application/vnd.ms-excel',
  '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  '.ppt': 'application/vnd.ms-powerpoint',
  '.pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  '.ics': 'text/calendar',
  '.eml': 'message/rfc822',
  '.mp3': 'audio/mpeg',
  '.mp4': 'video/mp4'
};

function text(message) {
  return { content: [{ type: "text", text: message }] };
}

function mimeTypeFor(name) {
  return MIME_TYPES[path.extname(name).toLowerCase()] || 'application/octet-stream';
}

function formatSize(bytes) {
  if (!bytes) return '0 B';
  const units = ['B', 'KB', 'MB', 'GB'];
  const i = Math.min(Math.floor(Math.log(bytes) / Math.log(1024)), units.length - 1);
  return `${parseFloat((bytes / Math.pow(1024, i)).toFixed(1))} ${units[i]}`;
}

/**
 * Directories local files may come from: OUTLOOK_MCP_ATTACHMENT_DIRS
 * (separated by the platform path delimiter), or null for no restriction.
 */
function allowedRoots() {
  const value = process.env.OUTLOOK_MCP_ATTACHMENT_DIRS;
  if (!value) return null;
  return value.split(path.delimiter).filter(Boolean).map(dir => {
    try { return fs.realpathSync(dir); } catch { return path.resolve(dir); }
  });
}

/**
 * Reads a local file after checking it is a regular file, not hidden or
 * inside a hidden directory (tokens, keys, .env), and inside an allowed
 * directory when OUTLOOK_MCP_ATTACHMENT_DIRS is set.
 */
function readLocalFile(filePath) {
  if (!path.isAbsolute(filePath)) {
    throw new Error('filePath must be an absolute path.');
  }
  let real;
  try {
    real = fs.realpathSync(filePath);
  } catch {
    throw new Error(`File not found: ${filePath}`);
  }
  if (real.split(path.sep).some(part => part.startsWith('.'))) {
    throw new Error('Hidden files and files inside hidden directories cannot be attached.');
  }
  const roots = allowedRoots();
  if (roots && !roots.some(root => real === root || real.startsWith(root + path.sep))) {
    throw new Error('That file is outside the directories allowed by OUTLOOK_MCP_ATTACHMENT_DIRS.');
  }
  const stat = fs.statSync(real);
  if (!stat.isFile()) throw new Error(`Not a regular file: ${filePath}`);
  if (stat.size > MAX_SIZE) {
    throw new Error(`File is ${formatSize(stat.size)}; Outlook attachments are limited to ${formatSize(MAX_SIZE)}.`);
  }
  return { content: fs.readFileSync(real), name: path.basename(real), contentType: mimeTypeFor(real) };
}

async function readOneDriveFile(accessToken, { oneDrivePath, oneDriveItemId }) {
  const endpoint = oneDriveItemId
    ? `me/drive/items/${oneDriveItemId}`
    : `me/drive/root:/${oneDrivePath.replace(/^\/+|\/+$/g, '')}`;
  // No $select: personal OneDrive omits @microsoft.graph.downloadUrl when
  // it is requested through $select.
  const item = await callGraphAPI(accessToken, 'GET', endpoint);
  if (!item || !item.name) throw new Error('OneDrive file not found.');
  if (item.folder) throw new Error(`"${item.name}" is a folder, not a file.`);
  if (item.size > MAX_SIZE) {
    throw new Error(`"${item.name}" is ${formatSize(item.size)}; Outlook attachments are limited to ${formatSize(MAX_SIZE)}.`);
  }
  const downloadUrl = item['@microsoft.graph.downloadUrl'];
  if (!downloadUrl) throw new Error(`Could not get a download URL for "${item.name}".`);
  const content = await downloadToBuffer(downloadUrl, MAX_SIZE);
  return {
    content,
    name: item.name,
    contentType: (item.file && item.file.mimeType) || mimeTypeFor(item.name)
  };
}

async function uploadLarge(accessToken, messageId, { content, name, contentType }) {
  const session = await callGraphAPI(
    accessToken,
    'POST',
    `me/messages/${encodeURIComponent(messageId)}/attachments/createUploadSession`,
    { AttachmentItem: { attachmentType: 'file', name, size: content.length, contentType } }
  );
  if (!session || !session.uploadUrl) throw new Error('Could not create an attachment upload session.');

  for (let offset = 0; offset < content.length; offset += CHUNK_SIZE) {
    const chunk = content.subarray(offset, Math.min(offset + CHUNK_SIZE, content.length));
    await putChunk(session.uploadUrl, chunk, offset, content.length);
  }
}

/**
 * @param {object} args
 * @param {string} args.id - Draft ID
 * @param {string} [args.filePath] - Absolute path of a local file
 * @param {string} [args.oneDrivePath] - OneDrive path, e.g. '/Documents/report.pdf'
 * @param {string} [args.oneDriveItemId] - OneDrive item ID
 * @param {string} [args.name] - Attachment name (defaults to the file name)
 */
async function handleAddAttachment(args = {}) {
  if (!args.id) return text("Draft ID is required (id).");
  const sources = ['filePath', 'oneDrivePath', 'oneDriveItemId'].filter(k => args[k]);
  if (sources.length !== 1) {
    return text("Give exactly one source: filePath, oneDrivePath or oneDriveItemId.");
  }

  try {
    const accessToken = await ensureAuthenticated();
    const endpoint = `me/messages/${encodeURIComponent(args.id)}`;
    const draft = await callGraphAPI(accessToken, 'GET', endpoint, null, { $select: 'id,isDraft,subject' });
    if (draft.isDraft !== true) {
      return text("That message is not a draft; nothing was attached.");
    }

    let file;
    try {
      file = args.filePath
        ? readLocalFile(args.filePath)
        : await readOneDriveFile(accessToken, args);
    } catch (error) {
      if (error.message === 'UNAUTHORIZED') throw error;
      return text(`Could not read the file: ${error.message}`);
    }
    if (args.name) file.name = args.name;
    if (file.content.length === 0) return text("The file is empty; nothing was attached.");

    if (file.content.length < SMALL_LIMIT) {
      await callGraphAPI(accessToken, 'POST', `${endpoint}/attachments`, {
        '@odata.type': '#microsoft.graph.fileAttachment',
        name: file.name,
        contentType: file.contentType,
        contentBytes: file.content.toString('base64')
      });
    } else {
      await uploadLarge(accessToken, args.id, file);
    }

    let listing = '';
    try {
      const attachments = await callGraphAPI(accessToken, 'GET', `${endpoint}/attachments`, null, { $select: 'name,size' });
      if (attachments && Array.isArray(attachments.value)) {
        listing = `\n\nAttachments on this draft (${attachments.value.length}):\n` +
          attachments.value.map(a => `- ${a.name} (${formatSize(a.size)})`).join('\n');
      }
    } catch {
      // The attachment is in place; the listing is only a convenience
    }

    return text(
      `Attached "${file.name}" (${formatSize(file.content.length)}) to the draft "${draft.subject || '(no subject)'}" (not sent).` +
      listing
    );
  } catch (error) {
    if (error.message === 'Authentication required' || error.message === 'UNAUTHORIZED') {
      return text("Authentication required. Please use the 'authenticate' tool first.");
    }
    return text(`Error adding attachment: ${error.message}`);
  }
}

module.exports = handleAddAttachment;
module.exports.readLocalFile = readLocalFile;
module.exports.SMALL_LIMIT = SMALL_LIMIT;
module.exports.CHUNK_SIZE = CHUNK_SIZE;
