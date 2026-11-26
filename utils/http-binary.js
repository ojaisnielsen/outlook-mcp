/**
 * Raw HTTPS helpers for binary transfers that do not go through the JSON
 * Graph client: downloading a file from a pre-authenticated URL and PUTting
 * a byte range to an upload session.
 */
const https = require('https');

/**
 * Downloads a URL into a Buffer, following redirects.
 * @param {string} url
 * @param {number} maxBytes - Fail once the body grows past this size
 */
function downloadToBuffer(url, maxBytes = Infinity, redirectsLeft = 5) {
  return new Promise((resolve, reject) => {
    const req = https.get(url, res => {
      if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
        res.resume();
        if (redirectsLeft <= 0) return reject(new Error('Too many redirects'));
        const next = new URL(res.headers.location, url).toString();
        return resolve(downloadToBuffer(next, maxBytes, redirectsLeft - 1));
      }
      if (res.statusCode < 200 || res.statusCode >= 300) {
        res.resume();
        return reject(new Error(`Download failed with status ${res.statusCode}`));
      }
      const chunks = [];
      let size = 0;
      res.on('data', chunk => {
        size += chunk.length;
        if (size > maxBytes) {
          req.destroy();
          reject(new Error(`File is larger than ${maxBytes} bytes`));
          return;
        }
        chunks.push(chunk);
      });
      res.on('end', () => resolve(Buffer.concat(chunks)));
      res.on('error', reject);
    });
    req.on('error', reject);
  });
}

/**
 * PUTs one byte range to an upload session URL (no Authorization header:
 * the URL carries its own token).
 * @returns {Promise<{status: number, body: string}>}
 */
function putChunk(uploadUrl, chunk, start, totalSize) {
  return new Promise((resolve, reject) => {
    const end = start + chunk.length - 1;
    const req = https.request(uploadUrl, {
      method: 'PUT',
      headers: {
        'Content-Type': 'application/octet-stream',
        'Content-Length': chunk.length,
        'Content-Range': `bytes ${start}-${end}/${totalSize}`
      }
    }, res => {
      let body = '';
      res.on('data', d => { body += d; });
      res.on('end', () => {
        if (res.statusCode >= 200 && res.statusCode < 300) {
          resolve({ status: res.statusCode, body });
        } else {
          reject(new Error(`Upload failed at byte ${start} with status ${res.statusCode}: ${body}`));
        }
      });
    });
    req.on('error', reject);
    req.write(chunk);
    req.end();
  });
}

module.exports = { downloadToBuffer, putChunk };
