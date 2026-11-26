const fs = require('fs');
const os = require('os');
const path = require('path');
const handleAddAttachment = require('../../email/attachments');
const { readLocalFile, SMALL_LIMIT, CHUNK_SIZE } = handleAddAttachment;
const { callGraphAPI } = require('../../utils/graph-api');
const { ensureAuthenticated } = require('../../auth');
const { downloadToBuffer, putChunk } = require('../../utils/http-binary');

jest.mock('../../utils/graph-api');
jest.mock('../../auth');
jest.mock('../../utils/http-binary');

let dir;

beforeAll(() => {
  dir = fs.mkdtempSync(path.join(os.tmpdir(), 'attach-test-'));
  fs.writeFileSync(path.join(dir, 'report.pdf'), 'PDFDATA');
  fs.mkdirSync(path.join(dir, '.secret'));
  fs.writeFileSync(path.join(dir, '.secret', 'token.json'), '{}');
  fs.writeFileSync(path.join(dir, '.env'), 'X=1');
  fs.writeFileSync(path.join(dir, 'empty.txt'), '');
});

afterAll(() => fs.rmSync(dir, { recursive: true, force: true }));

beforeEach(() => {
  jest.resetAllMocks();
  ensureAuthenticated.mockResolvedValue('token');
  delete process.env.OUTLOOK_MCP_ATTACHMENT_DIRS;
  jest.spyOn(console, 'error').mockImplementation(() => {});
});

afterEach(() => console.error.mockRestore());

// os.tmpdir() may itself be under a hidden directory on some machines
const tmpIsHidden = fs.realpathSync(os.tmpdir()).split(path.sep).some(p => p.startsWith('.'));
const maybe = tmpIsHidden ? test.skip : test;

describe('readLocalFile', () => {
  maybe('reads a file with its name and MIME type', () => {
    const file = readLocalFile(path.join(dir, 'report.pdf'));
    expect(file.name).toBe('report.pdf');
    expect(file.contentType).toBe('application/pdf');
    expect(file.content.toString()).toBe('PDFDATA');
  });

  test('refuses relative paths, hidden files and hidden directories', () => {
    expect(() => readLocalFile('report.pdf')).toThrow('absolute');
    expect(() => readLocalFile(path.join(dir, '.env'))).toThrow('Hidden');
    expect(() => readLocalFile(path.join(dir, '.secret', 'token.json'))).toThrow('Hidden');
  });

  test('refuses directories and missing files', () => {
    expect(() => readLocalFile(path.join(dir, 'nope.pdf'))).toThrow('File not found');
  });

  maybe('honours OUTLOOK_MCP_ATTACHMENT_DIRS', () => {
    process.env.OUTLOOK_MCP_ATTACHMENT_DIRS = path.join(dir, 'elsewhere');
    expect(() => readLocalFile(path.join(dir, 'report.pdf'))).toThrow('outside');
    process.env.OUTLOOK_MCP_ATTACHMENT_DIRS = dir;
    expect(readLocalFile(path.join(dir, 'report.pdf')).name).toBe('report.pdf');
  });
});

describe('add-attachment', () => {
  test('needs exactly one source', async () => {
    expect((await handleAddAttachment({ id: 'd1' })).content[0].text).toContain('exactly one source');
    expect((await handleAddAttachment({ id: 'd1', filePath: '/a', oneDrivePath: '/b' })).content[0].text).toContain('exactly one source');
    expect(callGraphAPI).not.toHaveBeenCalled();
  });

  test('refuses a message that is not a draft', async () => {
    callGraphAPI.mockResolvedValueOnce({ id: 'm1', isDraft: false });
    const result = await handleAddAttachment({ id: 'm1', filePath: path.join(dir, 'report.pdf') });
    expect(result.content[0].text).toContain('not a draft');
    expect(callGraphAPI).toHaveBeenCalledTimes(1);
  });

  maybe('attaches a small local file in one request and lists attachments', async () => {
    callGraphAPI
      .mockResolvedValueOnce({ id: 'd1', isDraft: true, subject: 'Report' })
      .mockResolvedValueOnce({})
      .mockResolvedValueOnce({ value: [{ name: 'Q3.pdf', size: 7 }] });
    const result = await handleAddAttachment({ id: 'd1', filePath: path.join(dir, 'report.pdf'), name: 'Q3.pdf' });
    expect(callGraphAPI).toHaveBeenNthCalledWith(2, 'token', 'POST', 'me/messages/d1/attachments', {
      '@odata.type': '#microsoft.graph.fileAttachment',
      name: 'Q3.pdf',
      contentType: 'application/pdf',
      contentBytes: Buffer.from('PDFDATA').toString('base64')
    });
    expect(result.content[0].text).toContain('Attached "Q3.pdf"');
    expect(result.content[0].text).toContain('- Q3.pdf');
  });

  maybe('refuses an empty file', async () => {
    callGraphAPI.mockResolvedValueOnce({ id: 'd1', isDraft: true });
    const result = await handleAddAttachment({ id: 'd1', filePath: path.join(dir, 'empty.txt') });
    expect(result.content[0].text).toContain('empty');
    expect(callGraphAPI).toHaveBeenCalledTimes(1);
  });

  test('reports a hidden file without calling the attachment API', async () => {
    callGraphAPI.mockResolvedValueOnce({ id: 'd1', isDraft: true });
    const result = await handleAddAttachment({ id: 'd1', filePath: path.join(dir, '.env') });
    expect(result.content[0].text).toContain('Could not read the file: Hidden');
    expect(callGraphAPI).toHaveBeenCalledTimes(1);
  });

  test('copies a large OneDrive file through an upload session in chunks', async () => {
    const size = 2 * CHUNK_SIZE + 10; // > SMALL_LIMIT, so three PUTs
    callGraphAPI
      .mockResolvedValueOnce({ id: 'd1', isDraft: true, subject: 'Big' })
      .mockResolvedValueOnce({ name: 'video.mp4', size, file: { mimeType: 'video/mp4' }, '@microsoft.graph.downloadUrl': 'https://dl/x' })
      .mockResolvedValueOnce({ uploadUrl: 'https://upload/session' })
      .mockResolvedValueOnce({ value: [] });
    downloadToBuffer.mockResolvedValue(Buffer.alloc(size));
    putChunk.mockResolvedValue({ status: 200, body: '' });

    const result = await handleAddAttachment({ id: 'd1', oneDrivePath: '/Videos/video.mp4' });

    expect(callGraphAPI).toHaveBeenNthCalledWith(2, 'token', 'GET', 'me/drive/root:/Videos/video.mp4');
    expect(callGraphAPI).toHaveBeenNthCalledWith(3, 'token', 'POST', 'me/messages/d1/attachments/createUploadSession', {
      AttachmentItem: { attachmentType: 'file', name: 'video.mp4', size, contentType: 'video/mp4' }
    });
    const offsets = putChunk.mock.calls.map(c => c[2]);
    expect(offsets).toEqual([0, CHUNK_SIZE, 2 * CHUNK_SIZE]);
    expect(putChunk.mock.calls.reduce((n, c) => n + c[1].length, 0)).toBe(size);
    expect(result.content[0].text).toContain('Attached "video.mp4"');
  });

  test('refuses a OneDrive folder', async () => {
    callGraphAPI
      .mockResolvedValueOnce({ id: 'd1', isDraft: true })
      .mockResolvedValueOnce({ name: 'Docs', folder: {} });
    const result = await handleAddAttachment({ id: 'd1', oneDriveItemId: 'abc' });
    expect(result.content[0].text).toContain('is a folder');
    expect(downloadToBuffer).not.toHaveBeenCalled();
  });
});
