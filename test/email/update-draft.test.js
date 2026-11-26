const handleUpdateDraft = require('../../email/update-draft');
const { splitHtmlBody, buildBody, mergeRecipients, parseRecipients } = require('../../email/draft-utils');
const { callGraphAPI } = require('../../utils/graph-api');
const { ensureAuthenticated } = require('../../auth');

jest.mock('../../utils/graph-api');
jest.mock('../../auth');

const REPLY_HTML =
  '<html><head><meta charset="utf-8"></head><body>Old text <hr tabindex="-1" style="display:inline-block">' +
  '<div id="divRplyFwdMsg" dir="ltr"><b>From:</b> Ann</div><div>Original message</div></body></html>';

function draft(overrides = {}) {
  return {
    id: 'd1',
    isDraft: true,
    subject: 'RE: Hello',
    importance: 'normal',
    body: { contentType: 'html', content: REPLY_HTML },
    toRecipients: [{ emailAddress: { name: 'Ann', address: 'ann@example.com' } }],
    ccRecipients: [],
    bccRecipients: [],
    ...overrides
  };
}

beforeEach(() => {
  jest.resetAllMocks();
  ensureAuthenticated.mockResolvedValue('token');
  jest.spyOn(console, 'error').mockImplementation(() => {});
});

afterEach(() => console.error.mockRestore());

describe('draft-utils', () => {
  test('splitHtmlBody separates the quoted thread including its <hr>', () => {
    const { head, own, quote } = splitHtmlBody(REPLY_HTML);
    expect(head.endsWith('<body>')).toBe(true);
    expect(own).toBe('Old text ');
    expect(quote.startsWith('<hr tabindex="-1"')).toBe(true);
    expect(quote).toContain('Original message');
  });

  test('splitHtmlBody without a quote returns everything as own text', () => {
    const { quote, own } = splitHtmlBody('<p>Hi</p>');
    expect(quote).toBeNull();
    expect(own).toBe('<p>Hi</p>');
  });

  test('auto mode keeps the quote; replace drops it; prepend keeps old text', () => {
    const current = { contentType: 'html', content: REPLY_HTML };
    const auto = buildBody(current, 'New', 'auto').body.content;
    expect(auto).toContain('<body>New<br><hr');
    expect(auto).not.toContain('Old text');
    expect(auto).toContain('Original message');

    expect(buildBody(current, 'New', 'replace').body.content).toBe('New');

    const prepend = buildBody(current, 'New', 'prepend').body.content;
    expect(prepend).toContain('<body>New<br><br>Old text');
    expect(prepend).toContain('Original message');
  });

  test('plain-text drafts stay plain text', () => {
    const current = { contentType: 'text', content: 'Old' };
    expect(buildBody(current, 'New', 'auto').body).toEqual({ contentType: 'text', content: 'New' });
    expect(buildBody(current, 'New', 'prepend').body).toEqual({ contentType: 'text', content: 'New\n\nOld' });
  });

  test('HTML input into a plain-text draft switches it to HTML', () => {
    const result = buildBody({ contentType: 'text', content: 'a < b' }, '<p>Hi</p>', 'prepend').body;
    expect(result.contentType).toBe('html');
    expect(result.content).toBe('<p>Hi</p><br><br>a &lt; b');
  });

  test('HTML input is stripped of document wrappers', () => {
    const result = buildBody({ contentType: 'html', content: REPLY_HTML }, '<html><head><title>x</title></head><body><p>Hi</p></body></html>', 'auto');
    expect(result.body.content).toContain('<body><p>Hi</p><br><hr');
  });

  test('mergeRecipients skips duplicates regardless of case', () => {
    const merged = mergeRecipients(parseRecipients('a@x.com'), parseRecipients('A@x.com, b@x.com'));
    expect(merged.map(r => r.emailAddress.address)).toEqual(['a@x.com', 'b@x.com']);
  });
});

describe('update-draft', () => {
  test('requires an id and at least one change', async () => {
    expect((await handleUpdateDraft({})).content[0].text).toContain('Draft ID is required');
    expect((await handleUpdateDraft({ id: 'd1' })).content[0].text).toContain('Nothing to change');
    expect(callGraphAPI).not.toHaveBeenCalled();
  });

  test('rejects an unknown bodyMode', async () => {
    const result = await handleUpdateDraft({ id: 'd1', body: 'x', bodyMode: 'append' });
    expect(result.content[0].text).toContain('bodyMode must be one of');
    expect(callGraphAPI).not.toHaveBeenCalled();
  });

  test('refuses a message that is not a draft', async () => {
    callGraphAPI.mockResolvedValueOnce(draft({ isDraft: false }));
    const result = await handleUpdateDraft({ id: 'm1', subject: 'x' });
    expect(result.content[0].text).toContain('not a draft');
    expect(callGraphAPI).toHaveBeenCalledTimes(1);
  });

  test('patches only the fields given', async () => {
    callGraphAPI.mockResolvedValueOnce(draft()).mockResolvedValueOnce({});
    const result = await handleUpdateDraft({ id: 'd1', subject: 'New subject', addCc: 'bob@example.com' });
    expect(callGraphAPI).toHaveBeenLastCalledWith('token', 'PATCH', 'me/messages/d1', {
      subject: 'New subject',
      ccRecipients: [{ emailAddress: { address: 'bob@example.com' } }]
    });
    expect(result.content[0].text).toContain('Draft updated (not sent)');
    expect(result.content[0].text).toContain('Cc: bob@example.com');
    expect(result.content[0].text).toContain('To: Ann <ann@example.com>');
  });

  test('to replaces the list and addTo appends to it', async () => {
    callGraphAPI.mockResolvedValueOnce(draft()).mockResolvedValueOnce({});
    await handleUpdateDraft({ id: 'd1', to: 'c@example.com', addTo: 'd@example.com' });
    const patch = callGraphAPI.mock.calls[1][3];
    expect(patch.toRecipients.map(r => r.emailAddress.address)).toEqual(['c@example.com', 'd@example.com']);
  });

  test('empty to clears the list', async () => {
    callGraphAPI.mockResolvedValueOnce(draft()).mockResolvedValueOnce({});
    await handleUpdateDraft({ id: 'd1', to: '' });
    expect(callGraphAPI.mock.calls[1][3]).toEqual({ toRecipients: [] });
  });

  test('body in auto mode keeps the quoted thread of a reply', async () => {
    callGraphAPI.mockResolvedValueOnce(draft()).mockResolvedValueOnce({});
    const result = await handleUpdateDraft({ id: 'd1', body: 'Hello\nthere' });
    const patch = callGraphAPI.mock.calls[1][3];
    expect(patch.body.content).toContain('<body>Hello<br>there<br><hr');
    expect(patch.body.content).toContain('Original message');
    expect(result.content[0].text).toContain('quoted thread kept');
  });

  test('reports Graph errors', async () => {
    callGraphAPI.mockResolvedValueOnce(draft()).mockRejectedValueOnce(new Error('API call failed with status 400: bad'));
    const result = await handleUpdateDraft({ id: 'd1', subject: 'x' });
    expect(result.content[0].text).toContain('Error updating draft: API call failed with status 400');
  });
});
