const handleFlagEmail = require('../../email/flag');
const { handleDraftReply, handleSendDraft } = require('../../email/reply');
const { callGraphAPI } = require('../../utils/graph-api');
const { ensureAuthenticated } = require('../../auth');

jest.mock('../../utils/graph-api');
jest.mock('../../auth');

beforeEach(() => {
  jest.resetAllMocks();
  ensureAuthenticated.mockResolvedValue('token');
  jest.spyOn(console, 'error').mockImplementation(() => {});
});

afterEach(() => console.error.mockRestore());

describe('flag-email', () => {
  test('flags each ID with a PATCH', async () => {
    callGraphAPI.mockResolvedValue({});
    const result = await handleFlagEmail({ ids: 'a, b' });
    expect(callGraphAPI).toHaveBeenCalledTimes(2);
    expect(callGraphAPI).toHaveBeenCalledWith('token', 'PATCH', 'me/messages/a', { flag: { flagStatus: 'flagged' } });
    expect(result.content[0].text).toBe('2 email(s) flagged.');
  });

  test('unflags and reports partial failures', async () => {
    callGraphAPI.mockResolvedValueOnce({}).mockRejectedValueOnce(new Error('404'));
    const result = await handleFlagEmail({ ids: 'a,b', status: 'notFlagged' });
    expect(result.content[0].text).toContain('1 email(s) unflagged.');
    expect(result.content[0].text).toContain('Failed for 1 email(s)');
  });

  test('rejects an unknown status without calling the API', async () => {
    const result = await handleFlagEmail({ ids: 'a', status: 'red' });
    expect(result.content[0].text).toContain('status must be one of');
    expect(callGraphAPI).not.toHaveBeenCalled();
  });
});

describe('draft-reply', () => {
  test('creates a reply-all draft by default and shows recipients', async () => {
    callGraphAPI.mockResolvedValue({
      id: 'draft1',
      subject: 'RE: Hello',
      toRecipients: [{ emailAddress: { name: 'Ann', address: 'ann@example.com' } }],
      ccRecipients: []
    });
    const result = await handleDraftReply({ id: 'm1', body: 'Thanks!' });
    expect(callGraphAPI).toHaveBeenCalledWith('token', 'POST', 'me/messages/m1/createReplyAll', { comment: '' });
    expect(result.content[0].text).toContain('not sent');
    expect(result.content[0].text).toContain('Ann <ann@example.com>');
  });

  test('writes the body above the quoted thread, keeping line breaks', async () => {
    callGraphAPI.mockResolvedValueOnce({
      id: 'draft1',
      subject: 'RE: Hello',
      body: {
        contentType: 'html',
        content: '<html><body> <hr tabindex="-1"><div id="divRplyFwdMsg"><b>From:</b> Ann</div><div>Hi</div></body></html>'
      }
    }).mockResolvedValueOnce({});
    await handleDraftReply({ id: 'm1', body: 'Line 1\nLine 2 <ok>' });
    const [, method, path, patch] = callGraphAPI.mock.calls[1];
    expect(method).toBe('PATCH');
    expect(path).toBe('me/messages/draft1');
    expect(patch.body.contentType).toBe('html');
    expect(patch.body.content).toBe(
      '<html><body>Line 1<br>Line 2 &lt;ok&gt;<br><hr tabindex="-1"><div id="divRplyFwdMsg"><b>From:</b> Ann</div><div>Hi</div></body></html>'
    );
  });

  test('replyAll: false creates a plain reply', async () => {
    callGraphAPI.mockResolvedValue({ id: 'd', subject: 'RE: x' });
    await handleDraftReply({ id: 'm1', replyAll: false });
    expect(callGraphAPI).toHaveBeenCalledWith('token', 'POST', 'me/messages/m1/createReply', { comment: '' });
  });
});

describe('send-draft', () => {
  test('refuses a message that is not a draft', async () => {
    callGraphAPI.mockResolvedValueOnce({ id: 'm1', isDraft: false });
    const result = await handleSendDraft({ id: 'm1' });
    expect(result.content[0].text).toContain('not a draft');
    expect(callGraphAPI).toHaveBeenCalledTimes(1);
  });

  test('sends a draft', async () => {
    callGraphAPI
      .mockResolvedValueOnce({ id: 'd1', isDraft: true, subject: 'RE: x', toRecipients: [] })
      .mockResolvedValueOnce({});
    const result = await handleSendDraft({ id: 'd1' });
    expect(callGraphAPI).toHaveBeenLastCalledWith('token', 'POST', 'me/messages/d1/send');
    expect(result.content[0].text).toContain('Draft sent.');
  });
});
