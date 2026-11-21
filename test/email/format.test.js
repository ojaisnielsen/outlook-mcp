const {
  buildMessageFilter,
  matchesMessageFilter,
  formatEmailList,
  clampCount
} = require('../../email/format');

describe('buildMessageFilter', () => {
  test('returns empty string with no filters', () => {
    expect(buildMessageFilter({})).toBe('');
  });

  test('leads with receivedDateTime so $orderby receivedDateTime is accepted', () => {
    expect(buildMessageFilter({ unreadOnly: true }))
      .toBe('receivedDateTime ge 1900-01-01T00:00:00Z and isRead eq false');
  });

  test('combines dates, unread and flag status', () => {
    expect(buildMessageFilter({
      receivedAfter: '2026-11-20T06:00:00Z',
      receivedBefore: '2026-11-21',
      unreadOnly: true,
      flagStatus: 'notFlagged'
    })).toBe(
      'receivedDateTime ge 2026-11-20T06:00:00.000Z and receivedDateTime lt 2026-11-21T00:00:00.000Z' +
      " and isRead eq false and flag/flagStatus eq 'notFlagged'"
    );
  });

  test('rejects invalid dates and flag statuses', () => {
    expect(() => buildMessageFilter({ receivedAfter: 'yesterday-ish' })).toThrow('receivedAfter');
    expect(() => buildMessageFilter({ flagStatus: 'red' })).toThrow('flagStatus');
  });
});

describe('matchesMessageFilter', () => {
  const email = {
    receivedDateTime: '2026-11-20T10:00:00Z',
    isRead: false,
    hasAttachments: false,
    flag: { flagStatus: 'flagged' }
  };

  test('applies date, unread and flag filters', () => {
    expect(matchesMessageFilter(email, { receivedAfter: '2026-11-20T09:00:00Z', unreadOnly: true })).toBe(true);
    expect(matchesMessageFilter(email, { receivedAfter: '2026-11-20T11:00:00Z' })).toBe(false);
    expect(matchesMessageFilter(email, { flagStatus: 'notFlagged' })).toBe(false);
    expect(matchesMessageFilter({ ...email, flag: undefined }, { flagStatus: 'notFlagged' })).toBe(true);
    expect(matchesMessageFilter(email, { hasAttachments: true })).toBe(false);
  });
});

describe('formatEmailList', () => {
  test('shows ISO date, markers, preview and ID', () => {
    const text = formatEmailList([{
      id: 'abc',
      subject: 'Invoice',
      from: { emailAddress: { name: 'Shop', address: 'shop@example.com' } },
      receivedDateTime: '2026-11-20T10:00:00Z',
      isRead: false,
      hasAttachments: true,
      flag: { flagStatus: 'flagged' },
      bodyPreview: 'Your   invoice\n is attached'
    }]);
    expect(text).toContain('1. [UNREAD] [FLAGGED] [ATTACHMENT] 2026-11-20T10:00:00.000Z - From: Shop (shop@example.com)');
    expect(text).toContain('Preview: Your invoice is attached');
    expect(text).toContain('ID: abc');
  });

  test('can omit preview and truncates long previews', () => {
    const email = { id: 'x', subject: 's', receivedDateTime: '2026-11-20T10:00:00Z', isRead: true, bodyPreview: 'a'.repeat(500) };
    expect(formatEmailList([email], { includePreview: false })).not.toContain('Preview:');
    expect(formatEmailList([email])).toContain(`${'a'.repeat(200)}…`);
  });
});

describe('clampCount', () => {
  test('defaults and caps', () => {
    expect(clampCount(undefined)).toBe(10);
    expect(clampCount(25)).toBe(25);
    expect(clampCount(10000)).toBe(500);
    expect(clampCount(-3)).toBe(10);
  });
});
