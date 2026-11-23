const { buildSearchParams, kqlField } = require('../../email/search');

describe('KQL quoting for $search', () => {
  test('backslash-escapes the quotes around field values', () => {
    expect(kqlField('from', 'squarespace')).toBe('from:\\"squarespace\\"');
  });

  test('strips quotes and backslashes from user values', () => {
    expect(kqlField('subject', 'say "hi" \\ there')).toBe('subject:\\"say hi  there\\"');
  });

  test('builds a valid combined $search value', () => {
    const params = buildSearchParams({ from: 'squarespace', subject: 'domain renewal' }, {}, 5);
    expect(params.$search).toBe('"subject:\\"domain renewal\\" from:\\"squarespace\\""');
    expect(params.$orderby).toBeUndefined();
  });
});
