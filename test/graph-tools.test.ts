import { describe, expect, it } from 'vitest';
import { getTool } from '../src/tools/index.js';
import { runWithContext } from '../src/utils/context.js';
import { header, jsonResponse, mockGraph } from './helpers.js';

function callTool(name: string, args: Record<string, unknown> = {}) {
  const tool = getTool(name);
  if (!tool) throw new Error(`Unknown tool ${name}`);
  return runWithContext({ accessToken: 'graph-token', userId: 'user' }, () => tool.handler(args));
}

describe('Graph requests', () => {
  it('sends the caller token to Graph', async () => {
    const calls = mockGraph();
    await callTool('list-mail-folders');

    expect(calls).toHaveLength(1);
    expect(calls[0].url.origin).toBe('https://graph.microsoft.com');
    expect(header(calls[0], 'Authorization')).toBe('Bearer graph-token');
    expect(calls[0].init.signal).toBeInstanceOf(AbortSignal);
  });

  it('returns a tool error when Graph times out', async () => {
    mockGraph(() => {
      throw new DOMException('The operation was aborted due to timeout', 'TimeoutError');
    });
    const result = await callTool('list-mail-folders');

    expect(result.isError).toBe(true);
    expect(result.content[0].text).toContain('timeout');
  });

  it('fails without a request context', async () => {
    const calls = mockGraph();
    const result = await getTool('list-mail-folders')!.handler({});

    expect(result.isError).toBe(true);
    expect(calls).toHaveLength(0);
  });

  it('returns Graph errors as tool errors', async () => {
    mockGraph(() => jsonResponse({ error: { code: 'ErrorAccessDenied', message: 'Access is denied.' } }, 403));
    const result = await callTool('list-mail-folders');

    expect(result.isError).toBe(true);
    expect(result.content[0].text).toContain('Access is denied.');
  });

  it('strips OData metadata from responses', async () => {
    mockGraph(() => jsonResponse({ '@odata.context': 'x', value: [{ id: '1', '@odata.etag': 'y' }] }));
    const result = await callTool('list-mail-folders');

    expect(JSON.parse(result.content[0].text)).toEqual({ value: [{ id: '1' }] });
  });
});

describe('list-mail-messages', () => {
  it('lists the inbox newest first by default', async () => {
    const calls = mockGraph();
    await callTool('list-mail-messages');

    expect(calls[0].url.pathname).toBe('/v1.0/me/messages');
    expect(calls[0].url.searchParams.get('$orderby')).toBe('receivedDateTime desc');
    expect(calls[0].url.searchParams.has('$filter')).toBe(false);
  });

  it('escapes quotes in the sender filter', async () => {
    const calls = mockGraph();
    await callTool('list-mail-messages', { senderEmail: "o'brien@example.test" });

    expect(calls[0].url.searchParams.get('$filter')).toBe(
      "startswith(from/emailAddress/address, 'o''brien@example.test')",
    );
  });

  it.each([
    ['receivedAfter', { receivedAfter: '2026-01-01T00:00:00Z or isRead eq true' }],
    ['folderId', { folderId: '../users/someone-else' }],
  ])('rejects an injected %s without calling Graph', async (_label, args) => {
    const calls = mockGraph();
    const result = await callTool('list-mail-messages', args);

    expect(result.isError).toBe(true);
    expect(calls).toHaveLength(0);
  });
});

describe('get-mail-message', () => {
  it('asks Graph for a plain-text body', async () => {
    const calls = mockGraph(() => jsonResponse({ id: 'm1', uniqueBody: { contentType: 'text', content: 'hi' } }));
    await callTool('get-mail-message', { messageId: 'm1' });

    expect(calls[0].url.pathname).toBe('/v1.0/me/messages/m1');
    expect(header(calls[0], 'Prefer')).toBe('outlook.body-content-type="text"');
  });

  it('rejects a path traversal in the message id', async () => {
    const calls = mockGraph();
    const result = await callTool('get-mail-message', { messageId: '../../me/drive' });

    expect(result.isError).toBe(true);
    expect(calls).toHaveLength(0);
  });
});

describe('calendar', () => {
  it('uses calendarView when both dates are given', async () => {
    const calls = mockGraph();
    await callTool('list-calendar-events', {
      startAfter: '2026-01-19T00:00:00Z',
      startBefore: '2026-01-26T00:00:00Z',
    });

    expect(calls[0].url.pathname).toBe('/v1.0/me/calendarView');
    expect(calls[0].url.searchParams.get('startDateTime')).toBe('2026-01-19T00:00:00Z');
    expect(calls[0].url.searchParams.get('endDateTime')).toBe('2026-01-26T00:00:00Z');
  });

  it('rejects a path traversal in the calendar id', async () => {
    const calls = mockGraph();
    const result = await callTool('get-calendar-view', {
      calendarId: '../../users/x',
      startDateTime: '2026-01-19T00:00:00Z',
      endDateTime: '2026-01-26T00:00:00Z',
    });

    expect(result.isError).toBe(true);
    expect(calls).toHaveLength(0);
  });
});
