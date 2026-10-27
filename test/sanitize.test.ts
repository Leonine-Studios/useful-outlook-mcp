import { describe, expect, it } from 'vitest';
import {
  sanitizeODataDatetime,
  sanitizeODataString,
  sanitizePathSegment,
  sanitizeTimezone,
} from '../src/utils/sanitize.js';

describe('sanitizePathSegment', () => {
  it.each([
    'AAMkAGI2TG93AAA=',
    'AQMkADAwATM0MDAAMS1iNTcwLWI2NTEtMDACLTAwCgBGAAAD-_+.',
  ])('accepts Graph id %s', id => {
    expect(sanitizePathSegment(id)).toBe(id);
  });

  it.each([
    '',
    '../me',
    'a..b',
    'abc/def',
    'abc?$top=999',
    'abc#x',
    'abc&x=1',
    'abc def',
    'abc%2F',
  ])('rejects %j', value => {
    expect(() => sanitizePathSegment(value, 'messageId')).toThrow(/Invalid messageId/);
  });
});

describe('sanitizeODataString', () => {
  it('doubles single quotes', () => {
    expect(sanitizeODataString("o'brien' or 1 eq 1")).toBe("o''brien'' or 1 eq 1");
  });
});

describe('sanitizeODataDatetime', () => {
  it.each([
    '2026-01-20',
    '2026-01-20T10:00',
    '2026-01-20T10:00:00Z',
    '2026-01-20T10:00:00.123+02:00',
  ])('accepts %s', value => {
    expect(sanitizeODataDatetime(value)).toBe(value);
  });

  it.each([
    '2026-01-20T10:00:00Z or isRead eq true',
    "2026-01-20'",
    'yesterday',
  ])('rejects %j', value => {
    expect(() => sanitizeODataDatetime(value, 'receivedAfter')).toThrow(/Invalid receivedAfter/);
  });
});

describe('sanitizeTimezone', () => {
  it.each(['Europe/Berlin', 'America/Argentina/Buenos_Aires', 'Etc/GMT+2', 'UTC'])('accepts %s', value => {
    expect(sanitizeTimezone(value)).toBe(value);
  });

  it.each(['Europe/Berlin"', 'Europe/Berlin, x', 'Europe Berlin', 'a\r\nX-Injected: 1'])('rejects %j', value => {
    expect(() => sanitizeTimezone(value)).toThrow(/Invalid timeZone/);
  });
});
