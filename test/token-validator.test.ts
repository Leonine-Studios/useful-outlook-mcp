import { describe, expect, it } from 'vitest';
import { getUserIdentifier, parseToken, TokenValidationError } from '../src/auth/token-validator.js';
import { makeToken, setEnv } from './helpers.js';

function expectTokenError(token: string, code: TokenValidationError['code']): void {
  let caught: unknown;
  try {
    parseToken(token);
  } catch (error) {
    caught = error;
  }
  expect(caught).toBeInstanceOf(TokenValidationError);
  expect((caught as TokenValidationError).code).toBe(code);
}

describe('parseToken', () => {
  it('extracts identity claims', () => {
    const parsed = parseToken(makeToken({ oid: 'abc', tid: 'tenant-a' }));
    expect(parsed).toMatchObject({ oid: 'abc', sub: 'abc', tid: 'tenant-a' });
    expect(getUserIdentifier(parsed)).toBe('user@example.test');
  });

  it('falls back to sub when oid is missing', () => {
    const parsed = parseToken(makeToken({ oid: undefined, sub: 'sub-only' }));
    expect(parsed.oid).toBe('sub-only');
  });

  it.each(['not-a-jwt', 'a.b', 'a.b.c.d'])('rejects malformed token %j', token => {
    expectTokenError(token, 'invalid_token');
  });

  it('rejects a payload that is not JSON', () => {
    expectTokenError(`x.${Buffer.from('not json').toString('base64url')}.y`, 'invalid_token');
  });

  it.each([
    ['iss', { iss: undefined }],
    ['aud', { aud: undefined }],
    ['tid', { tid: undefined }],
    ['oid and sub', { oid: undefined, sub: undefined }],
  ])('rejects a token without %s', (_label, claims) => {
    expectTokenError(makeToken(claims), 'invalid_token');
  });

  it('rejects an expired token beyond the clock skew', () => {
    expectTokenError(makeToken({ exp: Math.floor(Date.now() / 1000) - 120 }), 'expired_token');
  });

  it('accepts a token inside the clock skew', () => {
    expect(() => parseToken(makeToken({ exp: Math.floor(Date.now() / 1000) - 30 }))).not.toThrow();
  });

  it('enforces the tenant allowlist', () => {
    setEnv({ MS365_MCP_ALLOWED_TENANTS: 'tenant-a, tenant-b' });
    expect(() => parseToken(makeToken({ tid: 'tenant-b' }))).not.toThrow();
    expectTokenError(makeToken({ tid: 'tenant-c' }), 'tenant_not_allowed');
  });
});
