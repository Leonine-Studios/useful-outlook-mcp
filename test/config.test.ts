import { describe, expect, it } from 'vitest';
import { getConfig } from '../src/config.js';
import { setEnv } from './helpers.js';

describe('getConfig', () => {
  it('requires a client id', () => {
    delete process.env.MS365_MCP_CLIENT_ID;
    setEnv({});
    expect(() => getConfig()).toThrow(/MS365_MCP_CLIENT_ID/);
  });
});
