import { randomUUID } from 'crypto';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { checkRateLimit, recordRequest } from '../src/middleware/rate-limiter.js';
import { setEnv } from './helpers.js';

describe('per-user rate limit', () => {
  afterEach(() => {
    vi.useRealTimers();
  });

  it('limits a user after the configured number of requests', () => {
    setEnv({ MS365_MCP_RATE_LIMIT_REQUESTS: '2', MS365_MCP_RATE_LIMIT_WINDOW_MS: '60000' });
    const user = randomUUID();

    expect(checkRateLimit(user)).toMatchObject({ isLimited: false, remaining: 2 });
    recordRequest(user);
    recordRequest(user);
    expect(checkRateLimit(user)).toMatchObject({ isLimited: true, remaining: 0 });
  });

  it('keeps users independent', () => {
    setEnv({ MS365_MCP_RATE_LIMIT_REQUESTS: '1' });
    const first = randomUUID();
    recordRequest(first);

    expect(checkRateLimit(first).isLimited).toBe(true);
    expect(checkRateLimit(randomUUID()).isLimited).toBe(false);
  });

  it('frees the user once the window has passed', () => {
    vi.useFakeTimers();
    setEnv({ MS365_MCP_RATE_LIMIT_REQUESTS: '1', MS365_MCP_RATE_LIMIT_WINDOW_MS: '1000' });
    const user = randomUUID();
    recordRequest(user);
    expect(checkRateLimit(user).isLimited).toBe(true);

    vi.advanceTimersByTime(1001);
    expect(checkRateLimit(user).isLimited).toBe(false);
  });
});
