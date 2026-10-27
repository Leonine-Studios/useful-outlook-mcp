import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { getFilteredTools } from '../src/tools/index.js';
import {
  initializeRequest,
  makeToken,
  mcpRequest,
  setEnv,
  startTestServer,
  type TestServer,
} from './helpers.js';

let server: TestServer;

beforeEach(async () => {
  server = await startTestServer();
});

afterEach(async () => {
  await server.close();
});

describe('health', () => {
  it('reports healthy', async () => {
    const response = await fetch(`${server.url}/health`);
    expect(response.status).toBe(200);
    expect(await response.json()).toMatchObject({ status: 'healthy' });
  });
});

describe('/mcp authentication', () => {
  it('rejects a request without a token', async () => {
    const response = await fetch(`${server.url}/mcp`, { method: 'POST' });
    expect(response.status).toBe(401);
  });

  it('rejects a non-bearer scheme', async () => {
    const response = await fetch(`${server.url}/mcp`, {
      method: 'POST',
      headers: { Authorization: 'Basic dXNlcjpwYXNz' },
    });
    expect(response.status).toBe(401);
  });

  it('rejects a malformed token', async () => {
    const { status } = await mcpRequest(server.url, initializeRequest(), 'not-a-jwt');
    expect(status).toBe(401);
  });

  it('rejects an expired token', async () => {
    const token = makeToken({ exp: Math.floor(Date.now() / 1000) - 3600 });
    const { status } = await mcpRequest(server.url, initializeRequest(), token);
    expect(status).toBe(401);
  });

  it('rejects a tenant outside the allowlist with 403', async () => {
    setEnv({ MS365_MCP_ALLOWED_TENANTS: 'tenant-a' });
    const { status } = await mcpRequest(server.url, initializeRequest(), makeToken({ tid: 'tenant-b' }));
    expect(status).toBe(403);
  });
});

describe('/mcp protocol', () => {
  it('initializes', async () => {
    const { status, message } = await mcpRequest(server.url, initializeRequest());
    expect(status).toBe(200);
    expect(message).toMatchObject({
      id: 1,
      result: { serverInfo: { name: 'outlook-oauth-mcp' } },
    });
  });

  it('serves requests while another client holds a stream open', async () => {
    const stream = new AbortController();
    const openStream = await fetch(`${server.url}/mcp`, {
      headers: {
        Authorization: `Bearer ${makeToken({ oid: 'streaming-user' })}`,
        Accept: 'text/event-stream',
      },
      signal: stream.signal,
    });

    try {
      expect(openStream.status).toBe(200);
      const responses = await Promise.all(
        [1, 2, 3].map(id => mcpRequest(server.url, initializeRequest(id))),
      );
      for (const response of responses) {
        expect(response.status).toBe(200);
        expect(response.message).toHaveProperty('result');
      }
    } finally {
      stream.abort();
    }
  });

  it('lists the enabled tools', async () => {
    const { status, message } = await mcpRequest(server.url, {
      jsonrpc: '2.0',
      id: 7,
      method: 'tools/list',
    });
    expect(status).toBe(200);
    const tools = (message?.result as { tools: Array<{ name: string }> }).tools;
    expect(tools.map(tool => tool.name).sort()).toEqual(getFilteredTools().map(tool => tool.name).sort());
  });

  it('rate limits per user', async () => {
    setEnv({ MS365_MCP_RATE_LIMIT_REQUESTS: '1' });
    const token = makeToken({ oid: 'rate-limited-user' });

    expect((await mcpRequest(server.url, initializeRequest(), token)).status).toBe(200);
    expect((await mcpRequest(server.url, initializeRequest(), token)).status).toBe(429);
  });
});

describe('OAuth endpoints', () => {
  it('serves protected resource metadata', async () => {
    const response = await fetch(`${server.url}/.well-known/oauth-protected-resource`);
    expect(response.status).toBe(200);
    expect(await response.json()).toMatchObject({
      resource: `${server.url}/mcp`,
      authorization_servers: [server.url],
    });
  });

  it('serves authorization server metadata', async () => {
    const response = await fetch(`${server.url}/.well-known/oauth-authorization-server`);
    expect(await response.json()).toMatchObject({
      issuer: server.url,
      authorization_endpoint: `${server.url}/authorize`,
      token_endpoint: `${server.url}/token`,
      code_challenge_methods_supported: ['S256'],
    });
  });

  it('redirects /authorize to Entra ID with our client id', async () => {
    const response = await fetch(
      `${server.url}/authorize?client_id=attacker&redirect_uri=http%3A%2F%2F127.0.0.1%2Fcb&state=s&prompt=none`,
      { redirect: 'manual' },
    );
    expect(response.status).toBe(302);

    const location = new URL(response.headers.get('location') ?? '');
    expect(location.origin).toBe('https://login.microsoftonline.com');
    expect(location.pathname).toBe('/test-tenant/oauth2/v2.0/authorize');
    expect(location.searchParams.get('client_id')).toBe('test-client-id');
    expect(location.searchParams.get('state')).toBe('s');
    expect(location.searchParams.has('prompt')).toBe(false);
    expect(location.searchParams.get('scope')).toContain('offline_access');
  });

  it('validates token requests before contacting Entra ID', async () => {
    const response = await fetch(`${server.url}/token`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: 'grant_type=password&username=a&password=b',
    });
    expect(response.status).toBe(400);
    expect(await response.json()).toMatchObject({ error: 'unsupported_grant_type' });
  });
});
