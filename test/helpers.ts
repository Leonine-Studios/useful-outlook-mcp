import type { AddressInfo } from 'net';
import type { Server } from 'http';
import { vi } from 'vitest';
import { createApp } from '../src/server.js';
import { resetConfig } from '../src/config.js';

export function makeToken(claims: Record<string, unknown> = {}): string {
  const header = Buffer.from(JSON.stringify({ alg: 'RS256', typ: 'JWT' })).toString('base64url');
  const payload = Buffer.from(JSON.stringify({
    iss: 'https://sts.windows.net/test-tenant/',
    aud: '00000003-0000-0000-c000-000000000000',
    tid: 'test-tenant',
    oid: 'user-oid',
    preferred_username: 'user@example.test',
    exp: Math.floor(Date.now() / 1000) + 3600,
    ...claims,
  })).toString('base64url');
  return `${header}.${payload}.signature`;
}

export function setEnv(values: Record<string, string>): void {
  Object.assign(process.env, values);
  resetConfig();
}

export interface TestServer {
  url: string;
  close: () => Promise<void>;
}

export async function startTestServer(): Promise<TestServer> {
  const app = createApp();
  const server: Server = await new Promise(resolve => {
    const listening = app.listen(0, '127.0.0.1', () => resolve(listening));
  });
  const { port } = server.address() as AddressInfo;
  return {
    url: `http://127.0.0.1:${port}`,
    close: () => new Promise((resolve, reject) => {
      server.closeAllConnections();
      server.close(error => (error ? reject(error) : resolve()));
    }),
  };
}

export async function mcpRequest(
  url: string,
  body: Record<string, unknown>,
  token: string = makeToken(),
): Promise<{ status: number; message: Record<string, unknown> | undefined }> {
  const response = await fetch(`${url}/mcp`, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: 'application/json, text/event-stream',
      'Content-Type': 'application/json',
    },
    body: JSON.stringify(body),
  });
  const text = await response.text();
  const contentType = response.headers.get('content-type') ?? '';

  if (contentType.includes('text/event-stream')) {
    const dataLine = text.split('\n').find(line => line.startsWith('data: '));
    return {
      status: response.status,
      message: dataLine ? JSON.parse(dataLine.slice('data: '.length)) : undefined,
    };
  }
  return { status: response.status, message: text ? JSON.parse(text) : undefined };
}

export function initializeRequest(id: number = 1): Record<string, unknown> {
  return {
    jsonrpc: '2.0',
    id,
    method: 'initialize',
    params: {
      protocolVersion: '2025-03-26',
      capabilities: {},
      clientInfo: { name: 'vitest', version: '0' },
    },
  };
}

export interface GraphCall {
  url: URL;
  init: RequestInit;
}

export function jsonResponse(body: unknown, status: number = 200): Response {
  return new Response(JSON.stringify(body), {
    status,
    headers: { 'content-type': 'application/json' },
  });
}

export function mockGraph(
  responder: (call: GraphCall) => Response | Promise<Response> = () => jsonResponse({ value: [] }),
): GraphCall[] {
  const calls: GraphCall[] = [];
  vi.stubGlobal('fetch', vi.fn(async (input: string | URL, init: RequestInit = {}) => {
    const call = { url: new URL(String(input)), init };
    calls.push(call);
    return responder(call);
  }));
  return calls;
}

export function header(call: GraphCall, name: string): string | undefined {
  return (call.init.headers as Record<string, string> | undefined)?.[name];
}
