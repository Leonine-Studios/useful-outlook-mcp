/**
 * Microsoft Graph API client
 * 
 * Provides a simple interface for making Graph API calls
 * using the access token from the request context.
 */

import { getContextToken } from '../utils/context.js';
import logger from '../utils/logger.js';
import { GRAPH_API_BASE } from '../config.js';
import { serializeResponse } from '../utils/tonl.js';

export interface GraphRequestOptions {
  method?: 'GET' | 'POST' | 'PATCH' | 'DELETE';
  body?: unknown;
  headers?: Record<string, string>;
}

export interface GraphResponse<T = unknown> {
  data: T;
  status: number;
  ok: boolean;
}

export interface GraphError {
  error: {
    code: string;
    message: string;
    innerError?: {
      'request-id': string;
      date: string;
    };
  };
}

/**
 * Make a request to Microsoft Graph API
 */
export async function graphRequest<T = unknown>(
  endpoint: string,
  options: GraphRequestOptions = {}
): Promise<GraphResponse<T>> {
  const accessToken = getContextToken();
  
  if (!accessToken) {
    throw new Error('No access token available in request context');
  }
  
  const url = endpoint.startsWith('http') 
    ? endpoint 
    : `${GRAPH_API_BASE}${endpoint.startsWith('/') ? endpoint : `/${endpoint}`}`;
  
  const headers: Record<string, string> = {
    'Authorization': `Bearer ${accessToken}`,
    'Content-Type': 'application/json',
    ...options.headers,
  };
  
  const requestOptions: RequestInit = {
    method: options.method || 'GET',
    headers,
  };
  
  if (options.body && options.method !== 'GET') {
    requestOptions.body = JSON.stringify(options.body);
  }
  
  logger.debug('Graph API request', { 
    method: requestOptions.method, 
    url: url.replace(GRAPH_API_BASE, ''),
  });
  
  const response = await fetch(url, requestOptions);
  
  let data: T;
  const contentType = response.headers.get('content-type');
  
  if (contentType?.includes('application/json')) {
    data = await response.json() as T;
  } else {
    const text = await response.text();
    data = (text || { success: true }) as T;
  }
  
  if (!response.ok) {
    logger.warn('Graph API error', {
      status: response.status,
      statusText: response.statusText,
      url: url.replace(GRAPH_API_BASE, ''),
    });
  }
  
  return {
    data,
    status: response.status,
    ok: response.ok,
  };
}

/**
 * Format a successful MCP tool response
 */
export function formatToolResponse(data: unknown): {
  content: Array<{ type: 'text'; text: string }>;
  isError?: boolean;
} {
  return {
    content: [{
      type: 'text',
      text: serializeResponse(data),
    }],
  };
}

/**
 * Format an error MCP tool response
 */
export function formatErrorResponse(error: unknown): {
  content: Array<{ type: 'text'; text: string }>;
  isError: boolean;
} {
  const message = error instanceof Error ? error.message : String(error);
  
  return {
    content: [{
      type: 'text',
      text: serializeResponse({ error: message }),
    }],
    isError: true,
  };
}

/**
 * Handle Graph API response and format for MCP
 */
export function handleGraphResponse<T>(
  response: GraphResponse<T>
): {
  content: Array<{ type: 'text'; text: string }>;
  isError?: boolean;
} {
  if (!response.ok) {
    const graphError = response.data as unknown as GraphError;
    return formatErrorResponse(
      graphError?.error?.message || `Graph API error: ${response.status}`
    );
  }
  
  return formatToolResponse(response.data);
}
