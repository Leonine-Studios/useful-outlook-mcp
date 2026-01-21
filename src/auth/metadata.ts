/**
 * OAuth metadata endpoints for MCP spec compliance
 * 
 * Implements:
 * - RFC 9728: OAuth 2.0 Protected Resource Metadata
 * - RFC 8414: OAuth 2.0 Authorization Server Metadata
 * - RFC 7591: OAuth 2.0 Dynamic Client Registration
 */

import { Request, Response, Router } from 'express';
import { randomUUID } from 'crypto';
import { getConfig, getAuthEndpoints } from '../config.js';
import { getRequiredScopes } from '../tools/index.js';
import logger from '../utils/logger.js';

const router = Router();

/**
 * In-memory store for dynamically registered clients
 * In production, this should be persisted to a database
 */
interface RegisteredClient {
  client_id: string;
  client_secret?: string;
  client_name?: string;
  redirect_uris: string[];
  client_id_issued_at: number;
}

const registeredClients = new Map<string, RegisteredClient>();

/**
 * Get the base URL from the request
 */
function getBaseUrl(req: Request): string {
  const protocol = req.secure || req.headers['x-forwarded-proto'] === 'https' 
    ? 'https' 
    : 'http';
  const host = req.headers['x-forwarded-host'] || req.headers.host || 'localhost';
  return `${protocol}://${host}`;
}

/**
 * RFC 9728: OAuth 2.0 Protected Resource Metadata
 * 
 * This endpoint tells MCP clients:
 * - What resource this server protects (the /mcp endpoint)
 * - Which authorization servers can issue tokens for it
 * - What scopes are supported
 */
router.get('/.well-known/oauth-protected-resource', (req: Request, res: Response) => {
  const baseUrl = getBaseUrl(req);
  const scopes = getRequiredScopes();
  
  logger.debug('Protected resource metadata requested', { baseUrl, scopes });
  
  res.json({
    resource: `${baseUrl}/mcp`,
    authorization_servers: [baseUrl],
    scopes_supported: scopes,
    bearer_methods_supported: ['header'],
    resource_documentation: `${baseUrl}/docs`,
  });
});

/**
 * RFC 8414: OAuth 2.0 Authorization Server Metadata
 * 
 * This endpoint describes the authorization server capabilities.
 * Since we proxy to Microsoft Entra ID, we advertise our proxy endpoints
 * but they redirect to/forward to Microsoft's actual endpoints.
 */
router.get('/.well-known/oauth-authorization-server', (req: Request, res: Response) => {
  const baseUrl = getBaseUrl(req);
  const scopes = getRequiredScopes();
  
  logger.debug('Authorization server metadata requested', { baseUrl, scopes });
  
  res.json({
    issuer: baseUrl,
    authorization_endpoint: `${baseUrl}/authorize`,
    token_endpoint: `${baseUrl}/token`,
    registration_endpoint: `${baseUrl}/register`,
    response_types_supported: ['code'],
    response_modes_supported: ['query'],
    grant_types_supported: ['authorization_code', 'refresh_token'],
    token_endpoint_auth_methods_supported: ['none', 'client_secret_post'],
    code_challenge_methods_supported: ['S256'],
    scopes_supported: scopes,
  });
});

/**
 * RFC 7591: OAuth 2.0 Dynamic Client Registration
 * 
 * This endpoint allows MCP clients to register themselves dynamically.
 * The registration is stored in memory (for simplicity) and doesn't affect
 * the actual authentication - we still use our Azure AD app's credentials
 * for the Microsoft OAuth flow.
 */
router.post('/register', (req: Request, res: Response) => {
  const body = req.body as {
    client_name?: string;
    redirect_uris?: string[];
    grant_types?: string[];
    response_types?: string[];
    token_endpoint_auth_method?: string;
  };
  
  logger.debug('Dynamic client registration request', { 
    client_name: body.client_name,
    redirect_uris: body.redirect_uris,
  });
  
  // Generate client credentials
  const clientId = randomUUID();
  const clientSecret = randomUUID(); // Optional for public clients
  const issuedAt = Math.floor(Date.now() / 1000);
  
  // Store the registered client
  const registeredClient: RegisteredClient = {
    client_id: clientId,
    client_secret: clientSecret,
    client_name: body.client_name,
    redirect_uris: body.redirect_uris || [],
    client_id_issued_at: issuedAt,
  };
  
  registeredClients.set(clientId, registeredClient);
  
  logger.debug('Client registered successfully', { 
    client_id: clientId,
    client_name: body.client_name,
  });
  
  // Return the registration response per RFC 7591
  res.status(201).json({
    client_id: clientId,
    client_secret: clientSecret,
    client_name: body.client_name,
    redirect_uris: body.redirect_uris || [],
    grant_types: body.grant_types || ['authorization_code', 'refresh_token'],
    response_types: body.response_types || ['code'],
    token_endpoint_auth_method: body.token_endpoint_auth_method || 'client_secret_post',
    client_id_issued_at: issuedAt,
  });
});

/**
 * Get a registered client by ID
 */
export function getRegisteredClient(clientId: string): RegisteredClient | undefined {
  return registeredClients.get(clientId);
}

/**
 * Authorization endpoint - redirects to Microsoft Entra ID
 */
router.get('/authorize', (req: Request, res: Response) => {
  const config = getConfig();
  const { authorizationEndpoint } = getAuthEndpoints(config.tenantId);
  
  // Build Microsoft authorization URL
  const microsoftAuthUrl = new URL(authorizationEndpoint);
  
  // Forward allowed OAuth parameters
  const allowedParams = [
    'response_type',
    'redirect_uri',
    'scope',
    'state',
    'code_challenge',
    'code_challenge_method',
    // 'prompt',
    'login_hint',
    'domain_hint',
  ];
  
  for (const param of allowedParams) {
    const value = req.query[param];
    if (value && typeof value === 'string') {
      microsoftAuthUrl.searchParams.set(param, value);
    }
  }
  
  // Always use our registered client_id
  microsoftAuthUrl.searchParams.set('client_id', config.clientId);
  
  // Ensure we have required scopes if none provided (based on enabled tools)
  if (!microsoftAuthUrl.searchParams.get('scope')) {
    const requiredScopes = getRequiredScopes();
    microsoftAuthUrl.searchParams.set('scope', requiredScopes.join(' '));
  }
  
  logger.debug('Redirecting to Microsoft authorization', {
    redirect_uri: req.query.redirect_uri,
    scope: microsoftAuthUrl.searchParams.get('scope'),
  });
  
  res.redirect(microsoftAuthUrl.toString());
});

/**
 * Token endpoint - proxies to Microsoft Entra ID
 */
router.post('/token', async (req: Request, res: Response) => {
  const config = getConfig();
  const { tokenEndpoint } = getAuthEndpoints(config.tenantId);
  
  const body = req.body as Record<string, string>;
  
  if (!body.grant_type) {
    res.status(400).json({
      error: 'invalid_request',
      error_description: 'grant_type parameter is required',
    });
    return;
  }
  
  // Build token request
  const params = new URLSearchParams();
  params.set('client_id', config.clientId);
  params.set('grant_type', body.grant_type);
  
  // Add client_secret if configured (confidential client)
  if (config.clientSecret) {
    params.set('client_secret', config.clientSecret);
  }
  
  if (body.grant_type === 'authorization_code') {
    if (!body.code || !body.redirect_uri) {
      res.status(400).json({
        error: 'invalid_request',
        error_description: 'code and redirect_uri are required for authorization_code grant',
      });
      return;
    }
    
    params.set('code', body.code);
    params.set('redirect_uri', body.redirect_uri);
    
    if (body.code_verifier) {
      params.set('code_verifier', body.code_verifier);
    }
    
    logger.debug('Token exchange: authorization_code');
  } else if (body.grant_type === 'refresh_token') {
    if (!body.refresh_token) {
      res.status(400).json({
        error: 'invalid_request',
        error_description: 'refresh_token is required for refresh_token grant',
      });
      return;
    }
    
    params.set('refresh_token', body.refresh_token);
    
    logger.debug('Token exchange: refresh_token');
  } else {
    res.status(400).json({
      error: 'unsupported_grant_type',
      error_description: `Grant type '${body.grant_type}' is not supported`,
    });
    return;
  }
  
  try {
    const response = await fetch(tokenEndpoint, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
      },
      body: params,
    });
    
    const data = await response.json();
    
    if (!response.ok) {
      logger.warn('Token exchange failed', { 
        status: response.status,
        error: (data as Record<string, unknown>).error,
      });
      res.status(response.status).json(data);
      return;
    }
    
    logger.debug('Token exchange successful');
    res.json(data);
  } catch (error) {
    logger.error('Token endpoint error', { 
      error: error instanceof Error ? error.message : String(error),
    });
    res.status(500).json({
      error: 'server_error',
      error_description: 'Token exchange failed',
    });
  }
});

export default router;
