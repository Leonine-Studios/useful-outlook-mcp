/**
 * Configuration management for the Outlook OAuth MCP Server
 */

export interface Config {
  /** Azure AD Application (client) ID */
  clientId: string;
  /** Azure AD client secret (optional, for confidential clients) */
  clientSecret?: string;
  /** Azure AD tenant ID (default: 'common' for multi-tenant) */
  tenantId: string;
  /** Server port */
  port: number;
  /** Server bind address */
  host: string;
  /** Log level */
  logLevel: 'debug' | 'info' | 'warn' | 'error';
  /** CORS allowed origins */
  corsOrigin: string;
  /** Rate limit: max requests per window */
  rateLimitRequests: number;
  /** Rate limit: window size in milliseconds */
  rateLimitWindowMs: number;
  /** Allowed tenant IDs (comma-separated, empty = all tenants allowed) */
  allowedTenants: string[];
  /** Read-only mode: when true, disables all write tools */
  readOnlyMode: boolean;
  /** Enabled tools: comma-separated list of tool names (empty = all tools) */
  enabledTools: string[];
}

let cachedConfig: Config | null = null;

/**
 * Load configuration from environment variables
 */
export function getConfig(): Config {
  if (cachedConfig) {
    return cachedConfig;
  }

  const clientId = process.env.MS365_MCP_CLIENT_ID;
  
  if (!clientId) {
    throw new Error('MS365_MCP_CLIENT_ID environment variable is required');
  }

  // Parse allowed tenants from comma-separated string
  const allowedTenantsEnv = process.env.MS365_MCP_ALLOWED_TENANTS || '';
  const allowedTenants = allowedTenantsEnv
    .split(',')
    .map(t => t.trim())
    .filter(t => t.length > 0);

  // Parse enabled tools from comma-separated string
  const enabledToolsEnv = process.env.MS365_MCP_ENABLED_TOOLS || '';
  const enabledTools = enabledToolsEnv
    .split(',')
    .map(t => t.trim())
    .filter(t => t.length > 0);

  cachedConfig = {
    clientId,
    clientSecret: process.env.MS365_MCP_CLIENT_SECRET || undefined,
    tenantId: process.env.MS365_MCP_TENANT_ID || 'common',
    port: parseInt(process.env.MS365_MCP_PORT || '3000', 10),
    host: process.env.MS365_MCP_HOST || '0.0.0.0',
    logLevel: (process.env.MS365_MCP_LOG_LEVEL || 'info') as Config['logLevel'],
    corsOrigin: process.env.MS365_MCP_CORS_ORIGIN || '*',
    rateLimitRequests: parseInt(process.env.MS365_MCP_RATE_LIMIT_REQUESTS || '30', 10),
    rateLimitWindowMs: parseInt(process.env.MS365_MCP_RATE_LIMIT_WINDOW_MS || '60000', 10),
    allowedTenants,
    readOnlyMode: process.env.MS365_MCP_READ_ONLY_MODE === 'true',
    enabledTools,
  };

  return cachedConfig;
}

/**
 * Microsoft Entra ID endpoints
 */
export function getAuthEndpoints(tenantId: string) {
  const authority = 'https://login.microsoftonline.com';
  
  return {
    authority,
    authorizationEndpoint: `${authority}/${tenantId}/oauth2/v2.0/authorize`,
    tokenEndpoint: `${authority}/${tenantId}/oauth2/v2.0/token`,
  };
}

/**
 * Microsoft Graph API base URL
 */
export const GRAPH_API_BASE = 'https://graph.microsoft.com/v1.0';

/**
 * Supported OAuth scopes for Outlook access
 */
export const SUPPORTED_SCOPES = [
  'Mail.Read',
  'Mail.ReadWrite',
  'Mail.Send',
  'Calendars.Read',
  'Calendars.ReadWrite',
  'offline_access',
  'User.Read',
] as const;

export type SupportedScope = (typeof SUPPORTED_SCOPES)[number];
