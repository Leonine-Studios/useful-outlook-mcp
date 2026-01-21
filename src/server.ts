/**
 * Express server setup for the MCP server
 */

import express, { Request, Response } from 'express';
import { StreamableHTTPServerTransport } from '@modelcontextprotocol/sdk/server/streamableHttp.js';
import { getConfig } from './config.js';
import { createMcpServer } from './mcp/handler.js';
import authRouter from './auth/metadata.js';
import { bearerAuthMiddleware, AuthenticatedRequest } from './auth/middleware.js';
import { rateLimitMiddleware, getRateLimitStats } from './middleware/rate-limiter.js';
import { runWithContext } from './utils/context.js';
import logger, { setLogLevel, printStartupBanner } from './utils/logger.js';
import { allTools, getFilteredTools } from './tools/index.js';

const VERSION = '1.0.0';

/**
 * Create and configure the Express application
 */
export function createApp() {
  const config = getConfig();
  
  // Set log level from config
  setLogLevel(config.logLevel);
  
  const app = express();
  
  // Trust proxy for correct protocol detection behind reverse proxies
  app.set('trust proxy', true);
  
  // Parse request bodies
  app.use(express.json());
  app.use(express.urlencoded({ extended: true }));
  
  // CORS configuration
  app.use((req, res, next) => {
    res.header('Access-Control-Allow-Origin', config.corsOrigin);
    res.header('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS');
    res.header(
      'Access-Control-Allow-Headers',
      'Origin, X-Requested-With, Content-Type, Accept, Authorization, mcp-protocol-version'
    );
    // Expose rate limit headers to clients
    res.header(
      'Access-Control-Expose-Headers',
      'X-RateLimit-Limit, X-RateLimit-Remaining, X-RateLimit-Reset, Retry-After'
    );
    
    // Handle preflight requests
    if (req.method === 'OPTIONS') {
      res.sendStatus(200);
      return;
    }
    
    next();
  });
  
  // Health check endpoint with rate limit stats
  app.get('/health', (_req: Request, res: Response) => {
    const rateLimitStats = getRateLimitStats();
    
    res.json({
      status: 'healthy',
      version: VERSION,
      timestamp: new Date().toISOString(),
      rateLimiting: {
        activeUsers: rateLimitStats.activeUsers,
        trackedRequests: rateLimitStats.totalTrackedRequests,
      },
    });
  });
  
  // Root endpoint
  app.get('/', (_req: Request, res: Response) => {
    res.json({
      name: 'outlook-oauth-mcp',
      version: VERSION,
      description: 'MCP server for Outlook with OAuth2 delegated access',
      endpoints: {
        mcp: '/mcp',
        health: '/health',
        oauth_protected_resource: '/.well-known/oauth-protected-resource',
        oauth_authorization_server: '/.well-known/oauth-authorization-server',
      },
    });
  });
  
  // OAuth metadata and proxy endpoints
  app.use(authRouter);
  
  // Create MCP server
  const mcpServer = createMcpServer(VERSION);
  
  // Create rate limit middleware instance
  const rateLimit = rateLimitMiddleware();
  
  // MCP endpoint - requires authentication and rate limiting
  const handleMcpRequest = async (req: AuthenticatedRequest, res: Response) => {
    if (!req.auth?.token) {
      res.status(401).json({
        error: 'invalid_token',
        error_description: 'No valid access token provided',
      });
      return;
    }
    
    // Create per-request context with the OAuth token
    const context = {
      accessToken: req.auth.token,
      userId: req.auth.userId,
    };
    
    await runWithContext(context, async () => {
      try {
        const transport = new StreamableHTTPServerTransport({
          sessionIdGenerator: undefined, // Stateless mode
        });
        
        res.on('close', () => {
          transport.close();
        });
        
        await mcpServer.connect(transport);
        await transport.handleRequest(req as never, res as never, req.body);
      } catch (error) {
        logger.error('MCP request error', {
          error: error instanceof Error ? error.message : String(error),
        });
        
        if (!res.headersSent) {
          res.status(500).json({
            jsonrpc: '2.0',
            error: {
              code: -32603,
              message: 'Internal server error',
            },
            id: null,
          });
        }
      }
    });
  };
  
  // Handle both GET (SSE) and POST for MCP
  // Middleware chain: auth -> rate limit -> handler
  app.get('/mcp', bearerAuthMiddleware, rateLimit, handleMcpRequest);
  app.post('/mcp', bearerAuthMiddleware, rateLimit, handleMcpRequest);
  
  return app;
}

/**
 * Start the server
 */
export async function startServer(): Promise<void> {
  const config = getConfig();
  const app = createApp();
  
  // Gather tool info for the banner
  const filteredTools = getFilteredTools();
  const enabledToolNames = filteredTools.map(t => t.name);
  const disabledToolNames = allTools
    .filter(t => !filteredTools.some(ft => ft.name === t.name))
    .map(t => t.name);
  
  app.listen(config.port, config.host, () => {
    printStartupBanner({
      version: VERSION,
      host: config.host,
      port: config.port,
      logLevel: config.logLevel,
      enabledTools: enabledToolNames,
      disabledTools: disabledToolNames,
      totalTools: allTools.length,
      rateLimitRequests: config.rateLimitRequests,
      rateLimitWindowMs: config.rateLimitWindowMs,
      allowedTenants: config.allowedTenants.length > 0 ? config.allowedTenants.length : 'all',
      readOnlyMode: config.readOnlyMode,
    });
  });
}
