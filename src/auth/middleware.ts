/**
 * Bearer token authentication middleware
 * 
 * Extracts and parses the access token from the Authorization header.
 * 
 * NOTE: Microsoft Graph API tokens cannot be cryptographically validated
 * by third parties. We parse the token to extract user identity for
 * logging and rate limiting. Graph API validates the token when called.
 */

import { Request, Response, NextFunction } from 'express';
import { parseToken, TokenValidationError, getUserIdentifier, ParsedToken } from './token-validator.js';
import logger from '../utils/logger.js';

/**
 * Extended request type with auth information
 */
export interface AuthenticatedRequest extends Request {
  auth?: {
    /** Raw access token for Graph API calls */
    token: string;
    /** User identifier (from parsed claims) */
    userId: string;
    /** Tenant ID the token was issued for */
    tenantId: string;
    /** Full parsed token payload */
    parsedToken: ParsedToken;
  };
  /** User ID for rate limiting */
  rateLimitUserId?: string;
}

/**
 * Map token validation error codes to HTTP status codes
 */
function getStatusCode(code: TokenValidationError['code']): number {
  switch (code) {
    case 'expired_token':
    case 'invalid_token':
      return 401;
    case 'tenant_not_allowed':
      return 403;
    default:
      return 401;
  }
}

/**
 * Middleware that requires a valid Bearer token in the Authorization header.
 * 
 * This middleware parses the token to extract user identity for:
 * - Audit logging
 * - Rate limiting  
 * - Tenant allowlist enforcement
 * 
 * The actual token validation is performed by Graph API when we make calls.
 */
export function bearerAuthMiddleware(
  req: AuthenticatedRequest,
  res: Response,
  next: NextFunction
): void {
  const authHeader = req.headers.authorization;

  if (!authHeader) {
    // This is expected during MCP OAuth discovery - client probes without auth first
    logger.debug('Missing Authorization header', { path: req.path });
    res.status(401).json({
      error: 'invalid_request',
      error_description: 'Missing Authorization header',
    });
    return;
  }

  if (!authHeader.startsWith('Bearer ')) {
    logger.warn('Invalid Authorization header format', { path: req.path });
    res.status(401).json({
      error: 'invalid_request',
      error_description: 'Authorization header must use Bearer scheme',
    });
    return;
  }

  const token = authHeader.substring(7).trim();

  if (!token) {
    logger.warn('Empty Bearer token', { path: req.path });
    res.status(401).json({
      error: 'invalid_token',
      error_description: 'Bearer token is empty',
    });
    return;
  }

  try {
    const parsedToken = parseToken(token);
    const userId = getUserIdentifier(parsedToken);
    
    // Store parsed auth info in request
    req.auth = {
      token,
      userId,
      tenantId: parsedToken.tid,
      parsedToken,
    };
    
    // Set user ID for rate limiting
    req.rateLimitUserId = parsedToken.oid;

    logger.debug('Token parsed successfully', { 
      userId, 
      tenantId: parsedToken.tid,
      path: req.path 
    });
    
    next();
  } catch (error) {
    if (error instanceof TokenValidationError) {
      logger.warn('Token validation failed', { 
        code: error.code, 
        message: error.message,
        path: req.path 
      });
      
      res.status(getStatusCode(error.code)).json({
        error: error.code,
        error_description: error.message,
      });
      return;
    }

    // Unexpected error
    logger.error('Unexpected authentication error', {
      error: error instanceof Error ? error.message : String(error),
      path: req.path,
    });
    
    res.status(500).json({
      error: 'server_error',
      error_description: 'Authentication failed due to internal error',
    });
  }
}

/**
 * Optional middleware that allows unauthenticated requests
 * but still parses and extracts auth info if present.
 */
export function optionalAuthMiddleware(
  req: AuthenticatedRequest,
  _res: Response,
  next: NextFunction
): void {
  const authHeader = req.headers.authorization;

  if (!authHeader?.startsWith('Bearer ')) {
    next();
    return;
  }

  const token = authHeader.substring(7).trim();
  if (!token) {
    next();
    return;
  }

  try {
    const parsedToken = parseToken(token);
    const userId = getUserIdentifier(parsedToken);
    
    req.auth = {
      token,
      userId,
      tenantId: parsedToken.tid,
      parsedToken,
    };
    
    req.rateLimitUserId = parsedToken.oid;
  } catch {
    // For optional auth, silently ignore parsing errors
  }
  
  next();
}
