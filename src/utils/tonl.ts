/**
 * TONL (Token Optimized Notation Language) serialization utility
 * 
 * Provides TONL encoding for MCP tool responses to reduce token usage
 * by 32-50% compared to JSON.
 */

import { encodeSmart, type TONLValue } from 'tonl';
import { getConfig } from '../config.js';

/**
 * Serialize data for MCP tool responses
 * 
 * Uses TONL encoding when enabled (default), falls back to JSON otherwise.
 * TONL achieves significant token savings especially for arrays of objects
 * by using a tabular format that eliminates key repetition.
 */
export function serializeResponse(data: unknown): string {
  const config = getConfig();
  
  if (config.useTonl) {
    return encodeSmart(data as TONLValue);
  }
  
  return JSON.stringify(data, null, 2);
}
