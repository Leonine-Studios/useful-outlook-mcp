/**
 * Tool registry - combines all tools for the MCP server
 */

import { mailToolDefinitions } from './mail.js';
import { calendarToolDefinitions } from './calendar.js';
import { getConfig } from '../config.js';
import logger from '../utils/logger.js';

export interface ToolDefinition {
  name: string;
  description: string;
  /** Whether this tool is read-only (doesn't modify data) */
  readOnly: boolean;
  /** OAuth scopes required by this tool */
  requiredScopes: string[];
  inputSchema: {
    type: 'object';
    properties: Record<string, unknown>;
    required?: string[];
  };
  handler: (params: Record<string, unknown>) => Promise<{
    content: Array<{ type: 'text'; text: string }>;
    isError?: boolean;
  }>;
}

/**
 * All available tools
 */
export const allTools: ToolDefinition[] = [
  ...mailToolDefinitions as ToolDefinition[],
  ...calendarToolDefinitions as ToolDefinition[],
];

/**
 * Get tool by name
 */
export function getTool(name: string): ToolDefinition | undefined {
  return allTools.find(t => t.name === name);
}

/**
 * Get all tool names
 */
export function getToolNames(): string[] {
  return allTools.map(t => t.name);
}

/**
 * Get filtered tools based on configuration
 * Applies both read-only mode and tool allowlist filtering.
 * When both are enabled: tool filter is applied first, then read-only removes write tools.
 */
export function getFilteredTools(): ToolDefinition[] {
  const config = getConfig();
  let tools = [...allTools];

  // Apply tool allowlist filter first (if configured)
  if (config.enabledTools.length > 0) {
    const enabledSet = new Set(config.enabledTools);
    tools = tools.filter(t => enabledSet.has(t.name));
  }

  // Apply read-only mode filter (removes write tools)
  if (config.readOnlyMode) {
    tools = tools.filter(t => t.readOnly);
  }

  return tools;
}

/** Base scopes always required for the server to function */
const BASE_SCOPES = ['User.Read', 'offline_access'];

/**
 * Get required OAuth scopes based on enabled tools
 * Returns only the scopes needed for the currently enabled tools.
 */
export function getRequiredScopes(): string[] {
  const tools = getFilteredTools();
  
  // Collect unique scopes from all enabled tools
  const scopeSet = new Set<string>(BASE_SCOPES);
  
  for (const tool of tools) {
    for (const scope of tool.requiredScopes) {
      scopeSet.add(scope);
    }
  }
  
  const scopes = Array.from(scopeSet);
  
  logger.debug('Required scopes computed', {
    enabledTools: tools.length,
    scopes,
  });
  
  return scopes;
}

/**
 * Export individual tool modules for direct access
 */
export * from './mail.js';
export * from './calendar.js';
