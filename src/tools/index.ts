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
    const beforeCount = tools.length;
    tools = tools.filter(t => enabledSet.has(t.name));
    
    const disabledByFilter = beforeCount - tools.length;
    if (disabledByFilter > 0) {
      logger.info('Tool filter applied', {
        enabledTools: config.enabledTools,
        disabledCount: disabledByFilter,
      });
    }
  }

  // Apply read-only mode filter (removes write tools)
  if (config.readOnlyMode) {
    const beforeCount = tools.length;
    const disabledTools = tools.filter(t => !t.readOnly).map(t => t.name);
    tools = tools.filter(t => t.readOnly);
    
    if (disabledTools.length > 0) {
      logger.info('Read-only mode enabled', {
        disabledWriteTools: disabledTools,
      });
    }
  }

  return tools;
}

/**
 * Export individual tool modules for direct access
 */
export * from './mail.js';
export * from './calendar.js';
