/**
 * MCP Server handler
 * 
 * Sets up the MCP server with all tools and handles requests.
 */

import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { allTools, getFilteredTools } from '../tools/index.js';
import logger from '../utils/logger.js';

/**
 * Create and configure the MCP server
 */
export function createMcpServer(version: string): McpServer {
  const server = new McpServer({
    name: 'outlook-oauth-mcp',
    version,
  });

  // Get filtered tools based on config (read-only mode and/or tool allowlist)
  const tools = getFilteredTools();
  
  // Log tool filtering summary
  if (tools.length < allTools.length) {
    const disabledTools = allTools
      .filter(t => !tools.some(ft => ft.name === t.name))
      .map(t => t.name);
    logger.info('Tool filtering active', {
      totalTools: allTools.length,
      enabledTools: tools.length,
      disabledTools,
    });
  }

  // Register filtered tools
  for (const tool of tools) {
    logger.debug('Registering tool', { name: tool.name, readOnly: tool.readOnly });
    
    // Convert JSON Schema properties to Zod schema
    const zodSchema = createZodSchema(tool.inputSchema);
    
    server.registerTool(
      tool.name,
      {
        description: tool.description,
        inputSchema: zodSchema,
      },
      async (args) => {
        logger.info('Tool called', { tool: tool.name });
        
        try {
          const result = await tool.handler(args);
          return result;
        } catch (error) {
          logger.error('Tool execution error', {
            tool: tool.name,
            error: error instanceof Error ? error.message : String(error),
          });
          return {
            content: [{
              type: 'text' as const,
              text: JSON.stringify({
                error: error instanceof Error ? error.message : 'Tool execution failed',
              }),
            }],
            isError: true,
          };
        }
      }
    );
  }

  logger.info('MCP server initialized', { 
    toolCount: tools.length,
    totalAvailable: allTools.length,
  });

  return server;
}

/**
 * Convert JSON Schema-like input schema to Zod schema
 */
function createZodSchema(inputSchema: {
  type: 'object';
  properties: Record<string, unknown>;
  required?: string[];
}): z.ZodObject<z.ZodRawShape> {
  const shape: z.ZodRawShape = {};
  
  for (const [key, prop] of Object.entries(inputSchema.properties)) {
    const propDef = prop as { type?: string; description?: string; enum?: string[]; items?: { type: string } };
    
    let zodType: z.ZodTypeAny;
    
    switch (propDef.type) {
      case 'string':
        if (propDef.enum) {
          zodType = z.enum(propDef.enum as [string, ...string[]]);
        } else {
          zodType = z.string();
        }
        break;
      case 'number':
        zodType = z.number();
        break;
      case 'boolean':
        zodType = z.boolean();
        break;
      case 'array':
        if (propDef.items?.type === 'string') {
          zodType = z.array(z.string());
        } else {
          zodType = z.array(z.unknown());
        }
        break;
      default:
        zodType = z.unknown();
    }
    
    // Make optional if not in required array
    if (!inputSchema.required?.includes(key)) {
      zodType = zodType.optional();
    }
    
    if (propDef.description) {
      zodType = zodType.describe(propDef.description);
    }
    
    shape[key] = zodType;
  }
  
  return z.object(shape);
}
