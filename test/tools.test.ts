import { describe, expect, it } from 'vitest';
import { allTools, getFilteredTools, getRequiredScopes } from '../src/tools/index.js';
import { setEnv } from './helpers.js';

describe('tool registry', () => {
  it('has unique tool names', () => {
    const names = allTools.map(tool => tool.name);
    expect(new Set(names).size).toBe(names.length);
  });

  it('declares scopes and an object schema for every tool', () => {
    for (const tool of allTools) {
      expect(tool.requiredScopes.length, tool.name).toBeGreaterThan(0);
      expect(tool.inputSchema.type, tool.name).toBe('object');
      for (const required of tool.inputSchema.required ?? []) {
        expect(tool.inputSchema.properties, `${tool.name}.${required}`).toHaveProperty(required);
      }
    }
  });

  it('only marks tools read-only when they need no write scope', () => {
    const writeScopes = ['Mail.ReadWrite', 'Mail.Send', 'Calendars.ReadWrite'];
    for (const tool of allTools.filter(t => t.readOnly)) {
      expect(tool.requiredScopes.filter(scope => writeScopes.includes(scope)), tool.name).toEqual([]);
    }
  });

  it('returns every tool by default', () => {
    expect(getFilteredTools()).toHaveLength(allTools.length);
  });

  it('applies the allowlist', () => {
    setEnv({ MS365_MCP_ENABLED_TOOLS: 'list-mail-messages, get-calendar-view, not-a-tool' });
    expect(getFilteredTools().map(tool => tool.name)).toEqual(['list-mail-messages', 'get-calendar-view']);
  });

  it('removes write tools in read-only mode, after the allowlist', () => {
    setEnv({
      MS365_MCP_READ_ONLY_MODE: 'true',
      MS365_MCP_ENABLED_TOOLS: 'list-mail-messages,send-mail',
    });
    expect(getFilteredTools().map(tool => tool.name)).toEqual(['list-mail-messages']);
  });

  it('requests only the scopes of the enabled tools', () => {
    setEnv({ MS365_MCP_ENABLED_TOOLS: 'list-mail-messages' });
    expect(getRequiredScopes().sort()).toEqual(['Mail.Read', 'User.Read', 'offline_access']);
  });
});
