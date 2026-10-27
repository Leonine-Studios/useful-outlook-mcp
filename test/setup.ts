import { afterEach } from 'vitest';
import { resetConfig } from '../src/config.js';
import { setLogLevel } from '../src/utils/logger.js';

const baseEnv = {
  MS365_MCP_CLIENT_ID: 'test-client-id',
  MS365_MCP_TENANT_ID: 'test-tenant',
  MS365_MCP_LOG_LEVEL: 'error',
  MS365_MCP_USE_TONL: 'false',
};

function applyBaseEnv(): void {
  for (const key of Object.keys(process.env)) {
    if (key.startsWith('MS365_MCP_')) {
      delete process.env[key];
    }
  }
  Object.assign(process.env, baseEnv);
  resetConfig();
}

applyBaseEnv();
setLogLevel('error');
afterEach(applyBaseEnv);
