/**
 * Structured JSON logger for the MCP server
 */

type LogLevel = 'debug' | 'info' | 'warn' | 'error';

const LOG_LEVELS: Record<LogLevel, number> = {
  debug: 0,
  info: 1,
  warn: 2,
  error: 3,
};

let currentLevel: LogLevel = 'info';

// ANSI color codes
const colors = {
  reset: '\x1b[0m',
  bold: '\x1b[1m',
  dim: '\x1b[2m',
  cyan: '\x1b[36m',
  green: '\x1b[32m',
  yellow: '\x1b[33m',
  blue: '\x1b[34m',
  magenta: '\x1b[35m',
  gray: '\x1b[90m',
  white: '\x1b[37m',
};

/**
 * Set the minimum log level
 */
export function setLogLevel(level: LogLevel): void {
  currentLevel = level;
}

interface LogEntry {
  timestamp: string;
  level: LogLevel;
  message: string;
  [key: string]: unknown;
}

function shouldLog(level: LogLevel): boolean {
  return LOG_LEVELS[level] >= LOG_LEVELS[currentLevel];
}

function log(level: LogLevel, message: string, data?: Record<string, unknown>): void {
  if (!shouldLog(level)) {
    return;
  }

  const entry: LogEntry = {
    timestamp: new Date().toISOString(),
    level,
    message,
    ...data,
  };

  const output = JSON.stringify(entry);
  
  if (level === 'error') {
    console.error(output);
  } else if (level === 'warn') {
    console.warn(output);
  } else {
    console.log(output);
  }
}

export interface StartupConfig {
  version: string;
  host: string;
  port: number;
  logLevel: string;
  enabledTools: string[];
  disabledTools: string[];
  totalTools: number;
  rateLimitRequests: number;
  rateLimitWindowMs: number;
  allowedTenants: string | number;
  readOnlyMode: boolean;
  useTonl: boolean;
}

/**
 * Print a pretty startup banner
 */
export function printStartupBanner(config: StartupConfig): void {
  const { cyan, green, yellow, blue, magenta, gray, white, bold, dim, reset } = colors;
  
  const baseUrl = `http://${config.host === '0.0.0.0' ? 'localhost' : config.host}:${config.port}`;
  
  console.log();
  console.log(`${cyan}${bold}  ╔═══════════════════════════════════════════════════════════╗${reset}`);
  console.log(`${cyan}${bold}  ║${reset}                                                           ${cyan}${bold}║${reset}`);
  console.log(`${cyan}${bold}  ║${reset}   ${blue}${bold}📧 Outlook OAuth MCP Server${reset}                            ${cyan}${bold}║${reset}`);
  console.log(`${cyan}${bold}  ║${reset}                                                           ${cyan}${bold}║${reset}`);
  console.log(`${cyan}${bold}  ╚═══════════════════════════════════════════════════════════╝${reset}`);
  console.log();
  
  // Server info
  console.log(`${white}${bold}  Server${reset}`);
  console.log(`${gray}  ├─${reset} Version:    ${green}${config.version}${reset}`);
  console.log(`${gray}  ├─${reset} Host:       ${white}${config.host}${reset}`);
  console.log(`${gray}  ├─${reset} Port:       ${white}${config.port}${reset}`);
  console.log(`${gray}  └─${reset} Log level:  ${white}${config.logLevel}${reset}`);
  console.log();
  
  // Tools info
  console.log(`${white}${bold}  Tools${reset}`);
  console.log(`${gray}  ├─${reset} Mode:       ${config.readOnlyMode ? `${yellow}read-only${reset}` : `${green}read-write${reset}`}`);
  console.log(`${gray}  ├─${reset} Output:     ${config.useTonl ? `${magenta}TONL${reset} ${dim}(token-optimized)${reset}` : `${white}JSON${reset}`}`);
  console.log(`${gray}  ├─${reset} Count:      ${green}${config.enabledTools.length}${reset}${dim}/${config.totalTools}${reset}`);
  console.log(`${gray}  └─${reset} Enabled:    ${dim}${config.enabledTools.join(', ')}${reset}`);
  console.log();
  
  // Rate limiting
  console.log(`${white}${bold}  Rate Limiting${reset}`);
  console.log(`${gray}  ├─${reset} Requests:   ${white}${config.rateLimitRequests}${reset} ${dim}per window${reset}`);
  console.log(`${gray}  └─${reset} Window:     ${white}${config.rateLimitWindowMs / 1000}s${reset}`);
  console.log();
  
  // Security
  console.log(`${white}${bold}  Security${reset}`);
  console.log(`${gray}  └─${reset} Tenants:    ${typeof config.allowedTenants === 'number' ? `${white}${config.allowedTenants}${reset} ${dim}allowed${reset}` : `${magenta}all${reset}`}`);
  console.log();
  
  // Endpoints
  console.log(`${white}${bold}  Endpoints${reset}`);
  console.log(`${gray}  ├─${reset} MCP:        ${cyan}${baseUrl}/mcp${reset}`);
  console.log(`${gray}  ├─${reset} Health:     ${cyan}${baseUrl}/health${reset}`);
  console.log(`${gray}  └─${reset} OAuth:      ${cyan}${baseUrl}/.well-known/oauth-protected-resource${reset}`);
  console.log();
  
  console.log(`${green}${bold}  ✓ Server ready${reset}`);
  console.log();
}

export const logger = {
  debug: (message: string, data?: Record<string, unknown>) => log('debug', message, data),
  info: (message: string, data?: Record<string, unknown>) => log('info', message, data),
  warn: (message: string, data?: Record<string, unknown>) => log('warn', message, data),
  error: (message: string, data?: Record<string, unknown>) => log('error', message, data),
};

export default logger;
