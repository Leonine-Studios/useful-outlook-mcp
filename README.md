# Outlook OAuth MCP Server

A minimal, spec-compliant MCP server for Microsoft Outlook with OAuth2 delegated access.

## Features

- **MCP Spec Compliant**: Implements RFC 9728 and RFC 8414
- **OAuth2 Delegated Access**: Users authenticate with their own Microsoft accounts
- **Stateless Design**: No token storage—tokens passed per-request
- **Rate Limiting**: Configurable per-user rate limiting

## Quick Start

### Prerequisites

- Node.js >= 20
- Azure AD App Registration with delegated permissions

### Installation

```bash
npm install
npm run build
```

### Configuration

Create a `.env` file:

```bash
MS365_MCP_CLIENT_ID=your-azure-ad-client-id
MS365_MCP_TENANT_ID=your-tenant-id  # or 'common' for multi-tenant
MS365_MCP_CORS_ORIGIN=https://your-app.com  # set in production
```

### Run

```bash
npm start
```

Server runs at `http://localhost:3000`

## Azure AD Setup

### 1. Create App Registration

1. Go to [Azure Portal](https://portal.azure.com) → Azure Active Directory → App registrations
2. New registration → Name: "Outlook MCP Server"
3. Choose supported account types based on your needs
4. Register

### 2. Add API Permissions

Add these delegated permissions: `User.Read`, `Mail.Read`, `Mail.ReadWrite`, `Mail.Send`, `Calendars.Read`, `Calendars.ReadWrite`, `offline_access`

### 3. Configure Redirect URIs

Add platform: Web

- `http://localhost:6274/oauth/callback` (MCP Inspector)
- `https://your-production-app.com/callback` (Production)

### 4. Get Credentials

Copy from Overview page:
- Application (client) ID → `MS365_MCP_CLIENT_ID`
- Directory (tenant) ID → `MS365_MCP_TENANT_ID`

## Environment Variables

| Variable | Required | Default | Description |
|----------|----------|---------|-------------|
| `MS365_MCP_CLIENT_ID` | Yes | - | Azure AD client ID |
| `MS365_MCP_CLIENT_SECRET` | No | - | Azure AD client secret |
| `MS365_MCP_TENANT_ID` | No | `common` | Azure AD tenant ID |
| `MS365_MCP_PORT` | No | `3000` | Server port |
| `MS365_MCP_HOST` | No | `0.0.0.0` | Bind address |
| `MS365_MCP_LOG_LEVEL` | No | `info` | Log level |
| `MS365_MCP_CORS_ORIGIN` | No | `*` | CORS allowed origins |
| `MS365_MCP_RATE_LIMIT_REQUESTS` | No | `30` | Max requests per window |
| `MS365_MCP_RATE_LIMIT_WINDOW_MS` | No | `60000` | Rate limit window (ms) |
| `MS365_MCP_ALLOWED_TENANTS` | No | - | Comma-separated tenant IDs |

## Docker

```bash
docker build -t outlook-oauth-mcp .

docker run -p 3000:3000 \
  -e MS365_MCP_CLIENT_ID=your-client-id \
  -e MS365_MCP_TENANT_ID=your-tenant-id \
  -e MS365_MCP_CORS_ORIGIN=https://your-app.com \
  outlook-oauth-mcp
```

## Testing with MCP Inspector

```bash
npm run dev
npx @modelcontextprotocol/inspector
```

Configure: Server URL `http://localhost:3000/mcp`

## Production Checklist

- [ ] Deploy behind HTTPS reverse proxy
- [ ] Set `MS365_MCP_CORS_ORIGIN` to your domain
- [ ] Set `MS365_MCP_ALLOWED_TENANTS` for multi-tenant
- [ ] Use client secret for confidential client flow
- [ ] Monitor `/health` endpoint

## License

MIT
