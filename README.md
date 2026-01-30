# useful-outlook-mcp

A remote MCP server for Microsoft Outlook with proper OAuth2 and agent-optimized tools.

## Features

- **Spec Compliant**: RFC 9728/8414 OAuth metadata, RFC 7591 dynamic client registration
- **Stateless OAuth**: Tokens managed by client, not server—as OAuth intended
- **Agent-Optimized Tools**: Extensive prompt engineering in tool descriptions
- **Dynamic Scopes**: OAuth scopes auto-adjust to enabled tools
- **Production Ready**: Docker, rate limiting, read-only mode, tool filtering

## Quick Start

```bash
npm install
npm run build
npm start  # Server at http://localhost:3000
```

Required environment:
```bash
MS365_MCP_CLIENT_ID=your-azure-ad-client-id
MS365_MCP_CLIENT_SECRET=your-azure-ad-client-secret
MS365_MCP_TENANT_ID=your-tenant-id  # or 'common'
```

## Azure AD Setup

1. [Azure Portal](https://portal.azure.com) → Microsoft Entra ID → App registrations → New
2. Add delegated permissions: `User.Read`, `Mail.Read`, `Mail.ReadWrite`, `Mail.Send`, `Calendars.Read`, `Calendars.ReadWrite`, `Calendars.Read.Shared`, `Place.Read.All`, `offline_access`
3. Add redirect URI: `http://localhost:6274/oauth/callback` (for MCP Inspector)
4. Certificates & secrets → New client secret → Copy the value
5. Copy Client ID, Client Secret, and Tenant ID to your `.env`

## Tools

### Mail
`list-mail-folders` · `list-mail-messages` · `search-mail` · `get-mail-message` · `send-mail` · `create-draft-mail` · `reply-mail` · `reply-all-mail` · `create-reply-draft` · `create-reply-all-draft` · `delete-mail-message` · `move-mail-message`

### Calendar
`list-calendars` · `list-calendar-events` · `search-calendar-events` · `find-meeting-times` · `get-calendar-event` · `get-calendar-view` · `create-calendar-event` · `update-calendar-event` · `delete-calendar-event`

## Room Search

For in-person meetings, `find-meeting-times` automatically:
- Fetches all available meeting rooms across your organization
- Includes them in availability checks alongside attendees
- Groups free rooms by location (city/building)
- Returns only available rooms with email addresses for booking

Set `isOnlineMeeting: false` to enable room search. For online meetings (default: `true`), Teams meeting links are automatically generated.

**Required scope**: `Place.Read.All` (add to Azure AD app permissions)

## Configuration

| Variable | Default | Description |
|----------|---------|-------------|
| `MS365_MCP_CLIENT_ID` | required | Azure AD client ID |
| `MS365_MCP_CLIENT_SECRET` | - | Client secret (optional) |
| `MS365_MCP_TENANT_ID` | `common` | Tenant ID |
| `MS365_MCP_PORT` | `3000` | Server port |
| `MS365_MCP_READ_ONLY_MODE` | `false` | Disable write operations |
| `MS365_MCP_ENABLED_TOOLS` | all | Comma-separated tool allowlist |
| `MS365_MCP_CORS_ORIGIN` | `*` | CORS origins |
| `MS365_MCP_RATE_LIMIT_REQUESTS` | `30` | Requests per window |
| `MS365_MCP_RATE_LIMIT_WINDOW_MS` | `60000` | Window size (ms) |
| `MS365_MCP_ALLOWED_TENANTS` | - | Restrict to specific tenants |

## Docker

From GitHub Container Registry:
```bash
docker run -p 3000:3000 \
  -e MS365_MCP_CLIENT_ID=xxx \
  -e MS365_MCP_CLIENT_SECRET=xxx \
  -e MS365_MCP_TENANT_ID=xxx \
  ghcr.io/Leonine-Studios/useful-outlook-mcp:latest
```

Or build locally:
```bash
docker build -t useful-outlook-mcp .
docker run -p 3000:3000 \
  -e MS365_MCP_CLIENT_ID=xxx \
  -e MS365_MCP_CLIENT_SECRET=xxx \
  -e MS365_MCP_TENANT_ID=xxx \
  useful-outlook-mcp
```

## Endpoints

| Endpoint | Description |
|----------|-------------|
| `POST /mcp` | MCP protocol |
| `GET /health` | Health check |
| `GET /.well-known/oauth-protected-resource` | RFC 9728 |
| `GET /.well-known/oauth-authorization-server` | RFC 8414 |
| `GET /authorize` | OAuth (proxies to Microsoft) |
| `POST /token` | Token exchange (proxies to Microsoft) |
| `POST /register` | Dynamic client registration |

## Testing

```bash
npm run dev
npx @modelcontextprotocol/inspector  # Connect to http://localhost:3000/mcp
```

---

## Design Notes

<details>
<summary>Why another Outlook MCP server?</summary>

### Problems with existing servers

1. **Legacy architecture**: Built as stdio with HTTP bolted on. This is HTTP-native.

2. **OAuth done wrong**: Most servers store tokens server-side. This server is stateless—tokens passed per-request via Authorization header, never stored.

3. **Tools without thought**: Typical servers map API endpoints 1:1 without guidance. Agents fail in practical use because they don't know API quirks or multi-step workflows.

### What's different

Every tool includes:
- When to use it vs alternatives
- Known Graph API quirks (there are many)
- Workflow guidance for multi-step tasks
- Parameter combinations that fail

Example: `find-meeting-times` explains that email addresses are required (names don't work), how to find emails from names using `search-mail`, what `OrganizerUnavailable` means, and that `isOrganizerOptional=true` needs user confirmation.

### Known Graph API quirks

- Mail sender filtering: `eq` on `from/emailAddress/address` is unreliable—uses `startswith()`
- Mail recipient filtering: Can't filter to/cc/bcc with `$filter`—must use `$search`
- Calendar organizer filtering: `$filter` on organizer email returns 500—filtered client-side
- Concurrency: Parallel calls can return `MailboxConcurrency` errors
- Search + sort: `$search` can't combine with `$orderby`

</details>

## License

MIT
