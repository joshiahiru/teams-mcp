# teams-mcp

An MCP server that posts messages to a Microsoft Teams channel via an Incoming Webhook or Power Automate Workflow webhook URL.

## Requirements

- Node.js 18+
- A Teams channel webhook URL. Two flavors are supported:
  - **Workflow webhook** (recommended) — set up via Teams → channel → "Workflows" → template "Post to a channel when a webhook request is received". Microsoft is retiring the classic Office 365 Connectors, so new setups should use Workflows.
  - **Classic Office 365 Connector** Incoming Webhook — still works in many tenants but is being phased out.

A webhook URL is bound to **one specific channel**. This server cannot DM users or send to multiple channels with a single URL — if you need that, use the Microsoft Graph API instead.

## Install / build

```bash
npm install
npm run build
```

## Run

```bash
TEAMS_WEBHOOK_URL='https://prod-XX.westeurope.logic.azure.com/...' npm start
```

The server speaks MCP over stdio.

## MCP client config

Add to your MCP client config (e.g. Claude Desktop / Claude Code):

```json
{
  "mcpServers": {
    "teams": {
      "command": "node",
      "args": ["/absolute/path/to/teams-mcp/dist/index.js"],
      "env": {
        "TEAMS_WEBHOOK_URL": "https://prod-XX.westeurope.logic.azure.com/..."
      }
    }
  }
}
```

`TEAMS_WEBHOOK_URL` is the default destination. Callers can also pass `webhook_url` per-call to override it (e.g. for posting to a different channel).

## Tools

### `send_teams_message`

Post an Adaptive Card to the channel bound to the webhook URL.

Arguments:
- `text` (string, required) — message body. Max 28000 chars.
- `title` (string, optional) — bold title rendered above the message. Max 200 chars.
- `webhook_url` (string, optional) — override the default webhook from `TEAMS_WEBHOOK_URL`. Must be `https://` and on a recognized Microsoft host (`*.webhook.office.com`, `outlook.office.com`, `*.logic.azure.com`, `*.logic.azure.us`).

### `get_configuration_status`

Reports server version and whether a (valid) webhook URL is configured. The URL is masked in the output.

## Notes & limitations

- **One-way.** Sending only — this server does not read messages.
- **One channel per webhook.** To target multiple channels, configure multiple MCP server instances or pass `webhook_url` explicitly per call.
- **No mentions / no replies.** The webhook payload posts as a generic Adaptive Card; @-mentions and threaded replies require Graph API.
- **Secrets.** The webhook URL is a bearer credential — anyone with the URL can post to the channel. Do not commit it.

## End-to-end test

After `npm run build`:

```bash
TEAMS_WEBHOOK_URL='https://...' node test-send.mjs "Hello from teams-mcp"
```
