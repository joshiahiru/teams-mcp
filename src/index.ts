#!/usr/bin/env node

import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
  Tool,
  CallToolRequest,
  CallToolResult,
  ErrorCode,
  McpError,
} from "@modelcontextprotocol/sdk/types.js";
import { VERSION } from "./version.js";

// ---------------------------------------------------------------------------
// Constants / config
// ---------------------------------------------------------------------------

const MAX_MESSAGE_LENGTH = 28_000;
const MAX_TITLE_LENGTH = 200;
const REQUEST_TIMEOUT_MS = 15_000;

// Allowlist of Microsoft host suffixes valid for a Teams webhook URL.
// Classic O365 Connector webhooks live on *.webhook.office.com /
// outlook.office.com; Power Automate Workflow webhooks live on
// *.logic.azure.com (commercial) or *.logic.azure.us (US Gov).
const ALLOWED_HOST_SUFFIXES = [
  ".webhook.office.com",
  "outlook.office.com",
  "outlook.office365.com",
  ".logic.azure.com",
  ".logic.azure.us",
];

interface SendArgs {
  webhookUrl: string;
  text: string;
  title?: string;
}

// ---------------------------------------------------------------------------
// Validation
// ---------------------------------------------------------------------------

function validateWebhookUrl(raw: unknown, source: string): string {
  if (typeof raw !== "string" || raw.trim().length === 0) {
    throw new McpError(
      ErrorCode.InvalidParams,
      `${source} is empty. Provide a Teams Incoming Webhook or Workflow URL.`,
    );
  }
  const trimmed = raw.trim();
  let url: URL;
  try {
    url = new URL(trimmed);
  } catch {
    throw new McpError(ErrorCode.InvalidParams, `${source} is not a valid URL.`);
  }
  if (url.protocol !== "https:") {
    throw new McpError(ErrorCode.InvalidParams, `${source} must use https://.`);
  }
  const host = url.hostname.toLowerCase();
  const ok = ALLOWED_HOST_SUFFIXES.some((suffix) =>
    suffix.startsWith(".") ? host.endsWith(suffix) : host === suffix,
  );
  if (!ok) {
    throw new McpError(
      ErrorCode.InvalidParams,
      `${source} host "${host}" is not a recognized Microsoft Teams webhook host. ` +
        `Expected one of: ${ALLOWED_HOST_SUFFIXES.join(", ")}.`,
    );
  }
  return trimmed;
}

function validateText(raw: unknown): string {
  if (typeof raw !== "string" || raw.length === 0) {
    throw new McpError(ErrorCode.InvalidParams, "text must be a non-empty string");
  }
  if (raw.length > MAX_MESSAGE_LENGTH) {
    throw new McpError(
      ErrorCode.InvalidParams,
      `text is too long (${raw.length} chars). Max: ${MAX_MESSAGE_LENGTH}.`,
    );
  }
  return raw;
}

function validateTitle(raw: unknown): string | undefined {
  if (raw === undefined || raw === null || raw === "") return undefined;
  if (typeof raw !== "string") {
    throw new McpError(ErrorCode.InvalidParams, "title must be a string");
  }
  if (raw.length > MAX_TITLE_LENGTH) {
    throw new McpError(
      ErrorCode.InvalidParams,
      `title is too long (${raw.length} chars). Max: ${MAX_TITLE_LENGTH}.`,
    );
  }
  return raw;
}

function parseArgs(args: Record<string, unknown>, envWebhook: string | undefined): SendArgs {
  const rawUrl = args.webhook_url ?? envWebhook;
  const source = args.webhook_url ? "webhook_url" : "TEAMS_WEBHOOK_URL env var";
  if (rawUrl === undefined || rawUrl === null) {
    throw new McpError(
      ErrorCode.InvalidParams,
      "No webhook URL provided. Set TEAMS_WEBHOOK_URL or pass webhook_url.",
    );
  }
  return {
    webhookUrl: validateWebhookUrl(rawUrl, source),
    text: validateText(args.text),
    title: validateTitle(args.title),
  };
}

// ---------------------------------------------------------------------------
// Adaptive Card payload
// ---------------------------------------------------------------------------

function buildAdaptiveCardPayload(text: string, title: string | undefined): unknown {
  const body: unknown[] = [];
  if (title) {
    body.push({
      type: "TextBlock",
      text: title,
      weight: "Bolder",
      size: "Medium",
      wrap: true,
    });
  }
  body.push({ type: "TextBlock", text, wrap: true });

  return {
    type: "message",
    attachments: [
      {
        contentType: "application/vnd.microsoft.card.adaptive",
        contentUrl: null,
        content: {
          $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
          type: "AdaptiveCard",
          version: "1.4",
          body,
        },
      },
    ],
  };
}

// ---------------------------------------------------------------------------
// HTTP send
// ---------------------------------------------------------------------------

async function postToWebhook(
  webhookUrl: string,
  payload: unknown,
): Promise<{ status: number; body: string }> {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), REQUEST_TIMEOUT_MS);
  try {
    const res = await fetch(webhookUrl, {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify(payload),
      signal: controller.signal,
    });
    const body = await res.text();
    return { status: res.status, body };
  } catch (err) {
    if (err instanceof Error && err.name === "AbortError") {
      throw new McpError(
        ErrorCode.InternalError,
        `Teams webhook request timed out after ${REQUEST_TIMEOUT_MS}ms`,
      );
    }
    const msg = err instanceof Error ? err.message : String(err);
    throw new McpError(ErrorCode.InternalError, `Teams webhook request failed: ${msg}`);
  } finally {
    clearTimeout(timer);
  }
}

function previewOf(text: string): string {
  const oneLine = text.replace(/\s+/g, " ").trim();
  return oneLine.length > 60 ? `${oneLine.slice(0, 57)}...` : oneLine;
}

function maskWebhookUrl(url: string): string {
  try {
    const u = new URL(url);
    const path = u.pathname.length > 12 ? `${u.pathname.slice(0, 8)}...` : u.pathname;
    return `${u.protocol}//${u.hostname}${path}`;
  } catch {
    return "<unparseable>";
  }
}

// ---------------------------------------------------------------------------
// Server
// ---------------------------------------------------------------------------

class TeamsMCP {
  private server: Server;

  constructor() {
    this.server = new Server(
      { name: "teams-mcp", version: VERSION },
      { capabilities: { tools: {} } },
    );
    this.setupHandlers();
  }

  private setupHandlers(): void {
    this.server.setRequestHandler(ListToolsRequestSchema, async () => ({
      tools: [
        {
          name: "send_teams_message",
          description:
            "Post a message to a Microsoft Teams channel via an Incoming Webhook or Power Automate Workflow webhook URL. " +
            "The webhook URL is bound to a single channel; this tool cannot DM users or target multiple channels. " +
            "If TEAMS_WEBHOOK_URL is set in the environment, webhook_url is optional.",
          inputSchema: {
            type: "object",
            properties: {
              text: {
                type: "string",
                description: `Message body. Plain text or basic markdown. Max ${MAX_MESSAGE_LENGTH} characters.`,
              },
              title: {
                type: "string",
                description: `Optional bold title rendered above the message. Max ${MAX_TITLE_LENGTH} characters.`,
              },
              webhook_url: {
                type: "string",
                description:
                  "Optional override of the Teams webhook URL. If omitted, the TEAMS_WEBHOOK_URL environment variable is used. " +
                  "Must be https and on a Microsoft host (*.webhook.office.com, outlook.office.com, *.logic.azure.com, *.logic.azure.us).",
              },
            },
            required: ["text"],
            additionalProperties: false,
          },
        },
        {
          name: "get_configuration_status",
          description:
            "Report whether this server can post messages: server version, whether a webhook URL is configured, and (masked) which host it targets.",
          inputSchema: { type: "object", properties: {}, additionalProperties: false },
        },
      ] as Tool[],
    }));

    this.server.setRequestHandler(
      CallToolRequestSchema,
      async (request: CallToolRequest): Promise<CallToolResult> => {
        try {
          switch (request.params.name) {
            case "send_teams_message":
              return await this.sendTeamsMessage(request);
            case "get_configuration_status":
              return await this.getConfigurationStatus();
            default:
              throw new McpError(
                ErrorCode.MethodNotFound,
                `Unknown tool: ${request.params.name}`,
              );
          }
        } catch (error) {
          if (error instanceof McpError) throw error;
          throw new McpError(
            ErrorCode.InternalError,
            `Tool execution failed: ${error instanceof Error ? error.message : String(error)}`,
          );
        }
      },
    );
  }

  private async sendTeamsMessage(request: CallToolRequest): Promise<CallToolResult> {
    const rawArgs = (request.params.arguments || {}) as Record<string, unknown>;
    const { webhookUrl, text, title } = parseArgs(rawArgs, process.env.TEAMS_WEBHOOK_URL);

    const payload = buildAdaptiveCardPayload(text, title);
    const { status, body } = await postToWebhook(webhookUrl, payload);

    if (status < 200 || status >= 300) {
      const snippet = body.length > 300 ? `${body.slice(0, 300)}...` : body;
      throw new McpError(
        ErrorCode.InternalError,
        `Teams webhook returned HTTP ${status}: ${snippet || "<empty body>"}`,
      );
    }

    return {
      content: [
        {
          type: "text",
          text:
            `Posted to Teams (${maskWebhookUrl(webhookUrl)}): ` +
            `"${previewOf(text)}"` +
            (title ? ` [title: "${previewOf(title)}"]` : ""),
        },
      ],
    };
  }

  private async getConfigurationStatus(): Promise<CallToolResult> {
    const envUrl = process.env.TEAMS_WEBHOOK_URL?.trim();
    const lines = [`Server: teams-mcp v${VERSION}`];

    if (!envUrl) {
      lines.push("TEAMS_WEBHOOK_URL: not set (callers must pass webhook_url explicitly)");
    } else {
      try {
        const valid = validateWebhookUrl(envUrl, "TEAMS_WEBHOOK_URL");
        lines.push(`TEAMS_WEBHOOK_URL: set, ${maskWebhookUrl(valid)}`);
      } catch (err) {
        const msg = err instanceof McpError ? err.message : String(err);
        lines.push(`TEAMS_WEBHOOK_URL: INVALID — ${msg}`);
      }
    }

    return { content: [{ type: "text", text: lines.join("\n") }] };
  }

  public async run(): Promise<void> {
    const transport = new StdioServerTransport();
    await this.server.connect(transport);
  }
}

const server = new TeamsMCP();
server.run().catch((err) => {
  console.error(err);
  process.exit(1);
});
