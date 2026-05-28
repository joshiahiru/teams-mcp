// One-off end-to-end test: spawns the built MCP server and sends a
// tools/call for send_teams_message. Prints the response and exits.
//
// Usage:
//   TEAMS_WEBHOOK_URL=https://... node test-send.mjs "Hello from teams-mcp"
//   node test-send.mjs "Hello" "Optional title" https://full-webhook-url
import { spawn } from "node:child_process";
import { fileURLToPath } from "node:url";
import path from "node:path";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const SERVER = path.join(__dirname, "dist", "index.js");

const MESSAGE = process.argv[2] ?? "Test message from teams-mcp end-to-end test.";
const TITLE = process.argv[3];
const WEBHOOK_ARG = process.argv[4];
const WEBHOOK = WEBHOOK_ARG || process.env.TEAMS_WEBHOOK_URL;

if (!WEBHOOK) {
  console.error(
    "usage: node test-send.mjs [message] [title] [webhook-url]\n" +
      "       or set TEAMS_WEBHOOK_URL in the environment",
  );
  process.exit(1);
}

const child = spawn("node", [SERVER], {
  stdio: ["pipe", "pipe", "inherit"],
  env: { ...process.env, TEAMS_WEBHOOK_URL: WEBHOOK },
});

let buf = "";
const pending = new Map(); // id -> resolver

child.stdout.on("data", (chunk) => {
  buf += chunk.toString();
  let nl;
  while ((nl = buf.indexOf("\n")) >= 0) {
    const line = buf.slice(0, nl).trim();
    buf = buf.slice(nl + 1);
    if (!line) continue;
    let msg;
    try {
      msg = JSON.parse(line);
    } catch {
      console.error("[server non-json]", line);
      continue;
    }
    if (msg.id != null && pending.has(msg.id)) {
      pending.get(msg.id)(msg);
      pending.delete(msg.id);
    } else {
      console.error("[server notify]", JSON.stringify(msg));
    }
  }
});

function send(obj) {
  child.stdin.write(JSON.stringify(obj) + "\n");
}

function request(id, method, params) {
  return new Promise((resolve) => {
    pending.set(id, resolve);
    send({ jsonrpc: "2.0", id, method, params });
  });
}

(async () => {
  const init = await request(1, "initialize", {
    protocolVersion: "2025-03-26",
    capabilities: {},
    clientInfo: { name: "test-send", version: "0.0.1" },
  });
  console.log("[initialize]", JSON.stringify(init.result ?? init.error));

  send({ jsonrpc: "2.0", method: "notifications/initialized" });

  const list = await request(2, "tools/list", {});
  console.log(
    "[tools]",
    (list.result?.tools ?? []).map((t) => t.name).join(", "),
  );

  const status = await request(3, "tools/call", {
    name: "get_configuration_status",
    arguments: {},
  });
  console.log("[status]", status.result?.content?.[0]?.text ?? JSON.stringify(status));

  const sendArgs = { text: MESSAGE };
  if (TITLE) sendArgs.title = TITLE;
  const send_res = await request(4, "tools/call", {
    name: "send_teams_message",
    arguments: sendArgs,
  });
  console.log(
    "[send]",
    send_res.result
      ? send_res.result.content?.[0]?.text
      : JSON.stringify(send_res.error),
  );

  child.kill();
  process.exit(send_res.error ? 1 : 0);
})();
