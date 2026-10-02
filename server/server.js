#!/usr/bin/env node
// Mobilise Mac Mailer - MCP server for sending and drafting email through
// Microsoft Outlook on macOS. No third-party dependencies: it speaks the MCP
// stdio protocol directly and drives Outlook with AppleScript (osascript).
//
// Nothing here talks to the network. Email leaves the Mac only through
// Outlook, from the account Outlook uses by default, to the recipients the
// tool call names.

"use strict";

const { execFile } = require("node:child_process");
const fs = require("node:fs");
const fsp = require("node:fs/promises");
const os = require("node:os");
const path = require("node:path");
const readline = require("node:readline");

const SERVER_NAME = "mobilise-mac-mailer";
const SERVER_VERSION = "1.0.0";
const SUPPORTED_PROTOCOLS = ["2025-06-18", "2025-03-26", "2024-11-05"];

const BRIDGE = path.join(__dirname, "outlook.applescript");
const DATA_DIR = expandHome(process.env.MAC_MAILER_DATA_DIR || "~/.mobilise-mac-mailer");
const ALLOW_FILE = path.join(DATA_DIR, "allowed-recipients.txt");
const LOG_FILE = path.join(DATA_DIR, "activity-log.tsv");
const MAX_SENDS_PER_HOUR = positiveInt(process.env.MAC_MAILER_MAX_SENDS_PER_HOUR, 20);
const MAX_RECIPIENTS = 50;
const MAX_ATTACHMENT_BYTES = 25 * 1024 * 1024;

const EMAIL_RE = /^[^\s@,;<>"'()\[\]\\]+@[^\s@,;<>"'()\[\]\\]+\.[^\s@,;<>"'()\[\]\\]+$/;

// ---------------------------------------------------------------- helpers

function expandHome(p) {
  if (p === "~") return os.homedir();
  if (p.startsWith("~/")) return path.join(os.homedir(), p.slice(2));
  return p;
}

function positiveInt(value, fallback) {
  const n = Number.parseInt(value, 10);
  return Number.isFinite(n) && n > 0 ? n : fallback;
}

class ToolError extends Error {}

function escapeHtml(s) {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function textToHtml(text) {
  return text
    .split(/\r?\n\s*\r?\n/)
    .map((para) => `<p>${escapeHtml(para).replace(/\r?\n/g, "<br>")}</p>`)
    .join("");
}

function asList(value, field) {
  if (value === undefined || value === null) return [];
  if (!Array.isArray(value)) throw new ToolError(`'${field}' must be a list.`);
  return value.map((v) => {
    if (typeof v !== "string") throw new ToolError(`Every entry in '${field}' must be text.`);
    return v.trim();
  }).filter(Boolean);
}

async function readAllowList() {
  let raw;
  try {
    raw = await fsp.readFile(ALLOW_FILE, "utf8");
  } catch {
    return null; // no file means no restriction
  }
  const entries = raw
    .split(/\r?\n/)
    .map((l) => l.trim().toLowerCase())
    .filter((l) => l && !l.startsWith("#"));
  return entries.length ? entries : null;
}

function isAllowed(address, allowList) {
  return allowList.some((entry) =>
    entry.startsWith("@") ? address.endsWith(entry) : address === entry
  );
}

async function readLogLines() {
  try {
    const raw = await fsp.readFile(LOG_FILE, "utf8");
    return raw.split(/\r?\n/).filter((l) => l.trim() && !l.startsWith("timestamp\t"));
  } catch {
    return [];
  }
}

async function sendsInLastHour() {
  const cutoff = Date.now() - 60 * 60 * 1000;
  const lines = await readLogLines();
  return lines.filter((l) => {
    const [stamp, action] = l.split("\t");
    return action === "send" && Date.parse(stamp) >= cutoff;
  }).length;
}

async function appendLog(action, message) {
  try {
    await fsp.mkdir(DATA_DIR, { recursive: true });
    if (!fs.existsSync(LOG_FILE)) {
      await fsp.writeFile(LOG_FILE, "timestamp\taction\tto\tcc\tbcc_count\tsubject\tattachments\n", "utf8");
    }
    const clean = (s) => String(s).replace(/[\t\r\n]+/g, " ");
    const row = [
      new Date().toISOString(),
      action,
      message.to.join(","),
      message.cc.join(","),
      String(message.bcc.length),
      clean(message.subject),
      String(message.attachments.length),
    ].join("\t");
    await fsp.appendFile(LOG_FILE, row + "\n", "utf8");
  } catch {
    // The log must never block an email.
  }
}

function runOsascript(args, timeoutMs = 120000) {
  return new Promise((resolve, reject) => {
    execFile("/usr/bin/osascript", args, { timeout: timeoutMs }, (err, stdout, stderr) => {
      if (err) {
        const detail = (stderr || err.message || "").toString().trim();
        reject(new ToolError(explainOsascriptError(detail)));
      } else {
        resolve(stdout.toString().trim());
      }
    });
  });
}

function explainOsascriptError(detail) {
  if (/-1743|not allowed to send Apple events|Not authorized/i.test(detail)) {
    return "macOS blocked access to Outlook. Open System Settings > Privacy & Security > Automation and allow this app to control Microsoft Outlook, then try again.";
  }
  if (/-1728|Can.t get application|-10814|isn.t running|not found/i.test(detail) && /Outlook/i.test(detail)) {
    return "Microsoft Outlook isn't installed or couldn't be found on this Mac.";
  }
  return `Outlook reported an error: ${detail}`;
}

function requireMac() {
  if (process.platform !== "darwin") {
    throw new ToolError("Mobilise Mac Mailer works only on macOS with Microsoft Outlook installed.");
  }
}

// ---------------------------------------------------------------- core

async function prepareMessage(args) {
  if (!args || typeof args !== "object") throw new ToolError("Missing arguments.");
  const to = asList(args.to, "to").map((a) => a.toLowerCase());
  const cc = asList(args.cc, "cc").map((a) => a.toLowerCase());
  const bcc = asList(args.bcc, "bcc").map((a) => a.toLowerCase());
  const attachments = asList(args.attachments, "attachments");
  const subject = typeof args.subject === "string" ? args.subject.trim() : "";
  const body = typeof args.body === "string" ? args.body : "";
  const format = args.body_format === "html" ? "html" : "text";

  if (to.length === 0) throw new ToolError("Add at least one address in 'to'.");
  if (!subject) throw new ToolError("The subject can't be empty.");
  if (subject.length > 255) throw new ToolError("The subject is longer than 255 characters.");
  if (!body.trim()) throw new ToolError("The body can't be empty.");

  const all = [...to, ...cc, ...bcc];
  if (all.length > MAX_RECIPIENTS) throw new ToolError(`Too many recipients (limit ${MAX_RECIPIENTS}).`);
  const invalid = all.filter((a) => !EMAIL_RE.test(a));
  if (invalid.length) throw new ToolError(`Not a valid email address: ${invalid.join(", ")}`);

  const allowList = await readAllowList();
  if (allowList) {
    const blocked = [...new Set(all.filter((a) => !isAllowed(a, allowList)))];
    if (blocked.length) {
      throw new ToolError(
        `Blocked by ${ALLOW_FILE}: ${blocked.join(", ")}. ` +
          "Only the user can add addresses to that file; do not try another way to send."
      );
    }
  }

  let totalBytes = 0;
  for (const file of attachments) {
    if (!path.isAbsolute(file)) throw new ToolError(`Attachment path must be absolute: ${file}`);
    let stat;
    try {
      stat = await fsp.stat(file);
    } catch {
      throw new ToolError(`Attachment not found: ${file}`);
    }
    if (!stat.isFile()) throw new ToolError(`Attachment is not a file: ${file}`);
    totalBytes += stat.size;
  }
  if (totalBytes > MAX_ATTACHMENT_BYTES) throw new ToolError("Attachments add up to more than 25 MB.");

  return {
    to: [...new Set(to)],
    cc: [...new Set(cc)],
    bcc: [...new Set(bcc)],
    subject,
    html: format === "html" ? body : textToHtml(body),
    attachments,
  };
}

async function runBridge(mode, message) {
  requireMac();
  const dir = await fsp.mkdtemp(path.join(os.tmpdir(), "mac-mailer-"));
  try {
    const bodyFile = path.join(dir, "body.html");
    const recipientsFile = path.join(dir, "recipients.txt");
    const attachmentsFile = path.join(dir, "attachments.txt");
    await fsp.writeFile(bodyFile, message.html, "utf8");
    const lines = [
      ...message.to.map((a) => `to:${a}`),
      ...message.cc.map((a) => `cc:${a}`),
      ...message.bcc.map((a) => `bcc:${a}`),
    ];
    await fsp.writeFile(recipientsFile, lines.join("\n") + "\n", "utf8");
    await fsp.writeFile(attachmentsFile, message.attachments.join("\n") + (message.attachments.length ? "\n" : ""), "utf8");
    return await runOsascript([BRIDGE, mode, message.subject, bodyFile, recipientsFile, attachmentsFile]);
  } finally {
    await fsp.rm(dir, { recursive: true, force: true });
  }
}

function describe(message) {
  const parts = [`to ${message.to.join(", ")}`];
  if (message.cc.length) parts.push(`cc ${message.cc.join(", ")}`);
  if (message.bcc.length) parts.push(`${message.bcc.length} bcc`);
  if (message.attachments.length) parts.push(`${message.attachments.length} attachment(s)`);
  return `"${message.subject}" ${parts.join("; ")}`;
}

// ---------------------------------------------------------------- tools

const messageProperties = {
  to: { type: "array", items: { type: "string" }, minItems: 1, description: "Recipient email addresses." },
  cc: { type: "array", items: { type: "string" }, description: "CC email addresses (optional)." },
  bcc: { type: "array", items: { type: "string" }, description: "BCC email addresses (optional)." },
  subject: { type: "string", description: "Subject line (max 255 characters)." },
  body: { type: "string", description: "Message body. Outlook adds the user's signature automatically, so don't include one." },
  body_format: { type: "string", enum: ["text", "html"], default: "text", description: "'text' (blank lines start new paragraphs) or 'html'." },
  attachments: { type: "array", items: { type: "string" }, description: "Absolute paths of files on this Mac to attach (optional, 25 MB total)." },
};

const TOOLS = [
  {
    name: "send_email",
    title: "Send email with Outlook",
    description:
      "Send an email immediately through Microsoft Outlook on this Mac, from Outlook's default account with the user's signature. " +
      "Sending can't be undone: only call this after the user has approved these exact recipients, subject, body and attachments. " +
      "If the user hasn't approved the final text, use create_draft instead.",
    inputSchema: { type: "object", properties: messageProperties, required: ["to", "subject", "body"], additionalProperties: false },
    annotations: { title: "Send email with Outlook", readOnlyHint: false, destructiveHint: false, idempotentHint: false, openWorldHint: true },
    async handler(args) {
      const message = await prepareMessage(args);
      const recent = await sendsInLastHour();
      if (recent >= MAX_SENDS_PER_HOUR) {
        throw new ToolError(`Hourly limit reached (${MAX_SENDS_PER_HOUR} emails). Try again later or use create_draft.`);
      }
      await runBridge("send", message);
      await appendLog("send", message);
      return `Sent through Outlook: ${describe(message)}.`;
    },
  },
  {
    name: "create_draft",
    title: "Open a draft in Outlook",
    description:
      "Open a new email in Microsoft Outlook on this Mac with the recipients, subject, body and attachments filled in, without sending it. " +
      "The user reviews it in Outlook and sends it themselves.",
    inputSchema: { type: "object", properties: messageProperties, required: ["to", "subject", "body"], additionalProperties: false },
    annotations: { title: "Open a draft in Outlook", readOnlyHint: false, destructiveHint: false, idempotentHint: false, openWorldHint: false },
    async handler(args) {
      const message = await prepareMessage(args);
      await runBridge("draft", message);
      await appendLog("draft", message);
      return `Opened a draft in Outlook (not sent): ${describe(message)}.`;
    },
  },
  {
    name: "check_setup",
    title: "Check Outlook connection",
    description: "Check that Microsoft Outlook is installed and that macOS lets this tool control it. Sends and creates nothing.",
    inputSchema: { type: "object", properties: {}, additionalProperties: false },
    annotations: { title: "Check Outlook connection", readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: false },
    async handler() {
      requireMac();
      const version = await runOsascript(["-e", 'tell application "Microsoft Outlook" to get version'], 30000);
      const allowList = await readAllowList();
      const recent = await sendsInLastHour();
      return [
        `Outlook ${version} is reachable and macOS allows control.`,
        allowList
          ? `Recipient allowlist is ON (${allowList.length} entries in ${ALLOW_FILE}).`
          : `Recipient allowlist is OFF (create ${ALLOW_FILE} to restrict recipients).`,
        `Sends in the last hour: ${recent} of ${MAX_SENDS_PER_HOUR}.`,
      ].join("\n");
    },
  },
  {
    name: "list_allowed_recipients",
    title: "List allowed recipients",
    description: "Show the recipient allowlist, if the user has set one up.",
    inputSchema: { type: "object", properties: {}, additionalProperties: false },
    annotations: { title: "List allowed recipients", readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: false },
    async handler() {
      const allowList = await readAllowList();
      if (!allowList) {
        return `No allowlist is set, so any valid address can be used. To restrict recipients, the user can list addresses (or @domain entries) one per line in ${ALLOW_FILE}.`;
      }
      return `Allowed recipients (${ALLOW_FILE}):\n${allowList.join("\n")}`;
    },
  },
  {
    name: "recent_activity",
    title: "Show recent email activity",
    description: "Show the most recent emails sent or drafted with this plugin, newest first.",
    inputSchema: {
      type: "object",
      properties: { limit: { type: "integer", minimum: 1, maximum: 100, default: 20, description: "How many entries to show." } },
      additionalProperties: false,
    },
    annotations: { title: "Show recent email activity", readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: false },
    async handler(args) {
      const limit = Math.min(Math.max(positiveInt(args && args.limit, 20), 1), 100);
      const lines = await readLogLines();
      if (!lines.length) return "No email activity recorded yet.";
      const rows = lines.slice(-limit).reverse().map((l) => {
        const [stamp, action, to, cc, bccCount, subject, att] = l.split("\t");
        const extra = [cc ? `cc ${cc}` : "", bccCount && bccCount !== "0" ? `${bccCount} bcc` : "", att && att !== "0" ? `${att} attachment(s)` : ""]
          .filter(Boolean)
          .join("; ");
        return `${stamp}  ${action}  to ${to}${extra ? "; " + extra : ""}  "${subject}"`;
      });
      return rows.join("\n");
    },
  },
];

// ---------------------------------------------------------------- MCP stdio

function write(message) {
  process.stdout.write(JSON.stringify(message) + "\n");
}

function reply(id, result) {
  write({ jsonrpc: "2.0", id, result });
}

function replyError(id, code, message) {
  write({ jsonrpc: "2.0", id, error: { code, message } });
}

async function handle(request) {
  const { id, method, params } = request;
  const isNotification = id === undefined || id === null;

  switch (method) {
    case "initialize": {
      const asked = params && params.protocolVersion;
      reply(id, {
        protocolVersion: SUPPORTED_PROTOCOLS.includes(asked) ? asked : SUPPORTED_PROTOCOLS[0],
        capabilities: { tools: { listChanged: false } },
        serverInfo: { name: SERVER_NAME, title: "Mobilise Mac Mailer", version: SERVER_VERSION },
        instructions:
          "Sends and drafts email through Microsoft Outlook on the user's Mac. Confirm recipients, subject, body and attachments with the user before send_email; use create_draft when the user wants to review in Outlook.",
      });
      return;
    }
    case "ping":
      if (!isNotification) reply(id, {});
      return;
    case "tools/list":
      reply(id, {
        tools: TOOLS.map(({ name, title, description, inputSchema, annotations }) => ({ name, title, description, inputSchema, annotations })),
      });
      return;
    case "tools/call": {
      const tool = TOOLS.find((t) => t.name === (params && params.name));
      if (!tool) {
        replyError(id, -32602, `Unknown tool: ${params && params.name}`);
        return;
      }
      try {
        const text = await tool.handler((params && params.arguments) || {});
        reply(id, { content: [{ type: "text", text }] });
      } catch (err) {
        const text = err instanceof ToolError ? err.message : `Unexpected error: ${err && err.message}`;
        reply(id, { content: [{ type: "text", text }], isError: true });
      }
      return;
    }
    default:
      if (method && method.startsWith("notifications/")) return;
      if (!isNotification) replyError(id, -32601, `Method not found: ${method}`);
  }
}

const rl = readline.createInterface({ input: process.stdin, crlfDelay: Infinity });
rl.on("line", (line) => {
  if (!line.trim()) return;
  let request;
  try {
    request = JSON.parse(line);
  } catch {
    replyError(null, -32700, "Parse error");
    return;
  }
  const requests = Array.isArray(request) ? request : [request];
  for (const r of requests) {
    handle(r).catch((err) => {
      if (r && r.id !== undefined && r.id !== null) replyError(r.id, -32603, String(err && err.message));
    });
  }
});
rl.on("close", () => process.exit(0));
