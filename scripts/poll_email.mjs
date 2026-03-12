#!/usr/bin/env node
/**
 * MS365 Monitor — Email Poller
 * Fetches unread inbox emails since last check via Microsoft Graph.
 * Outputs JSON for the agent to evaluate relevance.
 *
 * Usage: node poll_email.mjs
 * Output: {"new_emails": [...], "count": N}
 */

import fs from "fs";
import path from "path";
import os from "os";
import { fileURLToPath } from "url";
import { execFileSync } from "child_process";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const STATE_PATH = path.join(os.homedir(), ".openclaw/ms365-monitor/email_state.json");
const AUTH_SCRIPT = path.join(__dirname, "auth.mjs");
const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

function getToken() {
  const result = execFileSync(process.execPath, [AUTH_SCRIPT], {
    encoding: "utf-8",
    stdio: ["pipe", "pipe", "pipe"],
  });
  return result.trim();
}

async function graphGet(token, apiPath, params) {
  let url = `${GRAPH_BASE}${apiPath}`;
  if (params) {
    const query = Object.entries(params)
      .map(([k, v]) => `${k}=${encodeURIComponent(String(v))}`)
      .join("&");
    url = `${url}?${query}`;
  }
  try {
    const resp = await fetch(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: "application/json",
      },
      signal: AbortSignal.timeout(15000),
    });
    if (!resp.ok) {
      const text = await resp.text();
      return { error: text, code: resp.status };
    }
    return resp.json();
  } catch (e) {
    return { error: String(e) };
  }
}

function loadState() {
  try {
    if (fs.existsSync(STATE_PATH)) {
      return JSON.parse(fs.readFileSync(STATE_PATH, "utf-8"));
    }
  } catch {}
  return {};
}

function saveState(state) {
  fs.mkdirSync(path.dirname(STATE_PATH), { recursive: true });
  fs.writeFileSync(STATE_PATH, JSON.stringify(state));
}

async function main() {
  const token = getToken();
  const state = loadState();
  const lastCheck = state.last_email_check || "2000-01-01T00:00:00Z";

  const filterQuery = `isRead eq false and receivedDateTime gt ${lastCheck}`;
  const result = await graphGet(token, "/me/mailFolders/inbox/messages", {
    $top: "20",
    $filter: filterQuery,
    $select: "id,subject,from,receivedDateTime,importance,bodyPreview,webLink",
    $orderby: "receivedDateTime desc",
  });

  const now = new Date().toISOString().replace(/\.\d{3}Z$/, "Z");
  state.last_email_check = now;
  saveState(state);

  const messages = result.value || [];
  if (!messages.length) {
    console.log(JSON.stringify({ new_emails: [], count: 0 }));
    return;
  }

  const emails = messages.map((msg) => ({
    id: msg.id,
    subject: msg.subject || "(no subject)",
    from: (msg.from?.emailAddress?.address) || "",
    from_name: (msg.from?.emailAddress?.name) || "",
    received: msg.receivedDateTime,
    importance: msg.importance || "normal",
    preview: (msg.bodyPreview || "").slice(0, 300),
    link: msg.webLink || "",
  }));

  console.log(JSON.stringify({ new_emails: emails, count: emails.length }, null, 2));
}

main();
