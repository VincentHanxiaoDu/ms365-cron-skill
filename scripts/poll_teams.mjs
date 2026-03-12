#!/usr/bin/env node
/**
 * MS365 Monitor — Teams Poller
 * Fetches new Teams chat messages since last check.
 * Outputs JSON for the agent to evaluate relevance.
 *
 * Usage: node poll_teams.mjs
 * Output: {"new_messages": [...], "count": N, "my_name": "..."}
 */

import fs from "fs";
import path from "path";
import os from "os";
import { fileURLToPath } from "url";
import { execFileSync } from "child_process";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const STATE_PATH = path.join(os.homedir(), ".openclaw/ms365-monitor/teams_state.json");
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

function stripHtml(text) {
  return (text || "").replace(/<[^>]+>/g, "").trim().slice(0, 400);
}

async function getMyInfo(token) {
  const result = await graphGet(token, "/me", { $select: "id,displayName,mail" });
  return {
    id: result.id || "",
    displayName: result.displayName || "",
    mail: result.mail || "",
  };
}

async function main() {
  const token = getToken();
  const state = loadState();
  const lastCheck = state.last_teams_check || "2000-01-01T00:00:00Z";
  const { id: myId, displayName: myName, mail: myEmail } = await getMyInfo(token);

  const newMessages = [];

  // 1. Chats (1:1 and group)
  const chats = await graphGet(token, "/me/chats", { $top: "50", $expand: "members" });
  for (const chat of chats.value || []) {
    const chatId = chat.id;
    const chatType = chat.chatType || "";
    let topic = chat.topic || "";
    if (!topic && chatType === "oneOnOne") {
      let members = chat.members || [];
      if (!Array.isArray(members)) {
        members = members.value || [];
      }
      const others = members
        .filter((m) => m.userId !== myId)
        .map((m) => m.displayName || "");
      topic = others[0] || "1:1 Chat";
    }

    const msgs = await graphGet(token, `/me/chats/${chatId}/messages`, {
      $top: "10",
      $orderby: "createdDateTime desc",
    });
    for (const msg of msgs.value || []) {
      const created = msg.createdDateTime || "";
      if (created <= lastCheck) continue;

      const msgFrom = msg.from || {};
      const senderId = (msgFrom.user || {}).id || "";
      if (senderId === myId) continue;

      const senderName = (msgFrom.user || {}).displayName || "Unknown";
      const body = stripHtml((msg.body || {}).content || "");
      if (!body) continue;

      const mentions = (msg.mentions || []).map(
        (m) => ((m.mentioned || {}).user || {}).id || ""
      );
      const isMentioned = mentions.includes(myId);

      newMessages.push({
        type: "chat",
        chat_topic: topic,
        chat_type: chatType,
        sender: senderName,
        time: created,
        body,
        mentioned: isMentioned,
        link: `https://teams.microsoft.com/l/message/${chatId}/${msg.id || ""}`,
      });
    }
  }

  const now = new Date().toISOString().replace(/\.\d{3}Z$/, "Z");
  state.last_teams_check = now;
  saveState(state);

  newMessages.sort((a, b) => (b.time || "").localeCompare(a.time || ""));

  console.log(
    JSON.stringify(
      {
        new_messages: newMessages,
        count: newMessages.length,
        my_name: myName,
        my_email: myEmail,
      },
      null,
      2
    )
  );
}

main();
