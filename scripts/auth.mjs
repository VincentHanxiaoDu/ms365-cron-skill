#!/usr/bin/env node
/**
 * MS365 Monitor — Token Manager
 * Uses @azure/msal-node for device code auth + silent refresh.
 * Tokens stored in ~/.openclaw/ms365-monitor/token-cache.json
 */

import { PublicClientApplication } from "@azure/msal-node";
import fs from "fs";
import path from "path";
import os from "os";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url);

const CACHE_PATH = path.join(os.homedir(), ".openclaw/ms365-monitor/token-cache.json");
const CONFIG_PATH = path.join(os.homedir(), ".openclaw/ms365-monitor/config.json");

const DEFAULT_CLIENT_ID = "084a3e9f-a9f4-43f7-89f9-d229cf97853e";
const AUTHORITY = "https://login.microsoftonline.com/common";
// Use the full Softeria scope set — admin consent was granted org-wide for these.
// Individual users can self-consent via the device code "Continue" screen.
const SCOPES = [
  "Mail.ReadWrite", "Mail.Read.Shared", "Mail.Send", "Mail.Send.Shared",
  "Calendars.ReadWrite", "Calendars.Read.Shared",
  "Contacts.ReadWrite",
  "Files.ReadWrite", "Files.Read.All",
  "Tasks.ReadWrite",
  "Notes.Create", "Notes.Read",
  "Sites.Read.All",
  "People.Read",
  "User.Read", "User.Read.All",
  "Chat.Read", "ChatMessage.Read", "ChatMessage.Send",
  "Channel.ReadBasic.All", "ChannelMessage.Read.All", "ChannelMessage.Send",
  "Team.ReadBasic.All", "TeamMember.Read.All",
  "Group.Read.All", "Group.ReadWrite.All",
  "OnlineMeetings.Read", "OnlineMeetingTranscript.Read.All",
  "openid", "profile", "email", "offline_access",
];

function loadConfig() {
  try {
    if (fs.existsSync(CONFIG_PATH)) return JSON.parse(fs.readFileSync(CONFIG_PATH, "utf-8"));
  } catch {}
  return {};
}

function loadCacheData() {
  try {
    if (fs.existsSync(CACHE_PATH)) return fs.readFileSync(CACHE_PATH, "utf-8");
  } catch {}
  return null;
}

function saveCacheData(data) {
  fs.mkdirSync(path.dirname(CACHE_PATH), { recursive: true });
  fs.writeFileSync(CACHE_PATH, data);
}

function buildApp(clientId) {
  const cachePlugin = {
    beforeCacheAccess: async (ctx) => {
      const data = loadCacheData();
      if (data) ctx.tokenCache.deserialize(data);
    },
    afterCacheAccess: async (ctx) => {
      if (ctx.cacheHasChanged) saveCacheData(ctx.tokenCache.serialize());
    },
  };

  return new PublicClientApplication({
    auth: { clientId, authority: AUTHORITY },
    cache: { cachePlugin },
  });
}

export async function getToken() {
  const config = loadConfig();
  const clientId = config.client_id || DEFAULT_CLIENT_ID;
  const app = buildApp(clientId);

  // Try silent first
  const accounts = await app.getTokenCache().getAllAccounts();
  if (accounts.length > 0) {
    try {
      const result = await app.acquireTokenSilent({ scopes: SCOPES, account: accounts[0] });
      if (result?.accessToken) return result.accessToken;
    } catch {}
  }

  // Device code flow
  const result = await app.acquireTokenByDeviceCode({
    scopes: SCOPES,
    deviceCodeCallback: (response) => {
      process.stderr.write("\n" + response.message + "\n\n");
    },
  });

  if (!result?.accessToken) throw new Error("Authentication failed");
  return result.accessToken;
}

// When run directly, print token to stdout
const isMain = process.argv[1] && fs.existsSync(process.argv[1]) &&
  fs.realpathSync(process.argv[1]) === fs.realpathSync(__filename);

if (isMain) {
  getToken()
    .then((token) => process.stdout.write(token + "\n"))
    .catch((e) => { process.stderr.write("Error: " + e.message + "\n"); process.exit(1); });
}
