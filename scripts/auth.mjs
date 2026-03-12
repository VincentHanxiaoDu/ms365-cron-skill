#!/usr/bin/env node
/**
 * MS365 Monitor — Token Manager
 * Handles device code auth flow and token refresh for Microsoft Graph API.
 * Tokens stored in ~/.openclaw/ms365-monitor/token-cache.json
 */

import fs from "fs";
import path from "path";
import os from "os";

const CACHE_PATH = path.join(os.homedir(), ".openclaw/ms365-monitor/token-cache.json");
const CONFIG_PATH = path.join(os.homedir(), ".openclaw/ms365-monitor/config.json");

const AUTHORITY = "https://login.microsoftonline.com/common/oauth2/v2.0";
const SCOPES = "https://graph.microsoft.com/Mail.Read https://graph.microsoft.com/Chat.Read https://graph.microsoft.com/User.Read offline_access";

// Default public client ID from Softeria ms-365-mcp-server (pre-registered, no setup needed)
const DEFAULT_CLIENT_ID = "084a3e9f-a9f4-43f7-89f9-d229cf97853e";

function loadConfig() {
  try {
    if (fs.existsSync(CONFIG_PATH)) {
      return JSON.parse(fs.readFileSync(CONFIG_PATH, "utf-8"));
    }
  } catch {}
  return {};
}

function loadCache() {
  try {
    if (fs.existsSync(CACHE_PATH)) {
      return JSON.parse(fs.readFileSync(CACHE_PATH, "utf-8"));
    }
  } catch {}
  return {};
}

function saveCache(cache) {
  fs.mkdirSync(path.dirname(CACHE_PATH), { recursive: true });
  fs.writeFileSync(CACHE_PATH, JSON.stringify(cache, null, 2));
}

async function httpPost(url, data) {
  const body = new URLSearchParams(data).toString();
  const resp = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body,
    signal: AbortSignal.timeout(30000),
  });
  return resp.json();
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

export async function getToken() {
  const config = loadConfig();
  const clientId = config.client_id || DEFAULT_CLIENT_ID;

  const cache = loadCache();
  const now = Math.floor(Date.now() / 1000);

  // 1. Try cached access token (with 5 min buffer)
  if (cache.access_token && (cache.expires_at || 0) > now + 300) {
    return cache.access_token;
  }

  // 2. Try refresh token
  if (cache.refresh_token) {
    try {
      const result = await httpPost(`${AUTHORITY}/token`, {
        client_id: clientId,
        grant_type: "refresh_token",
        refresh_token: cache.refresh_token,
        scope: SCOPES,
      });
      if (result.access_token) {
        cache.access_token = result.access_token;
        cache.expires_at = now + (result.expires_in || 3600);
        if (result.refresh_token) {
          cache.refresh_token = result.refresh_token;
        }
        saveCache(cache);
        return cache.access_token;
      }
    } catch (e) {
      process.stderr.write(`Refresh failed: ${e}\n`);
    }
  }

  // 3. Device code flow
  process.stderr.write("Authentication required. Starting device code flow...\n");
  const result = await httpPost(`${AUTHORITY}/devicecode`, {
    client_id: clientId,
    scope: SCOPES,
  });

  process.stderr.write(`\n${result.message || "Visit the URL and enter the code shown."}\n\n`);
  const deviceCode = result.device_code;
  let interval = result.interval || 5;
  const expiresIn = result.expires_in || 900;
  const deadline = now + expiresIn;

  while (Math.floor(Date.now() / 1000) < deadline) {
    await sleep(interval * 1000);
    try {
      const tokenResult = await httpPost(`${AUTHORITY}/token`, {
        client_id: clientId,
        grant_type: "urn:ietf:params:oauth:grant-type:device_code",
        device_code: deviceCode,
      });
      if (tokenResult.access_token) {
        cache.access_token = tokenResult.access_token;
        cache.expires_at = Math.floor(Date.now() / 1000) + (tokenResult.expires_in || 3600);
        cache.refresh_token = tokenResult.refresh_token || "";
        saveCache(cache);
        process.stderr.write("Authentication successful.\n");
        return cache.access_token;
      }
      const err = tokenResult.error || "";
      if (err === "authorization_pending") {
        continue;
      } else if (err === "slow_down") {
        interval += 5;
      } else {
        process.stderr.write(`Auth error: ${err}: ${tokenResult.error_description || ""}\n`);
        process.exit(1);
      }
    } catch (e) {
      process.stderr.write(`Polling error: ${e}\n`);
    }
  }

  process.stderr.write("Authentication timed out.\n");
  process.exit(1);
}

// When run directly, print the token to stdout
const isMain = process.argv[1] && fs.realpathSync(process.argv[1]) === fs.realpathSync(new URL(import.meta.url).pathname);
if (isMain) {
  getToken().then((token) => process.stdout.write(token + "\n"));
}
