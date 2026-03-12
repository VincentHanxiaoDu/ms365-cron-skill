#!/usr/bin/env node
/**
 * MS365 Monitor — Setup Wizard
 * Guides user through: Azure App Registration → OAuth login → cron configuration.
 *
 * Usage: node setup.mjs [--reset-auth] [--reset-all]
 */

import fs from "fs";
import path from "path";
import os from "os";
import readline from "readline";
import { fileURLToPath } from "url";
import { execFileSync } from "child_process";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const CONFIG_PATH = path.join(os.homedir(), ".openclaw/ms365-monitor/config.json");
const CACHE_PATH = path.join(os.homedir(), ".openclaw/ms365-monitor/token-cache.json");

// Default public client ID from Softeria ms-365-mcp-server (pre-registered, no Azure setup needed)
const DEFAULT_CLIENT_ID = "084a3e9f-a9f4-43f7-89f9-d229cf97853e";

function loadConfig() {
  try {
    if (fs.existsSync(CONFIG_PATH)) {
      return JSON.parse(fs.readFileSync(CONFIG_PATH, "utf-8"));
    }
  } catch {}
  return {};
}

function saveConfig(config) {
  fs.mkdirSync(path.dirname(CONFIG_PATH), { recursive: true });
  fs.writeFileSync(CONFIG_PATH, JSON.stringify(config, null, 2));
  console.log(`Config saved: ${CONFIG_PATH}`);
}

function prompt(label, defaultValue) {
  return new Promise((resolve) => {
    const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
    const display = defaultValue ? `${label} [${defaultValue}]: ` : `${label}: `;
    rl.question(display, (answer) => {
      rl.close();
      resolve(answer.trim() || defaultValue || "");
    });
  });
}

function verifyAuth() {
  try {
    const result = execFileSync(process.execPath, [path.join(__dirname, "auth.mjs")], {
      encoding: "utf-8",
      stdio: ["inherit", "pipe", "inherit"],
    });
    const token = result.trim();

    // Get user info
    try {
      const resp = execFileSync(process.execPath, ["-e", `
        fetch("https://graph.microsoft.com/v1.0/me?$select=displayName,mail", {
          headers: { Authorization: "Bearer ${token}", Accept: "application/json" },
          signal: AbortSignal.timeout(10000),
        })
        .then(r => r.json())
        .then(d => process.stdout.write(JSON.stringify({ name: d.displayName || "", email: d.mail || "" })))
        .catch(e => process.stdout.write(JSON.stringify({ name: "", email: "", error: String(e) })));
      `], { encoding: "utf-8", stdio: ["pipe", "pipe", "pipe"] });
      return [true, JSON.parse(resp)];
    } catch (e) {
      return [true, { name: "", email: "", error: String(e) }];
    }
  } catch (e) {
    return [false, e.stderr ? e.stderr.toString().trim() : String(e)];
  }
}

function parseArgs(argv) {
  return {
    resetAuth: argv.includes("--reset-auth"),
    resetAll: argv.includes("--reset-all"),
  };
}

async function main() {
  const args = parseArgs(process.argv.slice(2));

  console.log("\n=== MS365 Monitor Setup ===\n");

  if (args.resetAll) {
    for (const p of [CONFIG_PATH, CACHE_PATH]) {
      if (fs.existsSync(p)) fs.unlinkSync(p);
    }
    console.log("Config and auth cleared.\n");
  }

  if (args.resetAuth) {
    if (fs.existsSync(CACHE_PATH)) fs.unlinkSync(CACHE_PATH);
    console.log("Auth cache cleared.\n");
  }

  let config = loadConfig();

  // --- Step 1: Client ID (optional — default works for most users) ---
  if (!config.client_id) {
    config.client_id = DEFAULT_CLIENT_ID;
    saveConfig(config);
    console.log("Step 1: Using default Azure App Client ID (Softeria ms-365-mcp-server, public).");
    console.log("        To use your own app, run: setup.mjs --reset-all and enter a custom Client ID.\n");
  }

  // --- Step 2: Authentication ---
  console.log("Step 2: Microsoft 365 Authentication");
  console.log("-".repeat(40));

  let info;
  if (fs.existsSync(CACHE_PATH) && !args.resetAuth) {
    console.log("Found cached credentials. Verifying...");
    let [ok, result] = verifyAuth();
    info = result;
    if (ok) {
      console.log(`✓ Authenticated as: ${info.name} <${info.email}>`);
    } else {
      console.log(`Cached auth failed: ${info}`);
      console.log("Re-authenticating...");
      if (fs.existsSync(CACHE_PATH)) fs.unlinkSync(CACHE_PATH);
      [ok, info] = verifyAuth();
      if (!ok) {
        console.log(`Authentication failed: ${info}`);
        process.exit(1);
      }
      console.log(`✓ Authenticated as: ${info.name} <${info.email}>`);
    }
  } else {
    console.log("Starting device code authentication...");
    const [ok, result] = verifyAuth();
    info = result;
    if (!ok) {
      console.log(`Authentication failed: ${info}`);
      process.exit(1);
    }
    console.log(`✓ Authenticated as: ${info.name} <${info.email}>`);
  }

  // Save user info to config
  if (typeof info === "object" && info.name) {
    config.user_name = info.name || "";
    config.user_email = info.email || "";
    saveConfig(config);
  }
  console.log();

  // --- Step 3: Monitoring Preferences ---
  if (!config.user_name || args.resetAll) {
    console.log("Step 3: Monitoring Preferences");
    console.log("-".repeat(40));
    config.user_name = await prompt("Your full name (for relevance filtering)", config.user_name || "");
    config.user_email = await prompt("Your work email", config.user_email || "");
    saveConfig(config);
    console.log();
  }

  // --- Summary ---
  console.log("=== Setup Complete ===");
  console.log(`User: ${config.user_name} <${config.user_email}>`);
  console.log(`Config: ${CONFIG_PATH}`);
  console.log();
  console.log("Next steps (run in your OpenClaw agent):");
  console.log("  • Test email polling:  node poll_email.mjs");
  console.log("  • Test Teams polling:  node poll_teams.mjs");
  console.log("  • Set up cron jobs:   ask your agent to configure polling crons");
  console.log();
  console.log("The agent will create cron jobs based on your preferred frequency.");
}

main();
