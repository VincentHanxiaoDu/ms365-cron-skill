#!/usr/bin/env python3
"""
MS365 Monitor — Setup Wizard
Guides user through: Azure App Registration → OAuth login → cron configuration.

Usage: python3 setup.py [--reset-auth] [--reset-all]
"""

import json
import os
import sys
import subprocess
import argparse

CONFIG_PATH = os.path.expanduser("~/.openclaw/ms365-monitor/config.json")
CACHE_PATH = os.path.expanduser("~/.openclaw/ms365-monitor/token-cache.json")
SCRIPTS_DIR = os.path.dirname(os.path.abspath(__file__))

# Default public client ID from Softeria ms-365-mcp-server (pre-registered, no Azure setup needed)
DEFAULT_CLIENT_ID = "084a3e9f-a9f4-43f7-89f9-d229cf97853e"


def load_config():
    if os.path.exists(CONFIG_PATH):
        with open(CONFIG_PATH) as f:
            return json.load(f)
    return {}


def save_config(config):
    os.makedirs(os.path.dirname(CONFIG_PATH), exist_ok=True)
    with open(CONFIG_PATH, "w") as f:
        json.dump(config, f, indent=2)
    print(f"Config saved: {CONFIG_PATH}")


def prompt(label, default=None, secret=False):
    display = f"{label}" + (f" [{default}]" if default else "") + ": "
    if secret:
        import getpass
        value = getpass.getpass(display)
    else:
        value = input(display).strip()
    return value or default


def verify_auth():
    """Run auth.py and return (success, user_info)."""
    auth_script = os.path.join(SCRIPTS_DIR, "auth.py")
    result = subprocess.run(["python3", auth_script], capture_output=True, text=True)
    if result.returncode != 0:
        return False, result.stderr.strip()

    # Get user info
    token = result.stdout.strip()
    try:
        import urllib.request
        req = urllib.request.Request(
            "https://graph.microsoft.com/v1.0/me?$select=displayName,mail",
            headers={"Authorization": f"Bearer {token}", "Accept": "application/json"}
        )
        with urllib.request.urlopen(req, timeout=10) as resp:
            import json as _json
            data = _json.loads(resp.read())
            return True, {"name": data.get("displayName", ""), "email": data.get("mail", "")}
    except Exception as e:
        return True, {"name": "", "email": "", "error": str(e)}


def main():
    parser = argparse.ArgumentParser(description="MS365 Monitor setup wizard")
    parser.add_argument("--reset-auth", action="store_true", help="Clear cached tokens and re-authenticate")
    parser.add_argument("--reset-all", action="store_true", help="Reset all config and re-run full setup")
    args = parser.parse_args()

    print("\n=== MS365 Monitor Setup ===\n")

    if args.reset_all:
        for path in [CONFIG_PATH, CACHE_PATH]:
            if os.path.exists(path):
                os.remove(path)
        print("Config and auth cleared.\n")

    if args.reset_auth:
        if os.path.exists(CACHE_PATH):
            os.remove(CACHE_PATH)
        print("Auth cache cleared.\n")

    config = load_config()

    # --- Step 1: Client ID (optional — default works for most users) ---
    if not config.get("client_id"):
        config["client_id"] = DEFAULT_CLIENT_ID
        save_config(config)
        print(f"Step 1: Using default Azure App Client ID (Softeria ms-365-mcp-server, public).")
        print(f"        To use your own app, run: setup.py --reset-all and enter a custom Client ID.\n")

    # --- Step 2: Authentication ---
    print("Step 2: Microsoft 365 Authentication")
    print("-" * 40)

    if os.path.exists(CACHE_PATH) and not args.reset_auth:
        print("Found cached credentials. Verifying...")
        ok, info = verify_auth()
        if ok:
            print(f"✓ Authenticated as: {info.get('name')} <{info.get('email')}>")
        else:
            print(f"Cached auth failed: {info}")
            print("Re-authenticating...")
            if os.path.exists(CACHE_PATH):
                os.remove(CACHE_PATH)
            ok, info = verify_auth()
            if not ok:
                print(f"Authentication failed: {info}")
                sys.exit(1)
            print(f"✓ Authenticated as: {info.get('name')} <{info.get('email')}>")
    else:
        print("Starting device code authentication...")
        ok, info = verify_auth()
        if not ok:
            print(f"Authentication failed: {info}")
            sys.exit(1)
        print(f"✓ Authenticated as: {info.get('name')} <{info.get('email')}>")

    # Save user info to config
    if isinstance(info, dict) and info.get("name"):
        config["user_name"] = info.get("name", "")
        config["user_email"] = info.get("email", "")
        save_config(config)
    print()

    # --- Step 3: Monitoring Preferences ---
    if not config.get("user_name") or args.reset_all:
        print("Step 3: Monitoring Preferences")
        print("-" * 40)
        user_name = prompt("Your full name (for relevance filtering)", config.get("user_name", ""))
        config["user_name"] = user_name

        user_email = prompt("Your work email", config.get("user_email", ""))
        config["user_email"] = user_email

        save_config(config)
        print()

    # --- Summary ---
    print("=== Setup Complete ===")
    print(f"User: {config.get('user_name')} <{config.get('user_email')}>")
    print(f"Config: {CONFIG_PATH}")
    print()
    print("Next steps (run in your OpenClaw agent):")
    print("  • Test email polling:  python3 poll_email.py")
    print("  • Test Teams polling:  python3 poll_teams.py")
    print("  • Set up cron jobs:   ask your agent to configure polling crons")
    print()
    print("The agent will create cron jobs based on your preferred frequency.")


if __name__ == "__main__":
    main()
