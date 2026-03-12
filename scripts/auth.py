#!/usr/bin/env python3
"""
MS365 Monitor — Token Manager
Handles device code auth flow and token refresh for Microsoft Graph API.
Tokens stored in ~/.openclaw/ms365-monitor/token-cache.json
"""

import json
import os
import sys
import time
import urllib.request
import urllib.parse

CACHE_PATH = os.path.expanduser("~/.openclaw/ms365-monitor/token-cache.json")
CONFIG_PATH = os.path.expanduser("~/.openclaw/ms365-monitor/config.json")

AUTHORITY = "https://login.microsoftonline.com/common/oauth2/v2.0"
SCOPES = "https://graph.microsoft.com/Mail.Read https://graph.microsoft.com/Chat.Read https://graph.microsoft.com/ChannelMessage.Read.All https://graph.microsoft.com/User.Read offline_access"


def load_config():
    if os.path.exists(CONFIG_PATH):
        with open(CONFIG_PATH) as f:
            return json.load(f)
    return {}


def load_cache():
    if os.path.exists(CACHE_PATH):
        with open(CACHE_PATH) as f:
            return json.load(f)
    return {}


def save_cache(cache):
    os.makedirs(os.path.dirname(CACHE_PATH), exist_ok=True)
    with open(CACHE_PATH, "w") as f:
        json.dump(cache, f, indent=2)


def http_post(url, data):
    body = urllib.parse.urlencode(data).encode()
    req = urllib.request.Request(url, data=body, headers={
        "Content-Type": "application/x-www-form-urlencoded"
    })
    with urllib.request.urlopen(req, timeout=30) as resp:
        return json.loads(resp.read())


def get_token():
    config = load_config()
    client_id = config.get("client_id")
    if not client_id:
        print("Error: Not configured. Run setup.py first.", file=sys.stderr)
        sys.exit(1)

    cache = load_cache()
    now = int(time.time())

    # 1. Try cached access token (with 5 min buffer)
    if cache.get("access_token") and cache.get("expires_at", 0) > now + 300:
        return cache["access_token"]

    # 2. Try refresh token
    if cache.get("refresh_token"):
        try:
            result = http_post(f"{AUTHORITY}/token", {
                "client_id": client_id,
                "grant_type": "refresh_token",
                "refresh_token": cache["refresh_token"],
                "scope": SCOPES,
            })
            if result.get("access_token"):
                cache["access_token"] = result["access_token"]
                cache["expires_at"] = now + result.get("expires_in", 3600)
                if result.get("refresh_token"):
                    cache["refresh_token"] = result["refresh_token"]
                save_cache(cache)
                return cache["access_token"]
        except Exception as e:
            print(f"Refresh failed: {e}", file=sys.stderr)

    # 3. Device code flow
    print("Authentication required. Starting device code flow...", file=sys.stderr)
    result = http_post(f"{AUTHORITY}/devicecode", {
        "client_id": client_id,
        "scope": SCOPES,
    })

    print(f"\n{result.get('message', 'Visit the URL and enter the code shown.')}\n", file=sys.stderr)
    device_code = result["device_code"]
    interval = result.get("interval", 5)
    expires_in = result.get("expires_in", 900)
    deadline = now + expires_in

    while time.time() < deadline:
        time.sleep(interval)
        try:
            token_result = http_post(f"{AUTHORITY}/token", {
                "client_id": client_id,
                "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
                "device_code": device_code,
            })
            if token_result.get("access_token"):
                cache["access_token"] = token_result["access_token"]
                cache["expires_at"] = int(time.time()) + token_result.get("expires_in", 3600)
                cache["refresh_token"] = token_result.get("refresh_token", "")
                save_cache(cache)
                print("Authentication successful.", file=sys.stderr)
                return cache["access_token"]
            err = token_result.get("error", "")
            if err == "authorization_pending":
                continue
            elif err == "slow_down":
                interval += 5
            else:
                print(f"Auth error: {err}: {token_result.get('error_description', '')}", file=sys.stderr)
                sys.exit(1)
        except Exception as e:
            print(f"Polling error: {e}", file=sys.stderr)

    print("Authentication timed out.", file=sys.stderr)
    sys.exit(1)


if __name__ == "__main__":
    print(get_token())
