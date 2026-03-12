#!/usr/bin/env python3
"""
MS365 Monitor — Email Poller
Fetches unread inbox emails since last check via Microsoft Graph.
Outputs JSON for the agent to evaluate relevance.

Usage: python3 poll_email.py
Output: {"new_emails": [...], "count": N}
"""

import json
import os
import sys
import urllib.request
import urllib.parse
import urllib.error
from datetime import datetime, timezone

STATE_PATH = os.path.expanduser("~/.openclaw/ms365-monitor/email_state.json")
AUTH_SCRIPT = os.path.join(os.path.dirname(__file__), "auth.py")
GRAPH_BASE = "https://graph.microsoft.com/v1.0"


def get_token():
    import subprocess
    result = subprocess.run(["python3", AUTH_SCRIPT], capture_output=True, text=True)
    if result.returncode != 0:
        print(f"Auth error: {result.stderr}", file=sys.stderr)
        sys.exit(1)
    return result.stdout.strip()


def graph_get(token, path, params=None):
    url = f"{GRAPH_BASE}{path}"
    if params:
        query = "&".join(f"{k}={urllib.parse.quote(str(v), safe='')}" for k, v in params.items())
        url = f"{url}?{query}"
    req = urllib.request.Request(url, headers={
        "Authorization": f"Bearer {token}",
        "Accept": "application/json",
    })
    try:
        with urllib.request.urlopen(req, timeout=15) as resp:
            return json.loads(resp.read())
    except urllib.error.HTTPError as e:
        return {"error": e.read().decode(), "code": e.code}


def load_state():
    if os.path.exists(STATE_PATH):
        with open(STATE_PATH) as f:
            return json.load(f)
    return {}


def save_state(state):
    os.makedirs(os.path.dirname(STATE_PATH), exist_ok=True)
    with open(STATE_PATH, "w") as f:
        json.dump(state, f)


def main():
    token = get_token()
    state = load_state()
    last_check = state.get("last_email_check", "2000-01-01T00:00:00Z")

    filter_query = f"isRead eq false and receivedDateTime gt {last_check}"
    result = graph_get(token, "/me/mailFolders/inbox/messages", {
        "$top": "20",
        "$filter": filter_query,
        "$select": "id,subject,from,receivedDateTime,importance,bodyPreview,webLink",
        "$orderby": "receivedDateTime desc",
    })

    now = datetime.now(timezone.utc).isoformat().replace("+00:00", "Z")
    state["last_email_check"] = now
    save_state(state)

    messages = result.get("value", [])
    if not messages:
        print(json.dumps({"new_emails": [], "count": 0}))
        return

    emails = []
    for msg in messages:
        emails.append({
            "id": msg.get("id"),
            "subject": msg.get("subject", "(no subject)"),
            "from": msg.get("from", {}).get("emailAddress", {}).get("address", ""),
            "from_name": msg.get("from", {}).get("emailAddress", {}).get("name", ""),
            "received": msg.get("receivedDateTime"),
            "importance": msg.get("importance", "normal"),
            "preview": msg.get("bodyPreview", "")[:300],
            "link": msg.get("webLink", ""),
        })

    print(json.dumps({"new_emails": emails, "count": len(emails)}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
