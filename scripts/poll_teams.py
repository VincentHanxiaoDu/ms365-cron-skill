#!/usr/bin/env python3
"""
MS365 Monitor — Teams Poller
Fetches new Teams messages (chats + channels) since last check.
Outputs JSON for the agent to evaluate relevance.

Usage: python3 poll_teams.py
Output: {"new_messages": [...], "count": N, "my_name": "..."}
"""

import json
import os
import sys
import urllib.request
import urllib.parse
import urllib.error
import re
from datetime import datetime, timezone

STATE_PATH = os.path.expanduser("~/.openclaw/ms365-monitor/teams_state.json")
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
        query = "&".join(f"{k}={urllib.parse.quote(str(v))}" for k, v in params.items())
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


def strip_html(text):
    return re.sub(r'<[^>]+>', '', text or '').strip()[:400]


def get_my_info(token):
    result = graph_get(token, "/me", {"$select": "id,displayName,mail"})
    return result.get("id", ""), result.get("displayName", ""), result.get("mail", "")


def main():
    token = get_token()
    state = load_state()
    last_check = state.get("last_teams_check", "2000-01-01T00:00:00Z")
    my_id, my_name, my_email = get_my_info(token)

    new_messages = []

    # 1. Chats (1:1 and group)
    chats = graph_get(token, "/me/chats", {"$top": "50", "$expand": "members"})
    for chat in chats.get("value", []):
        chat_id = chat.get("id")
        chat_type = chat.get("chatType", "")
        topic = chat.get("topic") or ""
        if not topic and chat_type == "oneOnOne":
            members = chat.get("members", [])
            if isinstance(members, dict):
                members = members.get("value", [])
            others = [m.get("displayName", "") for m in members if m.get("userId") != my_id]
            topic = others[0] if others else "1:1 Chat"

        msgs = graph_get(token, f"/me/chats/{chat_id}/messages", {
            "$top": "10",
            "$orderby": "createdDateTime desc",
        })
        for msg in msgs.get("value", []):
            created = msg.get("createdDateTime", "")
            if created <= last_check:
                continue
            msg_from = msg.get("from") or {}
            sender_id = (msg_from.get("user") or {}).get("id", "")
            if sender_id == my_id:
                continue
            sender_name = (msg_from.get("user") or {}).get("displayName", "Unknown")
            body = strip_html(msg.get("body", {}).get("content", ""))
            if not body:
                continue

            mentions = [(((m.get("mentioned") or {}).get("user")) or {}).get("id", "")
                        for m in (msg.get("mentions") or [])]
            is_mentioned = my_id in mentions

            new_messages.append({
                "type": "chat",
                "chat_topic": topic,
                "chat_type": chat_type,
                "sender": sender_name,
                "time": created,
                "body": body,
                "mentioned": is_mentioned,
                "link": f"https://teams.microsoft.com/l/message/{chat_id}/{msg.get('id', '')}",
            })

    # 2. Channel messages (joined teams)
    teams = graph_get(token, "/me/joinedTeams", {"$select": "id,displayName"})
    for team in teams.get("value", []):
        team_id = team.get("id")
        team_name = team.get("displayName", "")
        channels = graph_get(token, f"/teams/{team_id}/channels", {"$select": "id,displayName"})
        for channel in channels.get("value", [])[:5]:
            ch_id = channel.get("id")
            ch_name = channel.get("displayName", "")
            msgs = graph_get(token, f"/teams/{team_id}/channels/{ch_id}/messages", {
                "$top": "10",
                "$orderby": "createdDateTime desc",
            })
            for msg in msgs.get("value", []):
                created = msg.get("createdDateTime", "")
                if created <= last_check:
                    continue
                msg_from = msg.get("from") or {}
                sender_id = (msg_from.get("user") or {}).get("id", "")
                if sender_id == my_id:
                    continue
                sender_name = (msg_from.get("user") or {}).get("displayName", "Unknown")
                body = strip_html(msg.get("body", {}).get("content", ""))
                if not body:
                    continue

                mentions = [(((m.get("mentioned") or {}).get("user")) or {}).get("id", "")
                            for m in (msg.get("mentions") or [])]
                is_mentioned = my_id in mentions

                new_messages.append({
                    "type": "channel",
                    "team": team_name,
                    "channel": ch_name,
                    "sender": sender_name,
                    "time": created,
                    "body": body,
                    "mentioned": is_mentioned,
                    "link": f"https://teams.microsoft.com/l/message/{ch_id}/{msg.get('id', '')}",
                })

    now = datetime.now(timezone.utc).isoformat().replace("+00:00", "Z")
    state["last_teams_check"] = now
    save_state(state)

    new_messages.sort(key=lambda x: x.get("time", ""), reverse=True)

    print(json.dumps({
        "new_messages": new_messages,
        "count": len(new_messages),
        "my_name": my_name,
        "my_email": my_email,
    }, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
