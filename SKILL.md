---
name: ms365-monitor
description: "Monitor Microsoft 365 email and Microsoft Teams messages, push relevant notifications to the user via their connected channel (Telegram, Discord, etc.). Use when: (1) user asks to set up MS365/Outlook/Teams monitoring or notifications, (2) user wants to connect their work email or Teams to OpenClaw, (3) user asks to configure email or Teams polling crons, (4) user wants to authenticate with Microsoft 365, (5) user asks to check their email or Teams messages. Handles initial setup, re-authentication, cron configuration, and on-demand message checks."
---

# MS365 Monitor

Monitors Microsoft 365 email and Teams, pushes relevant messages to the user's connected channel.

## Scripts

All scripts are in the `scripts/` subdirectory relative to this SKILL.md.

| Script | Purpose |
|---|---|
| `setup.mjs` | One-time setup: Azure App auth + user config |
| `auth.mjs` | Token manager (called by poll scripts) |
| `poll_email.mjs` | Fetch unread inbox emails since last check |
| `poll_teams.mjs` | Fetch new Teams chat + channel messages |

Config and state stored in `~/.openclaw/ms365-monitor/`.

## Workflow

### First-time Setup

1. Run `setup.mjs` — uses built-in public client ID (no Azure setup needed for most users), runs device code auth, saves user profile
2. Test both pollers and verify output
3. Create cron jobs (see Cron Configuration below)

> **Note:** No Azure App Registration needed. The skill uses Softeria's pre-registered public client ID (`084a3e9f-a9f4-43f7-89f9-d229cf97853e`). To use your own app, run `setup.mjs --reset-all`.

### Re-authentication

```bash
node setup.mjs --reset-auth
```

### On-demand Check

Run pollers directly and summarize output for the user:

```bash
python3 poll_email.py
python3 poll_teams.py
```

## Cron Configuration

Create two cron jobs after setup. Use the user's connected channel for delivery.

**Email cron** (every 15 min recommended):

```
payload.message: Run: python3 <SKILL_DIR>/poll_email.py

Evaluate which emails need the user's attention (direct address, approval needed, urgent, mentions user by name/email). 

Output rules:
- Relevant email found: output report (sender, subject, summary, webLink)
- Nothing relevant: output only NO_REPLY (no other text)
- Links: always use webLink from poll output (outlook.office365.com/owa/?ItemID=... format)
```

**Teams cron** (every 5 min recommended):

```
payload.message: Run: python3 <SKILL_DIR>/poll_teams.py

Evaluate which messages need the user's attention. User's full name: <NAME>, email: <EMAIL>.
Push if: message explicitly mentions user, user is @mentioned, 1:1 chat message, needs their decision/approval, urgent.
Do NOT push for: other people with same surname, group chat noise, messages not involving user.

Output rules:
- Relevant message found: output report (sender, chat/channel, summary, link)
- Nothing relevant: output only NO_REPLY (no other text)
```

Replace `<SKILL_DIR>` with the absolute path to the skill's `scripts/` directory.
Replace `<NAME>` and `<EMAIL>` with values from `~/.openclaw/ms365-monitor/config.json`.

**Delivery config** for both crons:
```json
{
  "mode": "announce",
  "channel": "<user's channel>",
  "to": "<user's chat id>",
  "bestEffort": true
}
```

## Relevance Criteria

**Email — push if:**
- Addressed directly to the user
- User mentioned by name or email in body
- Approval/decision/review requested
- High importance flag
- From a known direct contact (1:1)

**Email — skip:**
- Mass newsletters, automated notifications, CC-only
- Marketing/promotional
- System alerts not requiring action

**Teams — push if:**
- 1:1 chat message (always relevant)
- User is @mentioned
- Message explicitly uses user's name
- Decision or approval requested

**Teams — skip:**
- Group chat messages not involving user
- Messages from other people with same surname
- Automated bot messages

## Report Format

**Email:**
```
📧 **[Sender Name]** `<email>`
**Subject:** ...
**Summary:** one or two sentences
🔗 [View/Reply](<webLink>)
```

**Teams:**
```
💬 **[Sender]** (chat/channel name)
message summary
🔗 [Open in Teams](<link>)
```

## Troubleshooting

- **Auth fails**: Run `python3 setup.py --reset-auth`
- **No messages appearing**: Check `~/.openclaw/ms365-monitor/` state files — timestamps may be too recent
- **ChannelMessage.Read.All error**: Requires admin consent in Azure — see `references/azure-setup.md`
- **Reset state** (re-scan all recent messages): Delete `email_state.json` or `teams_state.json` in `~/.openclaw/ms365-monitor/`
