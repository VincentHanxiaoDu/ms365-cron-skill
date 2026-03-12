# ms365-cron-skill

An OpenClaw skill that monitors Microsoft 365 email and Microsoft Teams, pushing relevant notifications to the user via Telegram, Discord, or any connected channel.

## Features

- 📧 **Email monitoring** — polls your inbox every 15 minutes, alerts on emails that need your attention
- 💬 **Teams monitoring** — polls chats and channels every 5 minutes, alerts on mentions and direct messages
- 🔐 **Standalone auth** — device code OAuth flow, no MCP server dependency
- ⚙️ **Easy setup** — guided wizard for Azure App Registration + cron configuration
- 🔔 **Smart filtering** — only notifies for relevant messages, suppresses noise

## Installation

Install via [ClaWHub](https://clawhub.com) or copy the `.skill` file to your OpenClaw skills directory.

Then tell your agent:
> "Set up MS365 monitoring"

The skill will guide you through the rest.

## Requirements

- Python 3.8+
- A free [Azure App Registration](references/azure-setup.md) (Microsoft account)
- OpenClaw with a connected channel (Telegram / Discord / etc.)

## File Structure

```
ms365-monitor/
├── SKILL.md                  # Skill definition + agent instructions
├── scripts/
│   ├── setup.py              # One-time setup wizard
│   ├── auth.py               # Token manager (device code + auto-refresh)
│   ├── poll_email.py         # Email poller
│   └── poll_teams.py         # Teams poller
└── references/
    └── azure-setup.md        # Azure App Registration guide
```

## License

MIT
