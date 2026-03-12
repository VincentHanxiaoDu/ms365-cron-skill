# Azure App Registration Setup

## Why you need this

MS365 Monitor uses Microsoft Graph API to read your email and Teams messages. You need a free Azure App Registration to authenticate.

## Steps

### 1. Create the app

1. Go to [portal.azure.com](https://portal.azure.com) → **Azure Active Directory** → **App registrations** → **New registration**
2. Name: `MS365 Monitor` (or anything)
3. Supported account types: **"Accounts in any organizational directory and personal Microsoft accounts"**
4. Redirect URI: Leave blank (not needed for device code flow)
5. Click **Register**

### 2. Copy the Client ID

On the app overview page, copy the **Application (client) ID** — this is what `setup.py` asks for.

### 3. Enable public client flow

Go to **Authentication** → scroll down → enable **"Allow public client flows"** → Save.

### 4. Add API permissions

Go to **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated permissions**, add:

| Permission | Purpose |
|---|---|
| `Mail.Read` | Read inbox emails |
| `Chat.Read` | Read Teams chat messages |
| `ChannelMessage.Read.All` | Read Teams channel messages |
| `User.Read` | Get your profile info |
| `offline_access` | Keep you logged in with refresh tokens |

Click **Grant admin consent** (or ask your tenant admin if required).

## Notes

- No client secret needed — this uses device code flow (public client)
- Tokens are cached in `~/.openclaw/ms365-monitor/token-cache.json`
- Run `setup.py --reset-auth` to re-authenticate without changing config
- Run `setup.py --reset-all` to start over completely

## Troubleshooting

**"Need admin approval"**: Your organization may require admin consent for `ChannelMessage.Read.All`. Ask your IT admin to grant it, or use the Azure portal to pre-consent.

**"Invalid client"**: Double-check the Client ID and ensure "Allow public client flows" is enabled.

**Token expires**: The script auto-refreshes. If refresh fails, run `setup.py --reset-auth`.
