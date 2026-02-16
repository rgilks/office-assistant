---
name: calendar-setup
description: >
  Set up or troubleshoot Office 365 calendar authentication. Use when the user
  needs to log in, authenticate, reconnect, or fix auth issues.
disable-model-invocation: true
---

Help the user connect their Microsoft account to the calendar assistant.

## How to call tools

The office-assistant MCP server is already registered and running. Call tools
**directly** using the `mcp__office-assistant__<tool_name>` functions available
in your tool list. Do **NOT** use Bash, Python scripts, or subprocess calls to
invoke tools. All tool names below use their short form (e.g. `get_my_profile`);
the actual callable tool is always `mcp__office-assistant__<short_name>`.

## Check current state

First, try calling `get_my_profile`. If it succeeds, authentication is already
working. Tell the user they're connected and show their name and email.

## If authentication fails

### Quick setup

Tell the user to run this single command in a terminal:

```
uv run python -m office_assistant.setup
```

It will walk them through everything interactively: creating the `.env` file
and signing in. They just need an Azure App Registration first (see below).

### Re-authentication (token expired, `.env` already exists)

When the `.env` file already exists but the token has expired, the setup script
skips the interactive `.env` prompts and goes straight to the device code flow.
In this case you can run it programmatically:

Run `uv run python -m office_assistant.setup` in the background with a long
timeout (`300000` ms / 5 minutes) using `run_in_background: true`. After it
starts, read its output to get the device code and URL, then show those to
the user. The script blocks until the user completes browser sign-in and caches
the token. Once the background task completes, authentication is restored.

### Azure App Registration (one-time prerequisite)

The user needs to create an app in Azure before they can authenticate.
**Ask first**: are you using a **work/school** account or a **personal**
Microsoft account (outlook.com, hotmail.com, live.com)?

1. Go to https://portal.azure.com
2. Search for "App registrations" → click it → "New registration"
3. Set:
   - Name: `Office Assistant`
   - Supported account types:
     - **Work/school**: "Accounts in this organizational directory only"
     - **Personal**: "Personal Microsoft accounts only"
   - Redirect URI: leave blank
4. Click "Register"
5. Copy the **Application (client) ID** from the overview page
6. Go to **Authentication** → set **Allow public client flows** to **Yes** → Save
7. Go to **API permissions** → "Add a permission" → Microsoft Graph → Delegated:
   - **Work/school**: `Calendars.ReadWrite`, `Calendars.ReadWrite.Shared`, `Place.Read.All`, `User.Read`
   - **Personal**: `Calendars.ReadWrite` and `User.Read` only
   - Click "Add permissions"

> **Tip:** Work/school accounts may need an Azure AD admin to grant consent.
> If you can't get admin approval, use a personal account instead — it doesn't
> require any admin.

Then run `uv run python -m office_assistant.setup` and follow the prompts.

## Troubleshooting

- **"CLIENT_ID and TENANT_ID must be set"**: Run
  `uv run python -m office_assistant.setup` to create the `.env` file.
- **"Application is configured for use by Microsoft Account users only"** /
  **"AADSTS9002346"**: Set `TENANT_ID=consumers` in `.env` for personal
  accounts.
- **Device code expires immediately**: The requested scopes may be wrong.
  Personal accounts don't support `Calendars.ReadWrite.Shared` — make sure
  you're on the latest code.
- **"ErrorAccessDenied"**: Check API permissions in the Azure Portal.
- **"Approval required" / admin consent screen**: The user's organisation
  requires admin approval. They can ask their IT admin to grant consent, or
  use a personal Microsoft account instead.
- **Want to start fresh?**: Run `uv run python -m office_assistant.setup` —
  it will detect the expired token and walk you through signing in again.
