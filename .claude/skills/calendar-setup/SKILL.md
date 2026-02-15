---
name: calendar-setup
description: >
  Set up or troubleshoot Office 365 calendar authentication. Use when the user
  needs to log in, authenticate, reconnect, or fix auth issues.
disable-model-invocation: true
---

Help the user connect their Office 365 account to the calendar assistant.

## Check current state

First, try calling `get_my_profile`. If it succeeds, authentication is already
working. Tell the user they're connected and show their name and email.

## If authentication fails

Walk the user through these steps:

### Step 1: Azure App Registration (one-time)

If the user doesn't already have a CLIENT_ID and TENANT_ID, they need to
register an app in Azure:

1. Go to https://portal.azure.com
2. Search for "App registrations" and click it
3. Click "New registration"
4. Set:
   - Name: "Office Assistant"
   - Supported account types: "Accounts in any organizational directory and personal Microsoft accounts"
   - Redirect URI: leave blank
5. Click "Register"
6. On the overview page, copy:
   - **Application (client) ID** → this is your CLIENT_ID
   - **Directory (tenant) ID** → this is your TENANT_ID
7. Go to **Authentication** in the left menu
   - Under "Advanced settings", set **Allow public client flows** to **Yes**
   - Click Save
8. Go to **API permissions** in the left menu
   - Click "Add a permission" → "Microsoft Graph" → "Delegated permissions"
   - Add: `Calendars.ReadWrite`, `Calendars.ReadWrite.Shared`, `User.Read`
   - Click "Add permissions"

### Step 2: Configure the .env file

Create or update the `.env` file in the project root with:

```
CLIENT_ID=<paste your Application (client) ID>
TENANT_ID=<paste your Directory (tenant) ID>
```

> If using a personal Microsoft account (@outlook.com, @hotmail.com), set
> `TENANT_ID=common` instead of the Directory (tenant) ID from the portal.

### Step 3: Authenticate

Call `get_my_profile` again. The MCP server will start a device code login flow.
The user will see instructions in their terminal to:

1. Open https://microsoft.com/devicelogin in their browser
2. Enter the code shown
3. Sign in with their Microsoft account (work, school, or personal)

Once complete, their token is cached and they won't need to do this again for
about 90 days.

## Troubleshooting

- **"CLIENT_ID and TENANT_ID must be set"**: The .env file is missing or
  doesn't have the right values. Check it exists in the project root.
- **"ErrorAccessDenied"**: The app doesn't have the right permissions. Go back
  to Azure Portal → App registrations → your app → API permissions and check
  that all three permissions are listed and granted.
- **"Approval required" / admin consent screen**: The user's organisation
  requires admin approval for calendar apps. They can either ask their IT admin
  to grant consent, or use a personal Microsoft account (@outlook.com) instead
  with `TENANT_ID=common`.
- **Token expired**: Just run any calendar command and it will automatically
  re-authenticate if the refresh token is still valid. If not, the device code
  flow will start again.
