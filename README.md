# Office Assistant

![Office Assistant Screenshot](screenshot.png)

Manage Office 365 calendars through a chat interface powered by [Claude Code](https://docs.anthropic.com/en/docs/agents-and-tools/claude-code/overview).

Type things like **"What's on my calendar today?"** or **"Schedule a meeting with Alice tomorrow at 2pm"** and the assistant handles the rest -- creating events, checking availability, finding times that work for everyone, and more.

---

## What you'll need

Before starting, make sure you have:

- **A computer running macOS, Linux, or Windows (with WSL)**
- **A Microsoft account** -- either a work/school account (Office 365) or a personal account (@outlook.com, @hotmail.com, etc.)
- **Claude Code** installed -- [installation guide](https://docs.anthropic.com/en/docs/agents-and-tools/claude-code/overview)
- **Python 3.11 or newer** -- [download here](https://www.python.org/downloads/) if you don't have it

You'll also need to register an app in the Azure portal (Step 2 below). You can do this with any Microsoft account -- a personal @outlook.com account works fine. If you're using a work account, your organisation may restrict app registrations; ask your IT department if you run into issues.

---

## Setup (one time only)

### Step 1: Download and install

Open a terminal and run:

```bash
git clone https://github.com/rgilks/office-assistant.git
cd office-assistant
./setup.sh
```

The setup script installs everything automatically, including [uv](https://docs.astral.sh/uv/) (a Python package manager) if you don't already have it.

### Step 2: Register an app in Azure

This step tells Microsoft that the Office Assistant is allowed to access calendars on your behalf. You only need to do this once.

1. Open your web browser and go to **[portal.azure.com](https://portal.azure.com)**
2. Sign in with your **Microsoft account** (work, school, or personal)
3. In the search bar at the top, type **App registrations** and click the result

   > If you don't see "App registrations", your organisation may restrict this. Ask your IT admin to do this step for you.

4. Click the **+ New registration** button
5. Fill in the form:
   - **Name**: `Office Assistant`
   - **Supported account types**: choose based on your account:
     - **Work/school account**: "Accounts in this organizational directory only"
     - **Personal account** (@outlook.com, @hotmail.com): "Personal Microsoft accounts only"
   - **Redirect URI**: leave this blank
6. Click **Register**

You'll now see an overview page for your new app. You need two values from here:

7. Copy the **Application (client) ID** -- it looks like `a1b2c3d4-e5f6-7890-abcd-ef1234567890`
8. Copy the **Directory (tenant) ID** -- same format, right below the client ID

Now configure two more settings:

9. In the left sidebar, click **Authentication**
   - Scroll down to **Advanced settings**
   - Set **Allow public client flows** to **Yes**
   - Click **Save** at the top

10. In the left sidebar, click **API permissions**
    - Click **+ Add a permission**
    - Click **Microsoft Graph**
    - Click **Delegated permissions**
    - Search for and tick these permissions:
      - `Calendars.ReadWrite`
      - `Calendars.ReadWrite.Shared` (work/school accounts only -- skip this for personal accounts)
      - `User.Read`
    - Click **Add permissions**

### Step 3: Sign in

Run the interactive setup command:

```bash
cd office-assistant
uv run python -m office_assistant.setup
```

It will prompt you for:
- Your **Application (client) ID** (from step 7 above)
- Whether you're using a **work/school** or **personal** account

Then it will show a sign-in message like:

> To sign in, use a web browser to open https://microsoft.com/devicelogin and enter the code **ABCD1234**

1. Open that link in your browser
2. Enter the code shown
3. Sign in with your Microsoft account
4. Approve the permissions when asked

That's it! Your login is saved for about 90 days, so you won't need to do this again for a while.

---

## How to use it

Once set up, just start Claude Code in the `office-assistant` folder and chat naturally about your calendar. Here are some examples:

### View your schedule

```
What's on my calendar today?
Show me my meetings for next week
What do I have on Thursday?
```

### View someone else's calendar

```
What's on Alice's calendar tomorrow?
Show me bob@company.com's schedule for Monday
```

> The other person needs to have shared their calendar with you in Outlook for this to work.

### Create meetings

```
Schedule a meeting called "Project Review" tomorrow at 3pm
Book a 1-hour meeting with alice@company.com and bob@company.com on Friday at 10am
Set up a meeting about budget planning next Tuesday at 2pm in the Board Room
```

By default, all meetings include a Microsoft Teams link. If you want an in-person-only meeting, just say so:

```
Schedule an in-person meeting in Room 4A tomorrow at 11am
```

### Change or cancel meetings

```
Move the Project Review to 4pm
Cancel tomorrow's budget meeting
Cancel the 3pm meeting and let everyone know it's been postponed
```

### Check availability

```
Is alice@company.com free tomorrow afternoon?
Check if bob@company.com and carol@company.com are available on Thursday
```

### Find a time that works for everyone

```
Find a time for a 30-minute meeting with alice@company.com next week
When can bob@company.com, carol@company.com and I all meet for an hour?
```

### Slash commands

You can also use these shortcuts:

| Command | What it does |
|---------|-------------|
| `/calendar` | Answer questions about your calendar |
| `/schedule-meeting` | Walk you through creating a meeting step by step |
| `/find-time` | Find available meeting slots across multiple people |
| `/check-availability` | Check if people are free or busy |
| `/calendar-setup` | Help with signing in or fixing connection issues |

---

## Troubleshooting

### "CLIENT_ID and TENANT_ID must be set"

The `.env` file is missing or empty. Make sure you completed Step 3 above and that the file contains both values.

### "Could not start device-code flow"

Your Azure app may not have public client flows enabled. Go back to Azure Portal > App registrations > your app > Authentication, and make sure "Allow public client flows" is set to **Yes**.

### "You don't have permission to view this person's calendar"

The other person hasn't shared their calendar with you. Ask them to share it: in Outlook, they go to **Calendar** > right-click their calendar > **Sharing and Permissions** > add your email address.

### "ErrorAccessDenied" on any calendar operation

The Azure app is missing the required permissions. Go to Azure Portal > App registrations > your app > API permissions and check that the right permissions are listed. Note: personal accounts don't support `Calendars.ReadWrite.Shared`.

### "Approval required" or admin consent screen

Your organisation requires an admin to approve apps that access calendar data. You have two options:

1. **Ask your IT admin** to grant consent for the app in the Azure portal
2. **Use a personal Microsoft account** (@outlook.com, @hotmail.com) instead -- register a new app under that account and sign in with it

### Some features don't work with my personal account

Personal Microsoft accounts have limited Graph API support. Checking other people's availability (`/check-availability`, `/find-time`) and viewing other people's calendars are only available with work/school accounts. Your own calendar, creating events, and managing events all work fine.

### The login expired

Just use any calendar command and it will automatically refresh. If the refresh token has also expired (after ~90 days), you'll be prompted to sign in again with the device code flow.

### Something else isn't working

Run `/calendar-setup` -- it will check your connection and guide you through fixing common issues.
