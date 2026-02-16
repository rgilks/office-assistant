---
name: check-availability
description: >
  Check free/busy status for one or more people. Use when the user asks if
  someone is free, busy, available, or what their availability looks like.
argument-hint: "[person emails] on [date/time range]"
disable-model-invocation: true
---

Check people's availability using their Office 365 free/busy schedule.

## How to call tools

The office-assistant MCP server is already registered and running. Call tools
**directly** using the `mcp__office-assistant__<tool_name>` functions available
in your tool list. Do **NOT** use Bash, Python scripts, or subprocess calls to
invoke tools. All tool names below use their short form (e.g. `get_my_profile`);
the actual callable tool is always `mcp__office-assistant__<short_name>`.

## Process

1. Call `get_my_profile` to get the user's timezone.
2. Parse the email addresses and time range from the request.
   Convert relative dates (today/tomorrow/next week) into exact calendar dates
   in that timezone.
3. If no time range is given, default to the current business day (9am-5pm)
   in the user's timezone.
4. Call `get_free_busy` with the parameters.
5. Present results clearly for each person:
   - Their scheduled items with times and status (free/busy/tentative/out of office)
   - A plain-English summary: "Alice is free from 10-11am and 2-4pm.
     Bob is in meetings until 3pm."
6. If the user wants to book a slot based on the results, offer to create the
   meeting using `create_event`.

**Note:** This tool requires a work/school Microsoft 365 account. Personal
Microsoft accounts (outlook.com, hotmail.com, live.com) cannot check other
people's availability. If the user is on a personal account (timezone is
`null` from `get_my_profile`), explain this limitation.
