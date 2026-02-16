---
name: find-time
description: >
  Find available meeting times across multiple people. Use when the user asks
  to find a time that works, check when people are free, or schedule across
  multiple calendars.
argument-hint: "[attendee emails] for [duration]"
disable-model-invocation: true
---

Find available meeting times for a group of people using Office 365.

## How to call tools

The office-assistant MCP server is already registered and running. Call tools
**directly** using the `mcp__office-assistant__<tool_name>` functions available
in your tool list. Do **NOT** use Bash, Python scripts, or subprocess calls to
invoke tools. All tool names below use their short form (e.g. `get_my_profile`);
the actual callable tool is always `mcp__office-assistant__<short_name>`.

## Process

1. Call `get_my_profile` to get the user's timezone.
2. Parse attendee email addresses and desired meeting duration from the request.
   Resolve any relative date words into absolute dates in that timezone before
   calling tools.
3. If no time window is specified, search the next 5 business days.
4. Call `find_meeting_times` with the parameters.
5. Present the suggested times clearly:
   - Date and time (in the user's timezone)
   - Duration
   - Confidence percentage
   - Each attendee's availability for that slot
6. If no times are found, explain why using `emptySuggestionsReason` and suggest:
   - Widening the time window
   - Shortening the meeting duration
   - Making some attendees optional
7. If the user picks a slot, offer to create the meeting using `create_event`.

**Note:** This tool requires a work/school Microsoft 365 account. Personal
Microsoft accounts (outlook.com, hotmail.com, live.com) cannot use the
`find_meeting_times` API. If the user is on a personal account (timezone is
`null` from `get_my_profile`), explain this limitation.
