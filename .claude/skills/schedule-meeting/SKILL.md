---
name: schedule-meeting
description: >
  Quickly schedule a new meeting. Use when the user wants to create a meeting,
  send a meeting invite, or book time on a calendar.
argument-hint: "[subject] with [attendees] at [time]"
disable-model-invocation: true
---

Create a meeting on the user's Office 365 calendar.

## How to call tools

The office-assistant MCP server is already registered and running. Call tools
**directly** using the `mcp__office-assistant__<tool_name>` functions available
in your tool list. Do **NOT** use Bash, Python scripts, or subprocess calls to
invoke tools. All tool names below use their short form (e.g. `get_my_profile`);
the actual callable tool is always `mcp__office-assistant__<short_name>`.

## Process

1. Call `get_my_profile` to get the user's timezone.
2. Parse the user's request for: subject, attendees (emails), start time,
   end time or duration, location, and whether it should be a Teams meeting.
3. If any required info is missing (at minimum: subject and start time), ask.
   - For **recurring meetings**, ask about the pattern (daily, weekly, monthly),
     frequency, and end condition (end date, number of occurrences, or no end).
   - For **room bookings** (work/school only), use `list_rooms` to find rooms
     or ask which room to book.
4. Default to a 30-minute meeting if no duration or end time is given.
5. For **work/school accounts**, default to a Teams meeting
   (`is_online_meeting=true`) unless told otherwise.
   For **personal accounts** (timezone is `null` from `get_my_profile`),
   always set `is_online_meeting=false` — Teams meetings are not supported.
6. If the user is scheduling **on behalf of someone else** (e.g. "Schedule a
   meeting for Sarah"), use the `user_email` parameter with that person's email.
   This requires delegate access (work/school accounts only).
7. Show the user a summary:
   - Subject
   - Date and time (in their timezone)
   - Duration
   - Recurrence (if applicable)
   - Attendees (if any)
   - Room (if any)
   - Location or "Teams meeting"
8. Ask "Shall I go ahead and create this?"
9. On confirmation, call `create_event` with the parameters.
10. Report success: show the created event details including any Teams link.
