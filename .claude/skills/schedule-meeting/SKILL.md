---
name: schedule-meeting
description: >
  Quickly schedule a new meeting. Use when the user wants to create a meeting,
  send a meeting invite, or book time on a calendar.
argument-hint: "[subject] with [attendees] at [time]"
disable-model-invocation: true
---

Create a meeting on the user's Office 365 calendar.

## Process

1. Call `get_my_profile` to get the user's timezone.
2. Parse the user's request for: subject, attendees (emails), start time,
   end time or duration, location, and whether it should be a Teams meeting.
3. If any required info is missing (at minimum: subject and start time), ask.
4. Default to a 30-minute meeting if no duration or end time is given.
5. Default to a Teams meeting (`is_online_meeting=true`) unless told otherwise.
6. Show the user a summary:
   - Subject
   - Date and time (in their timezone)
   - Duration
   - Attendees (if any)
   - Location or "Teams meeting"
7. Ask "Shall I go ahead and create this?"
8. On confirmation, call `create_event` with the parameters.
9. Report success: show the created event details including any Teams link.
