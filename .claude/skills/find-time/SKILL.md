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
