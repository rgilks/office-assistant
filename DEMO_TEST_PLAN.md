# Demo Test Plan

Manual test script for verifying Office Assistant against a real Microsoft account.
Run through these prompts in order in Claude Code from the `office-assistant` folder.

> **Account type matters.** Steps marked **(org only)** will fail gracefully on
> personal accounts (@outlook.com, @hotmail.com) — that's expected. The test
> verifies the error message is clear and helpful.

---

## 1. Opening — identity and timezone

| # | Prompt | Expected |
|---|--------|----------|
| 1 | `What's on my calendar today?` | Shows your name, email, and timezone (or asks your timezone if personal account where it's `null`). Lists today's events or says the day is clear. |

---

## 2. Build up the schedule

Create test events to work with. All times should be interpreted in your timezone.

| # | Prompt | Expected |
|---|--------|----------|
| 2 | `Schedule a meeting called "Team Standup" tomorrow at 9am for 15 minutes` | Confirms: subject, date, time, 15-min duration, no Teams link (personal) or Teams link (org). Asks to go ahead. Creates on confirmation. |
| 3 | `Book a 1-hour meeting called "Project Review" with alice@example.com on Thursday at 2pm` | Confirms: subject, date, time, 1 hour, attendee listed. Creates on confirmation. No Teams link for personal accounts. |
| 4 | `Set up a weekly meeting called "Weekly 1:1" every Tuesday at 10am for the next 4 weeks` | Confirms: subject, weekly recurrence on Tuesdays, 30-min default duration, 4 occurrences. Creates on confirmation. Recurrence pattern visible in confirmation. |
| 5 | `Create an event called "Budget Sync" tomorrow at 3pm with a note: Review Q1 figures` | Confirms: subject, date, time, body text. Creates on confirmation. |
| 6 | `Schedule a "Lunch & Learn" on Friday at 12pm for 90 minutes` | Confirms: subject, date, time, 90-min duration. Creates on confirmation. |

---

## 3. View the schedule

| # | Prompt | Expected |
|---|--------|----------|
| 7 | `What does my week look like?` | Lists all events for the current week grouped by day. Should include the events just created. |
| 8 | `What do I have on Thursday?` | Filters to Thursday only. Shows the Project Review and any other Thursday events. |

---

## 4. Make changes

| # | Prompt | Expected |
|---|--------|----------|
| 9 | `Move the Budget Sync to 3:30pm instead` | Finds the event, confirms the change (old time → new time), updates on confirmation. |
| 10 | `Change the Project Review subject to "Q1 Project Review" and add a note saying "Please bring your status updates"` | Finds the event, confirms subject change and body addition, updates on confirmation. |
| 11 | `Cancel the Lunch & Learn and let attendees know it's been postponed to next month` | Finds the event, confirms cancellation, deletes on confirmation. Notes if there are no attendees to notify. |

---

## 5. Respond to invitations

| # | Prompt | Expected |
|---|--------|----------|
| 12 | `Tentatively accept the Q1 Project Review with a note saying I might be a few minutes late` | If you are the organiser, returns a clear message: "You can't respond to this event because you are the organiser." If you received it as an invitation, tentatively accepts with the comment. |

> **Tip:** To test accept/decline properly, have another account send you an
> invitation first, then accept or decline it.

---

## 6. Personal account limitations (org only features)

These should fail gracefully on personal accounts with clear error messages.
On org accounts they should work normally.

| # | Prompt | Expected |
|---|--------|----------|
| 13 | `Is alice@example.com free tomorrow afternoon?` | **Personal:** "Free/busy lookup requires a work/school Microsoft 365 account." **Org:** Shows Alice's free/busy slots. |
| 14 | `Find a time that works for me and bob@example.com next week` | **Personal:** "Finding meeting times requires a work/school Microsoft 365 account." **Org:** Suggests available slots. |
| 15 | `What meeting rooms are available?` | **Personal:** Clear error about org-only feature. **Org:** Lists rooms with names, emails, capacity. |

---

## 7. Clean up

Delete the test events created during the session.

| # | Prompt | Expected |
|---|--------|----------|
| 16 | `Delete all the meetings I created during this test — Team Standup, Weekly 1:1, Q1 Project Review, and Budget Sync` | Finds each event, confirms the list, deletes them one by one. Should not trigger re-authentication. |

> **Verify:** Check Outlook/calendar app to confirm events are gone. Outlook may
> take 30–60 seconds to sync after deletions.

---

## What to watch for

- **No re-auth loops.** You should authenticate once at the start. Org-only
  features failing on personal accounts should NOT trigger re-authentication
  prompts.
- **Correct timezone.** All dates and times should be in your timezone. The
  assistant should confirm absolute dates ("Thursday, February 19, 2026") not
  just relative ones ("Thursday").
- **No Teams links on personal accounts.** Events created on personal accounts
  should not include `isOnlineMeeting` or a Teams link.
- **Graceful personal account errors.** Steps 13–15 should return helpful
  messages, not raw API errors or auth prompts.
- **Outlook sync delay.** After creating/updating/deleting events, Outlook may
  take 30–60 seconds to reflect changes. If events seem unchanged, wait a moment
  and refresh — verify via the assistant if needed ("Show me Thursday's events
  again").
- **Organiser can't respond.** If you try to accept/decline an event you
  created, the assistant should explain that you're the organiser (step 12).
