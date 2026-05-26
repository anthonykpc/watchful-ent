# Power Automate — Monthly Reminder Flows

With the shift from SharePoint Lists + Forms to direct Excel entry,
Power Automate's role is now limited to **monthly reminders** that
nudge each role owner to update their log workbook.

No Form-to-List routing is needed. No overdue-cadence logic. Just
five scheduled emails.

---

## Option A — Power Automate (one flow, five recipients)

Build one flow. Takes ~5 minutes.

### Setup

1. Go to <https://make.powerautomate.com>.
2. Click **+ Create** → **Scheduled cloud flow**.
3. Name: `HARVEST — Monthly KPI Reminder`.
4. Set start date to the 1st of next month, recurring **monthly** on
   the **1st**.
5. Click **Create**.

### Steps

1. **Trigger**: `Recurrence` → Monthly, Day 1, 8:00 AM ET.
2. **Send an email (V2)** — Outlook 365 connector:
   - To: `logistics-manager@rwjbh.org; culinary-manager@rwjbh.org;
     outreach-specialist@rwjbh.org; director@rwjbh.org`
   - Subject: `[HARVEST] Monthly KPI log reminder — [month/year]`
   - Body:

> Time to update your KPI log for last month.
>
> Open your log workbook from the HARVEST SharePoint folder,
> add your rows for the prior month, and save.
>
> **Your log:**
> - Logistics Manager → Kitchen_Operations_Log.xlsx
> - Culinary Manager → Culinary_Training_Log.xlsx
> - Community Outreach → Community_Engagement_Log.xlsx
> - Director → Director_Partnership_Log.xlsx + UAC_Farmer_Engagement_Log.xlsx
>
> The Director will refresh the master dashboard after logs are updated.
>
> Questions? Reply to this email.

3. **Optional second reminder**: Add a **Delay** action (7 days), then
   another **Send email** with subject `[HARVEST] KPI log reminder —
   follow-up` for anyone who hasn't updated yet.

### Error handling

If the send-email action fails (e.g., mailbox unavailable), Power
Automate will show a red X in the run history. No special error
handling needed — it'll retry on the next monthly trigger.

---

## Option B — Outlook recurring calendar events (no Power Automate)

If you'd rather skip Power Automate entirely:

1. Create a recurring calendar event in Outlook:
   - Title: `Update your HARVEST KPI log`
   - Date: 1st of each month, 8:00 AM
   - Invite: all four HARVEST staff members
   - Body: same as the email body above
2. Set a reminder for 1 day before.

This is simpler and requires no Power Automate license. The trade-off
is that it's a calendar event (easy to dismiss) rather than an inbox
email. For a team of four people, either approach works.

---

## What was removed

The previous version had four types of Power Automate flows:
- Form → List routing (5 flows)
- New entrepreneur onboarding alert
- Overdue cadence flag
- Grant entered alert

These are no longer needed because:
- **No Forms or Lists** — staff enter data directly in Excel logs.
- **Onboarding alerts** — the Logistics Manager sees new members in
  their own log; no automation needed.
- **Overdue cadence** — the monthly reminder covers this.
- **Grant alerts** — the Director enters grants in their own log and
  can notify the JSC directly.

If you later want to add back an alert when a grant is entered (e.g.,
auto-post to a Teams channel when the Director adds a row with
`Entry_Type = Grant`), that requires a Power Automate flow triggered
by "When a file is modified" on the Director's log workbook, with a
condition checking whether the last row is a Grant. This is doable but
fragile with Excel triggers — revisit when the team outgrows the
monthly-reminder model.
