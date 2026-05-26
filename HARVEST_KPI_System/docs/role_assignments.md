# Role Assignments — Who Enters What

Each staff member has their own Excel log workbook. Open it from
SharePoint/OneDrive, add rows, save. The master workbook reads from
all five logs via Power Query.

## Quick reference

| Role | Log workbook | Cadence | ~Time per entry |
|---|---|---|---|
| Logistics Manager | Kitchen_Operations_Log.xlsx | Weekly + incidents | ~5 min/week |
| Culinary Manager | Culinary_Training_Log.xlsx | Per workshop/training | ~3 min/event |
| Community Outreach Specialist | Community_Engagement_Log.xlsx | Per event | ~3 min/event |
| Director | Director_Partnership_Log.xlsx | Monthly + event-based | ~10 min/month |
| Director (for UAC) | UAC_Farmer_Engagement_Log.xlsx | Monthly | ~10 min/month |

The Director also refreshes the master workbook before each JSC meeting.

---

## By role

### Logistics Manager — Kitchen_Operations_Log.xlsx

**Weekly:**
- Active member count, new members this week
- Equipment status, any incidents
- Food waste diverted (lbs), composted (lbs)

**Do NOT enter here:** Kitchen rental hours — those come from the
Food Corridor CSV in the master workbook.

**Monthly task:** Export the Food Corridor CSV and give it to the
Director (or drop it into the master workbook's Food Corridor Import tab).

---

### Culinary Manager — Culinary_Training_Log.xlsx

**Per workshop, training, or TA session:**
- Event type, title, curriculum (for health demos), duration, attendance
- ServSafe pass/fail counts (when applicable)
- 1:1 TA topic and recipient business
- RD in-kind hours (hours Registered Dietitians provided free demos)
- Food captured for demos (lbs, $ value)

**New fields:**
- **Curriculum** — name of the curriculum used for health/nutrition
  demos (dropdown with common options, or type a new one). This feeds
  the "Health demo curriculums delivered" KPI.
- **RD_InKind_Hours** — hours the RD provided demos at no charge to
  the community. This is the in-kind hour tracking KPI.

---

### Community Outreach Specialist — Community_Engagement_Log.xlsx

**Per event (two tabs to fill in):**

**Log tab** — one row per event:
- Event name, date, type, attendance
- Volunteer hours, orgs engaged, sector count
- Multilingual/cultural flag and language
- Food recovery to Bridges or other providers (lbs, recipient)
- Direct service provider presence

**Survey Responses tab** — one row per survey question per event:
- Event name, survey instrument name
- Question text, response text, numeric score
- Add new survey questions as new rows — the structure never changes.

---

### Director — Director_Partnership_Log.xlsx

**Monthly:**
- JSC meetings held (date, notes)
- CAB recruitment (members added, sector)
- CAB meetings held
- Professional development activity
- Earned revenue % (quarterly)
- Avg % sales increase for members (quarterly)

**As it happens:**
- Grants, partnerships, sponsorships ($ amount, funder, period)
- Institutional connections (Rutgers, NPS, Audible, etc.)
- Purchase orders ($ value, institution)
- Menu placements

### Director — UAC_Farmer_Engagement_Log.xlsx

**Monthly (entered by Director on behalf of UAC):**
- One row per farm per month
- Produce moved (type, lbs, $), origin, distribution destinations
- Value-added products
- Farmer demographics

Source: monthly call or email with the UAC contact.

---

## Refresh and review rhythm

| When | Who | What |
|---|---|---|
| Weekly | Logistics Manager | Add rows to Kitchen Operations Log |
| Per event | Culinary / Outreach | Add rows within 48 hours |
| Monthly, 1st week | Director | Add rows to Director + UAC logs; get Food Corridor CSV |
| Monthly, 2nd week | Director | Open master workbook, Data > Refresh All, review dashboards |
| Before each JSC | Director | Refresh, screenshot 2026 Goals + Executive Summary |
| Quarterly | Director | Update earned revenue %, member sales growth |

## Monthly reminders

A Power Automate flow (or recurring calendar event) emails all staff
on the 1st of each month: "Time to update your KPI log." Setup in
`schemas/monthly_reminders.md`.
