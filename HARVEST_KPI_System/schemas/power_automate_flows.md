# Power Automate Flows — Specs

Minimal but essential automation. Build these in Power Automate
(<https://make.powerautomate.com>) using the **HARVEST service account**
or the Director's account — not a personal account, so flows survive
staff turnover.

All flows in this doc use **standard connectors only** (no premium tier
required): SharePoint, Microsoft Forms, Outlook 365, Teams, Schedule.

---

## Flow 1A — Kitchen Operations Log → Lists

**Trigger.** When a new response is submitted (Microsoft Forms)
**Source.** Kitchen Operations Log Form
**Destinations.** Food Entrepreneurs List (always); Environmental List
(conditionally, when waste/composted > 0)

### Steps

1. **Trigger**: `When a new response is submitted` → Form Id =
   `HARVEST — Kitchen Operations Log`.
2. **Get response details**: `Get response details` → Form Id same;
   Response Id = `List of response notifications Response Id` from
   trigger.
3. **Compose Title**: `Compose` → value:
   `"Kitchen Ops — week of " & [Reporting week start date]`
4. **Create item — Food Entrepreneurs List**:
   - Site Address: HARVEST SharePoint site
   - List Name: `Food Entrepreneurs List`
   - Field map (Form question → List column):

| Form question | List column |
|---|---|
| Reporting week — start date | `Reporting_Period_Start` |
| Reporting week — end date | `Reporting_Period_End` |
| (Composed Title) | `Title` |
| (Constant: "Operational Snapshot") | `Stage` |
| (Composed: type + active count + new count + equipment summary) | `Notes` |
| OFSA dimension(s) | `OFSA_Dimension` |
| (Constant: "Processing") | `Food_System_Sector` |
| (Constant: "Food Enterprise") | `Partner_Sector` |
| (Constant: "HARVEST Lead") | `Lead_Role` |
| Submitted by | `Submitted_By` |
| utcNow() | `Submitted_Date` |

5. **Condition**: `If "New members this week" > 0`:
   - **If yes**: Create another item in `Food Entrepreneurs List` with
     `Stage = "Onboarding"`, `Business_Name = [New member business name]`,
     `Demographic_Tags = [multi-select]`, `Date_Engaged = today`.
   - **If no**: skip.
6. **Condition**: `If "Food waste diverted this week" > 0 OR "Composted this week" > 0`:
   - **If yes**: Create item in `Environmental List` with
     `Activity_Type = "Food Waste Diverted"` (or `"Composted"`),
     `lbs_Value = [value]`, `Activity_Date = end of reporting week`,
     `OFSA_Dimension = ["Sustainability"]`.
7. **Error handling**: Configure the Create Item step's "Configure run
   after" to also run on **has failed** → email the Logistics Manager
   with the response ID and error details.

---

## Flow 1B — Culinary & Training Log → Lists

**Trigger.** New response on `HARVEST — Culinary & Training Log`
**Destinations.** Food Entrepreneurs List; Environmental List (if
food captured for demo)

### Steps

1. Trigger + Get response details (same pattern as 1A).
2. **Compose Title**: `[Event type] — [Event title/topic] — [Event date]`
3. **Create item — Food Entrepreneurs List**:

| Form question | List column |
|---|---|
| Event date | `Reporting_Period_Start` and `_End` |
| (Composed Title) | `Title` |
| 1:1 TA recipient business OR Wellness coaching recipient business | `Business_Name` (if applicable) |
| (Constant: "Active Member" if business specified, else blank) | `Stage` |
| Wellness coaching count → `+1` to running count | `Wellness_Coaching_Sessions` (via Update Item lookup) |
| (Composed: type, duration, attendance, ServSafe pass/fail) | `Notes` |
| OFSA dimension(s) | `OFSA_Dimension` |
| (Constant: "Processing") | `Food_System_Sector` |
| (Constant: "Education") | `Partner_Sector` |
| (Constant: "HARVEST Lead") | `Lead_Role` |
| Submitted by | `Submitted_By` |

4. **Condition**: `If "Food captured for demo — lbs" > 0`:
   - Create item in `Environmental List`:
     `Activity_Type = "Food Captured for Demo"`,
     `lbs_Value = [lbs]`, `Dollar_Value = [$ value]`,
     `Activity_Date = event date`,
     `OFSA_Dimension = ["Sustainability", "Utilization"]`.
5. Error handling per 1A.

---

## Flow 1C — Community Engagement Log → Lists

**Trigger.** New response on `HARVEST — Community Engagement Log`
**Destinations.** Community Orgs List; Environmental List (if food recovered)

### Steps

1. Trigger + Get response details.
2. **Create item — Community Orgs List**: map all questions directly to
   the columns named in `microsoft_forms.md` Q-to-List table for Form 3.
3. **Condition**: `If "Food recovered to direct service provider" > 0`:
   - Create item in `Environmental List`:
     `Activity_Type = "Food Waste Diverted"`,
     `Disposition = "Donated"`,
     `lbs_Value = [lbs]`,
     `Source_or_Recipient = [recipient]`,
     `Activity_Date = event date`,
     `OFSA_Dimension = ["Sustainability", "Access"]`.
4. Error handling.

---

## Flow 1D — Director & Partnership Log → Lists

**Trigger.** New response on `HARVEST — Director & Partnership Log`
**Destinations.** Investors-Funders List or Institution List, branched
by submission type.

### Steps

1. Trigger + Get response details.
2. **Switch on submission type** (Q1):
   - **Grant or Partnership** | **JSC Meeting** | **CAB Activity** |
     **Monthly Roll-Up** | **Professional Development** → Create item
     in **Investors-Funders List** with `Activity_Type` set per
     branch and field map per the question-to-column table.
   - **Institutional Connection** → Create item in **Institution List**.
3. **Set `Activity_Type` per branch** (overrides the Form's Q5):

| Form Q1 value | List | Activity_Type |
|---|---|---|
| Grant or Partnership | Investors-Funders | (use Form Q5) |
| JSC Meeting | Investors-Funders | `JSC Meeting` |
| CAB Activity | Investors-Funders | `CAB Meeting` if "Meeting Held" else `CAB Member Added` |
| Monthly Roll-Up | Investors-Funders | `Earned Revenue Snapshot` |
| Professional Development | Investors-Funders | `Professional Development` |
| Institutional Connection | Institution | (use Form Q5) |

4. Error handling.

---

## Flow 1E — UAC Farmer Engagement Log → Farmers List

**Trigger.** New response on `HARVEST UAC — Monthly Farmer Engagement Log`
**Destination.** Farmers List

### Steps

1. Trigger + Get response details.
2. **Create item — Farmers List**: map every question to its column per
   the Form 5 table. Set:
   - `Title = [Farm name] + " — " + [Reporting month]`
   - `Lead_Role = "UAC Lead"` (constant)
   - `Food_System_Sector = "Aggregation"` (constant)
   - `Partner_Sector = "Agriculture"` (constant)
   - `Submitted_By = [Q1 "Your name"]`
3. Error handling.

---

## Flow 2 — New Entrepreneur Onboarding Alert

**Trigger.** When an item is created (SharePoint)
**Source.** Food Entrepreneurs List, filter `Stage = "Onboarding"`

### Steps

1. **Trigger**: `When an item is created` → Site = HARVEST site,
   List Name = `Food Entrepreneurs List`.
2. **Condition**: `Stage equals "Onboarding"`. If no → terminate.
3. **Send email** (Outlook 365):
   - To: Director, Logistics Manager
   - Subject: `New member onboarding — [Business_Name]`
   - Body:
     > A new member business has started onboarding at HARVEST.
     >
     > **Business**: [Business_Name]
     > **Demographics**: [Demographic_Tags]
     > **Engaged**: [Date_Engaged]
     > **Submitted by**: [Submitted_By]
     >
     > Direct link: [Link to item]

---

## Flow 3 — Overdue Cadence Reminder

**Trigger.** Scheduled — weekly, Monday 8:00 AM
**Purpose.** Email each Form owner when their Form hasn't been
submitted within its expected window.

### Cadence windows

| Form | Cadence | Alert window |
|---|---|---|
| Kitchen Operations Log | Weekly | 10 days |
| Culinary & Training Log | Per event (no alert — opt out) | — |
| Community Engagement Log | Per event (no alert — opt out) | — |
| Director & Partnership Log | Monthly | 35 days |
| UAC Farmer Engagement Log | Monthly | 35 days |

### Steps

1. **Trigger**: `Recurrence` → Weekly, Monday 08:00.
2. For each Form with an alert window:
   - **Get items**: SharePoint List that this Form writes to, filtered to
     latest `Submitted_Date` for that Form (use a `Submitted_By` or a
     dedicated `Source_Form` column if you add one — recommended).
   - **Condition**: `today() - max(Submitted_Date) > [window days]`.
   - **If yes**: Send email to owner role:
     > **Reminder.** The [Form name] hasn't received a submission in over
     > [N] days. Expected cadence is [cadence]. Submit at: [Form link].

3. **Recommended schema addition**: Add a `Source_Form` choice column to
   each List, set automatically by Flows 1A–1E. Makes this flow trivial
   to implement and lets the dashboard count submissions by Form.

---

## Flow 4 — Grant Entered Alert

**Trigger.** When an item is created (SharePoint)
**Source.** Investors-Funders List, filter `Activity_Type = "Grant"`

### Steps

1. **Trigger**: `When an item is created` → Site, List =
   `Investors-Funders List`.
2. **Condition**: `Activity_Type equals "Grant"`. If no → terminate.
3. **Post message in a chat or channel** (Teams):
   - Team: HARVEST
   - Channel: JSC Updates (create if it doesn't exist)
   - Message:
     > 🌱 **New grant logged** — [Funder_Name]
     > **Amount**: $[Amount_$]
     > **Period**: [Period_Start] → [Period_End]
     > **Lead**: [Lead_Role]
     > **Notes**: [Notes]
     > Submitted by [Submitted_By] on [Submitted_Date]
4. **Also send email** (fallback for anyone not on Teams):
   - To: JSC distribution list (or Director + UAC Co-Lead)
   - Subject: `[HARVEST] New grant logged — [Funder_Name]`
   - Body: same content as Teams message

(If you don't want the emoji or RWJBH IT strips it from the channel
post, remove it — pure text works fine.)

---

## Field mapping reference — full

For each Form, the complete question-to-column map lives in
`microsoft_forms.md`. The flow steps above reference those maps by Form
name; if you change a question label or a column name, update both
files.

## Error handling — general pattern

For every Create Item action:
1. Click the `…` menu on the action → **Configure run after** → check
   **has failed**, **is skipped**, **has timed out**.
2. Add a parallel branch with `Send an email (V2)` that emails the
   owner role with:
   - The Run URL of the flow
   - The Response Id (so you can pull it from Forms manually)
   - The error message

This prevents silent data loss when Power Automate has hiccups (which
it will, periodically).

## Testing checklist

After building each flow:
1. Submit a test response via the Form (use "Test Test" as your name).
2. Within 60 seconds, confirm the row appears in the destination List.
3. Confirm all fields populated as expected.
4. For conditional branches (Flow 1A's waste insert, Flow 1B's demo
   insert, Flow 1D's switch), submit one test response per branch.
5. Delete test rows from the List when done.
