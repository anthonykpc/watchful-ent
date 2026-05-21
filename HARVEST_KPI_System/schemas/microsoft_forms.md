# Microsoft Forms — Specs and Setup

Five Forms, each owned by the role that enters that data. Each Form
writes to one or more SharePoint Lists via a Power Automate flow (see
`power_automate_flows.md`).

Build each Form at <https://forms.office.com>. The owner of the Form is
the role listed; the Director should be added as a co-owner on every
Form so submissions can be managed if the role owner is out.

---

## Form 1 — Kitchen Operations Log

**Owner.** Logistics Manager
**Cadence.** Weekly, plus event-based for incidents
**Sharing.** Tenant-restricted ("Only people in my organization")
**Destination List.** Kitchen Operations Log → also writes to
**Food Entrepreneurs List** when a new member is logged
**Estimated submission time.** ~5 minutes

### Setup

1. <https://forms.office.com> → **New Form**.
2. Title: `HARVEST — Kitchen Operations Log`.
3. Description:
   > Weekly log of kitchen operations. Hours are tracked via Food
   > Corridor — do not enter rental hours here. For an incident
   > (safety, conflict, equipment), submit a new entry the same day.
4. Settings (`…` → Settings):
   - Who can fill out: **Only people in my organization can respond**
   - Record name: **On**
   - One response per person: **Off**
   - Response receipts: **On** (optional)

### Questions

| # | Question | Type | Required | Choices / notes | Maps to List column |
|---|---|---|---|---|---|
| 1 | Reporting week — start date | Date | Yes | First Monday of the week being reported | `Reporting_Period_Start` |
| 2 | Reporting week — end date | Date | Yes | Sunday of the week being reported | `Reporting_Period_End` |
| 3 | Submission type | Choice (dropdown) | Yes | Weekly Log; Incident; New Member Onboarding | Routes flow logic |
| 4 | Active members this week | Number | Yes | Integer ≥ 0 | `Title` (composed) + custom |
| 5 | New members this week | Number | Yes | Integer ≥ 0 | Triggers Food Entrepreneurs List insert |
| 6 | New member business name (if any) | Short text | No | Shown if #5 > 0 (branching) | `Business_Name` (Food Entrepreneurs List) |
| 7 | New member demographics (multi-select) | Choice (multi) | No | Woman-owned; BIPOC-owned; Immigrant-owned; Veteran-owned; LGBTQ+-owned; First-time entrepreneur; Prefer not to say | `Demographic_Tags` |
| 8 | Equipment status | Choice (dropdown) | Yes | All operational; Minor issue (logged); Major issue (offline) | `Notes` (prefixed) |
| 9 | Equipment notes | Long text | No | Shown if #8 ≠ "All operational" | `Notes` |
| 10 | Incident — type | Choice (dropdown) | No | Shown if #3 = "Incident". Safety; Sanitation; Member conflict; Other | `Notes` (prefixed) |
| 11 | Incident — description | Long text | No | Shown if #3 = "Incident" | `Notes` |
| 12 | Food waste diverted this week (lbs) | Number | No | Decimal ≥ 0 | Inserts row in Environmental List |
| 13 | Composted this week (lbs) | Number | No | Decimal ≥ 0 | Inserts row in Environmental List |
| 14 | OFSA dimension(s) for this entry | Choice (multi) | Yes | Availability; Access; Utilization; Stability; Agency; Sustainability. Default: Availability; Stability | `OFSA_Dimension` |
| 15 | Notes (anything else) | Long text | No | — | `Notes` |
| 16 | Submitted by | Short text | Yes | Pre-filled from "Record name" but allow override | `Submitted_By` |

### Branching

- Q5 > 0 → show Q6, Q7
- Q3 = "Incident" → show Q10, Q11
- Q8 ≠ "All operational" → show Q9

---

## Form 2 — Culinary & Training Log

**Owner.** Culinary Manager
**Cadence.** Per workshop or training session
**Sharing.** Tenant-restricted
**Destination List.** Food Entrepreneurs List (training records);
optional row in **Environmental List** when food is captured for demos
**Estimated submission time.** ~3 minutes per event

### Setup

1. <https://forms.office.com> → **New Form**.
2. Title: `HARVEST — Culinary & Training Log`.
3. Description:
   > Log each workshop, training, or 1:1 technical assistance session.
   > Submit within 48 hours of the event.
4. Settings: same as Form 1 (tenant-restricted, record name on).

### Questions

| # | Question | Type | Required | Choices / notes | Maps to List column |
|---|---|---|---|---|---|
| 1 | Event date | Date | Yes | — | `Reporting_Period_Start` and `_End` (same date) |
| 2 | Event type | Choice (dropdown) | Yes | ServSafe Training; Business Building Seminar; Marketing Seminar; Capital Access Seminar; 1:1 Technical Assistance; Wellness Coaching; Other Workshop | `Title` (event name) |
| 3 | Event title / topic | Short text | Yes | — | `Title` |
| 4 | Duration (hours) | Number | Yes | Decimal ≥ 0 | `Notes` |
| 5 | Attendance count | Number | Yes | Integer ≥ 0 | `Notes` (parsed) |
| 6 | Member businesses present (multi-select) | Choice (multi) | No | Populated from existing Food Entrepreneurs List businesses; allow "Other — add new" | `Business_Name` references |
| 7 | ServSafe — # who passed (if applicable) | Number | No | Shown if Q2 = "ServSafe Training" | `ServSafe_Passed` (Food Entrepreneurs List) |
| 8 | ServSafe — # who attempted (if applicable) | Number | No | Shown if Q2 = "ServSafe Training" | `ServSafe_Attempted` (Food Entrepreneurs List) |
| 9 | 1:1 TA — topic (if applicable) | Choice (dropdown) | No | Shown if Q2 = "1:1 Technical Assistance". Business building; Capital access; Marketing; Operations; Other | `Notes` |
| 10 | 1:1 TA — recipient business name (if applicable) | Short text | No | Shown if Q2 = "1:1 Technical Assistance" | `Business_Name` |
| 11 | Wellness coaching — recipient business (if applicable) | Short text | No | Shown if Q2 = "Wellness Coaching" | `Business_Name` |
| 12 | Food captured for demo — lbs (if applicable) | Number | No | If demo used food that would have been wasted | Inserts row in Environmental List |
| 13 | Food captured for demo — $ value (if applicable) | Number | No | Estimated $ value | Inserts row in Environmental List |
| 14 | OFSA dimension(s) for this entry | Choice (multi) | Yes | Default: Utilization | `OFSA_Dimension` |
| 15 | Notes | Long text | No | — | `Notes` |
| 16 | Submitted by | Short text | Yes | Pre-filled | `Submitted_By` |

### Branching

- Q2 = "ServSafe Training" → show Q7, Q8
- Q2 = "1:1 Technical Assistance" → show Q9, Q10
- Q2 = "Wellness Coaching" → show Q11

---

## Form 3 — Community Engagement Log

**Owner.** Community Outreach Specialist
**Cadence.** Per event
**Sharing.** Tenant-restricted
**Destination List.** Community Orgs List; optional row in
**Environmental List** if food recovered
**Estimated submission time.** ~3 minutes per event

### Setup

1. Title: `HARVEST — Community Engagement Log`.
2. Description:
   > Log each community event, tour, workshop, or partner activity.
   > Submit within 48 hours of the event.
3. Settings: tenant-restricted.

### Questions

| # | Question | Type | Required | Choices / notes | Maps to List column |
|---|---|---|---|---|---|
| 1 | Event name | Short text | Yes | — | `Event_Name` / `Title` |
| 2 | Event date | Date | Yes | — | `Event_Date` / `Reporting_Period_Start` / `_End` |
| 3 | Event type | Choice (dropdown) | Yes | Teaching Kitchen; RD Workshop; Tour; Partner-Hosted Event; HARVEST-Hosted Event; Pop-Up; Other | `Event_Type` |
| 4 | Attendance count | Number | Yes | Integer ≥ 0 | `Attendance_Count` |
| 5 | Volunteer hours contributed | Number | No | Decimal ≥ 0 | `Volunteer_Hours` |
| 6 | Organizations engaged (one per line) | Long text | No | Type each org on a new line | `Orgs_Engaged` |
| 7 | Number of distinct sectors represented | Number | No | Integer ≥ 0 (e.g., healthcare, faith, education, etc.) | `Sector_Count` |
| 8 | Multilingual or cultural event? | Choice (dropdown) | Yes | Yes; No | `Multilingual_Cultural` |
| 9 | Language or cultural tradition (if Yes) | Short text | No | Shown if Q8 = Yes | `Language_Tradition` |
| 10 | Survey — avg satisfaction (1–5) | Number | No | 1.0–5.0 | `Survey_Satisfaction_Avg` |
| 11 | Survey — avg learning (1–5) | Number | No | 1.0–5.0 | `Survey_Learning_Avg` |
| 12 | Food recovered to direct service provider (lbs) | Number | No | Decimal ≥ 0 | `Food_Recovery_lbs` |
| 13 | Recipient (if Q12 > 0) | Short text | No | E.g., "Bridges" | `Food_Recovery_Recipient` |
| 14 | Direct service provider had recurring presence | Choice (dropdown) | No | Yes; No | `Direct_Service_Provider_Onsite` |
| 15 | OFSA dimension(s) for this entry | Choice (multi) | Yes | Default: Utilization; Agency | `OFSA_Dimension` |
| 16 | Notes | Long text | No | — | `Notes` |
| 17 | Submitted by | Short text | Yes | Pre-filled | `Submitted_By` |

### Branching

- Q8 = "Yes" → show Q9
- Q12 > 0 → show Q13

---

## Form 4 — Director & Partnership Log

**Owner.** Director
**Cadence.** Monthly, plus event-based for grants and partnerships
**Sharing.** Tenant-restricted
**Destination List.** Investors-Funders List, Institution List, and a
**Governance** sub-list (CAB and JSC tracked via this Form — see Notes
below)
**Estimated submission time.** ~10 minutes monthly

### Note on governance data

JSC meetings, CAB members, and CAB meetings are governance data, not
audience-category KPIs. To avoid creating a seventh List, governance
entries are written to the **Investors-Funders List** with
`Activity_Type` set to a governance-specific value (`JSC Meeting`,
`CAB Meeting`, `CAB Member Added`). This keeps the schema tight and
keeps governance visible alongside the funder relationships those
meetings advance.

### Setup

1. Title: `HARVEST — Director & Partnership Log`.
2. Description:
   > Monthly governance, grants, partnerships, and institutional
   > connections. Submit grants and partnerships as soon as they're
   > secured.
3. Settings: tenant-restricted.

### Questions

| # | Question | Type | Required | Choices / notes | Maps to List column |
|---|---|---|---|---|---|
| 1 | Submission type | Choice (dropdown) | Yes | Monthly Roll-Up; Grant or Partnership; Institutional Connection; JSC Meeting; CAB Activity; Professional Development | Routes flow logic |
| 2 | Reporting period start | Date | Yes | — | `Reporting_Period_Start` |
| 3 | Reporting period end | Date | Yes | — | `Reporting_Period_End` |
| 4 | Funder / partner / institution name | Short text | No | Shown for Grant, Institutional Connection | `Funder_Name` / `Institution_Name` |
| 5 | Activity type | Choice (dropdown) | No | Shown based on Q1. Grant; Partnership; Sponsorship; Leveraged Capital; Purchase Order; Menu Placement; TA Partnership; Cross-Institution Connection | `Activity_Type` |
| 6 | $ value | Number | No | Decimal ≥ 0 | `Amount_$` or `PO_Value_$` |
| 7 | Period start (grant) | Date | No | Shown for Grant | `Period_Start` |
| 8 | Period end (grant) | Date | No | Shown for Grant | `Period_End` |
| 9 | JSC meeting date | Date | No | Shown for JSC Meeting | `Title` (composed) |
| 10 | JSC meeting agenda summary | Long text | No | Shown for JSC Meeting | `Connection_Notes` |
| 11 | CAB activity type | Choice (dropdown) | No | Shown for CAB Activity. Member Added; Meeting Held; Other | `Title` (composed) |
| 12 | CAB member name (if applicable) | Short text | No | Shown if Q11 = "Member Added" | `Notes` |
| 13 | CAB member sector | Choice (dropdown) | No | Shown if Q11 = "Member Added". Healthcare; Agriculture; Education; Government; Philanthropy; Community-Based Org; Private Sector; Food Enterprise | `Notes` |
| 14 | Items on institutional menu (count) | Number | No | Shown for Menu Placement | `Items_On_Menu` |
| 15 | Earned revenue % (quarterly) | Number | No | Shown for Monthly Roll-Up. Leave blank except quarter-end | `Earned_Revenue_%` |
| 16 | Avg % sales increase for kitchen members (quarterly) | Number | No | Shown for Monthly Roll-Up. Leave blank except quarter-end | `Sales_YoY_Change_%` (aggregate) |
| 17 | Professional development activity | Choice (dropdown) | No | Shown for Professional Development. Shared Kitchen Summit; ServSafe Manager Cert; Community Presentation; Conference; Other | `Notes` |
| 18 | OFSA dimension(s) for this entry | Choice (multi) | Yes | Default: Stability | `OFSA_Dimension` |
| 19 | Notes | Long text | No | — | `Notes` |
| 20 | Submitted by | Short text | Yes | Pre-filled | `Submitted_By` |

### Branching

All Q4–Q17 visibility is controlled by Q1 and (in some cases) Q5/Q11.
Microsoft Forms supports "Add branching" from the `…` menu — define
each visibility rule there.

---

## Form 5 — UAC Farmer Engagement Log

**Owner.** UAC Partner (external — no M365 license required)
**Cadence.** Monthly
**Sharing.** **Anyone with the link** (public, anonymous)
**Destination List.** Farmers List
**Estimated submission time.** ~10 minutes

### Setup — different from the others

1. Title: `HARVEST UAC — Monthly Farmer Engagement Log`.
2. Description:
   > Monthly log of farms engaged, produce moved, and farmer
   > demographics. You do not need an RWJBarnabas login. Submit by the
   > 5th of each month for the prior month.
3. Settings (`…` → Settings) — **this is the key difference**:
   - Who can fill out: **Anyone can respond** (anonymous link)
   - Record name: **Off** (anonymous responses)
   - Response receipts: leave off (no email captured)
4. Add the Director and Logistics Manager as **co-owners** so they can
   view responses.

### Questions

| # | Question | Type | Required | Choices / notes | Maps to List column |
|---|---|---|---|---|---|
| 1 | Your name | Short text | Yes | UAC submitter's name (since "Record name" is off) | `Submitted_By` |
| 2 | Reporting month — first day | Date | Yes | First day of the month being reported | `Reporting_Period_Start` |
| 3 | Reporting month — last day | Date | Yes | Last day of the month being reported | `Reporting_Period_End` |
| 4 | Farm name | Short text | Yes | One farm per submission. Submit the form multiple times for multiple farms. | `Farm_Name` / `Title` |
| 5 | Farm location | Short text | No | City, State | `Farm_Location` |
| 6 | Farm contact (name / email / phone) | Short text | No | — | `Farm_Contact` |
| 7 | Produce types moved (one per line) | Long text | No | E.g., tomatoes, peppers, kale | `Produce_Types` |
| 8 | Total lbs moved this month | Number | No | Decimal ≥ 0 | `Produce_lbs` |
| 9 | Total $ value moved this month | Number | No | Decimal ≥ 0 | `Produce_Value_$` |
| 10 | Origin region | Short text | No | E.g., "Sussex County, NJ" | `Origin_Region` |
| 11 | Distribution destinations (one per line) | Long text | No | E.g., RWJBH cafeteria, Newark Public Schools | `Distribution_Destinations` |
| 12 | Did this farmer use HARVEST to make a value-added product? | Choice (dropdown) | Yes | Yes; No | `Value_Added_Product` |
| 13 | Value-added product description (if Yes) | Short text | No | E.g., "Hot sauce from peppers" | `Value_Added_Description` |
| 14 | Farmer race / ethnicity (multi-select) | Choice (multi) | No | Black/African American; Hispanic/Latino; Asian; White; Native American/Indigenous; Pacific Islander; Multiracial; Prefer not to say | `Demographic_Race_Ethnicity` |
| 15 | Farmer gender | Choice (dropdown) | No | Woman; Man; Non-binary; Prefer not to say | `Demographic_Gender` |
| 16 | Farmer age band | Choice (dropdown) | No | Under 25; 25–34; 35–44; 45–54; 55–64; 65+; Prefer not to say | `Demographic_Age_Band` |
| 17 | OFSA dimension(s) for this entry | Choice (multi) | Yes | Default: Availability; Access; Agency | `OFSA_Dimension` |
| 18 | Notes | Long text | No | — | `Notes` |

### Branching

- Q12 = "Yes" → show Q13

### Distribution

After publishing, copy the public link and share with UAC contacts via
email. Also embed the link in any UAC-facing onboarding doc. Test the
form in a private browser window (logged out of M365) to confirm it
works for unauthenticated users.

---

## Form-to-List mapping summary

| Form | Primary List | Secondary List(s) |
|---|---|---|
| Kitchen Operations Log | Kitchen Operations Log (sub-list of Food Entrepreneurs) | Food Entrepreneurs List (new members); Environmental List (waste) |
| Culinary & Training Log | Food Entrepreneurs List | Environmental List (food captured for demos) |
| Community Engagement Log | Community Orgs List | Environmental List (food recovery) |
| Director & Partnership Log | Investors-Funders List | Institution List |
| UAC Farmer Engagement Log | Farmers List | — |

**Note on the "Kitchen Operations Log" sub-list.** The Kitchen Operations
Log Form's primary destination is operational telemetry (members,
equipment, incidents). To keep schemas clean, the simplest implementation
writes those entries to a dedicated `Kitchen Operations Log` List that
mirrors the common columns plus a few weekly-snapshot fields. If you
want to keep the List count to six, you can instead write Kitchen
Operations rows to the Food Entrepreneurs List with `Stage = "Operational Snapshot"`
— but that conflates business records with weekly logs and is harder to
report on. The script and schema doc treat the Kitchen Operations Log as
the sub-list owned by the Food Entrepreneurs audience for dashboarding
purposes; create a 7th SharePoint List for it if your tenant allows.

For simplicity in the v1 build, the provisioning script creates the
**six core Lists** and routes Kitchen Operations data into the **Food
Entrepreneurs List** with `Stage = "Operational Snapshot"`. This can be
refactored into a 7th List later without changing the Forms.
