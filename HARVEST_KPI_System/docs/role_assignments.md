# Role Assignments — Who Enters What

This is the operating rhythm for the KPI system. Each role owns one
Form and is responsible for keeping its KPIs current.

## Quick reference

| Role | Owns Form | Cadence | Time per submission |
|---|---|---|---|
| Logistics Manager | Kitchen Operations Log | Weekly + event-based | ~5 min weekly |
| Culinary Manager | Culinary & Training Log | Per workshop/training | ~3 min per event |
| Community Outreach Specialist | Community Engagement Log | Per event | ~3 min per event |
| Director | Director & Partnership Log | Monthly + event-based | ~10 min monthly |
| UAC Partner (external) | UAC Farmer Engagement Log | Monthly | ~10 min monthly |

The Director also reviews the Executive Summary dashboard before each
JSC meeting and refreshes Power Query connections.

---

## By role — detail

### Logistics Manager

**Form**: Kitchen Operations Log

**What you enter, weekly:**
- Member counts (active members, new members this week)
- Equipment status (any equipment offline, repair needs)
- Incidents (safety, sanitation, member conflicts)
- Food waste diverted this week (lbs)
- Composted this week (lbs)

**What you do NOT enter:**
- Kitchen rental hours — these come from the Food Corridor CSV. Do
  not duplicate.

**Tools you also touch:**
- Food Corridor (export the kitchen rentals CSV monthly and drop it
  into the Excel workbook's `Food Corridor Import` tab)

---

### Culinary Manager

**Form**: Culinary & Training Log

**What you enter, per workshop or training session:**
- Workshop type (ServSafe, business building, marketing, capital, other)
- Date and duration
- Attendance count
- ServSafe pass/fail outcomes (when applicable)
- Technical assistance topic and recipient business (for 1:1 sessions)
- Food captured for nutritional/wellness demos (lbs, $ value)

---

### Community Outreach Specialist

**Form**: Community Engagement Log

**What you enter, per event:**
- Event name, date, type (teaching kitchen, RD workshop, tour, partner event, etc.)
- Attendance count
- Volunteer hours contributed
- Organizations engaged (multi-select from existing list, or add new)
- Multilingual / cultural flag and language
- Survey results (avg satisfaction 1–5, avg learning outcomes 1–5)
- Food recovery to Bridges or other providers (lbs, recipient)
- Direct service provider presence (if recurring)

---

### Director

**Form**: Director & Partnership Log

**What you enter, monthly:**
- JSC meetings held this month (date, agenda summary)
- CAB recruitment activity (members added, sector)
- CAB meetings held
- Professional development activity (Shared Kitchen Summit, ServSafe Manager Cert, community presentations)

**What you enter, as it happens (event-based):**
- New grants secured ($ amount, funder, period)
- New partnerships executed
- New institutional connections (Rutgers, NPS, Audible, etc.)
- New institutional purchase order ($, vendor, institution)

**What you also do:**
- Review Executive Summary dashboard before each JSC meeting
- Hit "Refresh All" in the Excel workbook (Data tab → Refresh All) before screenshotting for the supervisor / JSC
- Update `% earned revenue` quarterly from financials
- Update `Avg % sales increase for members` quarterly from member-reported data

---

### UAC Partner

**Form**: UAC Farmer Engagement Log

**You don't need an RWJBH login.** The form is shared as a public link
and works on phones. You can submit it from a tractor.

**What you enter, monthly:**
- Farms engaged this month (farm name, location, contact)
- Produce moved (type, lbs, $ value)
- Origin (which farm/region)
- Distribution destinations (which institution, org, or community)
- Farmers using HARVEST for value-added products (which farm, what product)
- Farmer demographics (rolled-up counts for the period, not per-person identifiers)

---

## Refresh and review rhythm

| When | Who | What |
|---|---|---|
| Weekly | Logistics Manager | Submit Kitchen Operations Log |
| Per event | Culinary / Outreach Mgrs | Submit relevant Form within 48 hours |
| Monthly, 1st week | UAC Partner | Submit UAC Farmer Engagement Log for prior month |
| Monthly, 1st week | Director | Submit Director & Partnership Log; export Food Corridor CSV and drop into Excel |
| Monthly, 2nd week | Director | Refresh workbook, review dashboards |
| Before each JSC meeting | Director | Refresh, screenshot 2026 Goals + Executive Summary |
| Quarterly | Director | Update SROI placeholders, % earned revenue, member sales growth |
| Annually | Director | Performance review using 2026 Goals tab |

## Overdue alerts

A Power Automate flow (see `schemas/power_automate_flows.md`) emails the
owner role when a Form hasn't been submitted within its expected cadence
window:
- Weekly Form: alert after 10 days of silence
- Monthly Form: alert after 35 days
- Quarterly Form: alert after 100 days

The alert is to the owner only — it doesn't escalate to the Director
unless a second window passes.
