# HARVEST KPI Catalog

This is the single source of truth for every metric HARVEST tracks.
Every Form question, SharePoint List column, and dashboard pivot
derives from this file. When a KPI changes, change it here first, then
propagate to the schemas and the workbook.

## How to read this catalog

Each row is one KPI. Columns are the **tagging schema** — every KPI
carries values for all of them.

- **KPI** — short name used on dashboards
- **Definition** — one-line plain-English description
- **Unit** — count, %, $, lbs, hours, etc.
- **Audience** — primary audience category (also drives which List the data lands in)
- **OFSA Dimension(s)** — one or more of Availability / Access / Utilization / Stability / Agency / Sustainability
- **Food System Sector** — Production / Aggregation / Processing / Distribution / Retail-Consumption / Recovery-Waste
- **Partner Sector** — Healthcare / Agriculture / Education / Government / Philanthropy / Community-Based Org / Private Sector / Food Enterprise
- **Cadence** — Real-time / Weekly / Monthly / Quarterly / Annual
- **Lead Role** — HARVEST Lead / UAC Lead / Joint
- **Data Source** — Form submission / Food Corridor CSV / Manual entry / System-calculated
- **2026 Target** — numeric goal if tied to a performance goal, else blank
- **Owner Role** — Director / Culinary Manager / Logistics Manager / Community Outreach Specialist / UAC Partner
- **Source Form** — which of the 5 Forms feeds this KPI
- **Destination List** — which SharePoint List stores the row
- **Status** — Active / Planned — methodology TBD / Retired

A KPI tagged **Planned — methodology TBD** appears in the catalog
and is documented for the OFSA logic model phase, but is suppressed
from the Executive Summary and Goals dashboards until a measurement
methodology exists.

---

## Food Entrepreneurs

| KPI | Definition | Unit | OFSA Dimension(s) | Food System Sector | Partner Sector | Cadence | Lead Role | Data Source | 2026 Target | Owner Role | Source Form | Destination List | Status |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| Entrepreneurs engaged | Unique individuals reached via outreach, intake, or info session | count | Agency | Processing | Food Enterprise | Monthly | HARVEST Lead | Form submission | 40 | Community Outreach Specialist | Community Engagement Log | Active |
| Entrepreneurs converted to renting members | Engaged entrepreneurs who signed a kitchen rental agreement | count | Agency; Availability | Processing | Food Enterprise | Monthly | HARVEST Lead | Form submission | 5–10 | Logistics Manager | Kitchen Operations Log | Active |
| % members achieving ServSafe through HARVEST | Members who pass ServSafe Food Handler or Manager via HARVEST programming | % | Utilization | Processing | Food Enterprise | Per training | HARVEST Lead | Form submission | 100% of cohort | Culinary Manager | Culinary & Training Log | Active |
| Avg % sales increase for kitchen members | Year-over-year revenue change for members using the kitchen ≥6 months | % | Stability; Agency | Processing | Food Enterprise | Quarterly | HARVEST Lead | Manual entry | Baseline year | Director | Director & Partnership Log | Active |
| Businesses supported | Distinct member businesses receiving any HARVEST service in the period | count | Agency | Processing | Food Enterprise | Monthly | HARVEST Lead | System-calculated | 5–10 | Director | Food Entrepreneurs List | Active |
| Workforce dev seminars held | Seminars on business building, food safety, marketing, capital access | count | Utilization; Agency | Processing | Education | Per event | HARVEST Lead | Form submission | 6 | Culinary Manager | Culinary & Training Log | Active |
| Kitchen hours rented | Total billable kitchen hours rented in the period | hours | Availability; Stability | Processing | Food Enterprise | Real-time | HARVEST Lead | Food Corridor CSV | 1,500 | Logistics Manager | Kitchen Operations Log | Active |
| Jobs created by member businesses | FTEs and PTEs hired by member businesses, self-reported | count | Stability; Agency | Processing | Food Enterprise | Quarterly | HARVEST Lead | Form submission |  | Director | Food Entrepreneurs List | Active |
| Growth of businesses by hours rented | Trend of avg hours/member over time (longitudinal) | hours | Stability | Processing | Food Enterprise | Quarterly | HARVEST Lead | System-calculated |  | Director | Food Entrepreneurs List | Active |
| Businesses buying local produce via UAC | Member businesses that sourced produce from UAC channels in the period | count | Access; Sustainability | Aggregation | Agriculture | Monthly | Joint | Form submission | 3 | UAC Partner | UAC Farmer Engagement Log | Active |
| Events showcasing local entrepreneurs | Public-facing events featuring member businesses (tastings, pop-ups, demos) | count | Agency | Retail-Consumption | Food Enterprise | Per event | HARVEST Lead | Form submission | 4 | Community Outreach Specialist | Community Engagement Log | Active |
| Technical assistance sessions delivered | 1:1 sessions for business building, capital access, marketing | count | Utilization; Agency | Processing | Food Enterprise | Per session | HARVEST Lead | Form submission | 30 | Culinary Manager | Culinary & Training Log | Active |
| Businesses receiving wellness coaching | Member businesses participating in HARVEST wellness/nutrition coaching | count | Utilization | Processing | Healthcare | Per session | HARVEST Lead | Form submission |  | Community Outreach Specialist | Food Entrepreneurs List | Active |

---

## Institution

| KPI | Definition | Unit | OFSA Dimension(s) | Food System Sector | Partner Sector | Cadence | Lead Role | Data Source | 2026 Target | Owner Role | Source Form | Destination List | Status |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| $ purchasing orders from HARVEST/UAC vendors | Total $ value of purchase orders by institutions for HARVEST/UAC products | $ | Access; Availability | Distribution | Healthcare | Monthly | Joint | Form submission |  | Director | Institution List | Active |
| HARVEST-produced items in institutional menus | Distinct HARVEST/UAC items appearing on institutional menus | count | Access; Availability | Retail-Consumption | Healthcare | Quarterly | HARVEST Lead | Form submission |  | Director | Institution List | Active |
| Connections with other institutions | Active relationships with Rutgers, NPS, Audible, etc. | count | Stability; Access | n/a | Education | Quarterly | HARVEST Lead | Form submission | 4 | Director | Institution List | Active |
| Partnerships for technical assistance | External institutions providing TA to HARVEST members | count | Utilization; Stability | n/a | Education | Quarterly | Joint | Form submission | 3 | Director | Institution List | Active |
| Reduction in readmission costs through HARVEST | $ reduction in RWJBH patient readmission costs attributable to HARVEST programming | $ | Utilization; Access | Retail-Consumption | Healthcare | Annual | Joint | Manual entry |  | Director | Institution List | Planned — methodology TBD |

---

## Community Organizations

| KPI | Definition | Unit | OFSA Dimension(s) | Food System Sector | Partner Sector | Cadence | Lead Role | Data Source | 2026 Target | Owner Role | Source Form | Destination List | Status |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| Community orgs engaged at HARVEST | Distinct organizations participating in any HARVEST activity in the period | count | Agency; Access | n/a | Community-Based Org | Monthly | HARVEST Lead | Form submission | 15 | Community Outreach Specialist | Community Orgs List | Active |
| Event participants | Total participants across teaching kitchen, RD workshops, etc. | count | Utilization; Agency | n/a | Community-Based Org | Per event | HARVEST Lead | Form submission | 300 | Community Outreach Specialist | Community Orgs List | Active |
| Community volunteer hours | Hours contributed by community volunteers at HARVEST | hours | Agency; Stability | n/a | Community-Based Org | Per event | HARVEST Lead | Form submission |  | Community Outreach Specialist | Community Orgs List | Active |
| Internal & external community events run by staff | Events run by HARVEST staff (on-site and external) | count | Agency | n/a | Community-Based Org | Per event | HARVEST Lead | Form submission | 24 | Community Outreach Specialist | Community Orgs List | Active |
| Event survey — satisfaction | Avg satisfaction rating from post-event survey (1–5) | rating | Utilization | n/a | Community-Based Org | Per event | HARVEST Lead | Form submission | ≥ 4.0 | Community Outreach Specialist | Community Orgs List | Active |
| Event survey — learning outcomes | Avg self-reported learning rating from post-event survey (1–5) | rating | Utilization | n/a | Community-Based Org | Per event | HARVEST Lead | Form submission | ≥ 4.0 | Community Outreach Specialist | Community Orgs List | Active |
| Tours given at HARVEST | Total facility tours given to external visitors | count | Agency | n/a | Community-Based Org | Per event | HARVEST Lead | Form submission | 24 | Community Outreach Specialist | Community Orgs List | Active |
| Community orgs hosting events at HARVEST | External orgs using HARVEST as their venue | count | Agency; Access | n/a | Community-Based Org | Per event | HARVEST Lead | Form submission | 6 | Community Outreach Specialist | Community Orgs List | Active |
| Sectors engaged at HARVEST | Distinct partner sectors represented across all engagements | count | Stability | n/a | Community-Based Org | Quarterly | HARVEST Lead | System-calculated | 5 | Director | Community Orgs List | Active |
| Multilingual / cultural events | Events delivered in a non-English language or built around a specific cultural tradition | count | Agency; Access | n/a | Community-Based Org | Per event | HARVEST Lead | Form submission | 6 | Community Outreach Specialist | Community Orgs List | Active |
| Food recovery to service providers | lbs of food diverted to Bridges and other direct service providers | lbs | Access; Sustainability | Recovery-Waste | Community-Based Org | Per event | HARVEST Lead | Form submission | 500 | Community Outreach Specialist | Community Orgs List | Active |
| Direct service providers with presence at HARVEST | Organizations like Bridges with a recurring on-site presence | count | Access | n/a | Community-Based Org | Quarterly | HARVEST Lead | Form submission | 3 | Community Outreach Specialist | Community Orgs List | Active |

---

## Investors / Funders

| KPI | Definition | Unit | OFSA Dimension(s) | Food System Sector | Partner Sector | Cadence | Lead Role | Data Source | 2026 Target | Owner Role | Source Form | Destination List | Status |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| Grants and partnerships secured | New grants or formal partnerships executed in the period | count | Stability | n/a | Philanthropy | Per event | HARVEST Lead | Form submission | 4 | Director | Investors-Funders List | Active |
| % earned revenue | Earned revenue (rentals, events, sales) as % of total revenue | % | Stability; Sustainability | Processing | Philanthropy | Quarterly | HARVEST Lead | Manual entry |  | Director | Investors-Funders List | Active |
| Funders and sponsors | Distinct active funders/sponsors in the period | count | Stability | n/a | Philanthropy | Quarterly | HARVEST Lead | Form submission |  | Director | Investors-Funders List | Active |
| Leveraged capital | $ in matching grants, investments, in-kind contributions leveraged by HARVEST | $ | Stability; Sustainability | n/a | Philanthropy | Quarterly | HARVEST Lead | Form submission |  | Director | Investors-Funders List | Active |
| Social return on investment (SROI) | Ratio of social value created to dollars invested | ratio | Sustainability; Stability | n/a | Philanthropy | Annual | HARVEST Lead | Manual entry |  | Director | Investors-Funders List | Planned — methodology TBD |

---

## Farmers

| KPI | Definition | Unit | OFSA Dimension(s) | Food System Sector | Partner Sector | Cadence | Lead Role | Data Source | 2026 Target | Owner Role | Source Form | Destination List | Status |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| Farms engaged | Distinct farms in active relationship with HARVEST/UAC | count | Agency; Stability | Production | Agriculture | Monthly | UAC Lead | Form submission | 4 | UAC Partner | Farmers List | Active |
| $ farmer business generated through HARVEST | Total $ value of farmer sales attributable to HARVEST/UAC channels | $ | Stability; Access | Aggregation | Agriculture | Monthly | UAC Lead | Form submission |  | UAC Partner | Farmers List | Active |
| Types of produce moved | Distinct produce types moved through HARVEST/UAC | count | Availability | Aggregation | Agriculture | Monthly | UAC Lead | Form submission |  | UAC Partner | Farmers List | Active |
| lbs of produce moved | Total lbs of produce aggregated/distributed via HARVEST/UAC | lbs | Availability; Access | Aggregation | Agriculture | Monthly | UAC Lead | Form submission |  | UAC Partner | Farmers List | Active |
| Produce distribution destinations | Where produce was distributed (city/org/recipient type) | text/list | Access | Distribution | Agriculture | Monthly | UAC Lead | Form submission |  | UAC Partner | Farmers List | Active |
| Produce origin | Where produce came from (farm/region) | text/list | Availability; Sustainability | Production | Agriculture | Monthly | UAC Lead | Form submission |  | UAC Partner | Farmers List | Active |
| Farmers making value-added products | Farmers using HARVEST to produce value-added products | count | Agency; Stability | Processing | Agriculture | Monthly | UAC Lead | Form submission |  | UAC Partner | Farmers List | Active |
| Farmer demographics | Counts by race/ethnicity, gender, age band, geography | counts | Agency | Production | Agriculture | Quarterly | UAC Lead | Form submission |  | UAC Partner | Farmers List | Active |

---

## Environmental Sustainability

| KPI | Definition | Unit | OFSA Dimension(s) | Food System Sector | Partner Sector | Cadence | Lead Role | Data Source | 2026 Target | Owner Role | Source Form | Destination List | Status |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| Food waste diversion | lbs of food waste diverted from landfill (any pathway) | lbs | Sustainability | Recovery-Waste | Community-Based Org | Monthly | HARVEST Lead | Form submission | 1,000 | Logistics Manager | Environmental List | Active |
| Composted lbs | lbs of food waste composted | lbs | Sustainability | Recovery-Waste | Community-Based Org | Monthly | HARVEST Lead | Form submission |  | Logistics Manager | Environmental List | Active |
| Local food miles | Avg distance (miles) from source to HARVEST for produce received | miles | Sustainability | Distribution | Agriculture | Quarterly | UAC Lead | System-calculated |  | UAC Partner | Environmental List | Active |
| Food captured for nutritional/wellness demos | lbs and $ value of food used in demos instead of being wasted | lbs / $ | Sustainability; Utilization | Recovery-Waste | Healthcare | Per event | HARVEST Lead | Form submission |  | Culinary Manager | Environmental List | Active |

---

## 2026 Performance Goals — KPI links

These are the metrics that roll up to the Director's annual review goals
and appear on the **2026 Performance Goals** tab of the master workbook.

| Goal | Sub-goal | Linked KPI | 2026 Target | Tracking |
|---|---|---|---|---|
| 1. Kitchen launch / operations | Founding cohort of members | Entrepreneurs converted to renting members | 5–10 | KPI system |
| 1. Kitchen launch / operations | ServSafe achievement | % members achieving ServSafe | 100% of cohort | KPI system |
| 1. Kitchen launch / operations | Rental hours utilization | Kitchen hours rented | 1,500 | KPI system (Food Corridor) |
| 2. Structural documentation | Guiding documents and SOPs | — | n/a — tracked in shared drive folder | **Outside KPI system** |
| 3. UAC partnership maintenance | JSC meeting cadence | JSC meetings held (Director Form) | Up to 10 | KPI system |
| 3. UAC partnership maintenance | Shared data framework | This system being operational | Live by Q1 | KPI system (binary status) |
| 3. UAC partnership maintenance | UAC-affiliated farmers engaged | Farms engaged | ≥ 4 | KPI system |
| 4. CAB development | CAB membership | CAB members across sectors | 7–12 across 4+ sectors | KPI system |
| 4. CAB development | CAB meetings | CAB meetings held | 2 by year-end | KPI system |
| 5. Professional development | Shared Kitchen Summit | Director's Form — PD log entry | 1 (attendance) | KPI system |
| 5. Professional development | ServSafe Manager Cert (Director) | Director's Form — PD log entry | 1 (pass) | KPI system |
| 5. Professional development | Community presentations | Director's Form — presentations | Up to 4 | KPI system |

**Note on Goal 2.** Structural documentation (SOPs, guiding docs) is
intentionally tracked outside this KPI system — a shared drive folder
with a simple checklist is the right tool. The README repeats this so
the omission isn't read as a gap.
