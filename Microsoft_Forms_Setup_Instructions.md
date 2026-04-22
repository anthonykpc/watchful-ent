# HARVEST Food Hub — Microsoft Forms Setup Instructions

This guide walks a non-technical user through creating the contact-intake form
in Microsoft Forms and connecting it to `HARVEST_Contact_Tracker.xlsx` so new
submissions appear automatically in the workbook.

Estimated time: **20–30 minutes, one-time setup.**

You only need a browser and your RWJBarnabas Health Microsoft 365 login.

---

## Step 0 — Before you start

1. Save `HARVEST_Contact_Tracker.xlsx` to the HARVEST team's SharePoint site or
   OneDrive for Business folder (not your local Downloads folder).
2. Open the file once in Excel Online to confirm it renders correctly.
   You should see three tabs: **All Contacts**, **Follow-Up Dashboard**,
   **How to Use This Tracker**.
3. Pick the path you want to use:
   - **Path A (recommended, simplest):** Create the form from *inside* the
     Excel workbook. Microsoft auto-wires the form to a sheet in the
     workbook and new submissions flow in with no extra setup.
   - **Path B:** Create a standalone form at forms.office.com, then use
     Power Automate to push each response into the `All Contacts` tab.
     See `Power_Automate_Optional_Setup.md`.

Follow **Path A** unless you specifically need the form to live outside the
workbook.

---

## Path A — Create the form from within the workbook (recommended)

### Step 1 — Open the workbook in Excel Online

1. In a browser, go to SharePoint or OneDrive and open
   `HARVEST_Contact_Tracker.xlsx`.
2. Make sure you are using **Excel for the web** (Excel Online) — not the
   desktop app. The "Insert > Forms" feature is only available in the web
   version.

### Step 2 — Insert a new form

1. Click the **Insert** tab in the Excel ribbon.
2. Click **Forms > New Form**.
3. A new browser tab opens in Microsoft Forms, with a blank form already
   linked to the workbook. Excel will also add a new sheet named something
   like `Form1` to hold incoming responses — we will use this sheet as the
   feeder into `All Contacts` (see Step 6).

### Step 3 — Set the form title and description

1. Click the placeholder title at the top of the form.
2. Enter **Title**:
   ```
   HARVEST Food Hub — Add a Contact
   ```
3. Enter **Description**:
   ```
   Use this form to add a new stakeholder, partner, funder, or community
   contact to the HARVEST network. Fields marked * are required.
   ```

### Step 4 — Add the questions (in this exact order)

For each question below, click **+ Add new**, pick the field type shown,
then copy/paste the question text and options. The **Required** toggle is
at the bottom-right of each question.

| # | Question text | Field type | Options | Required? |
|---|---|---|---|---|
| 1 | First Name | Text (short answer) | — | **Required** |
| 2 | Last Name | Text (short answer) | — | **Required** |
| 3 | Organization | Text (short answer) | — | Optional |
| 4 | Title / Role | Text (short answer) | — | Optional |
| 5 | Email | Text (short answer) — turn on **Restrictions > Email** in the `…` menu so Forms validates the address | — | Optional |
| 6 | Phone | Text (short answer) | — | Optional |
| 7 | Contact Category | **Choice** (dropdown) | Funders/Grantmakers; Government/Elected Officials; Community Organizations/Nonprofits; Food Entrepreneurs/Potential Kitchen Members; Healthcare/RWJBH Partners; Peer Food Hubs; Media/Press; Farmers/UAC-related; Vendors/Suppliers; Other | **Required** |
| 8 | HARVEST Pillar | **Choice** (dropdown) | Kitchen Incubator; Business Accelerator; Health & Wellness Hub; All Three; N/A | Optional |
| 9 | Relationship Warmth | **Choice** (dropdown) | Hot; Warm; Cold; Unknown | Optional |
| 10 | Connection Source — how did we meet this contact? | Text (short answer) | — | Optional |
| 11 | Last Contact Date | **Date** | — | Optional |
| 12 | Next Follow-Up Date | **Date** | — | Optional |
| 13 | Notes / Relationship History — include a date prefix, e.g. `2026-04-22: met at City Hall…` | Text (long answer) | — | Optional |
| 14 | Added By — your name | Text (short answer) | — | **Required** |
| 15 | Active? | **Choice** (dropdown) | Yes; No; Archived | **Required** — default to **Yes** (add it last in the list and type "Yes" into a default, or include instructions in the question text) |

**Tips for the dropdown (Choice) questions:**

- After clicking **Choice**, click the small arrow next to the "Multiple
  answers" toggle and turn on **Drop-down** so each question renders as a
  dropdown instead of a long list of radio buttons.
- Paste each option on its own line.
- Do **not** check "Multiple answers" — contacts belong to exactly one
  category/pillar/warmth at a time.

**Tip for the Date questions:** click the `…` menu on each date question and
turn on **Restrictions > Date** if available, so users can't type free text.

### Step 5 — Configure form settings

1. Click the `…` menu in the top-right of the Forms editor, then
   **Settings**.
2. Under **Who can fill out this form**, choose
   **Only people in my organization can respond**. This keeps submissions
   inside the RWJBH tenant.
3. Turn **on** "Record name" (so we automatically capture who submitted).
4. Turn **on** "One response per person" only if you want to prevent a
   single staffer from adding multiple contacts in a session — for
   HARVEST's use case, leave this **off**.
5. Optional: under **Response receipts**, turn on "Get email notification
   of each response" so the submitter gets a copy of what they entered.

### Step 6 — Confirm the Excel sync is working

1. Go back to the Excel Online tab. You should see a new sheet (e.g.
   `Form1`) with column headers matching your form questions, plus
   `ID`, `Start time`, `Completion time`, `Email`, and `Name` columns
   that Microsoft Forms adds automatically.
2. Submit a **test response** from the Forms preview (click **Preview** in
   the Forms editor, fill in a dummy contact like "Test Testerson /
   Delete me", submit).
3. Within ~1 minute the new row should appear in the Forms sheet inside
   the workbook. If it does, the sync is working. Delete the test row
   afterwards via the Forms Responses view (not by deleting the Excel
   row — see note below).

### Step 7 — Route responses into the `All Contacts` tab

Microsoft Forms writes submissions to the sheet it created (e.g. `Form1`),
not directly into `All Contacts`. You have two options:

**Option 7a — Use the Forms sheet as the live database (simplest).**
- Hide the original `All Contacts` tab (right-click the tab > **Hide**).
- Rename the Forms-generated sheet to `All Contacts`.
- Re-apply the dropdown data validations, date formatting, and header
  styling from the template (or leave defaults — the Forms sheet already
  comes formatted as a table with filters).
- Update the Follow-Up Dashboard formula's sheet name reference from
  `'All Contacts'` to match (it should already read `'All Contacts'` if
  you renamed the Forms sheet).

**Option 7b — Keep both sheets and use Power Automate to copy rows.**
- Follow the instructions in `Power_Automate_Optional_Setup.md` (adapting
  the "send email" flow into a "add row to Excel table" action). This is
  more robust long-term but requires Power Automate setup.

For a team of 3–6 people we recommend **Option 7a**. It's one sheet, one
source of truth, no extra automation to maintain.

### Step 8 — Share the form URL with the team

1. Click **Collect responses** (or **Send**) in the top-right of the Forms
   editor.
2. Copy the **link**.
3. Paste that link into:
   - the `How to Use This Tracker` tab, Section 1 ("Form URL" placeholder), and
   - any team email / Teams channel where staff will need it.

You're done. New contacts added via the form will appear in the workbook
automatically.

---

## Path B — Standalone form + Power Automate

If you prefer the form to live independently at forms.office.com (e.g. to
reuse it across multiple trackers, or to embed it on an external
webpage), use this path:

1. Go to <https://forms.office.com> and click **+ New Form**.
2. Follow **Steps 3, 4, 5, and 8** above (title, questions, settings,
   sharing).
3. Do NOT use "Insert > Forms" from Excel — this path keeps them separate.
4. Open `Power_Automate_Optional_Setup.md` and follow the flow labeled
   **"Copy form response into All Contacts"** (adapt the optional
   email-notification flow by adding a "Add a row into a table" Excel
   action before the "Send email" action).

---

## Troubleshooting

- **I don't see "Insert > Forms" in Excel.** You're in desktop Excel, not
  Excel Online. Open the file from OneDrive/SharePoint in a browser.
- **Form responses aren't showing up in Excel.** Check that you created
  the form via "Insert > Forms" from within the workbook (not as a
  standalone form). A standalone form requires Power Automate to
  push responses into Excel.
- **A dropdown shows up as radio buttons.** In the Choice question, click
  the `…` or arrow icon and enable **Drop-down**.
- **A teammate can't open the form.** The form is tenant-restricted. Make
  sure they're signed in with their RWJBarnabas Health account.
- **I accidentally deleted a response row in Excel.** Forms keeps its own
  copy of responses — go to the Forms editor, click **Responses**, and
  you can re-export.
