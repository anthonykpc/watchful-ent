# HARVEST Food Hub — Power Automate Setup (OPTIONAL)

> **This step is optional.** Skip it if your team doesn't want email
> notifications when new contacts are submitted. The contact tracker works
> fine without Power Automate — form responses still flow into the Excel
> workbook.
>
> Use this guide if you want an email to go out to the HARVEST team every
> time someone adds a new contact via the form.

Estimated time: **10–15 minutes, one-time setup.**

You need your RWJBarnabas Health Microsoft 365 login. No coding required.

---

## What this flow does

Every time someone submits the "HARVEST Food Hub — Add a Contact" form,
Power Automate sends a short email to a designated HARVEST team address
with the new contact's name, organization, and category.

Example email the team will receive:

> **Subject:** New HARVEST contact added: Jane Doe (Newark Community Foundation)
>
> A new contact was just added to the HARVEST tracker.
>
> • **Name:** Jane Doe
> • **Organization:** Newark Community Foundation
> • **Category:** Funders/Grantmakers
> • **Added by:** Maria Lopez
>
> View the tracker: [link to SharePoint workbook]

---

## Step 1 — Open Power Automate

1. In a browser, go to <https://make.powerautomate.com>.
2. Sign in with your RWJBarnabas Health Microsoft 365 account.
3. In the left sidebar, click **+ Create**.

## Step 2 — Choose the trigger

1. Click **Automated cloud flow**.
2. In the **Flow name** box, type:
   ```
   HARVEST — Notify team of new contact
   ```
3. In the **Choose your flow's trigger** search box, type `forms` and
   pick **When a new response is submitted** (Microsoft Forms).
4. Click **Create**.

## Step 3 — Point the trigger at the HARVEST form

1. In the first step's **Form Id** dropdown, pick
   **HARVEST Food Hub — Add a Contact**.
   (If you don't see it, make sure you created the form under the same
   account and that you've submitted at least one test response.)

## Step 4 — Get the response details

1. Click **+ New step**.
2. Search for `forms` and pick **Get response details** (Microsoft Forms).
3. **Form Id:** pick **HARVEST Food Hub — Add a Contact** again.
4. **Response Id:** click inside the box, then in the dynamic content
   panel on the right, pick **Response Id** (from the trigger step).

## Step 5 — Send the notification email

1. Click **+ New step**.
2. Search for `outlook` and pick **Send an email (V2)** (Office 365
   Outlook).
3. Fill in the fields as follows. When the instructions say "insert
   dynamic content," click into the field and pick the matching item
   from the right-hand panel — do NOT type the field name as plain text.

   - **To:** `[PASTE HARVEST TEAM EMAIL ADDRESS HERE]`
     (e.g. `harvest-team@rwjbh.org` — confirm the exact address with the
     team lead before saving)
   - **Subject:**
     ```
     New HARVEST contact added: [insert dynamic First Name] [insert dynamic Last Name] ([insert dynamic Organization])
     ```
   - **Body (switch the editor to plain text / HTML off, or use the rich
     editor — either works):**
     ```
     A new contact was just added to the HARVEST tracker.

     • Name:          [First Name] [Last Name]
     • Organization:  [Organization]
     • Category:      [Contact Category]
     • Added by:      [Added By]

     View the tracker:  [PASTE SHAREPOINT / ONEDRIVE LINK TO THE WORKBOOK HERE]

     — Automated by Power Automate. Reply to the person who added the
     contact if you have questions about the entry.
     ```

     Replace each `[Field Name]` bracket with the matching **dynamic
     content** item from the right-hand panel (they come from the
     "Get response details" step).

## Step 6 — Save and test

1. Click **Save** at the bottom-right.
2. Click **Test** in the top-right, pick **Manually**, then click **Test**.
3. In another browser tab, open the HARVEST contact form and submit a
   test entry.
4. Return to the Power Automate tab — the flow should show a green
   checkmark on each step within ~30 seconds.
5. Check the team inbox — the notification email should arrive.
6. Delete the test contact from the workbook when you're done (or mark it
   `Archived`).

## Step 7 — Turn it on

Back on the flow's main page, make sure the toggle in the top-right reads
**On**. Power Automate will now email the team every time the form is
submitted, in the background, forever.

---

## Optional extras

- **Post to a Teams channel instead of email.** Replace the "Send an
  email (V2)" step with **Post message in a chat or channel** (Microsoft
  Teams connector). Same dynamic-content fields work.
- **Only notify for specific categories.** After the "Get response
  details" step, add a **Condition** step: if `Contact Category` equals
  `Funders/Grantmakers` (for example), then send the email; otherwise do
  nothing.
- **Daily digest instead of per-submission.** Swap the trigger from
  "When a new response is submitted" to a **Recurrence** trigger (daily
  at 9am), then use **List rows present in a table** (Excel) to pull
  yesterday's entries, and send a single summary email. This is more
  complex — only build it if per-submission email becomes noisy.

---

## Turning the flow off or removing it

1. Go to <https://make.powerautomate.com> > **My flows**.
2. Find **HARVEST — Notify team of new contact**.
3. Use the toggle to turn it off, or click the `…` menu and **Delete**
   to remove it entirely.

Removing the flow has no effect on the form, the workbook, or existing
contacts — it only stops future email notifications.
