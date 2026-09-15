# Payment Voucher Generator — Project Roadmap & Requirements

## 1. Background / Problem Statement

The business currently makes combined (clubbed) cheque payments to multiple parties at once. For example, three parties are owed 52,146 / 78,145 / 9,655 — instead of writing three separate cheques, one single cheque is issued for the total (139,946), and an Excel sheet is prepared listing each party's bank details and individual amount. This sheet is printed and also emailed to the bank, so the bank knows how to split and route the combined cheque amount to each party's account.

Today this Excel sheet is built manually each time: looking up each party's bank details, typing them in, calculating the total, and converting the total into words.

**Goal:** A simple, local website (no login, single user, runs on one computer) that:
- Stores the party list (name + bank details) so it never has to be retyped.
- Lets the user pick parties for a given payment run and enter each one's amount.
- Auto-calculates the total and the amount-in-words.
- Generates a ready-to-print, ready-to-email Excel file in the exact bank-submission format already in use.

## 2. Reference Output (Ground Truth)

The user supplied a real example of the final sheet sent to the bank. The generated Excel file must match this layout and content exactly. Columns, in order:

| SR.NO. | PARTICULARS | BANK'S NAME | BANK ACCOUNT NUMBER | BANK'S LOCATION | IFSC CODE | AMOUNT |
|---|---|---|---|---|---|---|
| 1 | PCI PEST CONTROL PRIVATE LIMITED | STANDARD CHARTERED BANK | 9301811329 | VADODARA | SvBL5066085 | 50000 |
| 2 | SHREE MAHAVIR TIMBERS | BANK OF MAHARASTRA | 2053402589 | ANAND | MBHB0006456 | 137600 |

Below the party rows, the sheet contains:

- A merged row spanning the middle columns with the **amount in words** (e.g. "INT ONE LAKH EIGHTY SEVEN THOUSAND SIX HUNDRED ONLY"), with the **total amount** in the AMOUNT column, **highlighted yellow**.
- A blank spacer row.
- **CHEQUE NO -** `<value>` and **DATE:** `<value>` on the same row.
- A blank spacer row.
- Three lines giving the **paying account's own details** (the business's own account the cheque is drawn from):
  - NAME OF BANK A/C : `<account holder name>`
  - NAME OF BANK : `<bank name>`
  - BANK ACCOUNT NUMBER : `<account number>`

Formatting notes visible in the sample:
- Header row: bold, centered, with a header-style border/underline.
- SR.NO., AMOUNT columns are narrow; PARTICULARS and BANK'S NAME are wide.
- Amounts are right-aligned numbers, no currency symbol, no decimals (whole rupees).
- The amount-in-words + total row uses bold text and a yellow fill on the total cell.
- Indian numbering convention is used in words (Lakh, not Million).

This sample is the acceptance test for the export feature: a generated file, side by side with this one, should be structurally indistinguishable aside from the actual data.

## 3. Confirmed Requirements (from discussion with stakeholder)

| Topic | Decision |
|---|---|
| Who uses it | Just one person, on one computer. No multi-user, no login/auth needed. |
| Where party data lives | Inside the website itself (currently in Excel, to be migrated in). User can add/edit/delete parties going forward. |
| Party fields to store | SR.NO. (auto), PARTICULARS (party name), BANK'S NAME, BANK ACCOUNT NUMBER, BANK'S LOCATION (branch), IFSC CODE |
| Per-payment-run input | Select parties from saved list + enter that run's AMOUNT for each (amount is not stored on the party — it changes every time) |
| Auto-calculated | Total amount, amount-in-words (Indian format), SR.NO. sequencing |
| Final output | Downloadable Excel (.xlsx) file matching the reference format, which the user prints and attaches to an email to the bank |
| Hosting / install | Runs locally, single computer, no server/cloud dependency required |

## 4. Open Questions to Settle Before/During Build

These weren't pinned down yet and should be decided (either by the stakeholder, or as a reasonable default the build tool picks and documents):

1. **Paying (own) account details** — NAME OF BANK A/C, NAME OF BANK, ACCOUNT NUMBER at the bottom of the sheet. Is this always the same single account, or does the user sometimes pay from more than one of their own accounts? → If always the same, store it once in Settings and auto-fill. If it varies, it needs to be a field selected per payment run.
2. **Cheque No. & Date** — entered manually per run (most likely), since cheque numbers come from a physical chequebook. Confirm it's a simple text/date input on the export screen, not auto-generated.
3. **"INT" prefix** seen before the amount-in-words ("INT ONE LAKH...") — confirm if this is a fixed label/abbreviation that should always prepend the words line, or specific to this example.
4. **Multiple payment runs / history** — does the user need to look back at past payment sheets (a history/log), or is each run "use and discard" once the Excel is downloaded? Affects whether past runs need to be saved.
5. **Data backup** — since data lives in the browser (locally), is there a need for an "export all party data" / "import" backup button, in case the browser data is ever cleared? Recommended even if not explicitly requested, as a safety net.
6. **Editing a party's bank details** — should historical exported sheets be unaffected if a party's bank details are later edited? (Standard expectation: yes — exports are a snapshot at generation time.)

## 5. Functional Scope (What To Build)

### 5.1 Party Directory (Master List)
- Add new party: PARTICULARS, BANK'S NAME, BANK ACCOUNT NUMBER, BANK'S LOCATION, IFSC CODE.
- Edit existing party.
- Delete party (with confirmation).
- List/search view of all saved parties.
- One-time import of the existing Excel list into the system (so the user doesn't retype everyone).

### 5.2 New Payment Run
- Search/select multiple parties from the saved directory.
- Enter the AMOUNT for this run, per selected party.
- Live-updating total as amounts are entered.
- Live amount-in-words preview (Indian numbering system).
- Fields for Cheque No., Date, and the paying account details (see Open Question 1–2).
- Remove a party from the run before finalizing, without affecting the saved directory.

### 5.3 Excel Export
- Generates a `.xlsx` file matching the reference format described in Section 2 exactly (columns, merged cells, bold headers, yellow highlight on total, amount-in-words row, cheque/date row, paying-account block).
- File should be named predictably, e.g. `Payment-Voucher-<DATE>-<ChequeNo>.xlsx`.
- Downloads directly to the user's computer for printing/emailing — no need to open Excel manually to reformat anything first.

### 5.4 Data Storage
- All data (party directory) persisted locally so it survives closing/reopening the site.
- No external server, database, or internet dependency required for core functionality.

## 6. Non-Functional Requirements
- **Simplicity first** — single user, no accounts, no permissions system, no unnecessary screens.
- **Offline-friendly** — should work without an internet connection once loaded, since it's for one computer and one person.
- **Speed of entry** — the whole point is reducing manual work, so selecting parties and entering amounts should take seconds, not minutes.
- **Output fidelity** — the generated Excel must be visually and structurally faithful to the existing bank-accepted format, since the bank is used to seeing it a specific way.

## 7. Suggested Build Approach (for Antigravity / dev tool)

- **Type of app:** Single-page web app, runs locally in a browser.
- **Data storage:** Browser-local storage (e.g. IndexedDB/localStorage) is sufficient given single-user, single-computer use — avoids needing a backend or database server. (If the build tool prefers a tiny local backend with a file-based DB like SQLite for more robust data safety, that's a reasonable alternative — flagged here as a build decision, not a hard requirement.)
- **Excel generation:** Use a spreadsheet-generation library capable of cell merging, fills/highlighting, and number formatting (e.g. SheetJS/`xlsx`, or `exceljs` for richer formatting support like merged cells and fill colors — `exceljs` is likely the better fit given the formatting needs in Section 2).
- **Amount-in-words:** Implement Indian numbering (Lakh/Crore) conversion logic — this is a small, well-defined utility function, not a library dependency necessarily, but double-check edge cases (zero, exact lakhs, etc.).

## 8. Suggested Build Sequence

1. Set up the basic app shell (navigation between "Party Directory" and "New Payment Run").
2. Build Party Directory: add/edit/delete/list, backed by local storage.
3. Import existing Excel party list into the directory (one-time migration screen or script).
4. Build New Payment Run screen: party selection, amount entry, live total + words preview.
5. Resolve Open Questions in Section 4 (especially paying-account details and cheque/date handling).
6. Build Excel export matching the reference format pixel-for-pixel (column widths, merges, bold, yellow fill, words row, cheque/date row, paying-account block).
7. Test export against the original sample file for a side-by-side format match.
8. (Optional/nice-to-have) Add a backup export/import for the party directory data, and/or a simple history log of past payment runs.

## 9. Out of Scope (for now, unless requirested later)
- Multi-user accounts / login.
- Direct bank API integration or automatic emailing (user will continue to manually attach the file to an email).
- Mobile app version — this is a single-computer desktop tool.
- Cloud sync / multi-device access.

---
*This document reflects the requirements gathered through discussion and one reference sample file as of June 2026. Confirm Section 4's open questions before or during the build to avoid rework.*
