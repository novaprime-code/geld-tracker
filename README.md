# 💰 Geld-Tracker

A personal finance tracker built on **Google Sheets + Apps Script**.

Track daily expenses, monitor savings goals, manage loans, and get automated
email reminders — all inside a free Google Spreadsheet.

---

## 📋 Sheet Structure

| Sheet | Purpose |
|-------|---------|
| `Expenses` | Log every transaction |
| `Monthly_Summary` | Auto-calculated income / expense / savings overview |
| `Goals` | Track progress toward savings goals |
| `Loans` | Record money given/taken on loan |

### Expenses columns

| A | B | C | D | E | F | G |
|---|---|---|---|---|---|---|
| Date | Category | Subcategory | Amount | Payment Method | Necessary? | Notes |

**Category dropdown:** Personal · Household · Education · Test · Loan · Transport · Food  
**Payment Method dropdown:** Cash · UPI · Credit Card · Debit Card · Net Banking  
**Necessary? dropdown:** Yes · No

### Monthly_Summary columns

| A | B | C | D | E |
|---|---|---|---|---|
| Month (yyyy-MM) | Total Income | Total Expense | Savings | Savings % |

> Columns C–E are formula-driven. Only fill in **B (Total Income)** each month.

### Goals columns

| A | B | C | D | E |
|---|---|---|---|---|
| Goal Name | Target Amount | Saved Amount | Remaining | Deadline |

> Column D is formula-driven (`=B-C`). Column C is updated automatically by
> the `updateGoals` script.

### Loans columns

| A | B | C | D | E |
|---|---|---|---|---|
| Person | Amount | Given Date | Return Date | Status |

**Status dropdown:** Pending · Returned

---

## 🚀 Quick Start

### 1 – Copy the script into your spreadsheet

1. Open (or create) a Google Spreadsheet.
2. Click **Extensions → Apps Script**.
3. Delete any existing code in the editor.
4. Create the following files in the Apps Script editor and paste the
   corresponding file from this repo's `src/` folder:

   | Apps Script file | Source file |
   |------------------|-------------|
   | `setup.js` | `src/setup.js` |
   | `reminders.js` | `src/reminders.js` |
   | `summary.js` | `src/summary.js` |
   | `highlight.js` | `src/highlight.js` |
   | `goals.js` | `src/goals.js` |
   | `webapp.js` | `src/webapp.js` |
   | `index.html` | `src/index.html` |

   Also replace the contents of `appsscript.json` (click **Project Settings →
   Show "appsscript.json"**) with the `appsscript.json` from this repo.

5. Click **Save** (💾).

### 2 – Deploy as a Web App

1. In the Apps Script editor click **Deploy → New deployment**.
2. Click the gear icon ⚙️ and choose **Web app**.
3. Fill in the settings:
   - **Execute as:** Me *(your Google account)*
   - **Who has access:** Anyone with Google Account *(or Anyone for public access)*
4. Click **Deploy**.
5. Copy the **Web app URL** — open it in your browser to use Geld-Tracker. 🎉

> **First-time use:** After opening the web app, go to the **Actions** tab and
> click **Run Setup** to create all four sheets in the spreadsheet automatically.
> You do **not** need to run anything manually in the Apps Script editor.

### 3 – Set up time-driven triggers (optional)

For fully automated background processing, go to
**Extensions → Apps Script → Triggers (clock icon)** and add:

| Function | Trigger type | Frequency |
|----------|-------------|-----------|
| `sendDailyReminder` | Time-driven | Daily · 9 PM–10 PM |
| `generateMonthlySummary` | Time-driven | Month timer · Day 1 |
| `highlightUnnecessary` | Time-driven | Weekly · Sunday |
| `updateGoals` | Time-driven | Daily · midnight |

All of these functions can also be triggered manually from the web app's
**Actions** tab at any time.

---

## 🔢 Key Formulas

### Total Expense for a month (Monthly_Summary!C)

> Column C is written by the `generateMonthlySummary()` script.
> Equivalent formula for reference:

```excel
=IFERROR(SUMIF(TEXT(Expenses!A:A,"yyyy-MM"),A2,Expenses!D:D),0)
```

### Savings (Monthly_Summary!D)

```excel
=B2-C2
```

### Savings % (Monthly_Summary!E)

```excel
=IF(B2=0,0,(D2/B2)*100)
```

### Remaining Goal Amount (Goals!D)

```excel
=IF(B2="","",B2-C2)
```

---

## ⚙️ Scripts overview

### `doGet()` / server-side API — `src/webapp.js`

Web App entry point. `doGet()` serves the `index.html` single-page UI.
All other exported functions (`getExpenses`, `addExpense`, `getGoals`, etc.)
are callable from the browser via `google.script.run` and act as the REST-like
backend for the web app.

### `setupSpreadsheet()` — `src/setup.js`

Creates all four sheets with headers, dropdowns, column widths, frozen rows,
and pre-filled formulas. **Run once** after pasting the scripts.

### `sendDailyReminder()` — `src/reminders.js`

Sends an email to your Gmail with today's date and a link to the spreadsheet.
Activate via a **daily time-driven trigger** at 9 PM.

### `generateMonthlySummary()` — `src/summary.js`

Totals the current month's expenses from the Expenses sheet and writes the
result into `Monthly_Summary!C` for the matching month row.

### `highlightUnnecessary()` — `src/highlight.js`

Colours every "Necessary? = No" expense row in light red (`#ffcccc`).
Idempotent — safe to run multiple times.

### `updateGoals()` — `src/goals.js`

Iterates all goals and sums matching expenses by Category / Subcategory name,
writing the total into `Goals!C` (Saved Amount).

---

## 💡 Tips

- Fill in **Total Income** in `Monthly_Summary!B` each month manually.
- To add a new savings goal, just type the goal name in `Goals!A` and the
  script will match expenses whose **Category** or **Subcategory** equals that
  name.
- Re-run `highlightUnnecessary()` any time to refresh the colour coding after
  adding new expenses.

---

## 🔮 Future Upgrade Ideas

- 📊 Charts tab with monthly expense breakdown
- 📱 Google Form for quick phone-based expense entry
- 🤖 Telegram bot integration for expense logging
- 📈 Savings timeline projection chart
