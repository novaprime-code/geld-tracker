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

   Also replace the contents of `appsscript.json` (click **Project Settings →
   Show "appsscript.json"**) with the `appsscript.json` from this repo.

5. Click **Save** (💾).

### 2 – Run the one-click setup

1. In the Apps Script editor select the function **`setupSpreadsheet`**.
2. Click **▶ Run**.
3. Accept the required permissions when prompted.
4. Switch back to your spreadsheet – all four sheets are now ready. 🎉

### 3 – Set up triggers

Go to **Extensions → Apps Script → Triggers (clock icon)** and add:

| Function | Trigger type | Frequency |
|----------|-------------|-----------|
| `sendDailyReminder` | Time-driven | Daily · 9 PM–10 PM |
| `generateMonthlySummary` | Time-driven | Month timer · Day 1 |
| `highlightUnnecessary` | Time-driven | Weekly · Sunday |
| `updateGoals` | Time-driven | Daily · midnight |

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
