/**
 * webapp.js
 *
 * Web App entry point and server-side API for Geld-Tracker.
 *
 * How to deploy as a Web App:
 *   1. In the Apps Script editor click Deploy → New deployment.
 *   2. Type: Web app.
 *   3. Execute as: Me  (or "User accessing the web app").
 *   4. Who has access: Anyone with Google Account  (or Anyone).
 *   5. Click Deploy and copy the URL.
 *
 * All functions below are callable from the browser via google.script.run.
 * Functions that write to the spreadsheet reuse the private helpers already
 * defined in setup.js, summary.js, highlight.js, goals.js, and reminders.js
 * because every file in an Apps Script project shares the same global scope.
 */

// ---------------------------------------------------------------------------
// Web App entry point
// ---------------------------------------------------------------------------

/**
 * Serves the web app HTML page.
 * @param {GoogleAppsScript.Events.DoGet} e
 * @returns {GoogleAppsScript.HTML.HtmlOutput}
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile("index")
    .evaluate()
    .setTitle("💰 Geld-Tracker")
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

// ---------------------------------------------------------------------------
// Internal helpers
// ---------------------------------------------------------------------------

function _ss() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function _tz() {
  return Session.getScriptTimeZone();
}

function _fmtDate(val) {
  if (val instanceof Date && !isNaN(val)) {
    return Utilities.formatDate(val, _tz(), "yyyy-MM-dd");
  }
  return val ? String(val).substring(0, 10) : "";
}

function _sheet(name) {
  return _ss().getSheetByName(name);
}

// ---------------------------------------------------------------------------
// Meta / User info
// ---------------------------------------------------------------------------

/**
 * Returns the logged-in user's email and the spreadsheet URL/name.
 * @returns {{ email: string, spreadsheetUrl: string, spreadsheetName: string }}
 */
function getMetaInfo() {
  return {
    email: Session.getActiveUser().getEmail(),
    spreadsheetUrl: _ss().getUrl(),
    spreadsheetName: _ss().getName(),
  };
}

// ---------------------------------------------------------------------------
// Expenses
// ---------------------------------------------------------------------------

/**
 * Returns all expense rows (skipping empty rows).
 * @returns {Object[]|{error: string}}
 */
function getExpenses() {
  var sheet = _sheet("Expenses");
  if (!sheet) {
    return { error: "Expenses sheet not found. Please run Setup first." };
  }

  var data = sheet.getDataRange().getValues();
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (!r[0] && !r[1] && !r[3]) continue; // skip blank rows
    rows.push({
      rowIndex: i + 1, // 1-based sheet row (used for edits/deletes)
      date: _fmtDate(r[0]),
      category: r[1] || "",
      subcategory: r[2] || "",
      amount: Number(r[3]) || 0,
      paymentMethod: r[4] || "",
      necessary: r[5] || "",
      notes: r[6] || "",
    });
  }
  return rows;
}

/**
 * Appends a new expense row.
 * @param {{ date:string, category:string, subcategory:string, amount:number,
 *            paymentMethod:string, necessary:string, notes:string }} d
 * @returns {{ success: boolean }|{ error: string }}
 */
function addExpense(d) {
  var sheet = _sheet("Expenses");
  if (!sheet) return { error: "Expenses sheet not found." };
  sheet.appendRow([
    new Date(d.date),
    d.category || "",
    d.subcategory || "",
    parseFloat(d.amount) || 0,
    d.paymentMethod || "",
    d.necessary || "Yes",
    d.notes || "",
  ]);
  return { success: true };
}

/**
 * Deletes the expense row at the given 1-based sheet row index.
 * @param {number} rowIndex
 * @returns {{ success: boolean }|{ error: string }}
 */
function deleteExpense(rowIndex) {
  var sheet = _sheet("Expenses");
  if (!sheet) return { error: "Expenses sheet not found." };
  sheet.deleteRow(rowIndex);
  return { success: true };
}

// ---------------------------------------------------------------------------
// Monthly Summary
// ---------------------------------------------------------------------------

/**
 * Returns all Monthly_Summary rows.
 * @returns {Object[]|{ error: string }}
 */
function getMonthlySummary() {
  var sheet = _sheet("Monthly_Summary");
  if (!sheet) {
    return { error: "Monthly_Summary sheet not found. Please run Setup first." };
  }

  var data = sheet.getDataRange().getValues();
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (!r[0]) continue;
    rows.push({
      rowIndex: i + 1,
      month: r[0],
      income: Number(r[1]) || 0,
      expense: Number(r[2]) || 0,
      savings: Number(r[3]) || 0,
      savingsPct: Number(r[4]) || 0,
    });
  }
  return rows;
}

/**
 * Updates the Total Income cell for the given summary row.
 * @param {number} rowIndex  1-based sheet row
 * @param {number} income
 * @returns {{ success: boolean }|{ error: string }}
 */
function updateMonthlyIncome(rowIndex, income) {
  var sheet = _sheet("Monthly_Summary");
  if (!sheet) return { error: "Monthly_Summary sheet not found." };
  sheet.getRange(rowIndex, 2).setValue(parseFloat(income) || 0);
  return { success: true };
}

// ---------------------------------------------------------------------------
// Goals
// ---------------------------------------------------------------------------

/**
 * Returns all goals rows.
 * @returns {Object[]|{ error: string }}
 */
function getGoals() {
  var sheet = _sheet("Goals");
  if (!sheet) {
    return { error: "Goals sheet not found. Please run Setup first." };
  }

  var data = sheet.getDataRange().getValues();
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (!r[0]) continue;
    rows.push({
      rowIndex: i + 1,
      name: r[0],
      target: Number(r[1]) || 0,
      saved: Number(r[2]) || 0,
      remaining: Number(r[3]) || 0,
      deadline: _fmtDate(r[4]),
    });
  }
  return rows;
}

/**
 * Appends a new goal row and writes the Remaining formula.
 * @param {{ name:string, target:number, deadline:string }} d
 * @returns {{ success: boolean }|{ error: string }}
 */
function addGoal(d) {
  var sheet = _sheet("Goals");
  if (!sheet) return { error: "Goals sheet not found." };
  var nextRow = sheet.getLastRow() + 1;
  sheet.appendRow([
    d.name,
    parseFloat(d.target) || 0,
    0,
    "", // placeholder; formula set below
    d.deadline ? new Date(d.deadline) : "",
  ]);
  sheet
    .getRange(nextRow, 4)
    .setFormula(
      "=IF(B" + nextRow + '="","",B' + nextRow + "-C" + nextRow + ")"
    );
  return { success: true };
}

/**
 * Deletes the goal row at the given 1-based sheet row index.
 * @param {number} rowIndex
 * @returns {{ success: boolean }|{ error: string }}
 */
function deleteGoal(rowIndex) {
  var sheet = _sheet("Goals");
  if (!sheet) return { error: "Goals sheet not found." };
  sheet.deleteRow(rowIndex);
  return { success: true };
}

// ---------------------------------------------------------------------------
// Loans
// ---------------------------------------------------------------------------

/**
 * Returns all loans rows.
 * @returns {Object[]|{ error: string }}
 */
function getLoans() {
  var sheet = _sheet("Loans");
  if (!sheet) {
    return { error: "Loans sheet not found. Please run Setup first." };
  }

  var data = sheet.getDataRange().getValues();
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (!r[0]) continue;
    rows.push({
      rowIndex: i + 1,
      person: r[0],
      amount: Number(r[1]) || 0,
      givenDate: _fmtDate(r[2]),
      returnDate: _fmtDate(r[3]),
      status: r[4] || "Pending",
    });
  }
  return rows;
}

/**
 * Appends a new loan row.
 * @param {{ person:string, amount:number, givenDate:string,
 *            returnDate:string, status:string }} d
 * @returns {{ success: boolean }|{ error: string }}
 */
function addLoan(d) {
  var sheet = _sheet("Loans");
  if (!sheet) return { error: "Loans sheet not found." };
  sheet.appendRow([
    d.person,
    parseFloat(d.amount) || 0,
    d.givenDate ? new Date(d.givenDate) : "",
    d.returnDate ? new Date(d.returnDate) : "",
    d.status || "Pending",
  ]);
  return { success: true };
}

/**
 * Updates the Status cell of a loan row.
 * @param {number} rowIndex  1-based sheet row
 * @param {string} status
 * @returns {{ success: boolean }|{ error: string }}
 */
function updateLoanStatus(rowIndex, status) {
  var sheet = _sheet("Loans");
  if (!sheet) return { error: "Loans sheet not found." };
  sheet.getRange(rowIndex, 5).setValue(status);
  return { success: true };
}

/**
 * Deletes the loan row at the given 1-based sheet row index.
 * @param {number} rowIndex
 * @returns {{ success: boolean }|{ error: string }}
 */
function deleteLoan(rowIndex) {
  var sheet = _sheet("Loans");
  if (!sheet) return { error: "Loans sheet not found." };
  sheet.deleteRow(rowIndex);
  return { success: true };
}

// ---------------------------------------------------------------------------
// Actions  (wrappers for trigger-style functions, safe for web-app context)
// ---------------------------------------------------------------------------

/**
 * Creates/resets all four sheets without using SpreadsheetApp.getUi(),
 * which is unavailable in the web app execution context.
 * @returns {{ success: boolean, message: string }}
 */
function runSetup() {
  var ss = _ss();
  _setupExpensesSheet(ss);
  _setupMonthlySummarySheet(ss);
  _setupGoalsSheet(ss);
  _setupLoansSheet(ss);
  return {
    success: true,
    message:
      "✅ Setup complete! All 4 sheets created: Expenses, Monthly_Summary, Goals, Loans.",
  };
}

/**
 * Generates the monthly summary for the current calendar month.
 * @returns {{ success: boolean, message: string }}
 */
function runGenerateMonthlySummary() {
  generateMonthlySummary();
  return { success: true, message: "Monthly summary updated successfully." };
}

/**
 * Highlights "Necessary? = No" rows in the Expenses sheet.
 * @returns {{ success: boolean, message: string }}
 */
function runHighlightUnnecessary() {
  highlightUnnecessary();
  return {
    success: true,
    message: "Unnecessary expenses highlighted in the spreadsheet.",
  };
}

/**
 * Recalculates saved amounts for all goals.
 * @returns {{ success: boolean, message: string }}
 */
function runUpdateGoals() {
  updateGoals();
  return { success: true, message: "Goal saved amounts updated successfully." };
}

/**
 * Sends the daily reminder email.
 * @returns {{ success: boolean, message: string }}
 */
function runSendReminder() {
  sendDailyReminder();
  return { success: true, message: "Daily reminder email sent." };
}
