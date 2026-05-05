/**
 * setup.js
 *
 * One-click scaffolding: creates all four sheets (Expenses,
 * Monthly_Summary, Goals, Loans) with headers, column widths,
 * frozen rows, and data-validation dropdowns.
 *
 * Run this function once after opening a blank spreadsheet.
 */

var EXPENSE_CATEGORIES = [
  "Personal",
  "Household",
  "Education",
  "Test",
  "Loan",
  "Transport",
  "Food",
];

var PAYMENT_METHODS = [
  "Cash",
  "UPI",
  "Credit Card",
  "Debit Card",
  "Net Banking",
];

var NECESSARY_OPTIONS = ["Yes", "No"];

var LOAN_STATUS_OPTIONS = ["Pending", "Returned"];

/**
 * Creates (or resets) all sheets and applies structure.
 * Safe to re-run – existing sheets are cleared, not deleted.
 */
function setupSpreadsheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  _setupExpensesSheet(ss);
  _setupMonthlySummarySheet(ss);
  _setupGoalsSheet(ss);
  _setupLoansSheet(ss);

  SpreadsheetApp.getUi().alert(
    "✅ Geld-Tracker is ready!\n\n" +
      "All 4 sheets have been created:\n" +
      "• Expenses\n• Monthly_Summary\n• Goals\n• Loans\n\n" +
      "Start logging your expenses in the Expenses sheet."
  );
}

// ---------------------------------------------------------------------------
// Private helpers
// ---------------------------------------------------------------------------

function _getOrCreateSheet(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  } else {
    sheet.clearContents();
    sheet.clearFormats();
  }
  return sheet;
}

function _setDropdown(sheet, firstRow, lastRow, col, options) {
  var range = sheet.getRange(firstRow, col, lastRow - firstRow + 1, 1);
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(options, true)
    .setAllowInvalid(false)
    .build();
  range.setDataValidation(rule);
}

function _setupExpensesSheet(ss) {
  var sheet = _getOrCreateSheet(ss, "Expenses");

  var headers = [
    "Date",
    "Category",
    "Subcategory",
    "Amount",
    "Payment Method",
    "Necessary?",
    "Notes",
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Style header row
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange
    .setBackground("#4A90D9")
    .setFontColor("#FFFFFF")
    .setFontWeight("bold");

  // Freeze header row
  sheet.setFrozenRows(1);

  // Column widths
  sheet.setColumnWidth(1, 100); // Date
  sheet.setColumnWidth(2, 120); // Category
  sheet.setColumnWidth(3, 130); // Subcategory
  sheet.setColumnWidth(4, 90); // Amount
  sheet.setColumnWidth(5, 130); // Payment Method
  sheet.setColumnWidth(6, 100); // Necessary?
  sheet.setColumnWidth(7, 200); // Notes

  // Dropdowns for rows 2–1000
  _setDropdown(sheet, 2, 1000, 2, EXPENSE_CATEGORIES); // Category
  _setDropdown(sheet, 2, 1000, 5, PAYMENT_METHODS); // Payment Method
  _setDropdown(sheet, 2, 1000, 6, NECESSARY_OPTIONS); // Necessary?

  // Date format on column A
  sheet
    .getRange(2, 1, 999, 1)
    .setNumberFormat("yyyy-mm-dd");
}

function _setupMonthlySummarySheet(ss) {
  var sheet = _getOrCreateSheet(ss, "Monthly_Summary");

  var headers = [
    "Month",
    "Total Income",
    "Total Expense",
    "Savings",
    "Savings %",
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange
    .setBackground("#34A853")
    .setFontColor("#FFFFFF")
    .setFontWeight("bold");

  sheet.setFrozenRows(1);

  // Seed rows for the next 12 months starting from current month
  var today = new Date();
  var rows = [];
  for (var i = 0; i < 12; i++) {
    var d = new Date(today.getFullYear(), today.getMonth() + i, 1);
    var monthStr = Utilities.formatDate(
      d,
      Session.getScriptTimeZone(),
      "yyyy-MM"
    );
    rows.push([monthStr, 0, 0, 0, 0]);
  }
  sheet.getRange(2, 1, rows.length, 5).setValues(rows);

  // Formulas for D (Savings) and E (Savings %) only.
  // Column C (Total Expense) is written by generateMonthlySummary() so
  // that a single source of truth owns the value and there is no conflict.
  for (var r = 2; r <= rows.length + 1; r++) {
    // Savings = Income - Expense
    sheet.getRange(`D${r}`).setFormula(`=B${r}-C${r}`);
    // Savings %
    sheet.getRange(`E${r}`).setFormula(`=IF(B${r}=0,0,(D${r}/B${r})*100)`);
  }

  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(3, 130);
  sheet.setColumnWidth(4, 100);
  sheet.setColumnWidth(5, 100);

  // Format B:D as currency, E as percentage text
  sheet.getRange(2, 2, rows.length, 3).setNumberFormat("₹#,##0.00");
  sheet.getRange(2, 5, rows.length, 1).setNumberFormat("0.00");
}

function _setupGoalsSheet(ss) {
  var sheet = _getOrCreateSheet(ss, "Goals");

  var headers = [
    "Goal Name",
    "Target Amount",
    "Saved Amount",
    "Remaining",
    "Deadline",
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange
    .setBackground("#FBBC04")
    .setFontColor("#000000")
    .setFontWeight("bold");

  sheet.setFrozenRows(1);

  // Remaining formula template for first 50 rows
  for (var r = 2; r <= 51; r++) {
    sheet.getRange(`D${r}`).setFormula(`=IF(B${r}="","",B${r}-C${r})`);
  }

  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 130);
  sheet.setColumnWidth(3, 130);
  sheet.setColumnWidth(4, 120);
  sheet.setColumnWidth(5, 110);

  sheet.getRange(2, 2, 50, 3).setNumberFormat("₹#,##0.00");
  sheet.getRange(2, 5, 50, 1).setNumberFormat("yyyy-mm-dd");
}

function _setupLoansSheet(ss) {
  var sheet = _getOrCreateSheet(ss, "Loans");

  var headers = [
    "Person",
    "Amount",
    "Given Date",
    "Return Date",
    "Status",
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange
    .setBackground("#EA4335")
    .setFontColor("#FFFFFF")
    .setFontWeight("bold");

  sheet.setFrozenRows(1);

  _setDropdown(sheet, 2, 1000, 5, LOAN_STATUS_OPTIONS);

  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 110);
  sheet.setColumnWidth(3, 110);
  sheet.setColumnWidth(4, 110);
  sheet.setColumnWidth(5, 110);

  sheet.getRange(2, 2, 999, 1).setNumberFormat("₹#,##0.00");
  sheet.getRange(2, 3, 999, 2).setNumberFormat("yyyy-mm-dd");
}
