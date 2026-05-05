/**
 * summary.js
 *
 * Script 2 – Monthly Summary Generator
 *
 * Reads every row in the Expenses sheet, totals up the amount
 * for the current calendar month, and writes the result into
 * Monthly_Summary!C (the "Total Expense" column) for the matching
 * month row.
 *
 * How to activate:
 *   1. Open Extensions → Apps Script.
 *   2. Click the clock icon (Triggers) in the left sidebar.
 *   3. Add trigger → Choose function: generateMonthlySummary
 *      → Event source: Time-driven → Month timer → Day 1
 *        (runs on the 1st of each month to finalise the previous month)
 *      OR add a second trigger on Day 28–31 for a mid-month preview.
 */

/**
 * Totals expenses for the current month and updates Monthly_Summary.
 */
function generateMonthlySummary() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var expenseSheet = ss.getSheetByName("Expenses");
  var summarySheet = ss.getSheetByName("Monthly_Summary");

  if (!expenseSheet || !summarySheet) {
    Logger.log(
      "Required sheets not found. Run setupSpreadsheet() first."
    );
    return;
  }

  var tz = Session.getScriptTimeZone();
  var currentMonth = Utilities.formatDate(new Date(), tz, "yyyy-MM");

  // Sum all expense amounts whose date falls in the current month
  var data = expenseSheet.getDataRange().getValues();
  var total = 0;

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue; // skip empty rows

    var date = new Date(row[0]);
    if (isNaN(date)) continue;

    var rowMonth = Utilities.formatDate(date, tz, "yyyy-MM");
    if (rowMonth === currentMonth) {
      var amount = parseFloat(row[3]);
      if (!isNaN(amount)) {
        total += amount;
      }
    }
  }

  // Find the matching month row in Monthly_Summary (column A)
  var summaryData = summarySheet.getDataRange().getValues();
  var targetRow = -1;

  for (var j = 1; j < summaryData.length; j++) {
    if (summaryData[j][0] === currentMonth) {
      targetRow = j + 1; // 1-based sheet row
      break;
    }
  }

  if (targetRow === -1) {
    // Month row doesn't exist yet – append one
    targetRow = summarySheet.getLastRow() + 1;
    summarySheet.getRange(targetRow, 1).setValue(currentMonth);
  }

  summarySheet.getRange(targetRow, 3).setValue(total);
  Logger.log(
    "Monthly summary updated: " + currentMonth + " → ₹" + total.toFixed(2)
  );
}
