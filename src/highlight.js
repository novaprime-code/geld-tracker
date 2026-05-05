/**
 * highlight.js
 *
 * Script 3 – Highlight Unnecessary Spending
 *
 * Scans the Expenses sheet and colours every row where the
 * "Necessary?" column (column F, index 5) is "No" with a
 * light-red background so they stand out at a glance.
 *
 * Rows marked "Yes" have their background reset to white so
 * re-running the function is always safe and idempotent.
 *
 * How to activate:
 *   Option A – Run manually whenever you want a visual review.
 *   Option B – Add a trigger:
 *     Extensions → Apps Script → Triggers
 *     → highlightUnnecessary → Time-driven → Weekly (Sunday night)
 */

var HIGHLIGHT_COLOR = "#ffcccc"; // light red  – unnecessary rows
var CLEAR_COLOR = "#ffffff"; // white      – necessary / empty rows

/**
 * Highlights rows with Necessary? = "No" and clears others.
 */
function highlightUnnecessary() {
  var sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("Expenses");

  if (!sheet) {
    Logger.log("Expenses sheet not found. Run setupSpreadsheet() first.");
    return;
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log("No expense data found.");
    return;
  }

  // Read only the rows that contain data (A2:G<lastRow>)
  var dataRange = sheet.getRange(2, 1, lastRow - 1, 7);
  var data = dataRange.getValues();

  // Build background colour array in one pass to minimise API calls
  var backgrounds = [];
  var highlighted = 0;

  for (var i = 0; i < data.length; i++) {
    var isUnnecessary = String(data[i][5]).trim().toLowerCase() === "no";
    backgrounds.push(
      isUnnecessary
        ? [HIGHLIGHT_COLOR, HIGHLIGHT_COLOR, HIGHLIGHT_COLOR,
           HIGHLIGHT_COLOR, HIGHLIGHT_COLOR, HIGHLIGHT_COLOR,
           HIGHLIGHT_COLOR]
        : [CLEAR_COLOR, CLEAR_COLOR, CLEAR_COLOR,
           CLEAR_COLOR, CLEAR_COLOR, CLEAR_COLOR,
           CLEAR_COLOR]
    );
    if (isUnnecessary) highlighted++;
  }

  // Single batch write – much faster than row-by-row setBackground
  dataRange.setBackgrounds(backgrounds);

  Logger.log(
    "Highlight complete. " + highlighted + " unnecessary row(s) marked."
  );
}
