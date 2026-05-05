/**
 * goals.js
 *
 * Script 4 – Goal Progress Auto-Update
 *
 * Sums the "Saved Amount" for each goal by matching the goal's name
 * against the Category or Subcategory column of the Expenses sheet,
 * then writes the total into the "Saved Amount" (column C) of the
 * Goals sheet.
 *
 * Default behaviour (matching the problem statement):
 *   • "Education" goal ← sums Expenses rows where Category = "Education"
 *
 * For every other goal row in the Goals sheet the function looks for
 * expense rows whose Category OR Subcategory matches the goal name
 * (case-insensitive), giving a general-purpose matching strategy.
 *
 * How to activate:
 *   1. Open Extensions → Apps Script.
 *   2. Click the clock icon (Triggers) in the left sidebar.
 *   3. Add trigger → Choose function: updateGoals
 *      → Event source: Time-driven → Day timer → midnight
 *        (runs nightly to keep goal progress current)
 */

/**
 * Updates the "Saved Amount" column of the Goals sheet for every goal.
 */
function updateGoals() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var expenseSheet = ss.getSheetByName("Expenses");
  var goalSheet = ss.getSheetByName("Goals");

  if (!expenseSheet || !goalSheet) {
    Logger.log(
      "Required sheets not found. Run setupSpreadsheet() first."
    );
    return;
  }

  var expenses = expenseSheet.getDataRange().getValues();
  var goals = goalSheet.getDataRange().getValues();

  // Pre-build a map: goalName (lower) → total saved
  var totals = {};

  for (var i = 1; i < goals.length; i++) {
    var goalName = String(goals[i][0]).trim();
    if (!goalName) continue;
    totals[goalName.toLowerCase()] = { name: goalName, row: i + 1, total: 0 };
  }

  // Walk every expense row and accumulate amounts
  for (var e = 1; e < expenses.length; e++) {
    var expRow = expenses[e];
    if (!expRow[0]) continue; // skip empty

    var category = String(expRow[1]).trim().toLowerCase();
    var subcategory = String(expRow[2]).trim().toLowerCase();
    var amount = parseFloat(expRow[3]);
    if (isNaN(amount)) continue;

    // Match against any goal whose name equals Category or Subcategory
    for (var key in totals) {
      if (category === key || subcategory === key) {
        totals[key].total += amount;
      }
    }
  }

  // Write results back to Goals sheet column C (Saved Amount)
  for (var key in totals) {
    var entry = totals[key];
    goalSheet.getRange(entry.row, 3).setValue(entry.total);
    Logger.log(
      "Goal '" + entry.name + "' saved amount → ₹" + entry.total.toFixed(2)
    );
  }
}
