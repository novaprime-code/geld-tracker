/**
 * reminders.js
 *
 * Script 1 – Daily Reminder
 *
 * Sends a daily email nudge to the spreadsheet owner reminding
 * them to log their expenses for the day.
 *
 * How to activate:
 *   1. Open Extensions → Apps Script.
 *   2. Click the clock icon (Triggers) in the left sidebar.
 *   3. Add trigger → Choose function: sendDailyReminder
 *      → Event source: Time-driven → Day timer → 9 PM to 10 PM
 */

/**
 * Sends a reminder email to the active user's Gmail address.
 */
function sendDailyReminder() {
  var email = Session.getActiveUser().getEmail();
  var today = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "EEEE, MMMM d yyyy"
  );

  var subject = "💰 Update Your Geld-Tracker – " + today;
  var body =
    "Hey! 👋\n\n" +
    "Just a friendly reminder to log today's expenses in your Geld-Tracker.\n\n" +
    "📅 Date: " +
    today +
    "\n\n" +
    "Staying consistent is the key to reaching your financial goals. " +
    "Even one entry a day adds up to powerful insights over time. 🚀\n\n" +
    "Open your tracker here:\n" +
    SpreadsheetApp.getActiveSpreadsheet().getUrl() +
    "\n\n" +
    "— Geld-Tracker Bot 🤖";

  MailApp.sendEmail(email, subject, body);
  Logger.log("Daily reminder sent to " + email);
}
