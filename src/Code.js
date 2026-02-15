/**
 * Habit Tracker - Google Apps Script
 * Track exercise, alcohol consumption, and weight daily.
 * Features: daily reminders, streak calculation, weekly summary.
 * Version: 1.1.0
 */

// ============================================================
// Constants
// ============================================================

const SHEET_NAME_TRACKING = "Tracking";
const SHEET_NAME_DASHBOARD = "Dashboard";

const COL = {
  DATE: 1,
  EXERCISE_TYPE: 2,
  EXERCISE_MIN: 3,
  BEER_COUNT: 4,
  WEIGHT: 5,
  NOTES: 6,
};

const HEADER_ROW = [
  "Date",
  "Exercise Type",
  "Duration (min)",
  "Beers",
  "Weight (kg)",
  "Notes",
];

const EXERCISE_TYPES = [
  "Gym",
  "Running",
  "Walking",
  "Cycling",
  "Swimming",
  "Yoga",
  "Home Workout",
  "Other",
  "Rest Day",
];

const REMINDER_HOUR = 21; // 9 PM
const WEEKLY_SUMMARY_HOUR = 8; // 8 AM Monday

// ============================================================
// Menu & Initialization
// ============================================================

/**
 * Add custom menu when spreadsheet opens.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Habit Tracker")
    .addItem("Add Today's Row", "addTodayRow")
    .addItem("Update Dashboard", "updateDashboard")
    .addSeparator()
    .addItem("Setup Triggers", "setupTriggers")
    .addItem("Initialize Sheet", "initializeSheet")
    .addToUi();
}

/**
 * Initialize the spreadsheet with Tracking and Dashboard sheets.
 */
function initializeSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- Tracking Sheet ---
  let tracking = ss.getSheetByName(SHEET_NAME_TRACKING);
  if (!tracking) {
    tracking = ss.insertSheet(SHEET_NAME_TRACKING);
  }

  // Set headers
  const headerRange = tracking.getRange(1, 1, 1, HEADER_ROW.length);
  headerRange.setValues([HEADER_ROW]);
  headerRange.setFontWeight("bold");
  headerRange.setBackground("#4a86c8");
  headerRange.setFontColor("#ffffff");

  // Column widths
  tracking.setColumnWidth(COL.DATE, 110);
  tracking.setColumnWidth(COL.EXERCISE_TYPE, 140);
  tracking.setColumnWidth(COL.EXERCISE_MIN, 120);
  tracking.setColumnWidth(COL.BEER_COUNT, 80);
  tracking.setColumnWidth(COL.WEIGHT, 100);
  tracking.setColumnWidth(COL.NOTES, 250);

  // Data validation for exercise type
  const exerciseRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(EXERCISE_TYPES, true)
    .setAllowInvalid(false)
    .build();
  tracking
    .getRange(2, COL.EXERCISE_TYPE, 500)
    .setDataValidation(exerciseRule);

  // Data validation for beer count (0-20)
  const beerRule = SpreadsheetApp.newDataValidation()
    .requireNumberBetween(0, 20)
    .setAllowInvalid(false)
    .build();
  tracking.getRange(2, COL.BEER_COUNT, 500).setDataValidation(beerRule);

  // Date format
  tracking.getRange(2, COL.DATE, 500).setNumberFormat("yyyy-mm-dd");

  // Freeze header row
  tracking.setFrozenRows(1);

  // --- Dashboard Sheet ---
  let dashboard = ss.getSheetByName(SHEET_NAME_DASHBOARD);
  if (!dashboard) {
    dashboard = ss.insertSheet(SHEET_NAME_DASHBOARD);
  }
  buildDashboard_(dashboard);

  // Remove default "Sheet1" if it exists
  const defaultSheet = ss.getSheetByName("Sheet1");
  if (defaultSheet && ss.getSheets().length > 1) {
    ss.deleteSheet(defaultSheet);
  }

  // Add today's row
  addTodayRow();

  SpreadsheetApp.getUi().alert(
    "Initialization complete!\n\n" +
      "- Tracking sheet created\n" +
      "- Dashboard sheet created\n" +
      "- Today's row added\n\n" +
      'Run "Setup Triggers" to enable daily reminders.'
  );
}

// ============================================================
// Daily Row Management
// ============================================================

/**
 * Add a new row for today if it doesn't already exist.
 */
function addTodayRow() {
  const sheet = getTrackingSheet_();
  const today = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyy-MM-dd"
  );

  // Check if today already has a row
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][COL.DATE - 1] instanceof Date) {
      const rowDate = Utilities.formatDate(
        data[i][COL.DATE - 1],
        Session.getScriptTimeZone(),
        "yyyy-MM-dd"
      );
      if (rowDate === today) {
        return; // Already exists
      }
    }
  }

  // Append new row
  const newRow = sheet.getLastRow() + 1;
  sheet.getRange(newRow, COL.DATE).setValue(new Date());
  sheet
    .getRange(newRow, COL.DATE)
    .setNumberFormat("yyyy-mm-dd");
}

/**
 * Auto-add row at the start of each day (triggered by daily trigger).
 * Uses LockService to prevent duplicate execution.
 */
function dailyAutoAddRow() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return;
  try {
    addTodayRow();
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// Streak Calculation
// ============================================================

/**
 * Calculate current exercise streak (consecutive days with exercise).
 * @return {number} Number of consecutive exercise days.
 */
function calculateExerciseStreak() {
  const data = getTrackingData_();
  let streak = 0;

  for (let i = data.length - 1; i >= 0; i--) {
    const exerciseType = data[i][COL.EXERCISE_TYPE - 1];
    if (exerciseType && exerciseType !== "Rest Day" && exerciseType !== "") {
      streak++;
    } else {
      break;
    }
  }

  return streak;
}

/**
 * Calculate current no-alcohol streak (consecutive days with 0 beers).
 * @return {number} Number of consecutive no-alcohol days.
 */
function calculateNoAlcoholStreak() {
  const data = getTrackingData_();
  let streak = 0;

  for (let i = data.length - 1; i >= 0; i--) {
    const beerCount = data[i][COL.BEER_COUNT - 1];
    if (beerCount === 0 || beerCount === "" || beerCount === null) {
      streak++;
    } else {
      break;
    }
  }

  return streak;
}

/**
 * Calculate the longest exercise streak ever.
 * @return {number} Longest exercise streak.
 */
function calculateLongestExerciseStreak() {
  const data = getTrackingData_();
  let longest = 0;
  let current = 0;

  for (let i = 0; i < data.length; i++) {
    const exerciseType = data[i][COL.EXERCISE_TYPE - 1];
    if (exerciseType && exerciseType !== "Rest Day" && exerciseType !== "") {
      current++;
      longest = Math.max(longest, current);
    } else {
      current = 0;
    }
  }

  return longest;
}

// ============================================================
// Dashboard
// ============================================================

/**
 * Update the dashboard sheet with current stats.
 */
function updateDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let dashboard = ss.getSheetByName(SHEET_NAME_DASHBOARD);
  if (!dashboard) {
    dashboard = ss.insertSheet(SHEET_NAME_DASHBOARD);
  }

  dashboard.clear();
  buildDashboard_(dashboard);
}

/**
 * Build dashboard content.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Dashboard sheet.
 * @private
 */
function buildDashboard_(sheet) {
  const data = getTrackingData_();

  // Title
  sheet.getRange("A1").setValue("Habit Tracker Dashboard");
  sheet.getRange("A1").setFontSize(18).setFontWeight("bold");

  const lastUpdated = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyy-MM-dd HH:mm"
  );
  sheet.getRange("A2").setValue(`Last updated: ${lastUpdated}`);
  sheet.getRange("A2").setFontColor("#888888");

  // --- Streaks Section ---
  sheet.getRange("A4").setValue("Current Streaks").setFontSize(14).setFontWeight("bold");

  const exerciseStreak = calculateExerciseStreak();
  const noAlcoholStreak = calculateNoAlcoholStreak();
  const longestExercise = calculateLongestExerciseStreak();

  const streakData = [
    ["Exercise Streak", `${exerciseStreak} days`],
    ["No-Alcohol Streak", `${noAlcoholStreak} days`],
    ["Longest Exercise Streak", `${longestExercise} days`],
  ];
  sheet.getRange(5, 1, streakData.length, 2).setValues(streakData);

  // --- This Week Summary ---
  sheet.getRange("A9").setValue("This Week").setFontSize(14).setFontWeight("bold");
  const weekStats = getWeekStats_();

  const weekData = [
    ["Exercise Days", `${weekStats.exerciseDays} / 7`],
    ["Total Exercise Time", `${weekStats.totalMinutes} min`],
    ["Total Beers", `${weekStats.totalBeers}`],
    ["Avg Weight", weekStats.avgWeight ? `${weekStats.avgWeight.toFixed(1)} kg` : "No data"],
  ];
  sheet.getRange(10, 1, weekData.length, 2).setValues(weekData);

  // --- Monthly Trend ---
  sheet.getRange("A15").setValue("Last 30 Days").setFontSize(14).setFontWeight("bold");
  const monthStats = getMonthStats_();

  const monthData = [
    ["Exercise Days", `${monthStats.exerciseDays} / 30`],
    ["No-Alcohol Days", `${monthStats.noAlcoholDays} / 30`],
    ["Beer-Free Rate", `${monthStats.beerFreeRate}%`],
  ];
  sheet.getRange(16, 1, monthData.length, 2).setValues(monthData);

  // Formatting
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 150);

  // Conditional color for streaks
  colorStreakCell_(sheet.getRange("B5"), exerciseStreak);
  colorStreakCell_(sheet.getRange("B6"), noAlcoholStreak);
}

/**
 * Color a cell based on streak value.
 * @param {GoogleAppsScript.Spreadsheet.Range} range
 * @param {number} streak
 * @private
 */
function colorStreakCell_(range, streak) {
  if (streak >= 7) {
    range.setBackground("#c6efce").setFontColor("#006100"); // Green
  } else if (streak >= 3) {
    range.setBackground("#ffeb9c").setFontColor("#9c5700"); // Yellow
  } else {
    range.setBackground("#ffc7ce").setFontColor("#9c0006"); // Red
  }
}

// ============================================================
// Statistics Helpers
// ============================================================

/**
 * Get stats for the current week (Mon-Sun).
 * @return {{exerciseDays: number, totalMinutes: number, totalBeers: number, avgWeight: number|null}}
 * @private
 */
function getWeekStats_() {
  const now = new Date();
  const dayOfWeek = now.getDay();
  const mondayOffset = dayOfWeek === 0 ? 6 : dayOfWeek - 1;
  const monday = new Date(now);
  monday.setDate(now.getDate() - mondayOffset);
  monday.setHours(0, 0, 0, 0);

  const data = getTrackingData_();
  let exerciseDays = 0;
  let totalMinutes = 0;
  let totalBeers = 0;
  let weights = [];

  for (const row of data) {
    if (!(row[COL.DATE - 1] instanceof Date)) continue;
    const rowDate = normalizeDate_(row[COL.DATE - 1]);
    if (rowDate < monday) continue;

    const exerciseType = row[COL.EXERCISE_TYPE - 1];
    if (exerciseType && exerciseType !== "Rest Day" && exerciseType !== "") {
      exerciseDays++;
    }

    const minutes = row[COL.EXERCISE_MIN - 1];
    if (typeof minutes === "number") {
      totalMinutes += minutes;
    }

    const beers = row[COL.BEER_COUNT - 1];
    if (typeof beers === "number") {
      totalBeers += beers;
    }

    const weight = row[COL.WEIGHT - 1];
    if (typeof weight === "number" && weight > 0) {
      weights.push(weight);
    }
  }

  const avgWeight = weights.length > 0
    ? weights.reduce((a, b) => a + b, 0) / weights.length
    : null;

  return { exerciseDays, totalMinutes, totalBeers, avgWeight };
}

/**
 * Get stats for the last 30 days.
 * @return {{exerciseDays: number, noAlcoholDays: number, beerFreeRate: string}}
 * @private
 */
function getMonthStats_() {
  const now = new Date();
  const thirtyDaysAgo = new Date(now);
  thirtyDaysAgo.setDate(now.getDate() - 30);
  thirtyDaysAgo.setHours(0, 0, 0, 0);

  const data = getTrackingData_();
  let exerciseDays = 0;
  let noAlcoholDays = 0;
  let totalDays = 0;

  for (const row of data) {
    if (!(row[COL.DATE - 1] instanceof Date)) continue;
    const rowDate = normalizeDate_(row[COL.DATE - 1]);
    if (rowDate < thirtyDaysAgo) continue;

    totalDays++;

    const exerciseType = row[COL.EXERCISE_TYPE - 1];
    if (exerciseType && exerciseType !== "Rest Day" && exerciseType !== "") {
      exerciseDays++;
    }

    const beers = row[COL.BEER_COUNT - 1];
    if (beers === 0 || beers === "" || beers === null) {
      noAlcoholDays++;
    }
  }

  const beerFreeRate = totalDays > 0
    ? ((noAlcoholDays / totalDays) * 100).toFixed(0)
    : "0";

  return { exerciseDays, noAlcoholDays, beerFreeRate };
}

// ============================================================
// Reminders & Notifications
// ============================================================

/**
 * Send daily reminder email at 9 PM.
 * Uses LockService and PropertiesService to prevent duplicates.
 */
function sendDailyReminder() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return;

  try {
    const today = Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      "yyyy-MM-dd"
    );

    // Prevent duplicate emails via PropertiesService
    const props = PropertiesService.getScriptProperties();
    const lastReminder = props.getProperty("lastReminderDate");
    if (lastReminder === today) return;

    const sheet = getTrackingSheet_();
    const data = sheet.getDataRange().getValues();
    let todayFilled = false;

    for (let i = 1; i < data.length; i++) {
      if (data[i][COL.DATE - 1] instanceof Date) {
        const rowDate = Utilities.formatDate(
          data[i][COL.DATE - 1],
          Session.getScriptTimeZone(),
          "yyyy-MM-dd"
        );
        if (rowDate === today) {
          const hasExercise = data[i][COL.EXERCISE_TYPE - 1] !== "";
          if (hasExercise) {
            todayFilled = true;
          }
          break;
        }
      }
    }

    if (todayFilled) return; // Already recorded

    const exerciseStreak = calculateExerciseStreak();
    const noAlcoholStreak = calculateNoAlcoholStreak();
    const ssUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();

    const subject = "Habit Tracker: Record Today's Habits";
    const body = [
      "Today's habits haven't been recorded yet!",
      "",
      `Exercise Streak: ${exerciseStreak} days`,
      `No-Alcohol Streak: ${noAlcoholStreak} days`,
      "",
      `Record now: ${ssUrl}`,
      "",
      "Keep it up!",
    ].join("\n");

    MailApp.sendEmail({
      to: Session.getEffectiveUser().getEmail(),
      subject: subject,
      body: body,
    });

    props.setProperty("lastReminderDate", today);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Send weekly summary email on Monday morning.
 * Uses LockService to prevent duplicate execution.
 */
function sendWeeklySummary() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return;

  try {
  const weekStats = getWeekStats_();
  const exerciseStreak = calculateExerciseStreak();
  const noAlcoholStreak = calculateNoAlcoholStreak();
  const longestStreak = calculateLongestExerciseStreak();
  const ssUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();

  const subject = "Habit Tracker: Weekly Summary";
  const body = [
    "=== Weekly Summary ===",
    "",
    `Exercise Days: ${weekStats.exerciseDays} / 7`,
    `Total Exercise: ${weekStats.totalMinutes} min`,
    `Total Beers: ${weekStats.totalBeers}`,
    weekStats.avgWeight
      ? `Average Weight: ${weekStats.avgWeight.toFixed(1)} kg`
      : "Weight: No data",
    "",
    "--- Streaks ---",
    `Current Exercise Streak: ${exerciseStreak} days`,
    `Current No-Alcohol Streak: ${noAlcoholStreak} days`,
    `All-Time Best Exercise Streak: ${longestStreak} days`,
    "",
    `View details: ${ssUrl}`,
    "",
    "Have a great week!",
  ].join("\n");

  MailApp.sendEmail({
    to: Session.getEffectiveUser().getEmail(),
    subject: subject,
    body: body,
  });

  // Also update dashboard
  updateDashboard();
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// Trigger Management
// ============================================================

/**
 * Set up all time-based triggers.
 */
function setupTriggers() {
  // Remove existing triggers to avoid duplicates
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    const handlerName = trigger.getHandlerFunction();
    if (
      handlerName === "sendDailyReminder" ||
      handlerName === "dailyAutoAddRow" ||
      handlerName === "sendWeeklySummary"
    ) {
      ScriptApp.deleteTrigger(trigger);
    }
  }

  // Daily: add row at midnight
  ScriptApp.newTrigger("dailyAutoAddRow")
    .timeBased()
    .everyDays(1)
    .atHour(0)
    .create();

  // Daily: reminder at 9 PM
  ScriptApp.newTrigger("sendDailyReminder")
    .timeBased()
    .everyDays(1)
    .atHour(REMINDER_HOUR)
    .create();

  // Weekly: summary on Monday at 8 AM
  ScriptApp.newTrigger("sendWeeklySummary")
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(WEEKLY_SUMMARY_HOUR)
    .create();

  SpreadsheetApp.getUi().alert(
    "Triggers set up:\n\n" +
      "- Daily: Auto-add row at midnight\n" +
      "- Daily: Reminder email at 9 PM\n" +
      "- Weekly: Summary email on Monday 8 AM"
  );
}

// ============================================================
// Utility Functions
// ============================================================

/**
 * Normalize a date to midnight for safe comparison.
 * @param {Date} date
 * @return {Date} Date with time set to 00:00:00.
 * @private
 */
function normalizeDate_(date) {
  const normalized = new Date(date);
  normalized.setHours(0, 0, 0, 0);
  return normalized;
}

/**
 * Get the Tracking sheet.
 * @return {GoogleAppsScript.Spreadsheet.Sheet}
 * @private
 */
function getTrackingSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME_TRACKING);
  if (!sheet) {
    throw new Error(
      `"${SHEET_NAME_TRACKING}" sheet not found. Run "Initialize Sheet" first.`
    );
  }
  return sheet;
}

/**
 * Get all tracking data (excluding header).
 * @return {Array<Array<*>>}
 * @private
 */
function getTrackingData_() {
  const sheet = getTrackingSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  return sheet
    .getRange(2, 1, lastRow - 1, HEADER_ROW.length)
    .getValues();
}
