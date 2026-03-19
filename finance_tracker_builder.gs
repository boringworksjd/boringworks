/**
 * Small Business Finance Tracker — Sheet Builder
 * ================================================
 * Paste this entire file into the Apps Script editor (script.google.com
 * or Extensions > Apps Script from any spreadsheet), then run
 * buildFinanceTracker(). A new spreadsheet is created and its URL is
 * logged to the execution log.
 *
 * Named ranges created (for use by future scripts):
 *   TransactionData       — Transactions!A2:N10000
 *   CategoryList          — Categories!A2:A12
 *   LowBalanceThreshold   — Settings!B2
 *   WeeklySummaryEmail    — Settings!B3
 */

// ── Constants ──────────────────────────────────────────────────────────────────

var CATEGORY_LIST = [
  "Revenue", "Labour", "Materials", "Fuel", "Subcontractors",
  "Equipment", "Utilities", "Insurance", "Office", "Meals", "Misc"
];

var MONTHS = [
  "Jan", "Feb", "Mar", "Apr", "May", "Jun",
  "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"
];

var TX_HEADERS = [
  "Timestamp",       // A
  "Date",            // B
  "Amount",          // C
  "Type",            // D  — dropdown: Income / Expense / Transfer
  "Category",        // E  — dropdown: CATEGORY_LIST
  "Vendor/Client",   // F
  "Description",     // G
  "Payment Method",  // H  — dropdown: Cash / Bank Transfer / Card / Check
  "Invoice ID",      // I
  "Due Date",        // J
  "Receipt Attached",// K  — dropdown: Y / N
  "Status",          // L  — dropdown: Paid / Unpaid / Pending
  "Source",          // M
  "Needs Review"     // N  — dropdown: Y / N
];

// Number of category rows (used to size named range)
var CAT_LAST_ROW = CATEGORY_LIST.length + 1; // row 2..12 → last row = 12

// ── Entry Point ────────────────────────────────────────────────────────────────

function buildFinanceTracker() {
  var ss         = SpreadsheetApp.create("Small Business Finance Tracker");
  var blankSheet = ss.getSheets()[0]; // default "Sheet1" — will be deleted

  // Insert tabs in display order (left → right)
  var dashboard      = ss.insertSheet("Dashboard");
  var transactions   = ss.insertSheet("Transactions");
  var categories     = ss.insertSheet("Categories");
  var monthlySummary = ss.insertSheet("Monthly Summary");
  var settings       = ss.insertSheet("Settings");

  ss.deleteSheet(blankSheet);

  // Build each tab
  _setupSettings(settings);
  _setupCategories(categories);
  _setupTransactions(transactions);
  _setupMonthlySummary(monthlySummary);
  _setupDashboard(dashboard);

  // Named ranges (must come after sheets are populated)
  _createNamedRanges(ss);

  SpreadsheetApp.flush();

  var url = ss.getUrl();
  Logger.log("Finance Tracker created: " + url);
  console.log("Finance Tracker URL: " + url);
}

// ── Settings Tab ───────────────────────────────────────────────────────────────

function _setupSettings(sheet) {
  // Header row
  sheet.getRange("A1:B1")
    .setValues([["Setting", "Value"]])
    .setFontWeight("bold")
    .setBackground("#1a73e8")
    .setFontColor("#ffffff");

  // Row 2 — Low Balance Alert Threshold
  sheet.getRange("A2").setValue("Low Balance Alert Threshold");
  sheet.getRange("B2").setValue(500).setNumberFormat('"$"#,##0.00');

  // Row 3 — Weekly Summary Email
  sheet.getRange("A3").setValue("Weekly Summary Email");
  sheet.getRange("B3").setValue("");

  sheet.setColumnWidth(1, 240);
  sheet.setColumnWidth(2, 220);
  sheet.setFrozenRows(1);
}

// ── Categories Tab ─────────────────────────────────────────────────────────────

function _setupCategories(sheet) {
  sheet.getRange("A1")
    .setValue("Category")
    .setFontWeight("bold")
    .setBackground("#1a73e8")
    .setFontColor("#ffffff");

  for (var i = 0; i < CATEGORY_LIST.length; i++) {
    sheet.getRange(i + 2, 1).setValue(CATEGORY_LIST[i]);
  }

  sheet.setColumnWidth(1, 180);
  sheet.setFrozenRows(1);
}

// ── Transactions Tab ───────────────────────────────────────────────────────────

function _setupTransactions(sheet) {
  // Headers
  var headerRange = sheet.getRange(1, 1, 1, TX_HEADERS.length);
  headerRange
    .setValues([TX_HEADERS])
    .setFontWeight("bold")
    .setBackground("#1a73e8")
    .setFontColor("#ffffff");

  sheet.setFrozenRows(1);

  // Column widths (A–N)
  var widths = [160, 100, 100, 100, 140, 150, 210, 140, 100, 100, 110, 90, 100, 110];
  for (var i = 0; i < widths.length; i++) {
    sheet.setColumnWidth(i + 1, widths[i]);
  }

  // Number / date formats on data rows
  sheet.getRange("A2:A10000").setNumberFormat("yyyy-MM-dd HH:mm:ss"); // Timestamp
  sheet.getRange("B2:B10000").setNumberFormat("yyyy-MM-dd");           // Date
  sheet.getRange("C2:C10000").setNumberFormat('"$"#,##0.00');          // Amount
  sheet.getRange("J2:J10000").setNumberFormat("yyyy-MM-dd");           // Due Date

  // ── Data Validation Dropdowns ──

  // Type (D)
  _setDropdown(sheet, "D2:D10000", ["Income", "Expense", "Transfer"]);

  // Category (E) — inline list mirrors Categories tab
  _setDropdown(sheet, "E2:E10000", CATEGORY_LIST);

  // Payment Method (H)
  _setDropdown(sheet, "H2:H10000", ["Cash", "Bank Transfer", "Card", "Check"]);

  // Receipt Attached (K)
  _setDropdown(sheet, "K2:K10000", ["Y", "N"]);

  // Status (L)
  _setDropdown(sheet, "L2:L10000", ["Paid", "Unpaid", "Pending"]);

  // Needs Review (N)
  _setDropdown(sheet, "N2:N10000", ["Y", "N"]);

  // Alternating row banding
  var banding = sheet.getRange("A1:N10000")
    .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
  banding.setHeaderRowColor("#1a73e8");
}

// ── Monthly Summary Tab ────────────────────────────────────────────────────────

function _setupMonthlySummary(sheet) {
  // Headers
  sheet.getRange(1, 1, 1, 4)
    .setValues([["Month", "Total Income", "Total Expenses", "Net Cash Flow"]])
    .setFontWeight("bold")
    .setBackground("#1a73e8")
    .setFontColor("#ffffff");

  // Year selector (top-right) — change this cell to view a different year
  sheet.getRange("F1").setValue("Year").setFontWeight("bold").setBackground("#1a73e8").setFontColor("#ffffff");
  sheet.getRange("G1")
    .setFormula("=YEAR(TODAY())")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 70);
  sheet.setColumnWidth(2, 140);
  sheet.setColumnWidth(3, 150);
  sheet.setColumnWidth(4, 140);

  // Rows 2–13 (Jan–Dec)
  // Formulas reference $G$1 for the year so users can change it.
  var txDate = "Transactions!$B$2:$B$10000";
  var txType = "Transactions!$D$2:$D$10000";
  var txAmt  = "Transactions!$C$2:$C$10000";

  for (var m = 0; m < 12; m++) {
    var row      = m + 2;
    var monthNum = m + 1;

    sheet.getRange(row, 1).setValue(MONTHS[m]);

    // Total Income
    sheet.getRange(row, 2).setFormula(
      "=IFERROR(SUMPRODUCT(" +
        "(" + txDate + "<>\"\")*" +
        "(MONTH(" + txDate + ")=" + monthNum + ")*" +
        "(YEAR(" + txDate + ")=$G$1)*" +
        "(" + txType + "=\"Income\")*" +
        "(" + txAmt + ")" +
      "),0)"
    );

    // Total Expenses
    sheet.getRange(row, 3).setFormula(
      "=IFERROR(SUMPRODUCT(" +
        "(" + txDate + "<>\"\")*" +
        "(MONTH(" + txDate + ")=" + monthNum + ")*" +
        "(YEAR(" + txDate + ")=$G$1)*" +
        "(" + txType + "=\"Expense\")*" +
        "(" + txAmt + ")" +
      "),0)"
    );

    // Net Cash Flow
    sheet.getRange(row, 4).setFormula("=B" + row + "-C" + row);
  }

  // Currency format on B:D data rows
  sheet.getRange("B2:D13").setNumberFormat('"$"#,##0.00');

  // Conditional formatting: highlight current month row
  var currentMonthRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=AND($G$1=YEAR(TODAY()),ROW()=MONTH(TODAY())+1)")
    .setBackground("#e8f0fe")
    .setRanges([sheet.getRange("A2:D13")])
    .build();

  // Negative net = light red
  var negNetRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0)
    .setFontColor("#b00020")
    .setRanges([sheet.getRange("D2:D13")])
    .build();

  sheet.setConditionalFormatRules([currentMonthRule, negNetRule]);
}

// ── Dashboard Tab ──────────────────────────────────────────────────────────────

function _setupDashboard(sheet) {
  // ── Title ──
  sheet.getRange("A1:D1")
    .merge()
    .setValue("Small Business Finance Tracker")
    .setFontSize(18)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground("#1a73e8")
    .setFontColor("#ffffff");

  sheet.getRange("A2").setValue("Period:");
  sheet.getRange("B2")
    .setFormula('=TEXT(TODAY(),"MMMM YYYY")')
    .setFontWeight("bold");

  // ── Section: Current Month ──
  _sectionHeader(sheet, "A4:D4", "CURRENT MONTH OVERVIEW");

  sheet.getRange("A5").setValue("Income");
  sheet.getRange("B5")
    .setFormula(_currentMonthSumFormula("Income"))
    .setNumberFormat('"$"#,##0.00');

  sheet.getRange("A6").setValue("Expenses");
  sheet.getRange("B6")
    .setFormula(_currentMonthSumFormula("Expense"))
    .setNumberFormat('"$"#,##0.00');

  sheet.getRange("A7").setValue("Net Cash Flow");
  sheet.getRange("B7")
    .setFormula("=B5-B6")
    .setNumberFormat('"$"#,##0.00')
    .setFontWeight("bold");

  // ── Section: Invoices & Review ──
  _sectionHeader(sheet, "A9:D9", "INVOICES & REVIEW");

  sheet.getRange("A10").setValue("Unpaid Invoices (Income)");
  sheet.getRange("B10").setFormula(
    '=COUNTIFS(Transactions!$L$2:$L$10000,"Unpaid",' +
    'Transactions!$D$2:$D$10000,"Income")'
  ).setFontWeight("bold");

  sheet.getRange("A11").setValue("Pending Transactions");
  sheet.getRange("B11").setFormula(
    '=COUNTIF(Transactions!$L$2:$L$10000,"Pending")'
  );

  sheet.getRange("A12").setValue("Flagged for Review");
  sheet.getRange("B12").setFormula(
    '=COUNTIF(Transactions!$N$2:$N$10000,"Y")'
  );

  // ── Section: Alerts ──
  _sectionHeader(sheet, "A14:D14", "ALERTS");

  sheet.getRange("A15").setValue("Low Balance Threshold");
  sheet.getRange("B15")
    .setFormula("=LowBalanceThreshold")
    .setNumberFormat('"$"#,##0.00');

  sheet.getRange("A16").setValue("Balance Status");
  sheet.getRange("B16").setFormula(
    '=IF(B7<B15,"⚠ Below Threshold — update Settings tab","✓ OK")'
  );

  sheet.getRange("A17").setValue("Summary Email");
  sheet.getRange("B17").setFormula("=WeeklySummaryEmail");

  // ── Conditional Formatting ──
  var rules = [];

  // Net cash flow: green / red
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThanOrEqualTo(0)
    .setFontColor("#0d7a3e").setBackground("#e6f4ea")
    .setRanges([sheet.getRange("B7")]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0)
    .setFontColor("#b00020").setBackground("#fce8e6")
    .setRanges([sheet.getRange("B7")]).build());

  // Balance status: red warning / green OK
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains("Below Threshold")
    .setFontColor("#b00020").setBackground("#fce8e6")
    .setRanges([sheet.getRange("B16")]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains("✓")
    .setFontColor("#0d7a3e").setBackground("#e6f4ea")
    .setRanges([sheet.getRange("B16")]).build());

  // Unpaid invoices: yellow if > 0
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0)
    .setBackground("#fef9c3").setFontColor("#854d0e")
    .setRanges([sheet.getRange("B10")]).build());

  sheet.setConditionalFormatRules(rules);

  // ── Column Widths ──
  sheet.setColumnWidth(1, 210);
  sheet.setColumnWidth(2, 180);
  sheet.setColumnWidth(3, 160);
  sheet.setColumnWidth(4, 160);
}

// ── Named Ranges ───────────────────────────────────────────────────────────────

function _createNamedRanges(ss) {
  var tx   = ss.getSheetByName("Transactions");
  var cat  = ss.getSheetByName("Categories");
  var sett = ss.getSheetByName("Settings");

  // Full transaction data body (header excluded)
  ss.setNamedRange("TransactionData", tx.getRange("A2:N10000"));

  // Category list (A2:A12)
  ss.setNamedRange("CategoryList", cat.getRange("A2:A" + CAT_LAST_ROW));

  // Settings values
  ss.setNamedRange("LowBalanceThreshold", sett.getRange("B2"));
  ss.setNamedRange("WeeklySummaryEmail",  sett.getRange("B3"));
}

// ── Helpers ────────────────────────────────────────────────────────────────────

/** Apply a dropdown list validation to a range. */
function _setDropdown(sheet, rangeA1, options) {
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(options, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(rangeA1).setDataValidation(rule);
}

/** Style a merged range as a section header. */
function _sectionHeader(sheet, rangeA1, label) {
  sheet.getRange(rangeA1)
    .merge()
    .setValue(label)
    .setFontWeight("bold")
    .setFontSize(11)
    .setBackground("#e8f0fe")
    .setFontColor("#1a73e8");
}

/**
 * SUMPRODUCT formula that totals Amount on Transactions where
 * Date is in the current month/year and Type matches txType.
 */
function _currentMonthSumFormula(txType) {
  var date = "Transactions!$B$2:$B$10000";
  var type = "Transactions!$D$2:$D$10000";
  var amt  = "Transactions!$C$2:$C$10000";
  return (
    "=IFERROR(SUMPRODUCT(" +
      "(" + date + "<>\"\")*" +
      "(MONTH(" + date + ")=MONTH(TODAY()))*" +
      "(YEAR(" + date + ")=YEAR(TODAY()))*" +
      "(" + type + "=\"" + txType + "\")*" +
      "(" + amt + ")" +
    "),0)"
  );
}
