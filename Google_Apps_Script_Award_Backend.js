/**
 * EXIUM MUPS - SR. / MIO AWARD CATALOGUE CHOICE BACKEND
 * Google Apps Script Web App
 * 
 * Features:
 * 1. Multi-Month Sheet Management ('Award_May_2026', 'Award_June_2026', and dynamic future months).
 * 2. High-Performance Concurrency Lock (LockService) to prevent race conditions.
 * 3. Anti-Overwrite Protection: Blank/corrupt submissions will NEVER overwrite an existing choice.
 * 4. Row Matching by Current 'SAP Area Code' (Col 7) + 'SAP MIO Code' (Col 9) based on FF List Master.
 * 5. Instant summary & JSON export for web application sync.
 */

var VOUCHER_OPTIONS = [
  "Infinity Gift Voucher",
  "Apex Gift Voucher",
  "Bata Gift Voucher",
  "Aarong Gift Voucher",
  "Best Buy Gift Voucher",
  "Cats Eye Gift Voucher"
];

function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var p = (e && e.parameter) ? e.parameter : {};
  var action = p.action || "summary";

  if (action === "fix_formatting") {
    var msg = fixAllSheetsFormattingAndTimestamps(ss);
    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      message: msg,
      timestamp: Utilities.formatDate(new Date(), "Asia/Dhaka", "dd-MMM-yyyy hh:mm:ss a")
    })).setMimeType(ContentService.MimeType.JSON);
  }

  var props = PropertiesService.getScriptProperties();
  var isLocked = (props.getProperty("SUBMISSIONS_LOCKED") === "true");

  if (action === "fetch_data") {
    var month = p.month || "all";
    var result = {};
    if (month === "all" || month === "May_2026") {
      result["May_2026"] = getSheetChoices(ss, "Award_May_2026");
    }
    if (month === "all" || month === "June_2026") {
      result["June_2026"] = getSheetChoices(ss, "Award_June_2026");
    }
    var jsonStr = JSON.stringify({
      status: "success",
      data: result,
      submissions_locked: isLocked,
      timestamp: Utilities.formatDate(new Date(), "Asia/Dhaka", "dd-MMM-yyyy hh:mm:ss a")
    });
    if (p.callback) {
      return ContentService.createTextOutput(p.callback + "(" + jsonStr + ")")
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(jsonStr).setMimeType(ContentService.MimeType.JSON);
  }

  var summary = getAwardSummary(ss);
  return ContentService.createTextOutput(JSON.stringify({
    status: "ok",
    message: "Exium Award Choice Backend is Active",
    summary: summary,
    submissions_locked: isLocked,
    timestamp: Utilities.formatDate(new Date(), "Asia/Dhaka", "dd-MMM-yyyy hh:mm:ss a")
  })).setMimeType(ContentService.MimeType.JSON);
}

// Convert any input timestamp/ISO string into strict Bangladesh Standard Time (BDT / UTC+6)
function toBangladeshTimeString(input) {
  if (!input) {
    return Utilities.formatDate(new Date(), "Asia/Dhaka", "dd-MMM-yyyy hh:mm:ss a");
  }
  if (typeof input === "string") {
    var trimmed = input.trim();
    if (/^\d{2}-[A-Za-z]{3}-\d{4}\s+\d{1,2}:\d{2}:\d{2}\s+(AM|PM)/i.test(trimmed)) {
      return trimmed;
    }
    try {
      var d = new Date(trimmed);
      if (!isNaN(d.getTime())) {
        return Utilities.formatDate(d, "Asia/Dhaka", "dd-MMM-yyyy hh:mm:ss a");
      }
    } catch (err) {}
  } else if (input instanceof Date && !isNaN(input.getTime())) {
    return Utilities.formatDate(input, "Asia/Dhaka", "dd-MMM-yyyy hh:mm:ss a");
  }
  return Utilities.formatDate(new Date(), "Asia/Dhaka", "dd-MMM-yyyy hh:mm:ss a");
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000); // 20s lock to prevent race conditions
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: "Server busy processing other requests. Please retry in a few seconds."
    })).setMimeType(ContentService.MimeType.JSON);
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.setSpreadsheetTimeZone("Asia/Dhaka");

    if (!e || !e.postData || !e.postData.contents) {
      return ContentService.createTextOutput(JSON.stringify({
        status: "error",
        message: "No POST payload received."
      })).setMimeType(ContentService.MimeType.JSON);
    }

    var payload = JSON.parse(e.postData.contents);
    var action = payload.action || "save_choices";
    var updatedCount = 0;

    if (action === "save_choices" || action === "update_choice" || action === "remove_choice") {
      var choices = payload.choices || [];
      if (payload.choice) {
        choices.push(payload.choice);
      }

      for (var i = 0; i < choices.length; i++) {
        var item = choices[i];
        var month = item.month || "May_2026";
        var sheetName = (month.indexOf("Jun") !== -1) ? "Award_June_2026" : "Award_May_2026";
        var sheet = ss.getSheetByName(sheetName);

        if (!sheet) continue;

        var areaCode = String(item.area_code || "").trim();
        var mioCode = String(item.mio_code || "").trim();
        var voucher = String(item.voucher || "").trim();
        var timestamp = toBangladeshTimeString(item.timestamp);

        if (!areaCode) continue;

        var lastRow = sheet.getLastRow();
        if (lastRow < 2) continue;

        // Current Master Columns (with FF List):
        // Area Code is Col 7, MIO Code is Col 9
        // May: Voucher Col 15, Timestamp Col 16, Status Col 17
        // June: Voucher Col 17, Timestamp Col 18, Status Col 19
        var isJune = (sheetName === "Award_June_2026");
        var voucherCol = isJune ? 17 : 15;
        var timeCol = isJune ? 18 : 16;
        var statusCol = isJune ? 19 : 17;

        var areaValues = sheet.getRange(2, 7, lastRow - 1, 1).getValues();
        var mioValues = sheet.getRange(2, 9, lastRow - 1, 1).getValues();

        for (var r = 0; r < areaValues.length; r++) {
          var rowArea = String(areaValues[r][0] || "").trim();
          var rowMio = String(mioValues[r][0] || "").trim();

          var match = false;
          if (rowArea === areaCode) {
            if (mioCode && rowMio) {
              match = (rowMio === mioCode);
            } else {
              match = true;
            }
          }

          if (match) {
            var rowIdx = r + 2;
            var currentVoucher = String(sheet.getRange(rowIdx, voucherCol).getValue() || "").trim();

            // Explicit removal / deselect action:
            // "Jokhon deselect kora hoy, tokhon only choice ta remove hoy/ blank hoy, kintu timestamp theke jay."
            if (action === "remove_choice" || item.action === "remove") {
              sheet.getRange(rowIdx, voucherCol).setValue("");
              // Note: timeCol is NOT modified/erased. The recorded timestamp remains strictly preserved!
              sheet.getRange(rowIdx, statusCol).setValue("Pending");
              sheet.getRange(rowIdx, voucherCol, 1, 3).setBackground(null);
              updatedCount++;
              break;
            }

            // ANTI-OVERWRITE GUARD: Never overwrite an existing voucher with a blank one!
            if (!voucher && currentVoucher) {
              break;
            }

            if (voucher) {
              sheet.getRange(rowIdx, voucherCol).setValue(voucher);
              var timeCell = sheet.getRange(rowIdx, timeCol);
              timeCell.setNumberFormat('@'); // Plain text format so Google Sheets treats it as exact literal text
              timeCell.setValue(timestamp);
              sheet.getRange(rowIdx, statusCol).setValue("Complete");

              // Highlight completed row in light amber (#FEF3C7)
              sheet.getRange(rowIdx, voucherCol, 1, 3).setBackground("#FEF3C7");
              updatedCount++;
            }
            break;
          }
        }
      }

      SpreadsheetApp.flush();
      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        updated: updatedCount,
        message: updatedCount + " choice(s) saved successfully.",
        timestamp: Utilities.formatDate(new Date(), "Asia/Dhaka", "dd-MMM-yyyy hh:mm:ss a")
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // Toggle Input Submission Lock for Regional Heads
    if (action === "set_submission_lock") {
      var lockStatus = !!payload.locked;
      PropertiesService.getScriptProperties().setProperty("SUBMISSIONS_LOCKED", lockStatus ? "true" : "false");
      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        submissions_locked: lockStatus,
        message: lockStatus ? "Voucher submissions have been locked (Read-Only Mode)." : "Voucher submissions are now active and open.",
        timestamp: Utilities.formatDate(new Date(), "Asia/Dhaka", "dd-MMM-yyyy hh:mm:ss a")
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // Delete All Voucher Choices Nationwide (Admin Reset - Fast Batch Mode)
    if (action === "delete_all_data") {
      var sheetsToReset = ["Award_May_2026", "Award_June_2026"];
      var totalReset = 0;
      for (var s = 0; s < sheetsToReset.length; s++) {
        var sh = ss.getSheetByName(sheetsToReset[s]);
        if (!sh) continue;
        var lr = sh.getLastRow();
        if (lr < 2) continue;
        var isJun = (sheetsToReset[s] === "Award_June_2026");
        var vCol = isJun ? 17 : 15;
        var numRows = lr - 1;

        var emptyBatch = [];
        for (var i = 0; i < numRows; i++) {
          emptyBatch.push(["", "", "Pending"]);
        }
        var targetRange = sh.getRange(2, vCol, numRows, 3);
        targetRange.setValues(emptyBatch);
        targetRange.setBackground(null);
        totalReset += numRows;
      }
      SpreadsheetApp.flush();
      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        message: "All national voucher choice data cleared successfully (" + totalReset + " records reset).",
        timestamp: Utilities.formatDate(new Date(), "Asia/Dhaka", "dd-MMM-yyyy hh:mm:ss a")
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // Delete Voucher Choices for a Specific Region (Fast Batch Mode)
    if (action === "delete_region_data") {
      var targetRegion = String(payload.region || payload.region_name || "").trim();
      if (!targetRegion) {
        return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "No region specified." })).setMimeType(ContentService.MimeType.JSON);
      }
      var sheetsToReset = ["Award_May_2026", "Award_June_2026"];
      var regReset = 0;
      for (var s = 0; s < sheetsToReset.length; s++) {
        var sh = ss.getSheetByName(sheetsToReset[s]);
        if (!sh) continue;
        var lr = sh.getLastRow();
        if (lr < 2) continue;
        var isJun = (sheetsToReset[s] === "Award_June_2026");
        var vCol = isJun ? 17 : 15;
        var regCol = 4; // Region Name is Column 4
        var numRows = lr - 1;

        var regValues = sh.getRange(2, regCol, numRows, 1).getValues();
        var choiceRange = sh.getRange(2, vCol, numRows, 3);
        var choiceVals = choiceRange.getValues();
        var bgVals = choiceRange.getBackgrounds();
        var changed = false;

        for (var r = 0; r < regValues.length; r++) {
          var rowReg = String(regValues[r][0] || "").trim();
          if (rowReg.toLowerCase() === targetRegion.toLowerCase()) {
            choiceVals[r][0] = "";
            choiceVals[r][1] = "";
            choiceVals[r][2] = "Pending";
            bgVals[r][0] = null;
            bgVals[r][1] = null;
            bgVals[r][2] = null;
            regReset++;
            changed = true;
          }
        }
        if (changed) {
          choiceRange.setValues(choiceVals);
          choiceRange.setBackgrounds(bgVals);
        }
      }
      SpreadsheetApp.flush();
      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        message: "Data for " + targetRegion + " cleared successfully (" + regReset + " records reset).",
        timestamp: Utilities.formatDate(new Date(), "Asia/Dhaka", "dd-MMM-yyyy hh:mm:ss a")
      })).setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: "Unrecognized action: " + action
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (ex) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: ex.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function getSheetChoices(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  var choices = {};
  if (!sheet) return choices;

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return choices;

  var isJune = (sheetName === "Award_June_2026");
  var voucherCol = isJune ? 17 : 15;
  var timeCol = isJune ? 18 : 16;
  var statusCol = isJune ? 19 : 17;

  var areaVals = sheet.getRange(2, 7, lastRow - 1, 1).getValues();
  var mioVals = sheet.getRange(2, 9, lastRow - 1, 1).getValues();
  var voucherVals = sheet.getRange(2, voucherCol, lastRow - 1, 1).getValues();
  var timeVals = sheet.getRange(2, timeCol, lastRow - 1, 1).getValues();
  var statusVals = sheet.getRange(2, statusCol, lastRow - 1, 1).getValues();

  for (var i = 0; i < areaVals.length; i++) {
    var ac = String(areaVals[i][0] || "").trim();
    var mc = String(mioVals[i][0] || "").trim();
    var v = String(voucherVals[i][0] || "").trim();
    var s = String(statusVals[i][0] || "").trim();
    var rawT = timeVals[i][0];
    var t = rawT ? toBangladeshTimeString(rawT) : "";

    if (ac && (v || t)) {
      var key = ac + "_" + mc;
      choices[key] = { voucher: v, status: s, timestamp: t };
    }
  }
  return choices;
}

function getAwardSummary(ss) {
  var summary = {
    May: { total: 0, completed: 0, pending: 0, brands: {} },
    June: { total: 0, completed: 0, pending: 0, brands: {} },
    Combined: { total: 0, completed: 0, pending: 0, brands: {} }
  };

  VOUCHER_OPTIONS.forEach(function(b) {
    summary.May.brands[b] = 0;
    summary.June.brands[b] = 0;
    summary.Combined.brands[b] = 0;
  });

  ["May", "June"].forEach(function(m) {
    var sName = (m === "June") ? "Award_June_2026" : "Award_May_2026";
    var sheet = ss.getSheetByName(sName);
    if (!sheet) return;

    var lr = sheet.getLastRow();
    if (lr < 2) return;

    var isJune = (m === "June");
    var vCol = isJune ? 17 : 15;
    var stCol = isJune ? 19 : 17;

    var vVals = sheet.getRange(2, vCol, lr - 1, 1).getValues();
    var sVals = sheet.getRange(2, stCol, lr - 1, 1).getValues();

    summary[m].total = lr - 1;
    for (var i = 0; i < vVals.length; i++) {
      var v = String(vVals[i][0] || "").trim();
      var s = String(sVals[i][0] || "").trim();
      if (s === "Complete" || v) {
        summary[m].completed++;
        if (summary[m].brands[v] !== undefined) {
          summary[m].brands[v]++;
        }
      } else {
        summary[m].pending++;
      }
    }
  });

  summary.Combined.total = summary.May.total + summary.June.total;
  summary.Combined.completed = summary.May.completed + summary.June.completed;
  summary.Combined.pending = summary.May.pending + summary.June.pending;
  VOUCHER_OPTIONS.forEach(function(b) {
    summary.Combined.brands[b] = summary.May.brands[b] + summary.June.brands[b];
  });

  return summary;
}

/**
 * ONE-CLICK REPAIR TOOL FOR GOOGLE SHEETS
 * Fixes May Ach% (divides by 100 if raw > 1, sets format 0.0%),
 * fixes June Growth% (divides by 100 if raw > 1, sets format 0.0%),
 * sets spreadsheet timezone to Asia/Dhaka, and converts existing timestamps to BDT.
 */
function fixAllSheetsFormattingAndTimestamps(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone("Asia/Dhaka");

  var countMay = 0;
  var countJun = 0;
  var timeCountMay = 0;
  var timeCountJun = 0;

  // 1. Fix Award_May_2026
  var sheetMay = ss.getSheetByName("Award_May_2026");
  if (sheetMay) {
    var lastRowMay = sheetMay.getLastRow();
    if (lastRowMay >= 2) {
      // Col 13 is Ach%
      var achRange = sheetMay.getRange(2, 13, lastRowMay - 1, 1);
      var achVals = achRange.getValues();
      var achChanged = false;
      for (var r = 0; r < achVals.length; r++) {
        var v = achVals[r][0];
        if (typeof v === "number" && v > 1) {
          achVals[r][0] = v / 100.0;
          achChanged = true;
          countMay++;
        }
      }
      if (achChanged) {
        achRange.setValues(achVals);
      }
      achRange.setNumberFormat("0.0%");

      // Col 16 is Timestamp
      var timeRangeMay = sheetMay.getRange(2, 16, lastRowMay - 1, 1);
      var timeValsMay = timeRangeMay.getValues();
      var timeChangedMay = false;
      for (var r = 0; r < timeValsMay.length; r++) {
        var tv = timeValsMay[r][0];
        if (tv) {
          var bdStr = toBangladeshTimeString(tv);
          if (bdStr !== tv) {
            timeValsMay[r][0] = bdStr;
            timeChangedMay = true;
            timeCountMay++;
          }
        }
      }
      timeRangeMay.setNumberFormat("@");
      if (timeChangedMay) {
        timeRangeMay.setValues(timeValsMay);
      }
    }
  }

  // 2. Fix Award_June_2026
  var sheetJun = ss.getSheetByName("Award_June_2026");
  if (sheetJun) {
    var lastRowJun = sheetJun.getLastRow();
    if (lastRowJun >= 2) {
      // Col 15 is Growth%
      var grRange = sheetJun.getRange(2, 15, lastRowJun - 1, 1);
      var grVals = grRange.getValues();
      var grChanged = false;
      for (var r = 0; r < grVals.length; r++) {
        var gv = grVals[r][0];
        if (typeof gv === "number" && gv > 1) {
          grVals[r][0] = gv / 100.0;
          grChanged = true;
          countJun++;
        }
      }
      if (grChanged) {
        grRange.setValues(grVals);
      }
      grRange.setNumberFormat("0.0%");

      // Col 18 is Timestamp
      var timeRangeJun = sheetJun.getRange(2, 18, lastRowJun - 1, 1);
      var timeValsJun = timeRangeJun.getValues();
      var timeChangedJun = false;
      for (var r = 0; r < timeValsJun.length; r++) {
        var tv = timeValsJun[r][0];
        if (tv) {
          var bdStr = toBangladeshTimeString(tv);
          if (bdStr !== tv) {
            timeValsJun[r][0] = bdStr;
            timeChangedJun = true;
            timeCountJun++;
          }
        }
      }
      timeRangeJun.setNumberFormat("@");
      if (timeChangedJun) {
        timeRangeJun.setValues(timeValsJun);
      }
    }
  }

  return "Fixed May Ach% (" + countMay + " rows), June Growth% (" + countJun + " rows), and timestamps to BD Time.";
}

function lockSubmissionsMenu() {
  PropertiesService.getScriptProperties().setProperty("SUBMISSIONS_LOCKED", "true");
  SpreadsheetApp.getUi().alert("🔒 Submissions Locked", "Voucher choice submission has been locked for all Regional Heads. Portal is now in Read-Only mode.", SpreadsheetApp.getUi().ButtonSet.OK);
}

function unlockSubmissionsMenu() {
  PropertiesService.getScriptProperties().setProperty("SUBMISSIONS_LOCKED", "false");
  SpreadsheetApp.getUi().alert("🔓 Submissions Unlocked", "Voucher choice submission is now OPEN and ACTIVE for all Regional Heads.", SpreadsheetApp.getUi().ButtonSet.OK);
}

function clearAllChoicesMenu() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.alert("⚠️ Confirm Data Reset", "Are you sure you want to CLEAR ALL voucher choice selections for May & June? All rows will be reset to Pending.", ui.ButtonSet.YES_NO);
  if (resp === ui.Button.YES) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = ["Award_May_2026", "Award_June_2026"];
    var resetCount = 0;
    for (var s = 0; s < sheets.length; s++) {
      var sh = ss.getSheetByName(sheets[s]);
      if (!sh) continue;
      var lr = sh.getLastRow();
      if (lr < 2) continue;
      var isJun = (sheets[s] === "Award_June_2026");
      var vCol = isJun ? 17 : 15;
      var tCol = isJun ? 18 : 16;
      var stCol = isJun ? 19 : 17;
      for (var r = 2; r <= lr; r++) {
        sh.getRange(r, vCol).setValue("");
        sh.getRange(r, tCol).setValue("");
        sh.getRange(r, stCol).setValue("Pending");
        sh.getRange(r, vCol, 1, 3).setBackground(null);
        resetCount++;
      }
    }
    ui.alert("✅ Data Cleared", resetCount + " records reset to Pending.", ui.ButtonSet.OK);
  }
}

function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu("Exium Tools")
      .addItem("🛠️ Fix May Ach% & BD Timestamps", "fixAllSheetsFormattingAndTimestamps")
      .addSeparator()
      .addItem("🔒 Lock Regional Head Inputs (Read-Only)", "lockSubmissionsMenu")
      .addItem("🔓 Unlock Regional Head Inputs (Active)", "unlockSubmissionsMenu")
      .addSeparator()
      .addItem("🗑️ Clear All Voucher Choices (Reset All)", "clearAllChoicesMenu")
      .addToUi();
  } catch (err) {}
}
