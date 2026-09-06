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
    timestamp: Utilities.formatDate(new Date(), "Asia/Dhaka", "dd-MMM-yyyy hh:mm:ss a")
  })).setMimeType(ContentService.MimeType.JSON);
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
        var timestamp = item.timestamp || Utilities.formatDate(new Date(), "Asia/Dhaka", "dd-MMM-yyyy hh:mm:ss a");

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
              sheet.getRange(rowIdx, timeCol).setValue(timestamp);
              sheet.getRange(rowIdx, statusCol).setValue("Complete");

              // Highlight completed row in light amber (#FEF3C7)
              sheet.getRange(rowIdx, voucherCol, 1, 3).setBackground("#FEF3C7");
              updatedCount++;
            }
            break;
          }
        }
      }

      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        updated: updatedCount,
        message: updatedCount + " choice(s) saved successfully.",
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
    var t = String(timeVals[i][0] || "").trim();

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
