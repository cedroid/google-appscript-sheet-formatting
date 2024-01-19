function processSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var columnFormats = getColumnFormats();
 
  // Create or get the 'Formatting Logs' sheet
  var logSheet = ss.getSheetByName("Formatting Logs") || ss.insertSheet("Formatting Logs");
  // Set up log sheet headers
  setupLogSheet(logSheet);
 
  var processedCount = 0;
  var errorCount = 0;
  var successCount = 0;
 
  // Process each sheet
  sheets.forEach(function(sheet) {
    if (sheet.getName() !== "Formatting Logs" && !isSheetProcessed(logSheet, sheet.getName())) {
      var status = processSingleSheet(sheet, columnFormats);
      logStatus(logSheet, sheet, status);
 
      if (status.error) {
        errorCount++;
      } else {
        successCount++;
      }
 
      processedCount++;
      if (processedCount >= 30) {
        // Stop processing after 30 sheets
        return;
      }
    }
  });
 
  // Show notification after processing
  SpreadsheetApp.getUi().alert(
    'Processing completed. ' +
    'Success: ' + successCount + ', ' +
    'Errors: ' + errorCount
  );
}
 
function processSingleSheet(sheet, columnFormats) {
  var status = { error: false, notes: [] };
  var lastColumn = sheet.getLastColumn();
  var lastRow = sheet.getLastRow();
 
  // Check if the sheet has columns and rows
  if (lastColumn === 0 || lastRow === 0) {
    status.error = true;
    status.notes.push("No columns or rows found in sheet.");
    return status;
  }
 
  // Get headers and sort them along with their data
  var headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  var data = sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();
 
  // Create a combined array of headers and data for sorting
  var combined = headers.map(function(header, index) {
    return [header].concat(data.map(function(row) {
      return row[index];
    }));
  });
 
  // Sort the combined array alphabetically based on headers
  combined.sort(function(a, b) {
    return a[0].localeCompare(b[0]);
  });
 
  // Write the sorted data back to the sheet
  for (var col = 0; col < combined.length; col++) {
    sheet.getRange(1, col + 1, lastRow).setValues(combined[col].map(function(cell, index) {
      return [cell];
    }));
  }
 
  // Apply formats to the sorted columns and check for errors
  combined.forEach(function(column, index) {
    var header = column[0];
    if (columnFormats[header]) {
      try {
        var range = sheet.getRange(2, index + 1, lastRow - 1);
        applyFormat(range, columnFormats[header]);
      } catch (e) {
        status.error = true;
        status.notes.push("Error in column: " + header + " - " + e.message);
      }
    }
  });
 
  // Return processing status
  return status;
}
 
function applyFormat(range, formatInfo) {
  // Apply format based on type and specific format rules
  switch (formatInfo.type) {
    case "date":
      range.setNumberFormat(formatInfo.format);
      break;
    case "text":
      // Text formatting can be added here
      break;
    case "number":
      range.setNumberFormat(formatInfo.format);
      break;
    case "currency":
      range.setNumberFormat(formatInfo.format);
      break;
    // Add more cases as needed
  }
}
 
function logStatus(logSheet, sheet, status) {
  var timeStamp = new Date();
  var sheetUrl = sheet.getParent().getUrl() + "#gid=" + sheet.getSheetId();
  var statusText = status.error ? "Error" : "Success";
  var notesText = status.notes.join("\n");
 
  // Append log entry
  logSheet.appendRow([sheet.getName(), sheetUrl, timeStamp, statusText, notesText]);
}
 
function setupLogSheet(logSheet) {
  // Setup headers for the log sheet
  logSheet.getRange("A1:E1").setValues([["Sheet Name", "Hyperlink", "Timestamp", "Status", "Status Notes"]]);
}
 
function isSheetProcessed(logSheet, sheetName) {
  // Check if the sheet is already processed
  var data = logSheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === sheetName && data[i][3] === "Success") {
      return true;
    }
  }
  return false;
}
 
function getColumnFormats() {
  // Hardcoded column names and their formats
  return {
    "Date": { type: "date", format: "MM/dd/yyyy" },
    "Name": { type: "text", format: "none" },
    "Room": { type: "number", format: "0" },
    "Price": { type: "currency", format: "$#,##0.00" },
    "Check-in": { type: "date", format: "MM/dd/yyyy" },
    "Check-out": { type: "date", format: "MM/dd/yyyy" },
    // Add more column formats as needed
  };
}
