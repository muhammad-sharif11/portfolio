var sheetName = "contact";

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    // Aapki ID maine yahan pehle hi daal di hai
    var doc = SpreadsheetApp.openById(
      "18SHO8Y_T22waQKymslRr1y_qosNjyEF5u6Js4h5pCDY",
    );
    var sheet = doc.getSheetByName(sheetName);

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var nextRow = sheet.getLastRow() + 1;

    var newRow = headers.map(function (header) {
      if (header === "timestamp") {
        return new Date(); // Google khud time likhega
      }
      return e.parameter[header]; // Form se data lega
    });

    sheet.getRange(nextRow, 1, 1, newRow.length).setValues([newRow]);

    return ContentService.createTextOutput(
      JSON.stringify({ result: "success", row: nextRow }),
    ).setMimeType(ContentService.MimeType.JSON);
  } catch (e) {
    return ContentService.createTextOutput(
      JSON.stringify({ result: "error", error: e }),
    ).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}
