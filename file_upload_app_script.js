
var sheetName = 'Sheet1'
var uploadFolderID = '1VgKyRjMYBrAgwT5WTdL41JCpVknpc0lZ'  // Replace with your Drive folder ID
var scriptProp = PropertiesService.getScriptProperties()

function intialSetup () {
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  scriptProp.setProperty('key', activeSpreadsheet.getId())
}


function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    var doc = SpreadsheetApp.openById(scriptProp.getProperty("key"));
    var sheet = doc.getSheetByName(sheetName);

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var nextRow = sheet.getLastRow() + 1;

    var newRow = headers.map(function(header) {
      return header === "timestamp" ? new Date() : e.parameter[header];
    });

    // Handle media file upload
    if (e.parameter.media) {
      const mediaBlob = Utilities.newBlob(
        Utilities.base64Decode(e.parameter.media),
        MimeType.JPEG, // Change this MIME type if needed
        "uploaded_media.jpg"
      );

      // Save file to Google Drive
      const folder = DriveApp.getFolderById(uploadFolderID); // Replace with your Drive folder ID
      const file = folder.createFile(mediaBlob);
      const fileUrl = file.getUrl();

      // Find the "Media" column index in headers
      const mediaColumnIndex = headers.indexOf("media");
      if (mediaColumnIndex > -1) {
        newRow[mediaColumnIndex] = fileUrl;
      } else {
        // Add to the end if "media" column doesn't exist
        newRow.push(fileUrl);
      }
    } else {
      // Add an empty value if no media uploaded
      const mediaColumnIndex = headers.indexOf("media");
      if (mediaColumnIndex > -1) {
        newRow[mediaColumnIndex] = "";
      } else {
        newRow.push("");
      }
    }

    // Write the new row data to the sheet
    sheet.getRange(nextRow, 1, 1, newRow.length).setValues([newRow]);

    return ContentService
      .createTextOutput(JSON.stringify({ result: "success", row: nextRow }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: "error", error: error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}
