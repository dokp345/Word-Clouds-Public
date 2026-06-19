/* =========================================
   Code.gs - Paste in Google Apps Script Editor
   ========================================= */

function doGet() {
  // Pulls the ID securely from your hidden Project Settings
  var sheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  var ss = SpreadsheetApp.openById(sheetId); 
  
  return HtmlService.createHtmlOutputFromFile('index');
}

// Automatically detect the correct script URL to build QR Codes
function getWebAppUrl() {
  try {
    return ScriptApp.getService().getUrl();
  } catch(e) {
    return "";
  }
}

// Appends submissions to Google Sheet: GameData
function submitParticipantWord(pin, question, limit, word, participantId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("GameData");
    if (!sheet) {
      sheet = ss.insertSheet("GameData");
      sheet.appendRow(["RoomPIN", "HostQuestion", "TimeLimit", "Timestamp", "Word", "ParticipantID"]);
    }
    
    var cleanWord = word.toString().trim();
    if (cleanWord.length > 0 && cleanWord.length <= 20) {
      sheet.appendRow([pin, question, limit, new Date(), cleanWord, participantId]);
      return { status: "success" };
    }
    return { status: "error", message: "Word length must be under 20 characters." };
  } catch(err) {
    return { status: "error", message: err.toString() };
  }
}

// Pulls live statistics to compute active cloud counts
function getHostLiveData(pin) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("GameData");
    if (!sheet) return { words: [], participantCount: 0, question: "", limit: 60 };
    
    var rows = sheet.getDataRange().getValues();
    var wordCounts = {};
    var uniqueParticipants = new Set();
    var question = "";
    var limit = 60;
    
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0].toString() === pin.toString()) {
        if (!question && rows[i][1]) question = rows[i][1];
        if (rows[i][2]) limit = Number(rows[i][2]);
        var word = rows[i][4];
        var pId = rows[i][5];
        if (word) {
          var clean = word.toLowerCase().trim();
          wordCounts[clean] = (wordCounts[clean] || 0) + 1;
        }
        if (pId) uniqueParticipants.add(pId);
      }
    }
    
    var wordList = [];
    for (var key in wordCounts) {
      wordList.push([key, wordCounts[key]]);
    }
    return { words: wordList, participantCount: uniqueParticipants.size, question: question, limit: limit };
  } catch(err) {
    return { words: [], participantCount: 0, question: "", limit: 60 };
  }
}

// Archives the ended session details along with a visual screenshot directly into "GameArchive" sheet
function saveSessionSummary(pin, question, limit, dateLaunched, base64Image) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("GameArchive");
    if (!sheet) {
      sheet = ss.insertSheet("GameArchive");
      sheet.appendRow(["RoomPIN", "HostQuestion", "TimeLimit", "DateLaunched", "TimestampEnded", "Screenshot"]);
      // Style headers
      sheet.getRange(1, 1, 1, 6).setFontWeight("bold").setBackground("#cbd5e1");
    }
    
    var cleanBase64 = base64Image.replace(/^data:image\/(png|jpeg|jpg);base64,/, "");
    var blob = Utilities.newBlob(Utilities.base64Decode(cleanBase64), 'image/jpeg', 'screenshot_pin_' + pin + '.jpg');
    
    var lastRow = sheet.getLastRow();
    var newRow = lastRow + 1;
    
    // Append descriptive row metadata
    sheet.appendRow([pin, question, limit, new Date(dateLaunched), new Date(), "Image Inserted Floating Beside"]);
    
    // Set a matching tall height for the row to accommodate the floating image beautifully
    sheet.setRowHeight(newRow, 110);
    
    // Insert the screenshot image floating inline beside metadata row in Column F (6)
    sheet.insertImage(blob, 6, newRow);
    
    return { status: "success" };
  } catch(err) {
    return { status: "error", message: err.toString() };
  }
}
            
