// Global variables
var sheetId = '';  // Replace with your Google Sheet ID
var docId = DocumentApp.getActiveDocument().getId();  // Get the current Google Docs ID

// Function to get the word count from the Google Doc
function getWordCountFromDoc() {
  var doc = DocumentApp.openById(docId); // Open the Google Doc by ID
  var body = doc.getBody();
  var text = body.getText();
  return text.split(/\s+/).length; // Split by whitespace and count words
}

// Function to update the word count in the Google Sheet
function updateWordCountInSheet() {
  var sheet = SpreadsheetApp.openById(sheetId);  // Open the Google Sheet by ID
  var sheetObj = sheet.getSheets()[0]; // Assuming you want to update the first sheet
  
  var wordCount = getWordCountFromDoc();
  
  // Update B2 with the current word count (updated each time the document is opened or edited)
  sheetObj.getRange('B2').setValue(wordCount);
  
  // Update A2 with the total word count if it hasn't been updated today
  var currentDate = formatDate(new Date()); // Get the current date in 'yyyy-mm-dd' format
  var lastUpdatedDate = PropertiesService.getScriptProperties().getProperty('lastUpdatedDate');
  
  if (lastUpdatedDate !== currentDate) {
    sheetObj.getRange('A2').setValue(wordCount);  // Update A2 with the word count
    PropertiesService.getScriptProperties().setProperty('lastUpdatedDate', currentDate);  // Save today's date
  }

  // Now, let's update Column D to track the daily word count
  updateDailyWordCount(sheetObj, currentDate, wordCount);
}

// Function to format date to 'yyyy-mm-dd' for comparison purposes
function formatDate(date) {
  var year = date.getFullYear();
  var month = (date.getMonth() + 1).toString().padStart(2, '0'); // Adding leading zero
  var day = date.getDate().toString().padStart(2, '0'); // Adding leading zero
  return year + '-' + month + '-' + day; // Return formatted date in 'yyyy-mm-dd' format
}

// Function to calculate and update the daily word count in Column D
function updateDailyWordCount(sheetObj, currentDate, wordCount) {
  var lastRow = sheetObj.getLastRow();
  
  // Calculate the daily word count (subtract the total word count in A2)
  var dailyWordCount = wordCount - sheetObj.getRange('A2').getValue(); // Calculate the word count difference
  
  // Log the calculated daily word count
  Logger.log('Calculated Daily Word Count: ' + dailyWordCount);
  
  // Check if the current date already exists in Column C
  var dateExists = false;
  for (var i = 1; i <= lastRow; i++) {
    var dateCell = formatDate(new Date(sheetObj.getRange('C' + i).getValue())); // Convert to Date and format it
    Logger.log('Date in Row ' + i + ': ' + dateCell); // Log the date in the cell for debugging
    
    if (dateCell === currentDate) {
      // If the date exists, update the corresponding row's Column D
      sheetObj.getRange('D' + i).setValue(dailyWordCount);
      sheetObj.getRange('D' + i).setNumberFormat('0'); // Force Column D to be a number format
      applyConditionalFormatting(sheetObj, i);  // Apply color formatting based on daily word count
      dateExists = true;
      break; // Exit the loop since we found the row for today
    }
  }
  
  // If the date doesn't exist, create a new row and record the daily word count
  if (!dateExists) {
    var newRow = lastRow + 1;
    sheetObj.getRange('C' + newRow).setValue(currentDate);  // Set current date in Column C
    sheetObj.getRange('D' + newRow).setValue(dailyWordCount); // Record daily word count in Column D
    sheetObj.getRange('D' + newRow).setNumberFormat('0'); // Force Column D to be a number format
    applyConditionalFormatting(sheetObj, newRow);  // Apply color formatting based on daily word count
  }
}

// Function to apply conditional formatting to Column D
function applyConditionalFormatting(sheetObj, row) {
  var dailyWordCount = sheetObj.getRange('D' + row).getValue();
  
  var cell = sheetObj.getRange('D' + row);
  
  if (dailyWordCount < 500) {
    cell.setBackground('red');  // Red for less than 500
  } else if (dailyWordCount >= 500 && dailyWordCount < 1000) {
    cell.setBackground('yellow');  // Yellow for 500-999
  } else if (dailyWordCount >= 1000) {
    cell.setBackground('green');  // Green for 1000 and above
  } else {
    cell.setBackground('white');  // Default (no formatting)
  }
}

// Trigger function to update word count when the document is opened
function onOpen() {
  updateWordCountInSheet();  // Update the word count when the Google Doc is opened
}

// Trigger function to update word count on a time-based schedule
function createTimeDrivenTrigger() {
  ScriptApp.newTrigger('updateWordCountInSheet')
    .timeBased()
    .everyMinutes(1)  // Set this to trigger every 1 minute
    .create();
}

// Function to delete the time-driven trigger if needed
function deleteTimeDrivenTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
}
