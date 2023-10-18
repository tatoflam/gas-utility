function getConfig() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  var range = sheet.getRange('A:B'); // Assuming keys in Column A and values in Column B
  var values = range.getValues();
  
  var config = {};
  for (var i = 0; i < values.length; i++) {
    config[values[i][0]] = values[i][1];
  }  
  return config;
}

function columnToIndex(columnLetter) {
  var index = 0;
  for (var i = 0; i < columnLetter.length; i++) {
    index = index * 26 + (columnLetter.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
  }
  return index - 1;  // zero-based index
}

function getUniqueValues(values) {
  Logger.log("values length: "+ values.length);
  var is_2d_array = false;
  if (Array.isArray(values[0])){
    is_2d_array = true;
  } 
  var uniqueValues = [];
  var seen = new Set();  // To keep track of seen values
  
  for (var i = 0; i < values.length; i++) {
    var value;
    if (is_2d_array){
      value = values[i][0];  // Since we have a single column, we use [0] to get the value
    } else {
      value = values[i]
    }
    if (!seen.has(value) && !isBlank(value)) {  // If value has not been seen yet
      seen.add(value);       // Mark the value as seen
      uniqueValues.push([value]);  // Add it to the list of unique values
    }
  }  
  Logger.log("unique values length: "+uniqueValues.length)
  return uniqueValues;
}

function isBlank(str) {
  return (!str || str.trim().length === 0);
}

function getRawUniqueValues(values) {
  Logger.log("values length: "+ values.length);
  var is_2d_array = false;
  if (Array.isArray(values[0])){
    is_2d_array = true;
  } 
  var uniqueValues = [];
  var seen = new Set();  // To keep track of seen values
  
  for (var i = 0; i < values.length; i++) {
    var value;
    if (is_2d_array){
      value = values[i][0];  // Since we have a single column, we use [0] to get the value
    } else {
      value = values[i]
    }
    if (!seen.has(value) && !isBlank(value)) {  // If value has not been seen yet
      seen.add(value);       // Mark the value as seen
      uniqueValues.push(value);  // Add it to the list of unique values
    }
  }  
  Logger.log("unique values length: "+uniqueValues.length)
  return uniqueValues;
}


function copy_spreadsheet(spreadsheet_url, folder_id, new_spreadsheet_name) {
  try {
    Logger.log(spreadsheet_url)

    // Get the existing spreadsheet by name
    var original_spreadsheet = SpreadsheetApp.openByUrl(spreadsheet_url);
    Logger.log(original_spreadsheet.getActiveSheet().getName())

    // Create a copy of the spreadsheet
    var copied_spreadsheet_id = original_spreadsheet.copy(new_spreadsheet_name).getId();
    move_file(copied_spreadsheet_id, folder_id)
    Logger.log(new_spreadsheet_name)
    return copied_spreadsheet_id
 
  } catch (e) {
    Logger.log("Error: " + e.toString());
  }
}

function move_file(fileId, folderId) {
  let folder = DriveApp.getFolderById(folderId); 
  let file = DriveApp.getFileById(fileId);
  file.moveTo(folder);
}

function copy_spreadsheet_from_template(spreadsheet_url, template_sheet_name, folder_id, new_spreadsheet_name) {
  try {
    Logger.log(spreadsheet_url)

    // Get the existing spreadsheet by name
    var original_spreadsheet = SpreadsheetApp.openByUrl(spreadsheet_url);
    var source_sheet = original_spreadsheet.getSheetByName(template_sheet_name);
    Logger.log("Creating a sparedsheet from a sheet: " + original_spreadsheet.getActiveSheet().getName())

    // Create a copy of the spreadsheet
    // var copied_spreadsheet_id = original_spreadsheet.copy(new_spreadsheet_name).getId();
    var copied_sheet = source_sheet.copyTo(SpreadsheetApp.create(new_spreadsheet_name));

    // Remove empty sheet created by Spreadsheet
    var copied_ss = copied_sheet.getParent();
    var empty_sheet = copied_ss.getActiveSheet();
    copied_ss.deleteSheet(empty_sheet);

    // Set sheet name as the same name of spreadsheet
    copied_sheet.setName(new_spreadsheet_name);

    var copied_spreadsheet_id = copied_sheet.getParent().getId();
    move_file(copied_spreadsheet_id, folder_id)
    Logger.log(new_spreadsheet_name)
    return copied_spreadsheet_id
 
  } catch (e) {
    Logger.log("Error: " + e.toString());
  }
}

function move_file(fileId, folderId) {
  let folder = DriveApp.getFolderById(folderId); 
  let file = DriveApp.getFileById(fileId);
  file.moveTo(folder);
}


function CreateSSFromTemplateSpreadsheet(template_spreadsheet_url, template_sheet_name, folder_id, new_spreadsheet_name){
  // Copy new spreadsheet by using new name (invoice date in yyyymm format)
  const new_spreadsheet_id = copy_spreadsheet_from_template(template_spreadsheet_url, template_sheet_name, folder_id, new_spreadsheet_name);
  // Open new spreadsheet and get the template sheet
  const new_spreadsheet = SpreadsheetApp.openById(new_spreadsheet_id);
  Logger.log(new_spreadsheet.getActiveSheet().getName());
  return [new_spreadsheet, new_spreadsheet_name];
}

function formatCurrentDateTime(date) {
  
  // Timezone is important. You can use the script's timezone or specify one.
  var timeZone = Session.getScriptTimeZone();
  
  // Formatting the date
  var formattedDate = Utilities.formatDate(date, timeZone, "yyyy/MM/dd HH:mm:ss");
  
  Logger.log("Formatted date and time: " + formattedDate);
  return formattedDate; 
}

function format_date_from_string(dateString) {
  // Logger.log(dateString);
  // Parse the string into a date object
  var date = new Date(dateString);
  
  // Extract year, month, and day
  var year = date.getFullYear();
  var month = date.getMonth() + 1; // Months are zero-based
  var day = date.getDate();
  
  // Format in "yyyy/mm/dd" format
  var formattedDate = year + "/" + (month < 10 ? "0" : "") + month + "/" + (day < 10 ? "0" : "") + day;
  
  return formattedDate;
}

function string_to_date(dateString) {
  var parts = dateString.split("/");
  var year = parseInt(parts[0], 10);
  var month = parseInt(parts[1], 10) - 1; // JavaScript months are 0-based, so we subtract 1
  var day = parseInt(parts[2], 10);

  return new Date(year, month, day);
}


function FormatTimeHMSm(time){
  // Convert time difference to hours, minutes, seconds, and milliseconds
  var hours = Math.floor(time / (1000 * 60 * 60));
  time -= hours * (1000 * 60 * 60);

  var mins = Math.floor(time / (1000 * 60));
  time -= mins * (1000 * 60);

  var secs = Math.floor(time / 1000);
  time -= secs * 1000;

  var millis = time;
  return formattedTime = Utilities.formatString("%02d:%02d:%02d.%03d", hours, mins, secs, millis);
  
}

function getEndOfNextMonth(dateString) {
  // Parse the date string
  var dateParts = dateString.split("/");
  var year = parseInt(dateParts[0], 10);
  var month = parseInt(dateParts[1], 10) - 1; // Months are zero-based in JavaScript Date
  var day = parseInt(dateParts[2], 10);
  
  var date = new Date(year, month, day);
  
  // Increase month by 2 to get to the next-next month
  date.setMonth(date.getMonth() + 2);
  
  // Set to the first day of the next-next month
  date.setDate(1);
  
  // Decrease by one day to get the last day of the next month
  date.setDate(date.getDate() - 1);
  
  // Format the date back to "yyyy/mm/dd"
  year = date.getFullYear();
  month = date.getMonth() + 1; // Convert back to 1-based month
  day = date.getDate();
  var formattedDate = year + "/" + (month < 10 ? "0" : "") + month + "/" + (day < 10 ? "0" : "") + day;
  
  return formattedDate;
}

function copyFormat(srcSheet, destSheet, startRow, lastColumn) {  
  // Define the source range (e.g., entire 1st row)
  var sourceRange = srcSheet.getRange(startRow, 1, 1, srcSheet.getLastColumn());

  // Define the destination range (e.g., rows 2 to 10)
  var destRange = destSheet.getRange(startRow, 1, destSheet.getLastRow()-startRow+1, lastColumn);
  
  // Copy format from source to destination
  sourceRange.copyTo(destRange, {contentsOnly: false, formatOnly: true});
}

function getLastRowWithValue(sheet, range_val){
  var row = sheet.getRange(range_val).getRow();
  var col = sheet.getRange(range_val).getColumn();
  var last_row = sheet.getLastRow();
  var row_size = last_row - row + 1;

  for (var r = 0; r <= row_size; r++){
    if (sheet.getRange(row+r, col).getValue()===""){
      row = row + r -1;
      break;
    }
  }
  return row
}

function getLastColumnWithValue(sheet, range_val){
  var row = sheet.getRange(range_val).getRow();
  var col = sheet.getRange(range_val).getColumn();
  var last_col = sheet.getLastColumn()
  var col_size = last_col - col + 1;

  for (var c = 0; c <= col_size; c++){
    if (sheet.getRange(row, col+c).getValue()===""){
      col = col + c -1;
      break
    }
  }
  return col
}

/**
 * Custom SUMIFS function for Google Apps Script.
 * 
 * @param {Array} sumRange - The range of numbers to be summed.
 * @param {Array} criteriaRanges - An array of ranges where criteria are applied.
 * @param {Array} criteria - An array of criteria corresponding to criteriaRanges.
 * @return {number} - The summed value based on the criteria.
 * @customfunction
 */
function CUSTOM_SUMIFS(sumRange, criteriaRanges, criteria) {
  if (criteriaRanges.length !== criteria.length) {
    throw new Error("Number of criteria ranges and criteria should be the same.");
  }

  var sum = 0;

  // Check if the criteria is included in the criteriaranges
  // This did not improve performance. 
  /**
  var criteria_match = true;
  for (var j = 0; j < criteria.length; j++) {
    var criteria_values = criteriaRanges[j]; // 2-d array
    var criterion = criteria[j];
    if (!valueExists(criteria_values, criterion)){
      return sum;
      Logger.log("value does not exist.")
    }
  }
  */

  // Helper function to check if two dates are the same
  function areDatesEqual(date1, date2) {
    return date1.getDate() === date2.getDate() &&
           date1.getMonth() === date2.getMonth() &&
           date1.getFullYear() === date2.getFullYear();
  }

  // Loop through each row in the sumRange
  for (var i = 0; i < sumRange.length; i++) {
    var matchesAllCriteria = true;

    // Check each criterion
    for (var j = 0; j < criteria.length; j++) {
      var cellValue = criteriaRanges[j][i][0];
      var criterion = criteria[j];


      // Check if both the cellValue and criterion are Date objects
      if (cellValue instanceof Date && criterion instanceof Date) {
        if (!areDatesEqual(cellValue, criterion)) {
          matchesAllCriteria = false;
          break;
        }
      } else if (cellValue !== criterion) {
        matchesAllCriteria = false;
        break;
      }
    }

    if (matchesAllCriteria) {
      sum += sumRange[i][0];
    }
  }

  return sum;
}

function CopySheetToLast(ss, org_sheet_name, new_sheet_name) {
  // Step 1: Access the active spreadsheet
  // var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Step 2: Get the sheet you want to duplicate
  var sheetToDuplicate = ss.getSheetByName(org_sheet_name); // Assuming you want to duplicate a sheet named 'Sheet1'
  
  // Step 3: Duplicate the sheet
  var duplicatedSheet = sheetToDuplicate.copyTo(ss);
  
  // Step 4: Rename the duplicated sheet
  duplicatedSheet.setName(new_sheet_name); // Replace 'NewSheetName' with the desired name for the duplicated sheet
  
  // Step 5: Move the duplicated sheet to the last position
  var totalSheets = ss.getSheets().length;
  ss.setActiveSheet(duplicatedSheet);
  ss.moveActiveSheet(totalSheets);
}

function ClearRowByRange(sheet, range_address) {
  // Step 1: Identify the cell range
  // var ss = SpreadsheetApp.getActiveSpreadsheet();
  // var sheet = ss.getActiveSheet();
  var range = sheet.getRange(range_address);  // Assume 'B5' is the cell range you specify
  
  // Step 2: Get the row number of the range
  var rowNumber = range.getRow();
  
  // Step 3: Clear the entire row based on the row number
  sheet.getRange(rowNumber, 1, 1, sheet.getMaxColumns()).clearContent();
}

function GetLastDayOfPreviousMonth(date) {
  // If specificDate is a string, convert it to a Date object
  if (typeof date === 'string') {
    date = new Date(date);
  }

  // Step 2: Set the date to the first day of the current month
  date.setDate(1);

  // Step 3: Subtract one day from the first day of the current month
  date.setDate(date.getDate() - 1);
  
  // Step 4: Format the date string if needed (for example, in "yyyy/MM/dd" format)
  var formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy/MM/dd');
  
  return formattedDate;
}

function datesAreEqual(date1, date2) {
  return date1.getTime() === date2.getTime();
}

// matrix is (n, 0) format 2 dimensional array
function valueExists(matrix, value) {
  for (var i = 0; i < matrix.length; i++) {
    var element = matrix[i][0];

    if (element instanceof Date && value instanceof Date) {
      if (datesAreEqual(element, value)) {
        return true;
      }
    } else if (element == value) {
      return true;
    }
  }
  return false;
}

function isWithinFinancialTerm(dateString, year_term) {
  // Convert the given string to a Date object
  var parts = dateString.split("/");
  var year = parseInt(parts[0], 10);
  var month = parseInt(parts[1], 10) - 1; // JavaScript months are 0-based, so we subtract 1
  var day = parseInt(parts[2], 10);
  var givenDate = new Date(year, month, day);

  // Define the start and end dates
  var startDate = new Date(year_term, 8, 1); // 1st September 2022
  var endDate = new Date(year_term+1, 7, 31);  // 31st August 2023

  // Check if the given date is within the range
  return givenDate >= startDate && givenDate <= endDate;
}

// csv_url: your CSV file's (in Shift-JIS) full URL from Google Drive
function csvToSpreadsheet(csv_url) {
  var fileId = extractFileIdFromUrl(csv_url);
  
  var file = DriveApp.getFileById(fileId);
  var csvBlob = file.getBlob();

  // Try decoding with Shift-JIS
  var csvData = Utilities.newBlob(csvBlob.getBytes()).getDataAsString('Shift_JIS');

  var newSpreadsheet = SpreadsheetApp.create(file.getName().replace('.csv', ''));
  var sheet = newSpreadsheet.getActiveSheet();
  var csvArray = Utilities.parseCsv(csvData);

  sheet.getRange(1, 1, csvArray.length, csvArray[0].length).setValues(csvArray);
  return newSpreadsheet;
  /** 
  var newFile = DriveApp.getFileById(newSpreadsheet.getId());  
  newFile.setContent(csvBlob.getDataAsString());

  // Ensure the file is a Google Sheets file before proceeding
  if (newFile.getMimeType() === MimeType.GOOGLE_SHEETS) {
    var spreadsheet = SpreadsheetApp.open(newFile);
    return spreadsheet;
  } else {
    throw new Error("The file is not a Google Sheets document.");
  }
  */
}

function extractFileIdFromUrl(url) {
  var match = url.match(/file\/d\/([^\/]+)\//);
  if (match) {
    return match[1];
  } else {
    throw new Error('Invalid Google Drive URL');
  }
}
