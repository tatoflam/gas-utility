function GetConfig() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  var range = sheet.getRange('A:B'); // Assuming keys in Column A and values in Column B
  var values = range.getValues();
  
  var config = {};
  for (var i = 0; i < values.length; i++) {
    config[values[i][0]] = values[i][1];
  }  
  return config;
}

function ColumnToIndex(columnLetter) {
  var index = 0;
  for (var i = 0; i < columnLetter.length; i++) {
    index = index * 26 + (columnLetter.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
  }
  return index - 1;  // zero-based indxex
}

function GetUniqueValues(values) {
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

function IsBlank(str) {
  return (!str || str.trim().length === 0);
}

function GetRawUniqueValues(values) {
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


function CopySpreadsheet(spreadsheet_url, folder_id, new_spreadsheet_name) {
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

/** 

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
*/

function CopySpreadsheetFromTemplate(spreadsheet_url, template_sheet_name, folder_id, 
new_spreadsheet_name, backup = false) {
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
    move_file(copied_spreadsheet_id, folder_id, backup = backup)
    Logger.log(new_spreadsheet_name)
    return copied_spreadsheet_id
 
  } catch (e) {
    Logger.log("Error: " + e.toString());
  }
}

function MoveFile(fileId, folderId, backup = false) {
  let folder = DriveApp.getFolderById(folderId); 
  let file = DriveApp.getFileById(fileId);
  
  // Search for files with the same name in the destination folder
  let existingFiles = folder.searchFiles('title = "' + file.getName() + '"');
  
  if (existingFiles.hasNext() && backup) {
    // If a file with the same name exists and backup is true
    let backupFolderId = ensureAndGetBackupFolderId(folderId);
    let backupFolder = DriveApp.getFolderById(backupFolderId);
    let existingFile = existingFiles.next();
    existingFile.moveTo(backupFolder);
  }
  
  // Move the new file to the destination folder
  file.moveTo(folder);
}

function EnsureAndGetBackupFolderId(parentFolderId) {
  var parentFolder = DriveApp.getFolderById(parentFolderId);
  var folders = parentFolder.searchFolders('title = "backup"');

  // If the folder named "backup" exists, return its ID
  if (folders.hasNext()) {
    return folders.next().getId();
  } 
  // If not, create a new folder named "backup" and return its ID
  else {
    var backupFolder = parentFolder.createFolder("backup");
    return backupFolder.getId();
  }
}

function CreateSSFromTemplateSpreadsheet(template_spreadsheet_url, template_sheet_name, folder_id, new_spreadsheet_name){
  // Copy new spreadsheet by using new name (invoice date in yyyymm format)
  const new_spreadsheet_id = copy_spreadsheet_from_template(template_spreadsheet_url, template_sheet_name, folder_id, new_spreadsheet_name);
  // Open new spreadsheet and get the template sheet
  const new_spreadsheet = SpreadsheetApp.openById(new_spreadsheet_id);
  Logger.log(new_spreadsheet.getActiveSheet().getName());
  return [new_spreadsheet, new_spreadsheet_name];
}


function CopyFormat(srcSheet, destSheet, startRow, lastColumn) {  
  // Define the source range (e.g., entire 1st row)
  var sourceRange = srcSheet.getRange(startRow, 1, 1, srcSheet.getLastColumn());

  // Define the destination range (e.g., rows 2 to 10)
  var destRange = destSheet.getRange(startRow, 1, destSheet.getLastRow()-startRow+1, lastColumn);
  
  // Copy format from source to destination
  sourceRange.copyTo(destRange, {contentsOnly: false, formatOnly: true});
}

function GetLastRowWithValue(sheet, range_val){
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

function GetLastColumnWithValue(sheet, range_val){
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

// matrix is (n, 0) format 2 dimensional array
function ValueExists(matrix, value) {
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


// csv_url: your CSV file's (in Shift-JIS) full URL from Google Drive
function CsvToSpreadsheet(csv_url) {
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

function ExtractFileIdFromUrl(url) {
  var match = url.match(/file\/d\/([^\/]+)\//);
  if (match) {
    return match[1];
  } else {
    throw new Error('Invalid Google Drive URL');
  }
}

function ResetHoursForDateArray(dateArray) {
  return dateArray.map(function(row) {
    return row.map(function(cell) {
      if (cell instanceof Date) {
        cell.setHours(0, 0, 0, 0);
      // When the input is 2-dimensional array
      } else if (cell instanceof Array) {
        cell[0].setHours(0, 0, 0, 0);
      }
      return cell;
    });
  });
}

function ListSpreadsheets(folderId, nameIdentifier) {
  var folder = DriveApp.getFolderById(folderId);
  var allFiles = folder.getFiles();
  var spreadsheetIds = [];
  var regexPattern = new RegExp(nameIdentifier, 'i'); // 'i' is for case-insensitive matching

  while (allFiles.hasNext()) {
    var file = allFiles.next();
    var fileName = file.getName();
    var fileType = file.getMimeType();
    
    // Check if the file is a Google Spreadsheet and its name matches the regex pattern
    if (fileType === MimeType.GOOGLE_SHEETS && regexPattern.test(fileName)) {
      spreadsheetIds.push(file.getId());
    }
  }

  return spreadsheetIds;
}

function SumupValues(values){
  var sum = 0;
  for (var i = 0; i < values.length; i++) {
    for (var j = 0; j < values[i].length; j++) {
      if (typeof values[i][j] === 'number') {
        sum += values[i][j];
      }
    }
  }
  return sum
}