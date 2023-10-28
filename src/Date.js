function FormatCurrentDateTime(date) {  
  var timeZone = Session.getScriptTimeZone();
  
  // Formatting the date
  var formattedDate = Utilities.formatDate(date, timeZone, "yyyy/MM/dd HH:mm:ss");
  
  Logger.log("Formatted date and time: " + formattedDate);
  return formattedDate; 
}

function FormatDateFromString(dateString) {
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

function StringToDate(dateString) {
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

function GetEndOfNextMonth(dateString) {
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

function DatesAreEqual(date1, date2) {
  return date1.getTime() === date2.getTime();
}

// datesArray: 2-dimensional array retrieved from 1 column range
function GetLatestDate(datesArray) {
  return datesArray.reduce(function(maxDate, currentRow) {
    var currentDate = currentRow[0];
    return currentDate > maxDate ? currentDate : maxDate;
  }, new Date(0));  // Starting with a default very old date for comparison
}

function IsWithinFinancialTerm(dateString, year_term) {
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