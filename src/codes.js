// Average character width (adjusted for Times New Roman, size 10)
const AVG_CHART_WIDTH = 4.5;
const BASE_HEIGHT = 25; // Base height for a single line
const ADDITIONAL_HEIGHT_PER_LINE = 20; // Additional height per line

/**
 * Creates a custom menu in the Google Sheets UI when the spreadsheet is opened.
 */
function onOpen() {
  let ui = SpreadsheetApp.getUi();
  // Create a new menu in the Google Sheets UI
  ui.createMenu('Custom Menu')
    .addItem('Show Authorization', 'showAuthorizationDialog')
    .addSeparator()
    .addItem('Adjust Row Heights', 'adjustRowHeights')
    .addToUi();
}
/**
 * This function can be used to manually trigger the authorization flow for the script.
 */
function showAuthorizationDialog() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  Browser.msgBox(
    'Authorization has been granted. You can now use the script functionalities.',
    Browser.Buttons.OK
  );
}
// Function to adjust the heights of rows in a Google Sheet based on the content
function adjustRowHeights() {
  let sheetName = Browser.inputBox(
    'Please input the sheet name',
    Browser.Buttons.OK
  );
  // Get the specific sheet from the active spreadsheet
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    Browser.msgBox('Invalid sheet name. Try again.');
    return;
  }
  let lastRowIndex = sheet.getLastRow();
  let lastColIndex = sheet.getLastColumn();

  // Prompt user to input the starting row number
  let startRowInput = Browser.inputBox(
    'Input the number of the starting row.)',
    Browser.Buttons.OK
  );
  let startRow = parseInt(startRowInput, 10);

  // Check if the input is a valid number
  if (isNaN(startRow)) {
    Browser.msgBox('Please enter a valid number.');
    return;
  }

  // Define the range of data to adjust row heights for
  let dataRange = sheet.getRange(
    startRow,
    1,
    lastRowIndex - startRow + 1,
    lastColIndex
  );
  let data = dataRange.getValues(); // Get the data as a 2D array
  let columnWidths = getColumnWidths_(sheet, lastColIndex);

  // Iterate over each row in the data
  data.forEach(function (row, rowIndex) {
    // Calculate the line count for each cell in the row
    let lineCounts = row.map(function (cell, colIndex) {
      return estimateLineCount_(cell, columnWidths[colIndex]);
    });

    // Determine the largest line count in the row
    let largestLineCount = getLargestLineCount_(lineCounts);

    // Calculate the new row height based on the largest line count
    let newRowHeight = calculateRowHeight_(largestLineCount);

    // Set the new height for the current row
    sheet.setRowHeight(startRow + rowIndex, newRowHeight);
  });
}

// Function to get the widths of all columns in the sheet
function getColumnWidths_(sheet, lastColIndex) {
  let widths = [];
  for (let i = 1; i <= lastColIndex; i++) {
    widths.push(sheet.getColumnWidth(i));
  }
  return widths;
}

// Function to estimate the line count in a cell based on its content and column width
function estimateLineCount_(cell, columnWidth) {
  // Convert the cell content to a string
  let cellText = cell.toString();

  // Estimate the width of the cell's text if it were laid out in a single line
  let estimatedLineWidth = AVG_CHART_WIDTH * cellText.length;

  // Calculate the number of lines the text would take up in the cell.
  // We use Math.ceil to always round up to the nearest whole line.
  // console.log(`estimatedLineWidth is ${estimatedLineWidth} and columnWidth is ${columnWidth}`);
  let lineCount = Math.ceil(estimatedLineWidth / columnWidth);

  // The cell text might contain explicit line breaks ('\n'). We need to count these as well.
  // cellText.match(/\n/g) finds all the line breaks in the cell text.
  // If there are no line breaks, match() would return null, so we use || [] to use an empty array in that case.
  // The number of lines is always one more than the number of line breaks.
  let explicitLineBreakCount = (cellText.match(/\n/g) || []).length + 1;

  // The final line count is the maximum of lineCount and explicitLineBreakCount.
  // This accounts for both the physical width of the text and manual line breaks.
  return Math.max(lineCount, explicitLineBreakCount);
}

// Function to find the largest line count in an array of counts
function getLargestLineCount_(counts) {
  return Math.max.apply(null, counts);
}

// Function to calculate the height of a row based on the line count
function calculateRowHeight_(lineCount) {
  return BASE_HEIGHT + (lineCount - 1) * ADDITIONAL_HEIGHT_PER_LINE;
}
