function createDropdownsForRanges() {
  // Open the active spreadsheet and get the sheets
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const coverSheet = ss.getSheetByName("Cover");
  const listSheet = ss.getSheetByName("Lists");

  // Get the data from the "Lists" sheet
  const data = listSheet.getDataRange().getValues();
  const headers = data[0]; // The first row contains headers
  const suppliers = [];
  const months = [];
  const years = [];

  // Extract unique values for each category, ignoring headers
  data.slice(1).forEach(row => {
    if (row[0] && !suppliers.includes(row[0])) suppliers.push(row[0]);
    if (row[1] && !months.includes(row[1])) months.push(row[1]);
    if (row[2] && !years.includes(row[2])) years.push(row[2]);
  });

  // Get user input for the starting range
  const ui = SpreadsheetApp.getUi();
  const rangeResponse = ui.prompt(
    "Set Dropdown Range",
    "Enter the range for the Supplier dropdowns (e.g., B1:G1):",
    ui.ButtonSet.OK_CANCEL
  );
  if (rangeResponse.getSelectedButton() !== ui.Button.OK) return;
  const rangeInput = rangeResponse.getResponseText();

  // Parse the range input
  const rangeMatch = rangeInput.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
  if (!rangeMatch) {
    ui.alert("Invalid range format. Please use a format like B1:G1.");
    return;
  }
  const [_, startCol, startRow, endCol, endRow] = rangeMatch;
  const startRowNum = parseInt(startRow, 10);

  // Place headers dynamically in a column to the left of the dropdowns
  const headerColIndex = getColumnIndex(startCol) - 1; // Column left of the dropdowns
  if (headerColIndex < 1) {
    throw new Error("Cannot place headers. No column to the left of the dropdowns.");
  }
  coverSheet.getRange(startRowNum, headerColIndex).setValue(headers[0]); // Supplier header
  coverSheet.getRange(startRowNum + 1, headerColIndex).setValue(headers[1]); // Month header
  coverSheet.getRange(startRowNum + 2, headerColIndex).setValue(headers[2]); // Year header


  // Calculate ranges for Supplier, Month, and Year dropdowns
  const supplierRange = coverSheet.getRange(`${startCol}${startRowNum}:${endCol}${startRowNum}`);
  const monthRange = coverSheet.getRange(`${startCol}${startRowNum + 1}:${endCol}${startRowNum + 1}`);
  const yearRange = coverSheet.getRange(`${startCol}${startRowNum + 2}:${endCol}${startRowNum + 2}`);

  // Set Supplier dropdowns
  const supplierRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(suppliers, true)
    .build();
  supplierRange.setDataValidation(supplierRule);

  // Set Month dropdowns
  const monthRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(months, true)
    .build();
  monthRange.setDataValidation(monthRule);

  // Set Year dropdowns
  const yearRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(years, true)
    .build();
  yearRange.setDataValidation(yearRule);

  // copyFilteredHeadersFromAllSheets function call
  copyFilteredHeadersFromAllSheets();

  // Confirmation message
  // ui.alert("Headers and dropdowns have been created successfully!");
}


// Helper function to convert column letters to numbers
function getColumnIndex(letter) {
  let column = 0;
  for (let i = 0; i < letter.length; i++) {
    column = column * 26 + (letter.charCodeAt(i) - 64);
  }
  return column;
}

function copyFilteredHeadersFromAllSheets() {
  // Open the active spreadsheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const coverSheet = ss.getSheetByName("Cover");

  // List of monthly sheet names
  const monthlySheets = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

  // Headers to exclude
  const excludedHeaders = [
    "",
    "Link to dataset", 
    "Channel", 
    "Cost per Deal", 
    "Cost per Sale", 
    "Cost per lead"
  ];

  // Collect headers from all sheets, ensuring uniqueness while preserving order
  const allHeaders = [];
  monthlySheets.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      // Get the headers (first row)
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      headers.forEach(header => {
        // Add header to the array if it's not excluded and not already present
        if (header && !excludedHeaders.includes(header) && !allHeaders.includes(header)) {
          allHeaders.push(header);
        }
      });
    }
  });

  // Check if allHeaders is empty before proceeding
  if (allHeaders.length === 0) {
    SpreadsheetApp.getUi().alert("No valid headers found to copy.");
    return;
  }

  // Clear existing data in the target range in the Cover sheet (starting from A5)
  coverSheet.getRange("A5:Z").clearContent();

  // Write the headers into the Cover sheet starting at A5
  const targetRange = coverSheet.getRange(5, 1, allHeaders.length, 1); // A5 and downward
  targetRange.setValues(allHeaders.map(header => [header]));

  // Confirmation
  // SpreadsheetApp.getUi().alert("Filtered headers from all monthly sheets have been copied to the Cover sheet!");
}


function onEdit(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const coverSheet = ss.getSheetByName("Cover");

  // Get the first few rows to detect dropdowns
  const headerRange = coverSheet.getRange("1:3"); // Check the first three rows for dropdowns
  const headerValues = headerRange.getValues(); // Get the data as a 2D array

  // Find columns with dropdowns (non-empty cells in rows 1, 2, and 3)
  const dropdownColumns = [];
  for (let col = 1; col <= headerRange.getNumColumns(); col++) {
    if (headerValues[0][col - 1] && headerValues[1][col - 1] && headerValues[2][col - 1]) {
      dropdownColumns.push(col); // Store the column index if all three rows have values
    }
  }

  // Iterate through each detected dropdown column
  dropdownColumns.forEach(columnIndex => {
    const supplierCell = coverSheet.getRange(1, columnIndex); // Row 1
    const monthCell = coverSheet.getRange(2, columnIndex); // Row 2
    const yearCell = coverSheet.getRange(3, columnIndex); // Row 3

    // Trigger only if the edited cell matches the current group
    if (e.range.getA1Notation() === supplierCell.getA1Notation() ||
        e.range.getA1Notation() === monthCell.getA1Notation() ||
        e.range.getA1Notation() === yearCell.getA1Notation()) {

      // Get the selected values
      const selectedSupplier = supplierCell.getValue();
      const selectedMonth = monthCell.getValue();
      const selectedYear = yearCell.getValue();

      // If any field is empty, skip
      if (!selectedSupplier || !selectedMonth || !selectedYear) {
        return;
      }

      // Find the corresponding sheet for the selected month
      const monthSheet = ss.getSheetByName(selectedMonth);
      if (!monthSheet) {
        SpreadsheetApp.getUi().alert(`Month sheet for '${selectedMonth}' not found.`);
        return;
      }

      // Get the data from the month sheet
      const data = monthSheet.getDataRange().getValues();

      // Filter rows based on the supplier
      const filteredRows = data.filter(row => row[1] === selectedSupplier); // Assuming column B (index 1) contains supplier data

      // Get the filtered headers from the Cover sheet
      const filteredHeaders = coverSheet.getRange(5, 1, coverSheet.getLastRow() - 4, 1).getValues().map(row => row[0]);

      // Clear previous data for this column group (starting from row 5)
      const startColumn = columnIndex + 1; // Data starts in the column after the dropdown group
      // Clear rows 5 to 100 for a single column
      coverSheet.getRange(5, startColumn - 1, 96, 1).clearContent();

      // Rearrange the filtered rows based on the header order
      const rearrangedData = filteredRows.map(row => {
        return filteredHeaders.map(header => {
          const headerIndex = monthSheet.getRange(1, 1, 1, monthSheet.getLastColumn()).getValues()[0].indexOf(header);
          return headerIndex !== -1 ? row[headerIndex] : "";
        });
      });

      // Write the rearranged data into the Cover sheet for this group
      if (rearrangedData.length > 0) {
        // Transpose the data to fit into a column-wise format
        const transposedData = rearrangedData[0].map((_, colIndex) =>
          rearrangedData.map(row => row[colIndex])
        );

        // Write the transposed data starting at row 5 and the specified column
        coverSheet.getRange(5, startColumn - 1, transposedData.length, transposedData[0].length).setValues(transposedData);
      } else {
        SpreadsheetApp.getUi().alert(`No matching data found for the selected criteria in column ${String.fromCharCode(64 + columnIndex)}.`);
      }
    }
  });
}
