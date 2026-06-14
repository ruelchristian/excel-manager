function main(workbook: ExcelScript.Workbook) {
  let sheet = workbook.getWorksheet("ADB MASTER LIST MONTHLY");
  if (!sheet) {
    console.log('Sheet "ADB MASTER LIST MONTHLY" not found.');
    return;
  }
  
  // Find the last used row in the worksheet
  let rangeUsed = sheet.getUsedRange();
  if (!rangeUsed) {
    console.log("No used cells found.");
    return;
  }
  
  let lastRowRange = rangeUsed.getLastRow();
  let lastRowIndex = lastRowRange.getRowIndex() + 1; // 1-based row index
  
  // If no data rows (data starts at row 4)
  if (lastRowIndex < 4) {
    console.log("No data records found to sort.");
    return;
  }
  
  // Get the data range from A4 to L{lastRow}
  let range = sheet.getRange("A4:L" + lastRowIndex);
  let values = range.getValues();
  
  // Status weights (1 is top, 4 is bottom)
  let statusWeights: { [key: string]: number } = {
    'new': 1,
    'renewal': 2,
    'complete': 3,
    'cancelled': 4
  };

  // Chronological month weights (January to December)
  let monthOrder: { [key: string]: number } = {
    'january': 1,
    'february': 2,
    'march': 3,
    'april': 4,
    'may': 5,
    'june': 6,
    'july': 7,
    'august': 8,
    'september': 9,
    'october': 10,
    'november': 11,
    'december': 12
  };
  
  // Helper to parse excel dates (which are returned as numbers or string dates)
  let getTime = function(val: string | number | boolean) {
    if (typeof val === 'number') return val;
    if (!val) return 0;
    let parsed = new Date(val as string).getTime();
    return isNaN(parsed) ? 0 : parsed;
  };

  // Sort values: Status -> Month (Chronological) -> Start Date (Recent to Oldest) -> EU Name
  values.sort((a, b) => {
    // 1. Status (Index 1)
    let statusA = String(a[1] || '').trim().toLowerCase();
    let statusB = String(b[1] || '').trim().toLowerCase();
    
    let weightStatusA = statusWeights[statusA] || 5;
    let weightStatusB = statusWeights[statusB] || 5;
    
    if (weightStatusA !== weightStatusB) {
      return weightStatusA - weightStatusB;
    }

    // 2. Month (Index 0) - Recent to Oldest (December to January)
    let monthA = String(a[0] || '').trim().toLowerCase();
    let monthB = String(b[0] || '').trim().toLowerCase();
    
    let weightMonthA = monthOrder[monthA] || 13;
    let weightMonthB = monthOrder[monthB] || 13;
    
    if (weightMonthA !== weightMonthB) {
      return weightMonthA - weightMonthB;
    }
    
    // 3. Start Date (Index 5) - Recent to Oldest (Descending)
    let timeA = getTime(a[5]);
    let timeB = getTime(b[5]);
    
    if (timeA !== timeB) {
      if (timeA === 0) return 1;  // empty dates to bottom
      if (timeB === 0) return -1; // empty dates to bottom
      return timeB - timeA;       // descending order (recent to oldest)
    }
    
    // 4. EU Name (Index 7)
    let nameA = String(a[7] || '').trim().toLowerCase();
    let nameB = String(b[7] || '').trim().toLowerCase();
    return nameA.localeCompare(nameB);
  });
  
  // Write the sorted values back to the worksheet
  range.setValues(values);
  
  // Define premium pastel colors for formatting
  let colorNew = '#e8f0fe';        // Soft Blue
  let colorRenewal = '#f3e8ff';    // Soft Purple
  let colorComplete = '#e6f4ea';   // Soft Green
  let colorCancelled = '#fce8e6';  // Soft Red
  let colorDefault = '#ffffff';    // White
  
  // Apply formatting row-by-row
  for (let i = 0; i < values.length; i++) {
    let status = String(values[i][1] || '').trim().toLowerCase();
    let rowColor = colorDefault;
    
    if (status === 'new') {
      rowColor = colorNew;
    } else if (status === 'renewal') {
      rowColor = colorRenewal;
    } else if (status === 'complete') {
      rowColor = colorComplete;
    } else if (status === 'cancelled') {
      rowColor = colorCancelled;
    }
    
    let rowRange = sheet.getRange("A" + (i + 4) + ":L" + (i + 4));
    rowRange.getFormat().getFill().setColor(rowColor);
  }
  
  console.log("Sheet sorted by Month & Start Date (Recent to Oldest) and colors applied successfully!");
}
