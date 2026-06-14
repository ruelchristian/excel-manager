function main(workbook: ExcelScript.Workbook) {
  // 1. Process ADB MASTER LIST MONTHLY
  let monthlySheet = workbook.getWorksheet("ADB MASTER LIST MONTHLY");
  if (monthlySheet) {
    sortAndColorMonthly(monthlySheet);
  } else {
    console.log('Sheet "ADB MASTER LIST MONTHLY" not found.');
  }
  
  // 2. Process ADB MASTER LIST ANNUAL
  let annualSheet = workbook.getWorksheet("ADB MASTER LIST ANNUAL");
  if (annualSheet) {
    sortAndColorAnnual(annualSheet);
  } else {
    console.log('Sheet "ADB MASTER LIST ANNUAL" not found.');
  }
}

function sortAndColorMonthly(sheet: ExcelScript.Worksheet) {
  let rangeUsed = sheet.getUsedRange();
  if (!rangeUsed) return;
  
  let lastRowRange = rangeUsed.getLastRow();
  let lastRowIndex = lastRowRange.getRowIndex() + 1; // 1-based row index
  if (lastRowIndex < 4) return;
  
  let range = sheet.getRange("A4:L" + lastRowIndex);
  let values = range.getValues();
  
  let statusWeights: { [key: string]: number } = {
    'new': 1,
    'renewal': 2,
    'complete': 3,
    'cancelled': 4
  };

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
  
  let getTime = function(val: string | number | boolean) {
    if (typeof val === 'number') return val;
    if (!val) return 0;
    let parsed = new Date(val as string).getTime();
    return isNaN(parsed) ? 0 : parsed;
  };

  // Sort values: Status -> Month -> Start Date (Recent to Oldest) -> EU Name
  values.sort((a, b) => {
    // 1. Status (Index 1)
    let statusA = String(a[1] || '').trim().toLowerCase();
    let statusB = String(b[1] || '').trim().toLowerCase();
    let weightStatusA = statusWeights[statusA] || 5;
    let weightStatusB = statusWeights[statusB] || 5;
    if (weightStatusA !== weightStatusB) return weightStatusA - weightStatusB;

    // 2. Month (Index 0)
    let monthA = String(a[0] || '').trim().toLowerCase();
    let monthB = String(b[0] || '').trim().toLowerCase();
    let weightMonthA = monthOrder[monthA] || 13;
    let weightMonthB = monthOrder[monthB] || 13;
    if (weightMonthA !== weightMonthB) return weightMonthA - weightMonthB;
    
    // 3. Start Date (Index 5) - Descending
    let timeA = getTime(a[5]);
    let timeB = getTime(b[5]);
    if (timeA !== timeB) {
      if (timeA === 0) return 1;
      if (timeB === 0) return -1;
      return timeB - timeA;
    }
    
    // 4. EU Name (Index 7)
    let nameA = String(a[7] || '').trim().toLowerCase();
    let nameB = String(b[7] || '').trim().toLowerCase();
    return nameA.localeCompare(nameB);
  });
  
  range.setValues(values);
  
  let colorNew = '#e8f0fe';        // Soft Blue
  let colorRenewal = '#f3e8ff';    // Soft Purple
  let colorComplete = '#e6f4ea';   // Soft Green
  let colorCancelled = '#fce8e6';  // Soft Red
  let colorDefault = '#ffffff';    // White
  
  for (let i = 0; i < values.length; i++) {
    let status = String(values[i][1] || '').trim().toLowerCase();
    let rowColor = colorDefault;
    
    if (status === 'new') {
      rowColor = colorNew;
    } else if (status === 'renewal') {
      rowColor = colorRenewal;
    } else if (status === 'complete' || status.indexOf('complete') !== -1) {
      rowColor = colorComplete;
    } else if (status === 'cancelled' || status.indexOf('cancel') !== -1) {
      rowColor = colorCancelled;
    }
    
    let rowRange = sheet.getRange("A" + (i + 4) + ":L" + (i + 4));
    rowRange.getFormat().getFill().setColor(rowColor);
  }
  
  console.log("Monthly Master sheet sorted and colored successfully.");
}

function sortAndColorAnnual(sheet: ExcelScript.Worksheet) {
  let rangeUsed = sheet.getUsedRange();
  if (!rangeUsed) return;
  
  let lastRowRange = rangeUsed.getLastRow();
  let lastRowIndex = lastRowRange.getRowIndex() + 1;
  if (lastRowIndex < 5) return;
  
  let range = sheet.getRange("A5:F" + lastRowIndex);
  let values = range.getValues();
  
  let statusWeights: { [key: string]: number } = {
    'active': 1,
    'complete': 2,
    'cancelled': 3
  };
  
  let getTime = function(val: string | number | boolean) {
    if (typeof val === 'number') return val;
    if (!val) return 0;
    let parsed = new Date(val as string).getTime();
    return isNaN(parsed) ? 0 : parsed;
  };

  // Sort values: Status -> Start Date (Recent to Oldest) -> EU Name
  values.sort((a, b) => {
    // 1. Status (Index 0)
    let statusA = String(a[0] || '').trim().toLowerCase();
    let statusB = String(b[0] || '').trim().toLowerCase();
    let weightStatusA = statusWeights[statusA] || 4;
    let weightStatusB = statusWeights[statusB] || 4;
    if (weightStatusA !== weightStatusB) return weightStatusA - weightStatusB;
    
    // 2. Start Date (Index 1) - Descending
    let timeA = getTime(a[1]);
    let timeB = getTime(b[1]);
    if (timeA !== timeB) {
      if (timeA === 0) return 1;
      if (timeB === 0) return -1;
      return timeB - timeA;
    }
    
    // 3. EU Name (Index 3)
    let nameA = String(a[3] || '').trim().toLowerCase();
    let nameB = String(b[3] || '').trim().toLowerCase();
    return nameA.localeCompare(nameB);
  });
  
  range.setValues(values);
  
  let colorActive = '#e8f0fe';       // Soft Blue
  let colorComplete = '#e6f4ea';     // Soft Green
  let colorCancelled = '#fce8e6';    // Soft Red
  let colorDefault = '#ffffff';
  
  for (let i = 0; i < values.length; i++) {
    let status = String(values[i][0] || '').trim().toLowerCase();
    let rowColor = colorDefault;
    
    if (status === 'active' || status.indexOf('active') !== -1) {
      rowColor = colorActive;
    } else if (status === 'complete' || status.indexOf('complete') !== -1) {
      rowColor = colorComplete;
    } else if (status === 'cancelled' || status.indexOf('cancel') !== -1) {
      rowColor = colorCancelled;
    }
    
    let rowRange = sheet.getRange("A" + (i + 5) + ":F" + (i + 5));
    rowRange.getFormat().getFill().setColor(rowColor);
  }
  
  console.log("Annual Master sheet sorted and colored successfully.");
}
