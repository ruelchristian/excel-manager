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
  
  // 1. Detect column layout from Header Row 3
  let headerRange = sheet.getRange("A3:L3");
  let headerValues = headerRange.getValues()[0];
  
  let colBHeader = String(headerValues[1] || '').trim().toUpperCase();
  let colDHeader = String(headerValues[3] || '').trim().toUpperCase();
  
  let isOld10Column = (colBHeader === "MONTHS LEFT");
  let isOld12Column = (colBHeader === "STATUS" && colDHeader === "MONTHS LEFT");

  let values: any[][];
  
  if (isOld10Column || isOld12Column) {
    // Read the old data
    let oldValues: any[][];
    if (isOld10Column) {
      let oldRange = sheet.getRange("A4:J" + lastRowIndex);
      oldValues = oldRange.getValues();
    } else {
      let oldRange = sheet.getRange("A4:L" + lastRowIndex);
      oldValues = oldRange.getValues();
    }

    // Restructure to 9 columns in-memory
    let restructuredValues: any[][] = [];
    for (let i = 0; i < oldValues.length; i++) {
      let month: any, status: any, poStatus: any;
      let monthsLeft: any, totalMonths: any;
      let startDate: any, endDate: any;
      let eu: any, fa: any, so: any;
      let subId: any, subName: any;

      if (isOld10Column) {
        let rawStatus = oldValues[i][0];
        monthsLeft = oldValues[i][1];
        totalMonths = oldValues[i][2];
        startDate = oldValues[i][3];
        endDate = oldValues[i][4];
        eu = oldValues[i][5];
        fa = oldValues[i][6];
        so = oldValues[i][7];
        subId = oldValues[i][8];
        subName = oldValues[i][9];

        let parsed = parseStatusField(rawStatus, endDate);
        month = parsed.month;
        status = parsed.status;
        poStatus = parsed.poStatus;
      } else {
        month = oldValues[i][0];
        status = oldValues[i][1];
        poStatus = oldValues[i][2];
        monthsLeft = oldValues[i][3];
        totalMonths = oldValues[i][4];
        startDate = oldValues[i][5];
        endDate = oldValues[i][6];
        eu = oldValues[i][7];
        fa = oldValues[i][8];
        so = oldValues[i][9];
        subId = oldValues[i][10];
        subName = oldValues[i][11];
      }

      // Merge Months Left and Total Months into R/T
      let left = monthsLeft !== null && monthsLeft !== "" ? String(monthsLeft).trim() : "0";
      let total = totalMonths !== null && totalMonths !== "" ? String(totalMonths).trim() : "0";
      let rt = left + "/" + total;

      // Merge Start Date and End Date into Period
      let startStr = formatDateDDMMMYY(startDate);
      let endStr = formatDateDDMMMYY(endDate);
      let period = startStr + (endStr ? "  " + endStr : "");

      // Merge FA and SO into FA/SO
      let faStr = String(fa || '').trim();
      let soStr = String(so || '').trim();
      let faso = faStr + (faStr && soStr ? "  " : "") + soStr;

      restructuredValues.push([
        month,
        status,
        poStatus,
        rt,
        period,
        eu,
        faso,
        subId,
        subName
      ]);
    }

    // Write new 9-column headers
    let newHeaders = [
      "MONTH", "STATUS", "PO STATUS", "R/T", "Period",
      "EU", "FA/SO", "SUBSCRIPTION ID", "SUBSCRIPTION"
    ];
    // Clear columns A to L to avoid legacy data artifacts
    let clearRange = sheet.getRange("A3:L" + lastRowIndex);
    clearRange.clear(ExcelScript.ClearApplyTo.contents);

    // Set new headers
    let newHeaderRange = sheet.getRange("A3:I3");
    newHeaderRange.setValues([newHeaders]);

    // Set comments (notes/tooltips) to headers
    let setComment = (cell: ExcelScript.Range, text: string) => {
      let existing = sheet.getCommentByCell(cell);
      if (existing) {
        existing.delete();
      }
      sheet.addComment(cell, text);
    };
    setComment(sheet.getRange("D3"), "R/T = Remaining Months / Total Months");
    setComment(sheet.getRange("E3"), "Period = Start Date  End Date");
    setComment(sheet.getRange("G3"), "FA/SO = Financial Advisor  Sales Officer");

    values = restructuredValues;
  } else {
    // Already 9 columns, read range A4:I
    let range = sheet.getRange("A4:I" + lastRowIndex);
    values = range.getValues();
  }
  
  let statusWeights: { [key: string]: number } = {
    'new': 1,
    'renewal': 2,
    'complete': 3,
    'cancelled': 4
  };
  
  // Sort values: Status -> Start Date (Recent to Oldest) -> EU Name
  values.sort((a, b) => {
    // 1. Status (Index 1)
    let statusA = String(a[1] || '').trim().toLowerCase();
    let statusB = String(b[1] || '').trim().toLowerCase();
    let weightStatusA = statusWeights[statusA] || 5;
    let weightStatusB = statusWeights[statusB] || 5;
    if (weightStatusA !== weightStatusB) return weightStatusA - weightStatusB;

    // 2. Start Date (Index 4) - Descending
    let timeA = getTimeFromPeriod(a[4]);
    let timeB = getTimeFromPeriod(b[4]);
    if (timeA !== timeB) {
      if (timeA === 0) return 1;
      if (timeB === 0) return -1;
      return timeB - timeA;
    }
    
    // 3. EU Name (Index 5)
    let nameA = String(a[5] || '').trim().toLowerCase();
    let nameB = String(b[5] || '').trim().toLowerCase();
    return nameA.localeCompare(nameB);
  });
  
  let writeRange = sheet.getRange("A4:I" + lastRowIndex);
  writeRange.setValues(values);
  
  let colorNew = '#fef08a';        // Soft Yellow
  let colorRenewal = '#bbf7d0';    // Soft Green
  let colorComplete = '#bfdbfe';   // Soft Blue
  let colorCancelled = '#fecaca';  // Soft Red
  let colorDefault = '#ffffff';    // White
  
  for (let i = 0; i < values.length; i++) {
    let status = String(values[i][1] || '').trim().toLowerCase();
    let rowColor = colorDefault;
    
    if (status === 'new') {
      rowColor = colorNew;
    } else if (status === 'renewal') {
      rowColor = colorRenewal;
    } else if (status === 'complete' || status.indexOf('complete') !== -1 || status === 'closed') {
      rowColor = colorComplete;
    } else if (status === 'cancelled' || status.indexOf('cancel') !== -1) {
      rowColor = colorCancelled;
    }
    
    let rowRange = sheet.getRange("A" + (i + 4) + ":I" + (i + 4));
    rowRange.getFormat().getFill().setColor(rowColor);
  }

  // Clear any existing validations in columns B, C, D first
  sheet.getRange("B4:D10000").getDataValidation().clear();

  // Set number format of all columns (A to I) to Text to prevent Excel from converting R/T and Period to dates
  sheet.getRange("A4:I10000").setNumberFormatLocal("@");


  if (lastRowIndex >= 4) {
    let statusRange = sheet.getRange("B4:B" + lastRowIndex);
    statusRange.getDataValidation().setRule({
      list: {
        inCellDropDown: true,
        source: "New,Renewal,Complete,Cancelled"
      }
    });

    let poRange = sheet.getRange("C4:C" + lastRowIndex);
    poRange.getDataValidation().setRule({
      list: {
        inCellDropDown: true,
        source: "PO Done,PO Pending"
      }
    });
  }

  // Clear everything below the active rows to prevent infinite scroll/formatting artifacts
  if (lastRowIndex < 10000) {
    let clearRangeBelow = sheet.getRange("A" + (lastRowIndex + 1) + ":L10000");
    clearRangeBelow.clear(ExcelScript.ClearApplyTo.all);
  }
  
  console.log("Monthly Master sheet sorted and colored successfully.");
}

function getTimeFromPeriod(val: any): number {
  if (!val) return 0;
  let periodStr = String(val).trim();
  let firstDate = periodStr.split(' ')[0] || '';
  if (!firstDate) return 0;
  
  // Parse DD-MMM-YY
  let parts = firstDate.split('-');
  if (parts.length === 3) {
    let day = parseInt(parts[0], 10);
    let monthName = parts[1].toLowerCase();
    let yearShort = parseInt(parts[2], 10);
    
    const months = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];
    let monthIndex = months.indexOf(monthName);
    if (monthIndex !== -1) {
      let year = yearShort < 50 ? 2000 + yearShort : 1900 + yearShort;
      let d = new Date(year, monthIndex, day);
      return d.getTime();
    }
  }
  return 0;
}

function formatDateDDMMMYY(val: any): string {
  if (!val) return "";
  const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  let d: Date;
  if (typeof val === 'number') {
    d = new Date(Math.round((val - 25569) * 86400 * 1000));
  } else {
    d = new Date(val);
  }
  if (isNaN(d.getTime())) return "";

  if (typeof val === 'string') {
    let clean = val.trim();
    let matchISO = clean.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
    if (matchISO) {
      let y = matchISO[1].slice(-2);
      let mIndex = parseInt(matchISO[2], 10) - 1;
      let dayStr = String(parseInt(matchISO[3], 10)).padStart(2, '0');
      if (mIndex >= 0 && mIndex < 12) {
        return `${dayStr}-${monthNames[mIndex]}-${y}`;
      }
    }
    let matchUS = clean.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})/);
    if (matchUS) {
      let y = matchUS[3].slice(-2);
      let mIndex = parseInt(matchUS[1], 10) - 1;
      let dayStr = String(parseInt(matchUS[2], 10)).padStart(2, '0');
      if (mIndex >= 0 && mIndex < 12) {
        return `${dayStr}-${monthNames[mIndex]}-${y}`;
      }
    }
  }

  let day = String(d.getUTCDate()).padStart(2, '0');
  let month = monthNames[d.getUTCMonth()];
  let year = String(d.getUTCFullYear()).slice(-2);
  return `${day}-${month}-${year}`;
}

function getMonthFromDate(val: any): string | null {
  if (!val) return null;
  const monthNames = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
  ];
  if (typeof val === 'number') {
    let d = new Date(Math.round((val - 25569) * 86400 * 1000));
    return monthNames[d.getUTCMonth()];
  }
  if (typeof val === 'string') {
    let cleaned = val.trim();
    let match = cleaned.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
    if (match) {
      let monthIndex = parseInt(match[2], 10) - 1;
      if (monthIndex >= 0 && monthIndex < 12) {
        return monthNames[monthIndex];
      }
    }
    let d = new Date(cleaned);
    if (!isNaN(d.getTime())) {
      if (cleaned.includes('Z') || cleaned.includes('T')) {
        return monthNames[d.getUTCMonth()];
      }
      return monthNames[d.getMonth()];
    }
  }
  return null;
}

function parseStatusField(rawStatus: any, endDateVal: any) {
  let statusStr = String(rawStatus || '').trim().toLowerCase();
  
  let month: string | null = null;
  let status: string | null = null;
  let poStatus: string | null = null;

  const months = ["january", "february", "march", "april", "may", "june", "july", "august", "september", "october", "november", "december"];
  
  let foundMonth = months.find(m => statusStr.includes(m));
  if (foundMonth) {
    month = foundMonth.charAt(0).toUpperCase() + foundMonth.slice(1);
  }

  if (statusStr.includes('renewal')) {
    status = 'Renewal';
  } else if (statusStr.includes('new')) {
    status = 'New';
  } else if (statusStr.includes('complete') || statusStr.includes('closed') || statusStr.includes('active') || statusStr.includes('license')) {
    status = 'Complete';
  } else if (statusStr.includes('cancel')) {
    status = 'Cancelled';
  } else {
    status = 'Complete';
  }

  if (status === 'Complete' && !month) {
    month = getMonthFromDate(endDateVal);
  }

  if (statusStr.includes('done po') || statusStr.includes('done  po') || statusStr.includes('po done') || statusStr.includes('complete') || statusStr.includes('closed') || statusStr.includes('active') || statusStr === "") {
    poStatus = 'PO Done';
  } else if (statusStr.includes('pending po') || statusStr.includes('po pending')) {
    poStatus = 'PO Pending';
  }

  return { month, status, poStatus };
}

function sortAndColorAnnual(sheet: ExcelScript.Worksheet) {
  let rangeUsed = sheet.getUsedRange();
  if (!rangeUsed) return;
  
  let lastRowRange = rangeUsed.getLastRow();
  let lastRowIndex = lastRowRange.getRowIndex() + 1;
  if (lastRowIndex < 5) return;
  
  let range = sheet.getRange("A5:F" + lastRowIndex);
  let values = range.getValues();
  
  // Clean up and standardize statuses in memory
  for (let i = 0; i < values.length; i++) {
    let rawStatus = String(values[i][0] || '').trim();
    let statusLower = rawStatus.toLowerCase();
    
    // Check if the row has any actual data (Columns B to F / index 1 to 5)
    let hasData = false;
    for (let col = 1; col < 6; col++) {
      if (values[i][col] !== null && String(values[i][col]).trim() !== "") {
        hasData = true;
        break;
      }
    }
    
    if (hasData) {
      if (statusLower === "" || statusLower === "none" || statusLower === "null" || statusLower === "undefined") {
        values[i][0] = "Active";
      } else if (statusLower === "complete license" || statusLower.indexOf("complete") !== -1) {
        values[i][0] = "Complete";
      } else if (statusLower.indexOf("cancel") !== -1) {
        values[i][0] = "Cancelled";
      } else {
        if (rawStatus) {
          values[i][0] = rawStatus.charAt(0).toUpperCase() + rawStatus.slice(1);
        }
      }
    } else {
      // If row has no actual data, ensure status is blank so it goes to the bottom
      values[i][0] = "";
    }
  }
  
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
  
  let colorActive = '#fef08a';       // Soft Yellow
  let colorComplete = '#bfdbfe';     // Soft Blue
  let colorCancelled = '#fecaca';    // Soft Red
  let colorDefault = '#ffffff';
  
  for (let i = 0; i < values.length; i++) {
    let status = String(values[i][0] || '').trim().toLowerCase();
    let rowColor = colorDefault;
    
    if (status === 'active' || status.indexOf('active') !== -1) {
      rowColor = colorActive;
    } else if (status === 'complete' || status.indexOf('complete') !== -1 || status === 'closed') {
      rowColor = colorComplete;
    } else if (status === 'cancelled' || status.indexOf('cancel') !== -1) {
      rowColor = colorCancelled;
    }
    
    let rowRange = sheet.getRange("A" + (i + 5) + ":F" + (i + 5));
    rowRange.getFormat().getFill().setColor(rowColor);
  }

  // Set number formats for dates and text in Annual sheet
  sheet.getRange("B5:C10000").setNumberFormatLocal("yyyy-mm-dd");
  sheet.getRange("A5:A10000").setNumberFormatLocal("@");
  sheet.getRange("D5:F10000").setNumberFormatLocal("@");

  // Clear everything below the active rows to prevent infinite scroll/formatting artifacts
  if (lastRowIndex < 10000) {
    let clearRangeBelow = sheet.getRange("A" + (lastRowIndex + 1) + ":F10000");
    clearRangeBelow.clear(ExcelScript.ClearApplyTo.all);
  }
  
  console.log("Annual Master sheet sorted and colored successfully.");
}
