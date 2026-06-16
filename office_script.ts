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
  let colCHeader = String(headerValues[2] || '').trim().toUpperCase();
  let colDHeader = String(headerValues[3] || '').trim().toUpperCase();
  
  let is12Column = (colBHeader === "STATUS" && colDHeader === "MONTHS LEFT");
  let is10Column = (colBHeader === "MONTHS LEFT");
  let is9Column = (colBHeader === "STATUS" && colDHeader === "R/T");

  let values: any[][];
  let startRow = 4;
  
  if (is10Column || is9Column) {
    let oldValues: any[][];
    if (is10Column) {
      let oldRange = sheet.getRange("A4:J" + lastRowIndex);
      oldValues = oldRange.getValues();
    } else {
      let oldRange = sheet.getRange("A4:I" + lastRowIndex);
      oldValues = oldRange.getValues();
    }

    let restructuredValues: any[][] = [];
    for (let i = 0; i < oldValues.length; i++) {
      let month: any = "", status: any = "", poStatus: any = "";
      let monthsLeft: any = 0, totalMonths: any = 0;
      let startDate: any = "", endDate: any = "";
      let eu: any = "", fa: any = "", so: any = "";
      let subId: any = "", subName: any = "";

      if (is10Column) {
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
        // 9-column layout
        month = oldValues[i][0];
        status = oldValues[i][1];
        poStatus = oldValues[i][2];
        
        let rt = cleanRTValue(oldValues[i][3]);
        let rtParts = rt.split('/');
        monthsLeft = rtParts[0] ? parseInt(rtParts[0].trim(), 10) : 0;
        totalMonths = rtParts[1] ? parseInt(rtParts[1].trim(), 10) : 0;

        let period = cleanPeriodValue(oldValues[i][4]);
        let periodParts = period.trim().split(/\s+/);
        let startStr = periodParts[0] || '';
        let endStr = periodParts[1] || '';
        startDate = parseDDMMMYY(startStr);
        endDate = parseDDMMMYY(endStr);

        eu = oldValues[i][5];

        let faso = String(oldValues[i][6] || '').trim();
        let fasoParts = faso.split(/\s{2,}/);
        if (fasoParts.length >= 2) {
          fa = fasoParts[0].trim();
          so = fasoParts[1].trim();
        } else {
          let soIndex = faso.indexOf('SOUNI');
          if (soIndex !== -1) {
            fa = faso.substring(0, soIndex).trim();
            so = faso.substring(soIndex).trim();
          } else {
            fa = faso;
          }
        }

        subId = oldValues[i][7];
        subName = oldValues[i][8];
      }

      restructuredValues.push([
        month,
        status,
        poStatus,
        monthsLeft,
        totalMonths,
        startDate,
        endDate,
        eu,
        fa,
        so,
        subId,
        subName
      ]);
    }

    // Write new 12-column headers
    let newHeaders = [
      "MONTH", "STATUS", "PO STATUS", "MONTHS LEFT", "TOTAL MONTHS",
      "START DATE", "END DATE", "EU", "FA", "SO", "SUBSCRIPTION ID", "SUBSCRIPTION"
    ];
    // Clear data rows (contents, formats, validation) to remove date formats
    sheet.getRange("A4:L" + lastRowIndex).clear(ExcelScript.ClearApplyTo.all);
    // Clear header contents
    sheet.getRange("A3:L3").clear(ExcelScript.ClearApplyTo.contents);

    // Set new headers
    sheet.getRange("A3:L3").setValues([newHeaders]);

    // Delete any old comments/notes from headers (D3, E3, G3)
    let deleteComment = (cell: ExcelScript.Range) => {
      let existing = sheet.getCommentByCell(cell);
      if (existing) {
        existing.delete();
      }
    };
    deleteComment(sheet.getRange("D3"));
    deleteComment(sheet.getRange("E3"));
    deleteComment(sheet.getRange("G3"));

    values = restructuredValues;
  } else {
    // Already 12 columns, read range A4:L
    values = sheet.getRange("A4:L" + lastRowIndex).getValues();
  }

  // Filter out empty rows (where EU Name (index 7) and Subscription (index 11) are both blank)
  values = values.filter(row => {
    let eu = String(row[7] || '').trim();
    let sub = String(row[11] || '').trim();
    return eu !== '' || sub !== '';
  });

  let activeRowCount = values.length;
  let newLastRowIndex = activeRowCount >= 1 ? 3 + activeRowCount : 3;

  let statusWeights: { [key: string]: number } = {
    'new': 1,
    'renewal': 2,
    'complete': 3,
    'cancelled': 4
  };
  
  let getTime = function(val: any) {
    if (typeof val === 'number') return val;
    if (!val) return 0;
    let parsed = new Date(val).getTime();
    return isNaN(parsed) ? 0 : parsed;
  };

  // Sort values: Status -> Start Date (Recent to Oldest) -> EU Name
  values.sort((a, b) => {
    // 1. Status (Index 1)
    let statusA = String(a[1] || '').trim().toLowerCase();
    let statusB = String(b[1] || '').trim().toLowerCase();
    let weightStatusA = statusWeights[statusA] || 5;
    let weightStatusB = statusWeights[statusB] || 5;
    if (weightStatusA !== weightStatusB) return weightStatusA - weightStatusB;

    // 2. Start Date (Index 5) - Descending
    let timeA = getTime(a[5]);
    let timeB = getTime(b[5]);
    if (timeA !== timeB) {
      if (timeA === 0) return 1;
      if (timeB === 0) return -1;
      return timeB - timeA;
    }
    
    // 3. EU Name (Index 7)
    let nameA = String(a[7] || '').trim().toLowerCase();
    let nameB = String(b[7] || '').trim().toLowerCase();
    return nameA.localeCompare(nameB);
  });
  
  // Set number formats for columns BEFORE writing values to prevent Excel date auto-conversion
  sheet.getRange("A4:C10000").setNumberFormatLocal("@");
  sheet.getRange("D4:E10000").setNumberFormatLocal("General");
  sheet.getRange("F4:G10000").setNumberFormatLocal("[$-809]dddd\\,d\\ mmmm\\ yyyy;@");
  sheet.getRange("H4:L10000").setNumberFormatLocal("@");

  // Clear any existing validations in columns A to L
  sheet.getRange("A4:L10000").getDataValidation().clear();

  if (activeRowCount > 0) {
    let writeRange = sheet.getRange("A4:L" + newLastRowIndex);
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
      
      let rowRange = sheet.getRange("A" + (i + 4) + ":L" + (i + 4));
      rowRange.getFormat().getFill().setColor(rowColor);
    }

    let statusRange = sheet.getRange("B4:B" + newLastRowIndex);
    statusRange.getDataValidation().setRule({
      list: {
        inCellDropDown: true,
        source: "New,Renewal,Complete,Cancelled"
      }
    });

    let poRange = sheet.getRange("C4:C" + newLastRowIndex);
    poRange.getDataValidation().setRule({
      list: {
        inCellDropDown: true,
        source: "PO Done,PO Pending"
      }
    });
  }

  // Clear everything below the active rows to prevent infinite scroll/formatting artifacts
  let clearRangeBelow = sheet.getRange("A" + (newLastRowIndex + 1) + ":L10000");
  clearRangeBelow.clear(ExcelScript.ClearApplyTo.all);

  // Format the header/title row A3:L3 with Orange color and bold font
  let headerRange = sheet.getRange("A3:L3");
  headerRange.getFormat().getFill().setColor("#FFC000");
  headerRange.getFormat().getFont().setBold(true);

  // Set horizontal alignment to Center for the entire range (headers + data)
  sheet.getRange("A3:L10000").getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  
  // Auto-fit all columns to prevent text truncation
  sheet.getUsedRange().getFormat().getAutofitColumns();

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

  // Filter out empty rows (where EU Name (index 3) and Subscription (index 5) are both blank)
  values = values.filter(row => {
    let eu = String(row[3] || '').trim();
    let sub = String(row[5] || '').trim();
    return eu !== '' || sub !== '';
  });

  let activeRowCount = values.length;
  let newLastRowIndex = activeRowCount >= 1 ? 4 + activeRowCount : 4;
  
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
        values[i][0] = "ACTIVE";
      } else if (statusLower === "complete license" || statusLower.indexOf("complete") !== -1) {
        values[i][0] = "COMPLETE";
      } else if (statusLower.indexOf("cancel") !== -1) {
        values[i][0] = "CANCELLED";
      } else {
        if (rawStatus) {
          values[i][0] = rawStatus.toUpperCase();
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
  
  // Set number formats for dates and text in Annual sheet BEFORE writing values
  sheet.getRange("B5:C10000").setNumberFormatLocal("[$-809]dddd\\,d\\ mmmm\\ yyyy;@");
  sheet.getRange("A5:A10000").setNumberFormatLocal("@");
  sheet.getRange("D5:F10000").setNumberFormatLocal("@");

  if (activeRowCount > 0) {
    let writeRange = sheet.getRange("A5:F" + newLastRowIndex);
    writeRange.setValues(values);
    
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
  }

  // Clear everything below the active rows to prevent infinite scroll/formatting artifacts
  let clearRangeBelow = sheet.getRange("A" + (newLastRowIndex + 1) + ":F10000");
  clearRangeBelow.clear(ExcelScript.ClearApplyTo.all);

  // Set A4 header to STATUS and format the header/title row A4:F4 with Orange color and bold font
  sheet.getRange("A4").setValue("STATUS");
  let headerRange = sheet.getRange("A4:F4");
  headerRange.getFormat().getFill().setColor("#FFC000");
  headerRange.getFormat().getFont().setBold(true);

  // Set horizontal alignment to Center for the entire range (headers + data)
  sheet.getRange("A4:F10000").getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  
  // Auto-fit all columns to prevent text truncation
  sheet.getUsedRange().getFormat().getAutofitColumns();

  console.log("Annual Master sheet sorted and colored successfully.");
}

function cleanRTValue(val: any): string {
  if (val === null || val === undefined) return "0/0";
  let str = String(val).trim();
  if (!str) return "0/0";

  // If it's already in the format "X/Y"
  if (/^\d+\/\d+$/.test(str)) {
    return str;
  }

  // If it's a number (Excel serial number) or a Date object
  let d: Date;
  if (typeof val === 'number') {
    d = new Date(Math.round((val - 25569) * 86400 * 1000));
  } else {
    d = new Date(val);
  }

  if (!isNaN(d.getTime())) {
    // Reconstruct Remaining/Total from Month/Day of parsed date
    let monthsLeft = d.getUTCMonth() + 1;
    let totalMonths = d.getUTCDate();
    return `${monthsLeft}/${totalMonths}`;
  }

  return str;
}

function cleanPeriodValue(val: any): string {
  if (val === null || val === undefined) return "";
  let str = String(val).trim();
  if (!str) return "";

  // If it already contains two spaces
  if (str.indexOf('  ') !== -1) {
    return str;
  }

  let d: Date;
  if (typeof val === 'number') {
    d = new Date(Math.round((val - 25569) * 86400 * 1000));
  } else {
    d = new Date(val);
  }

  if (!isNaN(d.getTime())) {
    return formatDateDDMMMYY(val);
  }

  return str;
}

