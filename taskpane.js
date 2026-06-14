var allSubscriptions = [];

// Initialize Office.js
Office.onReady(function (info) {
  if (info.host === Office.HostType.Excel) {
    document.getElementById('loader').innerText = "Loading subscriptions from Excel...";
    loadSubscriptions();
  } else {
    showError("This add-in only works in Microsoft Excel.");
  }
});

// Switch tab view
function switchTab(tabName) {
  document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
  document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
  
  if (tabName === 'list') {
    document.querySelectorAll('.tab')[0].classList.add('active');
    document.getElementById('tab-list').classList.add('active');
    loadSubscriptions();
  } else {
    document.querySelectorAll('.tab')[1].classList.add('active');
    document.getElementById('tab-form').classList.add('active');
  }
}

// Helper to parse Excel dates to yyyy-MM-dd format for HTML inputs
function parseExcelDate(val, textVal) {
  if (!val) return "";
  
  // If serial number
  if (typeof val === 'number') {
    // Excel base date is Dec 31, 1899 (represented as 0)
    var date = new Date(Math.round((val - 25569) * 86400 * 1000));
    var y = date.getFullYear();
    var m = String(date.getMonth() + 1).padStart(2, '0');
    var d = String(date.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }
  
  // If formatted text is already yyyy-mm-dd
  if (textVal && /^\d{4}-\d{2}-\d{2}$/.test(textVal.trim())) {
    return textVal.trim();
  }
  
  // Attempt string parsing
  try {
    var parsed = new Date(val);
    if (!isNaN(parsed.getTime())) {
      var y = parsed.getFullYear();
      var m = String(parsed.getMonth() + 1).padStart(2, '0');
      var d = String(parsed.getDate()).padStart(2, '0');
      return `${y}-${m}-${d}`;
    }
  } catch(e) {}
  
  return String(val || "").trim();
}

// Load records from Excel Workbook
async function loadSubscriptions() {
  document.getElementById('loader').style.display = 'block';
  document.getElementById('subscription-container').innerHTML = '';
  document.getElementById('error-box').style.display = 'none';

  try {
    await Excel.run(async function (context) {
      var sheet = context.workbook.worksheets.getItemOrNullObject("ADB MASTER LIST MONTHLY");
      sheet.load("nullObject");
      await context.sync();

      if (sheet.isNullObject) {
        throw new Error('Sheet "ADB MASTER LIST MONTHLY" not found. Please create or rename your sheet.');
      }

      // Find the last row
      var rangeUsed = sheet.getUsedRange();
      var lastRowRange = rangeUsed.getLastRow();
      lastRowRange.load("rowIndex");
      await context.sync();

      var lastRowIndex = lastRowRange.rowIndex + 1; // 1-based row number
      if (lastRowIndex < 4) {
        // No data rows (starts at row 4)
        allSubscriptions = [];
        document.getElementById('row-count').innerText = "0";
        document.getElementById('loader').style.display = 'none';
        renderList([]);
        return;
      }

      // Read A4:L{lastRow}
      var range = sheet.getRange("A4:L" + lastRowIndex);
      range.load(["values", "text"]);
      await context.sync();

      var values = range.values;
      var text = range.text;
      var records = [];

      for (var i = 0; i < values.length; i++) {
        var rowVal = values[i];
        var rowText = text[i];
        var rowNum = i + 4; // Excel row index is 1-based, starts at 4

        // Ignore empty rows
        if (!rowVal[7] && !rowVal[11]) continue;

        records.push({
          rowNum: rowNum,
          month: String(rowVal[0] || '').trim(),
          status: String(rowVal[1] || '').trim(),
          poStatus: String(rowVal[2] || '').trim(),
          monthsLeft: rowVal[3] !== null ? String(rowVal[3]).trim() : '',
          totalMonths: rowVal[4] !== null ? String(rowVal[4]).trim() : '',
          startDate: parseExcelDate(rowVal[5], rowText[5]),
          endDate: parseExcelDate(rowVal[6], rowText[6]),
          eu: String(rowVal[7] || '').trim(),
          fa: String(rowVal[8] || '').trim(),
          so: String(rowVal[9] || '').trim(),
          subId: String(rowVal[10] || '').trim(),
          subscription: String(rowVal[11] || '').trim()
        });
      }

      allSubscriptions = records;
      document.getElementById('row-count').innerText = records.length;
      document.getElementById('loader').style.display = 'none';
      renderList(records);
    });
  } catch (err) {
    showError("Failed to load subscriptions: " + err.message);
  }
}

// Render list cards in Taskpane
function renderList(records) {
  var container = document.getElementById('subscription-container');
  container.innerHTML = '';
  
  if (records.length === 0) {
    container.innerHTML = '<div class="loading">No records found.</div>';
    return;
  }

  records.forEach(function(rec) {
    var cardClass = '';
    var badgeClass = '';
    var statusLabel = rec.status ? rec.status : 'Complete';
    var statusLower = rec.status.toLowerCase();

    if (statusLower === 'new') {
      cardClass = 'status-new';
      badgeClass = 'badge-new';
      statusLabel = 'NEW';
    } else if (statusLower === 'renewal') {
      cardClass = 'status-renewal';
      badgeClass = 'badge-renewal';
      statusLabel = 'RENEWAL';
    } else if (statusLower === 'complete') {
      cardClass = 'status-completed';
      badgeClass = 'badge-completed';
      statusLabel = 'COMPLETE';
    } else if (statusLower === 'cancelled') {
      cardClass = 'status-cancelled';
      badgeClass = 'badge-cancelled';
      statusLabel = 'CANCELLED';
    }

    var card = document.createElement('div');
    card.className = 'card ' + cardClass;
    card.onclick = function() { editRecord(rec); };

    card.innerHTML = 
      '<div>' +
        '<div class="card-header">' +
          '<div class="card-title">' + rec.eu + '</div>' +
          '<div class="card-badge ' + badgeClass + '">' + statusLabel + '</div>' +
        '</div>' +
        '<div style="font-size: 12px; font-weight: 500; color: #1e293b; margin-bottom: 4px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;">' + rec.subscription + '</div>' +
        '<div class="card-details">' +
          '<div><strong>FA/SO:</strong> ' + (rec.fa || '-') + ' / ' + (rec.so || '-') + '</div>' +
          '<div><strong>Sub ID:</strong> ' + (rec.subId || '-') + '</div>' +
          '<div><strong>Months:</strong> ' + (rec.monthsLeft || '0') + ' / ' + (rec.totalMonths || '0') + '</div>' +
        '</div>' +
      '</div>' +
      '<div class="card-footer">' +
        '<span>Month: <strong>' + (rec.month || 'N/A') + '</strong></span>' +
        '<span>PO: <strong>' + (rec.poStatus || 'N/A') + '</strong></span>' +
      '</div>';
    
    container.appendChild(card);
  });
}

// Filter cards by search input
function filterList() {
  var query = document.getElementById('search-input').value.toLowerCase();
  var filtered = allSubscriptions.filter(function(rec) {
    return rec.eu.toLowerCase().indexOf(query) !== -1 ||
           rec.so.toLowerCase().indexOf(query) !== -1 ||
           rec.subId.toLowerCase().indexOf(query) !== -1 ||
           rec.subscription.toLowerCase().indexOf(query) !== -1 ||
           rec.fa.toLowerCase().indexOf(query) !== -1;
  });
  renderList(filtered);
}

// Load card details into form
function editRecord(rec) {
  document.getElementById('form-row-num').value = rec.rowNum;
  document.getElementById('form-eu').value = rec.eu;
  
  var subSelect = document.getElementById('form-subscription');
  var matched = false;
  for (var i = 0; i < subSelect.options.length; i++) {
    if (subSelect.options[i].value.toUpperCase() === rec.subscription.toUpperCase()) {
      subSelect.selectedIndex = i;
      matched = true;
      break;
    }
  }
  if (!matched && rec.subscription) {
    var opt = document.createElement('option');
    opt.value = rec.subscription;
    opt.text = rec.subscription;
    subSelect.add(opt);
    subSelect.value = rec.subscription;
  } else if (!rec.subscription) {
    subSelect.value = '';
  }

  document.getElementById('form-months-left').value = rec.monthsLeft;
  document.getElementById('form-total-months').value = rec.totalMonths;
  document.getElementById('form-start-date').value = rec.startDate;
  document.getElementById('form-end-date').value = rec.endDate;
  document.getElementById('form-fa').value = rec.fa;
  document.getElementById('form-so').value = rec.so;
  document.getElementById('form-sub-id').value = rec.subId;
  
  document.getElementById('form-month').value = rec.month ? rec.month : '';
  document.getElementById('form-status').value = rec.status ? rec.status : 'Complete';
  document.getElementById('form-po-status').value = rec.poStatus ? rec.poStatus : 'N/A';

  document.getElementById('save-btn').innerText = 'Update (Row ' + rec.rowNum + ')';
  document.getElementById('cancel-edit').style.display = 'block';

  switchTab('form');
}

// Reset Form
function resetForm() {
  document.getElementById('record-form').reset();
  document.getElementById('form-row-num').value = '';
  document.getElementById('save-btn').innerText = 'Save Subscription';
  document.getElementById('cancel-edit').style.display = 'none';
}

// Save form data back to Excel Sheet
async function saveRecord(e) {
  e.preventDefault();
  document.getElementById('error-box').style.display = 'none';
  document.getElementById('save-btn').innerText = 'Saving...';
  document.getElementById('save-btn').disabled = true;

  var rowNumVal = document.getElementById('form-row-num').value;
  var rowNum = rowNumVal ? parseInt(rowNumVal, 10) : null;

  var formData = {
    month: document.getElementById('form-month').value,
    status: document.getElementById('form-status').value,
    poStatus: document.getElementById('form-po-status').value,
    monthsLeft: document.getElementById('form-months-left').value,
    totalMonths: document.getElementById('form-total-months').value,
    startDate: document.getElementById('form-start-date').value,
    endDate: document.getElementById('form-end-date').value,
    eu: document.getElementById('form-eu').value,
    fa: document.getElementById('form-fa').value,
    so: document.getElementById('form-so').value,
    subId: document.getElementById('form-sub-id').value,
    subscription: document.getElementById('form-subscription').value
  };

  try {
    await Excel.run(async function (context) {
      var sheet = context.workbook.worksheets.getItem("ADB MASTER LIST MONTHLY");
      var range;

      if (rowNum) {
        // Edit existing
        range = sheet.getRange("A" + rowNum + ":L" + rowNum);
      } else {
        // Append at end
        var rangeUsed = sheet.getUsedRange();
        var lastRowRange = rangeUsed.getLastRow();
        lastRowRange.load("rowIndex");
        await context.sync();
        
        var nextRow = lastRowRange.rowIndex + 2; // Index is 0-based
        if (nextRow < 4) nextRow = 4;
        range = sheet.getRange("A" + nextRow + ":L" + nextRow);
      }

      // Map data into columns A to L
      var values = [
        [
          formData.month,
          formData.status,
          formData.poStatus,
          formData.monthsLeft ? parseInt(formData.monthsLeft, 10) : '',
          formData.totalMonths ? parseInt(formData.totalMonths, 10) : '',
          formData.startDate,
          formData.endDate,
          formData.eu,
          formData.fa,
          formData.so,
          formData.subId,
          formData.subscription
        ]
      ];

      range.values = values;

      // Number formatting: Months Left / Total Months (Col D & E)
      var colD_E = range.getCell(0, 3).getResizedRange(0, 1);
      colD_E.numberFormat = [["0", "0"]];

      // Date formatting: Start & End Date (Col F & G)
      var colF_G = range.getCell(0, 5).getResizedRange(0, 1);
      colF_G.numberFormat = [["yyyy-mm-dd", "yyyy-mm-dd"]];

      await context.sync();
    });

    // Run custom sort and color formatting
    await sortSubscriptions();
    
    resetForm();
    switchTab('list');
    document.getElementById('save-btn').disabled = false;
  } catch (err) {
    showError("Failed to save record: " + err.message);
    document.getElementById('save-btn').disabled = false;
    document.getElementById('save-btn').innerText = 'Save Subscription';
  }
}

// In-Memory Custom Sort (weighting Statuses) and formatting colors
async function sortSubscriptions() {
  try {
    await Excel.run(async function (context) {
      var sheet = context.workbook.worksheets.getItem("ADB MASTER LIST MONTHLY");
      var rangeUsed = sheet.getUsedRange();
      var lastRowRange = rangeUsed.getLastRow();
      lastRowRange.load("rowIndex");
      await context.sync();

      var lastRowIndex = lastRowRange.rowIndex + 1;
      if (lastRowIndex < 4) return;

      var range = sheet.getRange("A4:L" + lastRowIndex);
      range.load("values");
      await context.sync();

      var values = range.values;

      var statusWeights = {
        'new': 1,
        'renewal': 2,
        'complete': 3,
        'cancelled': 4
      };

      // Custom multi-tier sort: Status -> End Date -> EU Name
      values.sort(function(a, b) {
        var statusA = String(a[1] || '').trim().toLowerCase();
        var statusB = String(b[1] || '').trim().toLowerCase();
        
        var weightA = statusWeights[statusA] || 5;
        var weightB = statusWeights[statusB] || 5;
        
        if (weightA !== weightB) {
          return weightA - weightB;
        }
        
        // Secondary sort: End Date (Index 6)
        var dateA = a[6] ? new Date(a[6]).getTime() : 0;
        var dateB = b[6] ? new Date(b[6]).getTime() : 0;
        
        if (dateA !== dateB) {
          if (dateA === 0) return 1;
          if (dateB === 0) return -1;
          return dateA - dateB;
        }
        
        // Tertiary sort: EU Name (Index 7)
        var nameA = String(a[7] || '').trim().toLowerCase();
        var nameB = String(b[7] || '').trim().toLowerCase();
        return nameA.localeCompare(nameB);
      });

      // Overwrite cells with sorted values
      range.values = values;
      await context.sync();
      
      // Re-apply background color formatting based on status
      await formatSheetColorsDirect(sheet, values);
    });
  } catch (err) {
    showError("Sorting failed: " + err.message);
  }
}

// Background row pastel coloring helper
async function formatSheetColorsDirect(sheet, values) {
  var colorNew = '#e8f0fe';        // Soft Blue
  var colorRenewal = '#f3e8ff';    // Soft Purple
  var colorComplete = '#e6f4ea';   // Soft Green
  var colorCancelled = '#fce8e6';  // Soft Red
  var colorDefault = '#ffffff';    // White

  for (var i = 0; i < values.length; i++) {
    var status = String(values[i][1] || '').trim().toLowerCase();
    var rowColor = colorDefault;

    if (status === 'new') {
      rowColor = colorNew;
    } else if (status === 'renewal') {
      rowColor = colorRenewal;
    } else if (status === 'complete') {
      rowColor = colorComplete;
    } else if (status === 'cancelled') {
      rowColor = colorCancelled;
    }

    var rowRange = sheet.getRange("A" + (i + 4) + ":L" + (i + 4));
    rowRange.format.fill.color = rowColor;
  }
  await sheet.context.sync();
}

// Trigger row formatting manually
async function runFormatter() {
  document.getElementById('loader').style.display = 'block';
  try {
    await Excel.run(async function (context) {
      var sheet = context.workbook.worksheets.getItem("ADB MASTER LIST MONTHLY");
      var rangeUsed = sheet.getUsedRange();
      var lastRowRange = rangeUsed.getLastRow();
      lastRowRange.load("rowIndex");
      await context.sync();

      var lastRowIndex = lastRowRange.rowIndex + 1;
      if (lastRowIndex < 4) return;

      var range = sheet.getRange("A4:L" + lastRowIndex);
      range.load("values");
      await context.sync();

      await formatSheetColorsDirect(sheet, range.values);
    });
    loadSubscriptions();
  } catch (err) {
    showError("Formatting failed: " + err.message);
  }
}

function showError(msg) {
  var errBox = document.getElementById('error-box');
  errBox.innerText = msg;
  errBox.style.display = 'block';
  document.getElementById('loader').style.display = 'none';
}
