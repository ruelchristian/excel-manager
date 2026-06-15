var allSubscriptions = [];
var filteredSubscriptions = [];
var currentStatusFilter = "all";
var currentPage = 1;
var itemsPerPage = 10;
var currentListType = "monthly"; // 'monthly' or 'annual'

var startPicker, endPicker;

// Initialize Office.js
Office.onReady(function (info) {
  // Initialize Flatpickr date pickers
  startPicker = flatpickr("#form-start-date", {
    dateFormat: "Y-m-d",
    allowInput: true
  });
  endPicker = flatpickr("#form-end-date", {
    dateFormat: "Y-m-d",
    allowInput: true
  });

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
  
  if (tabName === 'monthly') {
    currentListType = 'monthly';
    document.getElementById('tab-btn-monthly').classList.add('active');
    document.getElementById('tab-list').classList.add('active');
    document.getElementById('list-subtitle').innerText = "Active sheet: Monthly Master";
    document.getElementById('btn-show-generator').style.display = 'block';
    document.getElementById('search-input').placeholder = "Search by EU, SO, Product, or Sub ID...";
    loadSubscriptions();
  } else if (tabName === 'annual') {
    currentListType = 'annual';
    document.getElementById('tab-btn-annual').classList.add('active');
    document.getElementById('tab-list').classList.add('active');
    document.getElementById('list-subtitle').innerText = "Active sheet: Annual Master";
    document.getElementById('btn-show-generator').style.display = 'none'; // Generator only for Monthly list source
    document.getElementById('search-input').placeholder = "Search by EU, FA/SO, or Product...";
    loadSubscriptions();
  } else {
    // Add/Edit form tab
    document.getElementById('tab-btn-form').classList.add('active');
    document.getElementById('tab-form').classList.add('active');
    
    // Dynamically update status choices
    var statusSelect = document.getElementById('form-status');
    var currentVal = statusSelect.value;
    if (currentListType === 'monthly') {
      statusSelect.innerHTML = 
        '<option value="New">New</option>' +
        '<option value="Renewal">Renewal</option>' +
        '<option value="Complete">Complete</option>' +
        '<option value="Cancelled">Cancelled</option>';
    } else {
      statusSelect.innerHTML = 
        '<option value="Active">Active</option>' +
        '<option value="Complete">Complete</option>' +
        '<option value="Cancelled">Cancelled</option>';
    }
    if (currentVal) statusSelect.value = currentVal;
    
    // Toggle form field groups based on type
    if (currentListType === 'monthly') {
      document.getElementById('monthly-fields-group1').style.display = 'flex';
      document.getElementById('monthly-fields-group2').style.display = 'flex';
      document.getElementById('monthly-fields-month').style.display = 'block';
      document.getElementById('monthly-fields-postatus').style.display = 'block';
      document.getElementById('annual-fields-faso').style.display = 'none';
      
      // Make monthly fields required
      document.getElementById('form-month').setAttribute('required', 'required');
      document.getElementById('form-po-status').setAttribute('required', 'required');
      document.getElementById('form-fa-so').removeAttribute('required');
    } else {
      document.getElementById('monthly-fields-group1').style.display = 'none';
      document.getElementById('monthly-fields-group2').style.display = 'none';
      document.getElementById('monthly-fields-month').style.display = 'none';
      document.getElementById('monthly-fields-postatus').style.display = 'none';
      document.getElementById('annual-fields-faso').style.display = 'block';
      
      // Remove required attributes
      document.getElementById('form-month').removeAttribute('required');
      document.getElementById('form-po-status').removeAttribute('required');
      document.getElementById('form-fa-so').setAttribute('required', 'required');
    }
  }
}

// Helper to parse Excel dates to yyyy-MM-dd format for HTML inputs
function parseExcelDate(val, textVal) {
  if (!val) return "";
  
  // If serial number
  if (typeof val === 'number') {
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
    // 1. Detect and Restructure columns if Monthly (run sortSubscriptions first)
    // 1. Detect and Restructure columns if Monthly (run sortSubscriptions first)
    if (currentListType === 'monthly') {
      var isNeedRestructure = false;
      await Excel.run(async function (context) {
        var sheetName = "ADB MASTER LIST MONTHLY";
        var sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
        sheet.load("nullObject");
        await context.sync();
        if (!sheet.isNullObject) {
          var headerRange = sheet.getRange("A3:L3");
          headerRange.load("values");
          await context.sync();
          var colBHeader = String(headerRange.values[0][1] || '').trim().toUpperCase();
          var colDHeader = String(headerRange.values[0][3] || '').trim().toUpperCase();
          if (colBHeader === "MONTHS LEFT" || (colBHeader === "STATUS" && colDHeader === "MONTHS LEFT")) {
            isNeedRestructure = true;
          }
        }
      });
      
      if (isNeedRestructure) {
        document.getElementById('loader').innerText = "Consolidating monthly master sheet columns...";
        await sortSubscriptions();
      }
    }

    await Excel.run(async function (context) {
      var sheetName = currentListType === 'monthly' ? "ADB MASTER LIST MONTHLY" : "ADB MASTER LIST ANNUAL";
      var sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
      sheet.load("nullObject");
      await context.sync();

      if (sheet.isNullObject) {
        throw new Error('Sheet "' + sheetName + '" not found. Please verify the sheet name.');
      }

      // Find the last row
      var rangeUsed = sheet.getUsedRange();
      var lastRowRange = rangeUsed.getLastRow();
      lastRowRange.load("rowIndex");
      await context.sync();

      var lastRowIndex = lastRowRange.rowIndex + 1; // 1-based row number
      var startRow = currentListType === 'monthly' ? 4 : 5;

      if (lastRowIndex < startRow) {
        allSubscriptions = [];
        document.getElementById('row-count').innerText = "0";
        document.getElementById('loader').style.display = 'none';
        renderList([]);
        return;
      }

      // Read columns (12 columns for Monthly, 6 columns for Annual)
      var rangeColLetter = currentListType === 'monthly' ? "L" : "F";
      var range = sheet.getRange("A" + startRow + ":" + rangeColLetter + lastRowIndex);
      range.load(["values", "text"]);
      await context.sync();

      var values = range.values;
      var text = range.text;
      var records = [];

      for (var i = 0; i < values.length; i++) {
        var rowVal = values[i];
        var rowText = text[i];
        var rowNum = i + startRow;

        if (currentListType === 'monthly') {
          // Ignore empty rows (Column H is EU Name / index 7, Column L is Subscription Name / index 11)
          if (!rowVal[7] && !rowVal[11]) continue;

          records.push({
            rowNum: rowNum,
            month: String(rowVal[0] || '').trim(),
            status: String(rowVal[1] || '').trim(),
            poStatus: String(rowVal[2] || '').trim(),
            monthsLeft: String(rowVal[3] !== null && rowVal[3] !== undefined ? rowVal[3] : '').trim(),
            totalMonths: String(rowVal[4] !== null && rowVal[4] !== undefined ? rowVal[4] : '').trim(),
            startDate: parseExcelDate(rowVal[5], rowText[5]),
            endDate: parseExcelDate(rowVal[6], rowText[6]),
            eu: String(rowVal[7] || '').trim(),
            fa: String(rowVal[8] || '').trim(),
            so: String(rowVal[9] || '').trim(),
            subId: String(rowVal[10] || '').trim(),
            subscription: String(rowVal[11] || '').trim()
          });
        } else {
          // Annual
          if (!rowVal[3] && !rowVal[5]) continue;

          records.push({
            rowNum: rowNum,
            status: String(rowVal[0] || '').trim(),
            startDate: parseExcelDate(rowVal[1], rowText[1]),
            endDate: parseExcelDate(rowVal[2], rowText[2]),
            eu: String(rowVal[3] || '').trim(),
            fa_so: String(rowVal[4] || '').trim(),
            subscription: String(rowVal[5] || '').trim()
          });
        }
      }

      allSubscriptions = records;
      document.getElementById('row-count').innerText = records.length;
      document.getElementById('loader').style.display = 'none';
      filterList(); // Triggers the combined filter & pagination render
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

    if (statusLower === 'new' || statusLower === 'active') {
      cardClass = 'status-new';
      badgeClass = 'badge-new';
      statusLabel = statusLower === 'new' ? 'NEW' : 'ACTIVE';
    } else if (statusLower === 'renewal') {
      cardClass = 'status-renewal';
      badgeClass = 'badge-renewal';
      statusLabel = 'RENEWAL';
    } else if (statusLower === 'complete' || statusLower.indexOf('complete') !== -1 || statusLower === 'closed') {
      cardClass = 'status-completed';
      badgeClass = 'badge-completed';
      statusLabel = statusLower === 'closed' ? 'CLOSED' : 'COMPLETE';
    } else if (statusLower === 'cancelled' || statusLower.indexOf('cancel') !== -1) {
      cardClass = 'status-cancelled';
      badgeClass = 'badge-cancelled';
      statusLabel = 'CANCELLED';
    }

    var card = document.createElement('div');
    card.className = 'card ' + cardClass;
    card.onclick = function() { editRecord(rec); };

    if (currentListType === 'monthly') {
      card.innerHTML = 
        '<div>' +
          '<div class="card-header">' +
            '<div class="card-title">' + rec.eu + '</div>' +
            '<div class="card-badge ' + badgeClass + '">' + statusLabel + '</div>' +
          '</div>' +
          '<div class="card-subtitle">' + rec.subscription + '</div>' +
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
    } else {
      // Annual Layout
      card.innerHTML = 
        '<div>' +
          '<div class="card-header">' +
            '<div class="card-title">' + rec.eu + '</div>' +
            '<div class="card-badge ' + badgeClass + '">' + statusLabel + '</div>' +
          '</div>' +
          '<div class="card-subtitle">' + rec.subscription + '</div>' +
          '<div class="card-details">' +
            '<div><strong>FA/SO:</strong> ' + (rec.fa_so || '-') + '</div>' +
            '<div><strong>Start Date:</strong> ' + (rec.startDate || '-') + '</div>' +
            '<div><strong>End Date:</strong> ' + (rec.endDate || '-') + '</div>' +
          '</div>' +
        '</div>' +
        '<div class="card-footer">' +
          '<span>Billing: <strong>Annual</strong></span>' +
          '<span>Expiry: <strong>' + (rec.endDate || 'N/A') + '</strong></span>' +
        '</div>';
    }
    
    container.appendChild(card);
  });
}

// Filter cards by search input and status pills
function filterList() {
  var query = document.getElementById('search-input').value.toLowerCase();
  
  filteredSubscriptions = allSubscriptions.filter(function(rec) {
    // 1. Status Filter check
    var matchesStatus = true;
    if (currentStatusFilter !== 'all') {
      matchesStatus = rec.status.toLowerCase() === currentStatusFilter;
    }
    
    // 2. Search Text check
    var matchesSearch = false;
    if (currentListType === 'monthly') {
      matchesSearch = rec.eu.toLowerCase().indexOf(query) !== -1 ||
                      rec.so.toLowerCase().indexOf(query) !== -1 ||
                      rec.subId.toLowerCase().indexOf(query) !== -1 ||
                      rec.subscription.toLowerCase().indexOf(query) !== -1 ||
                      rec.fa.toLowerCase().indexOf(query) !== -1;
    } else {
      matchesSearch = rec.eu.toLowerCase().indexOf(query) !== -1 ||
                      rec.fa_so.toLowerCase().indexOf(query) !== -1 ||
                      rec.subscription.toLowerCase().indexOf(query) !== -1;
    }
                        
    return matchesStatus && matchesSearch;
  });
  
  currentPage = 1; // Reset to page 1 when filter changes
  renderCurrentPage();
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

  if (currentListType === 'monthly') {
    document.getElementById('form-months-left').value = rec.monthsLeft;
    document.getElementById('form-total-months').value = rec.totalMonths;
    document.getElementById('form-sub-id').value = rec.subId;
    document.getElementById('form-fa').value = rec.fa;
    document.getElementById('form-so').value = rec.so;
    document.getElementById('form-month').value = rec.month ? rec.month : '';
    document.getElementById('form-po-status').value = rec.poStatus ? rec.poStatus : 'N/A';
  } else {
    document.getElementById('form-fa-so').value = rec.fa_so;
  }

  if (startPicker) startPicker.setDate(rec.startDate);
  else document.getElementById('form-start-date').value = rec.startDate;

  if (endPicker) endPicker.setDate(rec.endDate);
  else document.getElementById('form-end-date').value = rec.endDate;

  document.getElementById('form-status').value = rec.status ? rec.status : 'Complete';

  document.getElementById('save-btn').innerText = 'Update (Row ' + rec.rowNum + ')';
  document.getElementById('cancel-edit').style.display = 'block';
  var deleteBtn = document.getElementById('delete-btn');
  if (deleteBtn) deleteBtn.style.display = 'block';

  switchTab('form');
}

// Reset Form
function resetForm() {
  document.getElementById('record-form').reset();
  if (startPicker) startPicker.clear();
  if (endPicker) endPicker.clear();
  document.getElementById('form-row-num').value = '';
  document.getElementById('save-btn').innerText = 'Save Subscription';
  document.getElementById('save-btn').style.display = 'block';
  document.getElementById('cancel-edit').style.display = 'none';
  document.getElementById('cancel-edit').style.opacity = '1';
  document.getElementById('cancel-edit').disabled = false;
  var deleteBtn = document.getElementById('delete-btn');
  if (deleteBtn) {
    deleteBtn.style.display = 'none';
    deleteBtn.style.opacity = '1';
    deleteBtn.disabled = false;
  }
  var confirmBox = document.getElementById('delete-confirm-box');
  if (confirmBox) confirmBox.style.display = 'none';
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
    eu: document.getElementById('form-eu').value,
    subscription: document.getElementById('form-subscription').value,
    startDate: document.getElementById('form-start-date').value,
    endDate: document.getElementById('form-end-date').value,
    status: document.getElementById('form-status').value
  };

  if (currentListType === 'monthly') {
    formData.month = document.getElementById('form-month').value;
    formData.poStatus = document.getElementById('form-po-status').value;
    formData.monthsLeft = document.getElementById('form-months-left').value;
    formData.totalMonths = document.getElementById('form-total-months').value;
    formData.fa = document.getElementById('form-fa').value;
    formData.so = document.getElementById('form-so').value;
    formData.subId = document.getElementById('form-sub-id').value;
  } else {
    formData.fa_so = document.getElementById('form-fa-so').value;
  }

  try {
    await Excel.run(async function (context) {
      var sheetName = currentListType === 'monthly' ? "ADB MASTER LIST MONTHLY" : "ADB MASTER LIST ANNUAL";
      var sheet = context.workbook.worksheets.getItem(sheetName);
      var range;
      var startRow = currentListType === 'monthly' ? 4 : 5;

      if (rowNum) {
        // Edit existing
        var colLetter = currentListType === 'monthly' ? "L" : "F";
        range = sheet.getRange("A" + rowNum + ":" + colLetter + rowNum);
      } else {
        // Append at end
        var rangeUsed = sheet.getUsedRange();
        var lastRowRange = rangeUsed.getLastRow();
        lastRowRange.load("rowIndex");
        await context.sync();
        
        var nextRow = lastRowRange.rowIndex + 2; // Index is 0-based
        if (nextRow < startRow) nextRow = startRow;
        var colLetter = currentListType === 'monthly' ? "L" : "F";
        range = sheet.getRange("A" + nextRow + ":" + colLetter + nextRow);
      }

      if (currentListType === 'monthly') {
        range.numberFormat = [["@", "@", "@", "General", "General", "[$-809]dddd\\,d\\ mmmm\\ yyyy;@", "[$-809]dddd\\,d\\ mmmm\\ yyyy;@", "@", "@", "@", "@", "@"]];
        range.values = [[
          formData.month,
          formData.status,
          formData.poStatus,
          formData.monthsLeft ? parseInt(formData.monthsLeft, 10) : 0,
          formData.totalMonths ? parseInt(formData.totalMonths, 10) : 0,
          formData.startDate || "",
          formData.endDate || "",
          formData.eu,
          formData.fa,
          formData.so,
          formData.subId,
          formData.subscription
        ]];
      } else {
        range.numberFormat = [["@", "[$-809]dddd\\,d\\ mmmm\\ yyyy;@", "[$-809]dddd\\,d\\ mmmm\\ yyyy;@", "@", "@", "@"]];
        range.values = [[
          formData.status,
          formData.startDate,
          formData.endDate,
          formData.eu,
          formData.fa_so,
          formData.subscription
        ]];
      }

      await context.sync();
    });

    // Run custom sort and color formatting
    await sortSubscriptions();
    
    resetForm();
    switchTab(currentListType);
    document.getElementById('save-btn').disabled = false;
  } catch (err) {
    showError("Failed to save record: " + err.message);
    document.getElementById('save-btn').disabled = false;
    document.getElementById('save-btn').innerText = 'Save Subscription';
  }
}

// Show inline delete confirmation panel
function showDeleteConfirm() {
  document.getElementById('save-btn').style.display = 'none';
  document.getElementById('cancel-edit').style.display = 'none';
  document.getElementById('delete-btn').style.display = 'none';
  
  var confirmBox = document.getElementById('delete-confirm-box');
  if (confirmBox) confirmBox.style.display = 'block';
}

// Hide inline delete confirmation panel and restore buttons
function hideDeleteConfirm() {
  var confirmBox = document.getElementById('delete-confirm-box');
  if (confirmBox) confirmBox.style.display = 'none';

  document.getElementById('save-btn').style.display = 'block';
  document.getElementById('cancel-edit').style.display = 'block';
  document.getElementById('delete-btn').style.display = 'block';
}

// Delete form data from Excel Sheet (called after inline confirmation)
async function deleteRecord() {
  var rowNumVal = document.getElementById('form-row-num').value;
  if (!rowNumVal) return;
  var rowNum = parseInt(rowNumVal, 10);

  document.getElementById('error-box').style.display = 'none';
  
  var confirmBox = document.getElementById('delete-confirm-box');
  var confirmButtons = confirmBox ? confirmBox.getElementsByTagName('button') : [];
  for (var i = 0; i < confirmButtons.length; i++) {
    confirmButtons[i].disabled = true;
    if (confirmButtons[i].classList.contains('btn-danger')) {
      confirmButtons[i].innerText = 'Deleting...';
    }
  }

  try {
    await Excel.run(async function (context) {
      var sheetName = currentListType === 'monthly' ? "ADB MASTER LIST MONTHLY" : "ADB MASTER LIST ANNUAL";
      var sheet = context.workbook.worksheets.getItem(sheetName);
      
      // Delete the entire row in Excel
      var range = sheet.getRange("A" + rowNum + ":A" + rowNum);
      var entireRow = range.getEntireRow();
      entireRow.delete(Excel.DeleteShiftDirection.up);

      await context.sync();
    });

    // Run custom sort and color formatting
    await sortSubscriptions();
    
    resetForm();
    switchTab(currentListType);
    showSuccess("Subscription entry deleted successfully!");
  } catch (err) {
    showError("Failed to delete record: " + err.message);
    // Re-enable confirmation box buttons on error
    for (var i = 0; i < confirmButtons.length; i++) {
      confirmButtons[i].disabled = false;
      if (confirmButtons[i].classList.contains('btn-danger')) {
        confirmButtons[i].innerText = 'Yes, Delete';
      }
    }
  }
}

// In-Memory Custom Sort (weighting Statuses) and formatting colors
async function sortSubscriptions() {
  try {
    await Excel.run(async function (context) {
      var sheetName = currentListType === 'monthly' ? "ADB MASTER LIST MONTHLY" : "ADB MASTER LIST ANNUAL";
      var sheet = context.workbook.worksheets.getItem(sheetName);
      var rangeUsed = sheet.getUsedRange();
      var lastRowRange = rangeUsed.getLastRow();
      lastRowRange.load("rowIndex");
      await context.sync();

      var lastRowIndex = lastRowRange.rowIndex + 1;
      var startRow = currentListType === 'monthly' ? 4 : 5;
      if (lastRowIndex < startRow) return;

      var values;

      if (currentListType === 'monthly') {
        var headerRange = sheet.getRange("A3:L3");
        headerRange.load("values");
        await context.sync();

        var headerValues = headerRange.values[0];
        var colBHeader = String(headerValues[1] || '').trim().toUpperCase();
        var colCHeader = String(headerValues[2] || '').trim().toUpperCase();
        var colDHeader = String(headerValues[3] || '').trim().toUpperCase();

        var is12Column = (colBHeader === "STATUS" && colDHeader === "MONTHS LEFT");
        var is10Column = (colBHeader === "MONTHS LEFT");
        var is9Column = (colBHeader === "STATUS" && colDHeader === "R/T");

        if (is10Column || is9Column) {
          var oldValues;
          if (is10Column) {
            var oldRange = sheet.getRange("A4:J" + lastRowIndex);
            oldRange.load("values");
            await context.sync();
            oldValues = oldRange.values;
          } else {
            var oldRange = sheet.getRange("A4:I" + lastRowIndex);
            oldRange.load("values");
            await context.sync();
            oldValues = oldRange.values;
          }

          var restructuredValues = [];
          for (var i = 0; i < oldValues.length; i++) {
            var month = "", status = "", poStatus = "";
            var monthsLeft = 0, totalMonths = 0;
            var startDate = "", endDate = "";
            var eu = "", fa = "", so = "";
            var subId = "", subName = "";

            if (is10Column) {
              var rawStatus = oldValues[i][0];
              monthsLeft = oldValues[i][1];
              totalMonths = oldValues[i][2];
              startDate = oldValues[i][3];
              endDate = oldValues[i][4];
              eu = oldValues[i][5];
              fa = oldValues[i][6];
              so = oldValues[i][7];
              subId = oldValues[i][8];
              subName = oldValues[i][9];

              var parsed = parseStatusFieldJS(rawStatus, endDate);
              month = parsed.month;
              status = parsed.status;
              poStatus = parsed.poStatus;
            } else {
              // 9-column layout
              month = oldValues[i][0];
              status = oldValues[i][1];
              poStatus = oldValues[i][2];
              
              var rt = cleanRTValueJS(oldValues[i][3]);
              var rtParts = rt.split('/');
              monthsLeft = rtParts[0] ? parseInt(rtParts[0].trim(), 10) : 0;
              totalMonths = rtParts[1] ? parseInt(rtParts[1].trim(), 10) : 0;

              var period = cleanPeriodValueJS(oldValues[i][4]);
              var periodParts = period.trim().split(/\s+/);
              var startStr = periodParts[0] || '';
              var endStr = periodParts[1] || '';
              startDate = parseDDMMMYY(startStr);
              endDate = parseDDMMMYY(endStr);

              eu = oldValues[i][5];

              var faso = String(oldValues[i][6] || '').trim();
              var fasoParts = faso.split(/\s{2,}/);
              if (fasoParts.length >= 2) {
                fa = fasoParts[0].trim();
                so = fasoParts[1].trim();
              } else {
                var soIndex = faso.indexOf('SOUNI');
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

          var newHeaders = [
            "MONTH", "STATUS", "PO STATUS", "MONTHS LEFT", "TOTAL MONTHS",
            "START DATE", "END DATE", "EU", "FA", "SO", "SUBSCRIPTION ID", "SUBSCRIPTION"
          ];
          
          // Clear columns A to L below header (contents, formats, validation) to remove date formats
          var clearRangeData = sheet.getRange("A4:L" + lastRowIndex);
          clearRangeData.clear(Excel.ClearApplyTo.all);
          
          // Clear header contents only (preserving header style)
          var clearRangeHeader = sheet.getRange("A3:L3");
          clearRangeHeader.clear(Excel.ClearApplyTo.contents);
          await context.sync();

          var newHeaderRange = sheet.getRange("A3:L3");
          newHeaderRange.values = [newHeaders];
          await context.sync();

          // Delete comments/notes from headers (D3, E3, G3)
          var deleteCommentJS = async function(cell) {
            try {
              var existing = context.workbook.comments.getItemByCell(cell);
              existing.delete();
              await context.sync();
            } catch (e) {}
          };
          await deleteCommentJS(sheet.getRange("D3"));
          await deleteCommentJS(sheet.getRange("E3"));
          await deleteCommentJS(sheet.getRange("G3"));

          values = restructuredValues;
        } else {
          var range = sheet.getRange("A4:L" + lastRowIndex);
          range.load("values");
          await context.sync();
          values = range.values;
        }
      } else {
        // Annual Master
        var range = sheet.getRange("A" + startRow + ":F" + lastRowIndex);
        range.load("values");
        await context.sync();
        values = range.values;

        // Clean up and standardize statuses in memory for Annual Master
        for (var i = 0; i < values.length; i++) {
          var rawStatus = String(values[i][0] || '').trim();
          var statusLower = rawStatus.toLowerCase();
          
          var hasData = false;
          for (var col = 1; col < 6; col++) {
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
            values[i][0] = "";
          }
        }
      }

      var statusWeights = {
        'new': 1,
        'renewal': 2,
        'active': 1,
        'complete': currentListType === 'monthly' ? 3 : 2,
        'cancelled': currentListType === 'monthly' ? 4 : 3
      };

      // Custom multi-tier sort
      values.sort(function(a, b) {
        if (currentListType === 'monthly') {
          // 1. Status (Index 1)
          var statusA = String(a[1] || '').trim().toLowerCase();
          var statusB = String(b[1] || '').trim().toLowerCase();
          var weightStatusA = statusWeights[statusA] || 5;
          var weightStatusB = statusWeights[statusB] || 5;
          if (weightStatusA !== weightStatusB) return weightStatusA - weightStatusB;

          var getTime = function(val) {
            if (typeof val === 'number') return val;
            if (!val) return 0;
            var parsed = new Date(val).getTime();
            return isNaN(parsed) ? 0 : parsed;
          };

          // 2. Start Date (Index 5) - Descending
          var timeA = getTime(a[5]);
          var timeB = getTime(b[5]);
          if (timeA !== timeB) {
            if (timeA === 0) return 1;
            if (timeB === 0) return -1;
            return timeB - timeA;
          }
          
          // 3. EU Name (Index 7)
          var nameA = String(a[7] || '').trim().toLowerCase();
          var nameB = String(b[7] || '').trim().toLowerCase();
          return nameA.localeCompare(nameB);
        } else {
          // Annual
          // 1. Status (Index 0)
          var statusA = String(a[0] || '').trim().toLowerCase();
          var statusB = String(b[0] || '').trim().toLowerCase();
          var weightStatusA = statusWeights[statusA] || 5;
          var weightStatusB = statusWeights[statusB] || 5;
          if (weightStatusA !== weightStatusB) return weightStatusA - weightStatusB;

          var getTime = function(val) {
            if (typeof val === 'number') return val;
            if (!val) return 0;
            var parsed = new Date(val).getTime();
            return isNaN(parsed) ? 0 : parsed;
          };

          // 2. Start Date (Index 1) - Descending
          var timeA = getTime(a[1]);
          var timeB = getTime(b[1]);
          if (timeA !== timeB) {
            if (timeA === 0) return 1;
            if (timeB === 0) return -1;
            return timeB - timeA;
          }
          
          // 3. EU Name (Index 3)
          var nameA = String(a[3] || '').trim().toLowerCase();
          var nameB = String(b[3] || '').trim().toLowerCase();
          return nameA.localeCompare(nameB);
        }
      });

      // Overwrite cells
      var colLetter = currentListType === 'monthly' ? "L" : "F";

      if (currentListType === 'monthly') {
        // Clear any existing validations in columns A to L
        var clearValRange = sheet.getRange("A4:L10000");
        clearValRange.dataValidation.clear();

        // Set number formats BEFORE writing values to prevent Excel date auto-conversion
        var textRangeAC = sheet.getRange("A4:C10000");
        textRangeAC.numberFormat = "@";
        var numRangeDE = sheet.getRange("D4:E10000");
        numRangeDE.numberFormat = "General";
        var dateRangeFG = sheet.getRange("F4:G10000");
        dateRangeFG.numberFormat = "[$-809]dddd\\,d\\ mmmm\\ yyyy;@";
        var textRangeHL = sheet.getRange("H4:L10000");
        textRangeHL.numberFormat = "@";
        await context.sync();

        if (lastRowIndex >= 4) {
          var statusRange = sheet.getRange("B4:B" + lastRowIndex);
          statusRange.dataValidation.rule = {
            list: {
              inCellDropDown: true,
              source: "New,Renewal,Complete,Cancelled"
            }
          };
          
          var poRange = sheet.getRange("C4:C" + lastRowIndex);
          poRange.dataValidation.rule = {
            list: {
              inCellDropDown: true,
              source: "PO Done,PO Pending"
            }
          };
          await context.sync();
        }

        // Clear everything below the active rows to prevent infinite scroll/formatting artifacts
        if (lastRowIndex < 10000) {
          var clearBelowRange = sheet.getRange("A" + (lastRowIndex + 1) + ":L10000");
          clearBelowRange.clear(Excel.ClearApplyTo.all);
          await context.sync();
        }
      } else {
        // Annual Master
        // Set number formats BEFORE writing values to prevent Excel date auto-conversion
        var textRangeA = sheet.getRange("A5:A10000");
        textRangeA.numberFormat = "@";
        var dateRangeBC = sheet.getRange("B5:C10000");
        dateRangeBC.numberFormat = "[$-809]dddd\\,d\\ mmmm\\ yyyy;@";
        var textRangeDF = sheet.getRange("D5:F10000");
        textRangeDF.numberFormat = "@";
        await context.sync();

        if (lastRowIndex < 10000) {
          var clearBelowRange = sheet.getRange("A" + (lastRowIndex + 1) + ":F10000");
          clearBelowRange.clear(Excel.ClearApplyTo.all);
          await context.sync();
        }
      }
      
      var writeRange = sheet.getRange("A" + startRow + ":" + colLetter + lastRowIndex);
      writeRange.values = values;
      await context.sync();

      // Re-apply background colors
      await formatSheetColorsDirect(sheet, values);
    });
  } catch (err) {
    showError("Sorting failed: " + err.message);
  }
}

// Background row pastel coloring helper
async function formatSheetColorsDirect(sheet, values) {
  var colorNew = '#fef08a';        // Soft Yellow
  var colorRenewal = '#bbf7d0';    // Soft Green
  var colorComplete = '#bfdbfe';   // Soft Blue
  var colorCancelled = '#fecaca';  // Soft Red
  var colorDefault = '#ffffff';    // White

  var startRow = currentListType === 'monthly' ? 4 : 5;
  var colLetter = currentListType === 'monthly' ? "L" : "F";

  for (var i = 0; i < values.length; i++) {
    var status = currentListType === 'monthly'
      ? String(values[i][1] || '').trim().toLowerCase()
      : String(values[i][0] || '').trim().toLowerCase();
      
    var rowColor = colorDefault;

    if (status === 'new' || status === 'active') {
      rowColor = colorNew;
    } else if (status === 'renewal') {
      rowColor = colorRenewal;
    } else if (status === 'complete' || status.indexOf('complete') !== -1 || status === 'closed') {
      rowColor = colorComplete;
    } else if (status === 'cancelled' || status.indexOf('cancel') !== -1) {
      rowColor = colorCancelled;
    }

    var rowNum = i + startRow;
    var rowRange = sheet.getRange("A" + rowNum + ":" + colLetter + rowNum);
    rowRange.format.fill.color = rowColor;
  }

  // Format header/title row and set alignment to center
  if (currentListType === 'monthly') {
    var headerRange = sheet.getRange("A3:L3");
    headerRange.format.fill.color = "#FFC000";
    headerRange.format.font.bold = true;
    
    var dataRange = sheet.getRange("A3:L10000");
    dataRange.format.horizontalAlignment = "Center";
  } else {
    var headerRange = sheet.getRange("A4:F4");
    headerRange.format.fill.color = "#FFC000";
    headerRange.format.font.bold = true;
    
    var dataRange = sheet.getRange("A4:F10000");
    dataRange.format.horizontalAlignment = "Center";
  }

  await sheet.context.sync();
}

// Trigger row formatting manually (now standardizes, sorts, and colors the sheet)
async function runFormatter() {
  document.getElementById('loader').style.display = 'block';
  try {
    await sortSubscriptions();
    await loadSubscriptions();
  } catch (err) {
    showError("Formatting failed: " + err.message);
  }
}

function showError(msg) {
  showMessage(msg, true);
}

function showSuccess(msg) {
  showMessage(msg, false);
}

function showMessage(msg, isError) {
  var box = document.getElementById('error-box');
  if (box) {
    box.innerText = msg;
    box.style.display = 'block';
    if (isError) {
      box.style.backgroundColor = '#fef2f2';
      box.style.color = '#b91c1c';
      box.style.border = '1px solid rgba(239, 68, 68, 0.15)';
    } else {
      box.style.backgroundColor = '#ecfdf5';
      box.style.color = '#047857';
      box.style.border = '1px solid rgba(16, 185, 129, 0.15)';
    }
  }
  var loader = document.getElementById('loader');
  if (loader) {
    loader.style.display = 'none';
  }
}

function normalizeDateStr(str) {
  if (!str) return "";
  var clean = str.replace(/expires on/i, "").trim();
  var parts = clean.split(/[\/\-]/);
  if (parts.length === 3) {
    var m = parseInt(parts[0], 10);
    var d = parseInt(parts[1], 10);
    var y = parseInt(parts[2], 10);
    return m + "/" + d + "/" + y;
  }
  return clean.toLowerCase();
}

// =========================================================================
// PAGINATION & FILTERING HELPERS
// =========================================================================

// Render the subset of filtered items for the current page
function renderCurrentPage() {
  var startIndex = (currentPage - 1) * itemsPerPage;
  var endIndex = startIndex + itemsPerPage;
  var pageItems = filteredSubscriptions.slice(startIndex, endIndex);
  
  renderList(pageItems);
  updatePaginationControls();
}

// Handle clicking a status filter pill
function setStatusFilter(status) {
  currentStatusFilter = status;
  
  document.querySelectorAll('.pill').forEach(function(pill) {
    pill.classList.remove('active');
  });
  
  var targetSelector = '.pill';
  if (status === 'all') targetSelector = '.pill:not([class*="status-"])';
  else if (status === 'new') targetSelector = '.pill.status-new';
  else if (status === 'renewal') targetSelector = '.pill.status-renewal';
  else if (status === 'complete') targetSelector = '.pill.status-completed';
  else if (status === 'cancelled') targetSelector = '.pill.status-cancelled';
  
  var targetPill = document.querySelector(targetSelector);
  if (targetPill) {
    targetPill.classList.add('active');
  }
  
  filterList();
}

function prevPage() {
  if (currentPage > 1) {
    currentPage--;
    renderCurrentPage();
  }
}

function nextPage() {
  var totalPages = Math.ceil(filteredSubscriptions.length / itemsPerPage) || 1;
  if (currentPage < totalPages) {
    currentPage++;
    renderCurrentPage();
  }
}

function updatePaginationControls() {
  var totalPages = Math.ceil(filteredSubscriptions.length / itemsPerPage) || 1;
  
  if (currentPage > totalPages) currentPage = totalPages;
  if (currentPage < 1) currentPage = 1;
  
  document.getElementById('page-indicator').innerText = `Page ${currentPage} of ${totalPages}`;
  document.getElementById('prev-page').disabled = (currentPage === 1);
  document.getElementById('next-page').disabled = (currentPage === totalPages);
}

// =========================================================================
// SHEET GENERATOR SECTION
// =========================================================================

function toggleGenerator() {
  var sec = document.getElementById('generator-section');
  if (sec.style.display === 'none') {
    sec.style.display = 'block';
  } else {
    sec.style.display = 'none';
  }
}

async function generateMonthlySheet() {
  var selectedMonth = document.getElementById('gen-month').value;
  var selectedYear = parseInt(document.getElementById('gen-year').value, 10);
  if (!selectedYear) {
    showError("Please enter a valid year.");
    return;
  }

  document.getElementById('generator-section').style.display = 'none';
  document.getElementById('loader').style.display = 'block';
  document.getElementById('loader').innerText = "Generating monthly sheet...";

  try {
    await Excel.run(async function (context) {
      // 1. Determine sheet name
      var monthUpper = selectedMonth.toUpperCase();
      var targetSheetName = "ADB_" + monthUpper;
      if (selectedYear !== 2025) {
        targetSheetName = "ADB_" + monthUpper + " " + selectedYear;
      }

      // 1b. Fetch PO numbers from the previous month's sheet
      var months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
      var currentMonthIndex = months.findIndex(m => m.toLowerCase() === selectedMonth.toLowerCase());
      var prevMonthIndex = currentMonthIndex - 1;
      var prevYear = selectedYear;
      if (prevMonthIndex < 0) {
        prevMonthIndex = 11;
        prevYear = selectedYear - 1;
      }
      var prevMonthName = months[prevMonthIndex];
      var prevSheetName = "ADB_" + prevMonthName.toUpperCase();
      if (prevYear !== 2025) {
        prevSheetName = "ADB_" + prevMonthName.toUpperCase() + " " + prevYear;
      }

      var prevSheet = context.workbook.worksheets.getItemOrNullObject(prevSheetName);
      prevSheet.load("nullObject");
      await context.sync();

      var prevPOMap = {};
      if (!prevSheet.isNullObject) {
        // Load headers first to see how many columns we have and what their names are
        var headerRange = prevSheet.getRange("B6:H6");
        headerRange.load("values");
        await context.sync();

        var headers = headerRange.values[0];
        var idxLicense = -1;
        var idxDate = -1;
        var idxOldPo = -1;
        var idxNewPo = -1;
        var idxPo = -1;

        for (var i = 0; i < headers.length; i++) {
          var h = String(headers[i] || '').trim().toUpperCase();
          if (h === "LICENSE") idxLicense = i;
          else if (h === "DATE") idxDate = i;
          else if (h === "OLD PO NUMBER") idxOldPo = i;
          else if (h === "NEW PO NUMBER") idxNewPo = i;
          else if (h === "PO NUMBER") idxPo = i;
        }

        if (idxLicense !== -1 && idxDate !== -1) {
          var prevRangeUsed = prevSheet.getUsedRange();
          var prevLastRowRange = prevRangeUsed.getLastRow();
          prevLastRowRange.load("rowIndex");
          await context.sync();

          var prevLastRow = prevLastRowRange.rowIndex + 1;
          if (prevLastRow >= 7) {
            var prevRange = prevSheet.getRange("B7:H" + prevLastRow);
            prevRange.load("values");
            await context.sync();

            var prevValues = prevRange.values;
            for (var i = 0; i < prevValues.length; i++) {
              var row = prevValues[i];
              var lName = String(row[idxLicense] || '').trim();
              var dVal = String(row[idxDate] || '').trim();
              
              var oldPoVal = idxOldPo !== -1 ? String(row[idxOldPo] || '').trim() : "";
              var newPoVal = idxNewPo !== -1 ? String(row[idxNewPo] || '').trim() : "";
              var poVal = idxPo !== -1 ? String(row[idxPo] || '').trim() : "";

              var inheritedPo = newPoVal || oldPoVal || poVal;

              if (lName && dVal) {
                var day = null;
                var dateMatch = dVal.match(/Expires on \d+\/(\d+)\/\d+/);
                if (dateMatch) {
                  day = parseInt(dateMatch[1], 10);
                } else {
                  var cleanDate = dVal.replace("Expires on ", "").trim();
                  var parts = cleanDate.split("/");
                  if (parts.length >= 2) {
                    day = parseInt(parts[1], 10);
                  }
                }

                if (day !== null && !isNaN(day)) {
                  var key = lName.toLowerCase() + "|" + day;
                  if (inheritedPo) {
                    prevPOMap[key] = inheritedPo;
                  }
                }
              }
            }
          }
        }
      }

      // 2. Fetch all Monthly Master subscriptions
      var monthlySheet = context.workbook.worksheets.getItem("ADB MASTER LIST MONTHLY");
      var monthlyRangeUsed = monthlySheet.getUsedRange();
      var monthlyLastRowRange = monthlyRangeUsed.getLastRow();
      monthlyLastRowRange.load("rowIndex");
      await context.sync();

      var monthlyLastRow = monthlyLastRowRange.rowIndex + 1;
      var monthlyValues = [];
      var monthlyText = [];
      if (monthlyLastRow >= 4) {
        var rangeM = monthlySheet.getRange("A4:L" + monthlyLastRow);
        rangeM.load(["values", "text"]);
        await context.sync();
        monthlyValues = rangeM.values;
        monthlyText = rangeM.text;
      }

      // 3. Fetch all Annual Master subscriptions
      var annualSheet = context.workbook.worksheets.getItem("ADB MASTER LIST ANNUAL");
      var annualRangeUsed = annualSheet.getUsedRange();
      var annualLastRowRange = annualRangeUsed.getLastRow();
      annualLastRowRange.load("rowIndex");
      await context.sync();

      var annualLastRow = annualLastRowRange.rowIndex + 1;
      var annualValues = [];
      var annualText = [];
      if (annualLastRow >= 5) {
        var rangeA = annualSheet.getRange("A5:F" + annualLastRow);
        rangeA.load(["values", "text"]);
        await context.sync();
        annualValues = rangeA.values;
        annualText = rangeA.text;
      }

      // Helper to check if a date string/serial belongs to selected month
      var isDateInMonth = function(val, monthName) {
        if (!val) return false;
        var date = null;
        if (typeof val === 'number') {
          date = new Date(Math.round((val - 25569) * 86400 * 1000));
        } else {
          date = new Date(val);
        }
        if (isNaN(date.getTime())) return false;
        
        var mNames = ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december'];
        var dateMonth = mNames[date.getMonth()];
        return dateMonth === monthName.toLowerCase();
      };

      // Helper to format date as "Expires on MM/DD/YYYY"
      var formatExpiryDate = function(val) {
        if (!val) return "";
        var date = null;
        if (typeof val === 'number') {
          date = new Date(Math.round((val - 25569) * 86400 * 1000));
        } else {
          date = new Date(val);
        }
        if (isNaN(date.getTime())) return String(val);
        var mm = String(date.getMonth() + 1).padStart(2, '0');
        var dd = String(date.getDate()).padStart(2, '0');
        var yyyy = date.getFullYear();
        return "Expires on " + mm + "/" + dd + "/" + yyyy;
      };

      // 4. Filter active subscriptions for the target month
      var filteredActive = [];

      // Process Monthly Master (Filter by Month column A)
      for (var i = 0; i < monthlyValues.length; i++) {
        var row = monthlyValues[i];
        if (!row[7] && !row[11]) continue; // Skip empty
        
        var mCol = String(row[0] || '').trim().toLowerCase();
        if (mCol === selectedMonth.toLowerCase()) {
          filteredActive.push({
            license: String(row[11] || '').trim(),
            billing: "Monthly",
            endDate: row[6], // End Date
            status: String(row[1] || '').trim(),
            so: String(row[9] || '').trim(),
            fa: String(row[8] || '').trim(),
            poStatus: String(row[2] || '').trim()
          });
        }
      }

      // Process Annual Master (Filter by End Date month name)
      for (var i = 0; i < annualValues.length; i++) {
        var row = annualValues[i];
        if (!row[3] && !row[5]) continue;
        
        var endDateVal = row[2];
        if (isDateInMonth(endDateVal, selectedMonth)) {
          filteredActive.push({
            license: String(row[5] || '').trim(),
            billing: "Annual",
            endDate: endDateVal,
            status: String(row[0] || '').trim(),
            so: "",
            fa: String(row[4] || '').trim(),
            poStatus: ""
          });
        }
      }

      // 5. Get or Create Monthly Sheet
      var targetSheet = context.workbook.worksheets.getItemOrNullObject(targetSheetName);
      targetSheet.load("nullObject");
      await context.sync();

      var existingRowsMap = {};
      var targetExists = !targetSheet.isNullObject;

      if (!targetExists) {
        targetSheet = context.workbook.worksheets.add(targetSheetName);
        await context.sync();
      }

      // Always write/verify headers in row 6 (B6:H6) to support the OLD PO / NEW PO layout and style/color them beautifully
      var headerRange = targetSheet.getRange("B6:H6");
      headerRange.values = [["LICENSE", "QTY", "DATE", "SUBSRIPTION", "OLD PO NUMBER", "NEW PO NUMBER", "STATUS"]];
      headerRange.format.font.bold = true;
      headerRange.format.font.color = "#FFFFFF"; // White text
      headerRange.format.fill.color = "#4472C4"; // Medium blue accent color
      await context.sync();

      if (targetExists) {
        // Read existing headers dynamically to handle older/newer columns safely
        var tHeaderRange = targetSheet.getRange("B6:H6");
        tHeaderRange.load("values");
        await context.sync();

        var tHeaders = tHeaderRange.values[0];
        var tIdxLicense = -1;
        var tIdxQty = -1;
        var tIdxDate = -1;
        var tIdxOldPo = -1;
        var tIdxNewPo = -1;
        var tIdxPo = -1;
        var tIdxStatus = -1;

        for (var i = 0; i < tHeaders.length; i++) {
          var h = String(tHeaders[i] || '').trim().toUpperCase();
          if (h === "LICENSE") tIdxLicense = i;
          else if (h === "QTY") tIdxQty = i;
          else if (h === "DATE") tIdxDate = i;
          else if (h === "OLD PO NUMBER") tIdxOldPo = i;
          else if (h === "NEW PO NUMBER") tIdxNewPo = i;
          else if (h === "PO NUMBER") tIdxPo = i;
          else if (h === "STATUS") tIdxStatus = i;
        }

        if (tIdxLicense !== -1 && tIdxDate !== -1) {
          var tRangeUsed = targetSheet.getUsedRange();
          var tLastRowRange = tRangeUsed.getLastRow();
          tLastRowRange.load("rowIndex");
          await context.sync();

          var tLastRow = tLastRowRange.rowIndex + 1;
          if (tLastRow >= 7) {
            var tRange = targetSheet.getRange("B7:H" + tLastRow);
            tRange.load("values");
            await context.sync();
            
            var tValues = tRange.values;
            for (var i = 0; i < tValues.length; i++) {
              var row = tValues[i];
              var lName = String(row[tIdxLicense] || '').trim();
              var dVal = String(row[tIdxDate] || '').trim();
              if (lName) {
                var normDate = normalizeDateStr(dVal);
                var key = (lName + "|" + normDate).toLowerCase();
                
                var oldPoVal = tIdxOldPo !== -1 ? String(row[tIdxOldPo] || '').trim() : "";
                var newPoVal = tIdxNewPo !== -1 ? String(row[tIdxNewPo] || '').trim() : "";
                var poVal = tIdxPo !== -1 ? String(row[tIdxPo] || '').trim() : "";
                var qtyVal = tIdxQty !== -1 ? String(row[tIdxQty] || '1 Licenses').trim() : "1 Licenses";
                var statusVal = tIdxStatus !== -1 ? String(row[tIdxStatus] || '').trim() : "";

                existingRowsMap[key] = {
                  qty: qtyVal,
                  oldPo: oldPoVal || poVal,
                  newPo: newPoVal,
                  status: statusVal
                };
              }
            }
          }
        }
      }

      // 6. Build the rows to write
      var rowsToWrite = [];
      for (var i = 0; i < filteredActive.length; i++) {
        var rec = filteredActive[i];
        
        var dateStr = formatExpiryDate(rec.endDate);
        var normDate = normalizeDateStr(dateStr);
        var key = (rec.license + "|" + normDate).toLowerCase();
        
        var qty = "1 Licenses";

        // Map status
        var statusMapped = "CLOSED";
        var stLower = rec.status.toLowerCase();
        if (stLower === 'new') {
          statusMapped = "PENDING";
        } else if (stLower === 'renewal') {
          statusMapped = rec.poStatus.toLowerCase() === 'po pending' ? "PENDING" : "RENEWED";
        } else if (stLower === 'cancelled') {
          statusMapped = "CANCELLED";
        }

        // Initialize variables for oldPo and newPo
        var oldPo = "";
        var newPo = "";

        // First, check if carried over from previous month
        var dateObj = null;
        if (typeof rec.endDate === 'number') {
          dateObj = new Date(Math.round((rec.endDate - 25569) * 86400 * 1000));
        } else {
          dateObj = new Date(rec.endDate);
        }
        var day = dateObj && !isNaN(dateObj.getTime()) ? dateObj.getDate() : null;
        if (day !== null) {
          var carryOverKey = rec.license.toLowerCase() + "|" + day;
          if (prevPOMap[carryOverKey]) {
            oldPo = prevPOMap[carryOverKey];
          }
        }

        // If exists in old sheet, preserve manual entries
        if (existingRowsMap[key]) {
          qty = existingRowsMap[key].qty;
          oldPo = existingRowsMap[key].oldPo || oldPo;
          newPo = existingRowsMap[key].newPo || newPo;
          statusMapped = existingRowsMap[key].status || statusMapped;
        }

        // If status is PENDING, both OLD PO and NEW PO should be empty
        if (statusMapped === "PENDING") {
          oldPo = "";
          newPo = "";
        }

        rowsToWrite.push([
          rec.license,
          qty,
          dateStr,
          rec.billing,
          oldPo,
          newPo,
          statusMapped
        ]);
      }

      // 7. Clear rows below header (B7:H1000)
      var rangeToClear = targetSheet.getRange("B7:H1000");
      rangeToClear.clear();
      await context.sync();

      // 8. Write new rows starting at row 7
      if (rowsToWrite.length > 0) {
        var writeRange = targetSheet.getRange("B7:H" + (6 + rowsToWrite.length));
        writeRange.values = rowsToWrite;
        
        // Add thin gray borders
        writeRange.format.borders.getItem('EdgeBottom').style = 'Continuous';
        writeRange.format.borders.getItem('EdgeTop').style = 'Continuous';
        writeRange.format.borders.getItem('EdgeLeft').style = 'Continuous';
        writeRange.format.borders.getItem('EdgeRight').style = 'Continuous';
        writeRange.format.borders.getItem('InsideHorizontal').style = 'Continuous';
        writeRange.format.borders.getItem('InsideVertical').style = 'Continuous';
        writeRange.format.borders.getItem('EdgeBottom').color = '#e2e8f0';
        writeRange.format.borders.getItem('EdgeTop').color = '#e2e8f0';
        writeRange.format.borders.getItem('EdgeLeft').color = '#e2e8f0';
        writeRange.format.borders.getItem('EdgeRight').color = '#e2e8f0';
        writeRange.format.borders.getItem('InsideHorizontal').color = '#e2e8f0';
        writeRange.format.borders.getItem('InsideVertical').color = '#e2e8f0';
        
        await context.sync();

        // Apply background color-coding based on status (same as master list soft pastel colors) using Conditional Formatting
        var colorNew = '#fef08a';        // Soft Yellow (PENDING)
        var colorRenewal = '#bbf7d0';    // Soft Green (RENEWED)
        var colorComplete = '#bfdbfe';   // Soft Blue (CLOSED)
        var colorCancelled = '#fecaca';  // Soft Red (CANCELLED)

        // 1. Add Data Validation (Dropdown) to the Status column (Column H)
        var statusColRange = targetSheet.getRange("H7:H" + (6 + rowsToWrite.length));
        statusColRange.dataValidation.rule = {
          list: {
            inCellDropDown: true,
            source: "PENDING,RENEWED,CLOSED,CANCELLED"
          }
        };

        // 2. Add Conditional Formatting for the entire data range (B7:H(6 + rowsToWrite.length))
        var fullDataRange = targetSheet.getRange("B7:H" + (6 + rowsToWrite.length));

        // PENDING Rule
        var formatPending = fullDataRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
        formatPending.custom.rule.formula = '=$H7="PENDING"';
        formatPending.custom.format.fill.color = colorNew;
        formatPending.custom.format.font.color = "#000000";

        // RENEWED Rule
        var formatRenewed = fullDataRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
        formatRenewed.custom.rule.formula = '=$H7="RENEWED"';
        formatRenewed.custom.format.fill.color = colorRenewal;
        formatRenewed.custom.format.font.color = "#000000";

        // CLOSED Rule
        var formatClosed = fullDataRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
        formatClosed.custom.rule.formula = '=$H7="CLOSED"';
        formatClosed.custom.format.fill.color = colorComplete;
        formatClosed.custom.format.font.color = "#000000";

        // CANCELLED Rule
        var formatCancelled = fullDataRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
        formatCancelled.custom.rule.formula = '=$H7="CANCELLED"';
        formatCancelled.custom.format.fill.color = colorCancelled;
        formatCancelled.custom.format.font.color = "#000000";
        
        // Set horizontal alignment to Center for columns B to H (headers and data)
        targetSheet.getRange("B6:H10000").format.horizontalAlignment = "Center";
        
        // Auto-fit columns B to H based only on the table range to prevent sheet titles from making columns wider
        targetSheet.getRange("B6:H" + (6 + rowsToWrite.length)).format.autofitColumns();
        await context.sync();
      }

      targetSheet.activate();
      await context.sync();

      document.getElementById('loader').style.display = 'none';
      showSuccess("Sheet '" + targetSheetName + "' generated/updated successfully with " + rowsToWrite.length + " active subscriptions!");
      loadSubscriptions();
    });
  } catch (err) {
    document.getElementById('loader').style.display = 'none';
    showError("Sheet generation failed: " + err.message);
  }
}

function getMonthFromDateJS(val) {
  if (!val) return null;
  var monthNames = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
  ];
  if (typeof val === 'number') {
    var d = new Date(Math.round((val - 25569) * 86400 * 1000));
    return monthNames[d.getUTCMonth()];
  }
  if (typeof val === 'string') {
    var cleaned = val.trim();
    var match = cleaned.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
    if (match) {
      var monthIndex = parseInt(match[2], 10) - 1;
      if (monthIndex >= 0 && monthIndex < 12) {
        return monthNames[monthIndex];
      }
    }
    var d = new Date(cleaned);
    if (!isNaN(d.getTime())) {
      if (cleaned.indexOf('Z') !== -1 || cleaned.indexOf('T') !== -1) {
        return monthNames[d.getUTCMonth()];
      }
      return monthNames[d.getMonth()];
    }
  }
  return null;
}

function parseStatusFieldJS(rawStatus, endDateVal) {
  var statusStr = String(rawStatus || '').trim().toLowerCase();
  
  var month = null;
  var status = null;
  var poStatus = null;

  var months = ["january", "february", "march", "april", "may", "june", "july", "august", "september", "october", "november", "december"];
  
  var foundMonth = months.find(function(m) { return statusStr.indexOf(m) !== -1; });
  if (foundMonth) {
    month = foundMonth.charAt(0).toUpperCase() + foundMonth.slice(1);
  }

  if (statusStr.indexOf('renewal') !== -1) {
    status = 'Renewal';
  } else if (statusStr.indexOf('new') !== -1) {
    status = 'New';
  } else if (statusStr.indexOf('complete') !== -1 || statusStr.indexOf('closed') !== -1 || statusStr.indexOf('active') !== -1 || statusStr.indexOf('license') !== -1) {
    status = 'Complete';
  } else if (statusStr.indexOf('cancel') !== -1) {
    status = 'Cancelled';
  } else {
    status = 'Complete';
  }

  if (status === 'Complete' && !month) {
    month = getMonthFromDateJS(endDateVal);
  }

  if (statusStr.indexOf('done po') !== -1 || statusStr.indexOf('done  po') !== -1 || statusStr.indexOf('po done') !== -1 || statusStr.indexOf('complete') !== -1 || statusStr.indexOf('closed') !== -1 || statusStr.indexOf('active') !== -1 || statusStr === "") {
    poStatus = 'PO Done';
  } else if (statusStr.indexOf('pending po') !== -1 || statusStr.indexOf('po pending') !== -1) {
    poStatus = 'PO Pending';
  }

  return { month: month, status: status, poStatus: poStatus };
}

// Helper to format Date objects or values into DD-MMM-YY format
function formatDateDDMMMYYJS(val) {
  if (!val) return "";
  var monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  var d;
  if (typeof val === 'number') {
    d = new Date(Math.round((val - 25569) * 86400 * 1000));
  } else {
    d = new Date(val);
  }
  if (isNaN(d.getTime())) return "";

  if (typeof val === 'string') {
    var clean = val.trim();
    var matchISO = clean.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
    if (matchISO) {
      var y = matchISO[1].slice(-2);
      var mIndex = parseInt(matchISO[2], 10) - 1;
      var dayStr = String(parseInt(matchISO[3], 10)).padStart(2, '0');
      if (mIndex >= 0 && mIndex < 12) {
        return `${dayStr}-${monthNames[mIndex]}-${y}`;
      }
    }
    var matchUS = clean.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})/);
    if (matchUS) {
      var y = matchUS[3].slice(-2);
      var mIndex = parseInt(matchUS[1], 10) - 1;
      var dayStr = String(parseInt(matchUS[2], 10)).padStart(2, '0');
      if (mIndex >= 0 && mIndex < 12) {
        return `${dayStr}-${monthNames[mIndex]}-${y}`;
      }
    }
  }

  var day = String(d.getUTCDate()).padStart(2, '0');
  var month = monthNames[d.getUTCMonth()];
  var year = String(d.getUTCFullYear()).slice(-2);
  return `${day}-${month}-${year}`;
}

// Helper to parse Period field ("Start Date  End Date") into a numeric timestamp for sorting
function getTimeFromPeriodJS(val) {
  if (!val) return 0;
  var periodStr = String(val).trim();
  var firstDate = periodStr.split(' ')[0] || '';
  if (!firstDate) return 0;

  var parts = firstDate.split('-');
  if (parts.length === 3) {
    var day = parseInt(parts[0], 10);
    var monthName = parts[1].toLowerCase();
    var yearShort = parseInt(parts[2], 10);

    var months = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];
    var monthIndex = months.indexOf(monthName);
    if (monthIndex !== -1 && !isNaN(day)) {
      var year = yearShort < 50 ? 2000 + yearShort : 1900 + yearShort;
      var d = new Date(year, monthIndex, day);
      return d.getTime();
    }
  }
  return 0;
}

// Helper to parse DD-MMM-YY back to YYYY-MM-DD for HTML inputs and flatpickr
function parseDDMMMYY(str) {
  if (!str) return "";
  var clean = String(str).trim();
  if (!clean) return "";

  if (/^\d{4}-\d{2}-\d{2}$/.test(clean)) {
    return clean;
  }

  var parts = clean.split('-');
  if (parts.length === 3) {
    var day = parseInt(parts[0], 10);
    var monthName = parts[1].toLowerCase();
    var yearShort = parseInt(parts[2], 10);

    var months = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];
    var monthIndex = months.indexOf(monthName);
    if (monthIndex !== -1 && !isNaN(day)) {
      var year = yearShort < 50 ? 2000 + yearShort : 1900 + yearShort;
      var mm = String(monthIndex + 1).padStart(2, '0');
      var dd = String(day).padStart(2, '0');
      return `${year}-${mm}-${dd}`;
    }
  }

  try {
    var parsed = new Date(clean);
    if (!isNaN(parsed.getTime())) {
      var y = parsed.getFullYear();
      var m = String(parsed.getMonth() + 1).padStart(2, '0');
      var d = String(parsed.getDate()).padStart(2, '0');
      return `${y}-${m}-${d}`;
    }
  } catch (e) {}

  return clean;
}

// Helper to clean and heal R/T values if Excel auto-converted them to dates
function cleanRTValueJS(val) {
  if (val === null || val === undefined) return "0/0";
  var str = String(val).trim();
  if (!str) return "0/0";

  // If it's already in the format "X/Y"
  if (/^\d+\/\d+$/.test(str)) {
    return str;
  }

  var d;
  if (typeof val === 'number') {
    d = new Date(Math.round((val - 25569) * 86400 * 1000));
  } else {
    d = new Date(val);
  }

  if (!isNaN(d.getTime())) {
    // Reconstruct Remaining/Total from Month/Day
    var monthsLeft = d.getUTCMonth() + 1;
    var totalMonths = d.getUTCDate();
    return monthsLeft + "/" + totalMonths;
  }

  return str;
}

// Helper to clean and heal Period values if Excel auto-converted them to dates
function cleanPeriodValueJS(val) {
  if (val === null || val === undefined) return "";
  var str = String(val).trim();
  if (!str) return "";

  // If it already contains two spaces
  if (str.indexOf('  ') !== -1) {
    return str;
  }

  var d;
  if (typeof val === 'number') {
    d = new Date(Math.round((val - 25569) * 86400 * 1000));
  } else {
    d = new Date(val);
  }

  if (!isNaN(d.getTime())) {
    return formatDateDDMMMYYJS(val);
  }

  return str;
}


