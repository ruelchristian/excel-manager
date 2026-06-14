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

      // Read columns
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
    } else if (statusLower === 'complete' || statusLower.indexOf('complete') !== -1) {
      cardClass = 'status-completed';
      badgeClass = 'badge-completed';
      statusLabel = 'COMPLETE';
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

  switchTab('form');
}

// Reset Form
function resetForm() {
  document.getElementById('record-form').reset();
  if (startPicker) startPicker.clear();
  if (endPicker) endPicker.clear();
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
        range.values = [[
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
        ]];
        
        var colD_E = range.getCell(0, 3).getResizedRange(0, 1);
        colD_E.numberFormat = [["0", "0"]];
        var colF_G = range.getCell(0, 5).getResizedRange(0, 1);
        colF_G.numberFormat = [["yyyy-mm-dd", "yyyy-mm-dd"]];
      } else {
        range.values = [[
          formData.status,
          formData.startDate,
          formData.endDate,
          formData.eu,
          formData.fa_so,
          formData.subscription
        ]];
        
        var colB_C = range.getCell(0, 1).getResizedRange(0, 1);
        colB_C.numberFormat = [["yyyy-mm-dd", "yyyy-mm-dd"]];
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

      var colLetter = currentListType === 'monthly' ? "L" : "F";
      var range = sheet.getRange("A" + startRow + ":" + colLetter + lastRowIndex);
      range.load("values");
      await context.sync();

      var values = range.values;

      // Clean up and standardize statuses in memory for Annual Master
      if (currentListType === 'annual') {
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

      var monthOrder = {
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
      range.values = values;
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
  var colorNew = '#e8f0fe';        // Soft Blue
  var colorRenewal = '#f3e8ff';    // Soft Purple
  var colorComplete = '#e6f4ea';   // Soft Green
  var colorCancelled = '#fce8e6';  // Soft Red
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
    } else if (status === 'complete' || status.indexOf('complete') !== -1) {
      rowColor = colorComplete;
    } else if (status === 'cancelled' || status.indexOf('cancel') !== -1) {
      rowColor = colorCancelled;
    }

    var rowNum = i + startRow;
    var rowRange = sheet.getRange("A" + rowNum + ":" + colLetter + rowNum);
    rowRange.format.fill.color = rowColor;
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
        var statusMapped = "DONE";
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
        var colorNew = '#e8f0fe';        // Soft Blue (PENDING)
        var colorRenewal = '#f3e8ff';    // Soft Purple (RENEWED)
        var colorComplete = '#e6f4ea';   // Soft Green (DONE)
        var colorCancelled = '#fce8e6';  // Soft Red (CANCELLED)

        // 1. Add Data Validation (Dropdown) to the Status column (Column H)
        var statusColRange = targetSheet.getRange("H7:H" + (6 + rowsToWrite.length));
        statusColRange.dataValidation.rule = {
          list: {
            inCellDropDown: true,
            source: "PENDING, RENEWED, DONE, CANCELLED"
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

        // DONE Rule
        var formatDone = fullDataRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
        formatDone.custom.rule.formula = '=$H7="DONE"';
        formatDone.custom.format.fill.color = colorComplete;
        formatDone.custom.format.font.color = "#000000";

        // CANCELLED Rule
        var formatCancelled = fullDataRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
        formatCancelled.custom.rule.formula = '=$H7="CANCELLED"';
        formatCancelled.custom.format.fill.color = colorCancelled;
        formatCancelled.custom.format.font.color = "#000000";
        
        // Auto-fit columns B to H to prevent text truncation and ensure they are wide enough
        targetSheet.getRange("B:H").format.autofitColumns();
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
