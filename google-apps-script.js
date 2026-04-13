// ============================================================
// GOOGLE APPS SCRIPT — paste this into your Google Sheet
// Extensions → Apps Script → paste → Deploy → Web App
// ============================================================
// SETUP:
// 1. Open your Google Sheet (must have these tabs):
//    - "Received Orders" with headers: PO Number | Location | Configuration | Status | Receive Date | Cabinet Start Date | Assembly Complete Date | QC Complete Date | Cabinet Complete Date | Transfer Out Date
//    - "Transaction History" with headers: PO Number | Timestamp | Application | From Status | To Status | User
//    - "Material Info" with headers: PO Number | Component | Serial Number | Timestamp
//    - "QC Notes" with headers: PO Number | Timestamp | Result | Note
//    - "Infirmary Notes" with headers: PO Number | Timestamp | Note
//    - "Cabinet Configs" with headers: Configuration | Major Material 1 | Major Material 2 | ... | Major Material 15
// 2. Go to Extensions → Apps Script
// 3. Delete any existing code and paste this entire file
// 4. Deploy → New Deployment → Web App → Execute as: Me → Who has access: Anyone
// ============================================================

var RECEIVED_SHEET = 'Received Orders';
var HISTORY_SHEET = 'Transaction History';
var TS_FORMAT = 'M/d/yyyy HH:mm:ss';
// Set this to the ID of the Google Drive folder where photos should be saved
// To get the ID: open the folder in Google Drive, the URL will be https://drive.google.com/drive/folders/XXXXX — copy the XXXXX part
var PHOTO_FOLDER_ID = '1vc83o3eEtJvNli3agcd4E5lt4Qpy8aco';

function _respond(obj, callback) {
  var json = JSON.stringify(obj);
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function _ts(now) {
  return Utilities.formatDate(now, Session.getScriptTimeZone(), TS_FORMAT);
}

function _findPO(sheet, po) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  var poCol = sheet.getRange('A2:A' + lastRow).getValues().flat();
  for (var i = 0; i < poCol.length; i++) {
    if (poCol[i].toString().trim() === po) return i;
  }
  return -1;
}

function doGet(e) {
  var callback = (e && e.parameter && e.parameter.callback) ? e.parameter.callback : null;
  var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : 'read';
  var user = (e && e.parameter && e.parameter.user) ? e.parameter.user.toString().trim() : '';

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var receivedSheet = ss.getSheetByName(RECEIVED_SHEET);
    var historySheet = ss.getSheetByName(HISTORY_SHEET);

    if (!receivedSheet) return _respond({ success: false, error: 'Sheet "Received Orders" not found' }, callback);

    // ---- RECEIVE ----
    if (action === 'receive') {
      if (!historySheet) return _respond({ success: false, error: 'Sheet "Transaction History" not found' }, callback);
      var po = (e.parameter.po || '').toString().trim();
      var config = (e.parameter.config || '').toString().trim();
      var loc = (e.parameter.location || '').toString().trim();
      if (!po) return _respond({ success: false, error: 'PO Number is required' }, callback);

      if (_findPO(receivedSheet, po) >= 0) {
        return _respond({ success: false, error: 'DUPLICATE', message: 'PO#' + po + ' has already been received' }, callback);
      }

      var now = new Date();
      var ts = _ts(now);
      receivedSheet.appendRow([po, loc, config, 'Received, Awaiting Start', ts, '', '', '', '', '']);
      historySheet.appendRow([po, ts, 'Receive APP', '-', 'Received, Awaiting Start', user]);
      return _respond({ success: true, po: po, receiveDate: ts }, callback);
    }

    // ---- START BUILD ----
    if (action === 'startbuild') {
      if (!historySheet) return _respond({ success: false, error: 'Sheet "Transaction History" not found' }, callback);
      var po = (e.parameter.po || '').toString().trim();
      if (!po) return _respond({ success: false, error: 'PO Number is required' }, callback);

      var idx = _findPO(receivedSheet, po);
      if (idx === -1) return _respond({ success: false, error: 'PO#' + po + ' not found' }, callback);

      var row = idx + 2;
      var status = receivedSheet.getRange(row, 4).getValue().toString().trim();
      if (status !== 'Received, Awaiting Start') {
        return _respond({ success: false, error: 'PO#' + po + ' status is "' + status + '", expected "Received, Awaiting Start"' }, callback);
      }

      var now = new Date();
      var ts = _ts(now);
      receivedSheet.getRange(row, 4).setValue('In Assembly');
      receivedSheet.getRange(row, 6).setValue(ts);
      historySheet.appendRow([po, ts, 'Start Build APP', 'Received, Awaiting Start', 'In Assembly', user]);
      return _respond({ success: true, po: po, startDate: ts }, callback);
    }

    // ---- BUILD COMPLETE ----
    if (action === 'buildcomplete') {
      if (!historySheet) return _respond({ success: false, error: 'Sheet "Transaction History" not found' }, callback);
      var po = (e.parameter.po || '').toString().trim();
      if (!po) return _respond({ success: false, error: 'PO Number is required' }, callback);

      var idx = _findPO(receivedSheet, po);
      if (idx === -1) return _respond({ success: false, error: 'PO#' + po + ' not found' }, callback);

      var row = idx + 2;
      var status = receivedSheet.getRange(row, 4).getValue().toString().trim();
      if (status !== 'In Assembly') {
        return _respond({ success: false, error: 'PO#' + po + ' status is "' + status + '", expected "In Assembly"' }, callback);
      }

      // Check material info completeness
      var config = receivedSheet.getRange(row, 3).getValue().toString().trim(); // Column C = Configuration
      if (config) {
        var cfgSheet = ss.getSheetByName('Cabinet Configs');
        if (cfgSheet) {
          var cfgData = cfgSheet.getDataRange().getValues();
          var cfgHeaders = cfgData[0];
          var requiredMats = [];
          for (var ci = 1; ci < cfgData.length; ci++) {
            if (cfgData[ci][0].toString().trim() === config) {
              for (var mi = 1; mi <= 15; mi++) {
                var matName = (cfgData[ci][mi] || '').toString().trim();
                if (matName && matName !== '-') requiredMats.push(matName);
              }
              break;
            }
          }

          if (requiredMats.length > 0) {
            var matSheet = ss.getSheetByName('Material Info');
            var recordedComponents = [];
            if (matSheet && matSheet.getLastRow() > 1) {
              var matData = matSheet.getDataRange().getValues();
              for (var mi = 1; mi < matData.length; mi++) {
                if (matData[mi][0].toString().trim() === po) {
                  recordedComponents.push(matData[mi][1].toString().trim());
                }
              }
            }

            var missing = [];
            for (var ri = 0; ri < requiredMats.length; ri++) {
              if (recordedComponents.indexOf(requiredMats[ri]) === -1) {
                missing.push(requiredMats[ri]);
              }
            }

            if (missing.length > 0) {
              return _respond({
                success: false,
                error: 'MISSING_MATERIALS',
                message: 'Material info required for: ' + missing.join(', '),
                missing: missing
              }, callback);
            }
          }
        }
      }

      var now = new Date();
      var ts = _ts(now);
      receivedSheet.getRange(row, 4).setValue('Awaiting Testing');
      receivedSheet.getRange(row, 7).setValue(ts);
      historySheet.appendRow([po, ts, 'Assembly Complete APP', 'In Assembly', 'Awaiting Testing', user]);
      return _respond({ success: true, po: po, completeDate: ts }, callback);
    }

    // ---- TEST RESULT (move from Awaiting Testing to Awaiting QC) ----
    if (action === 'testresult') {
      if (!historySheet) return _respond({ success: false, error: 'Sheet "Transaction History" not found' }, callback);
      var po = (e.parameter.po || '').toString().trim();
      if (!po) return _respond({ success: false, error: 'PO Number is required' }, callback);

      var idx = _findPO(receivedSheet, po);
      if (idx === -1) return _respond({ success: false, error: 'PO#' + po + ' not found' }, callback);

      var row = idx + 2;
      var status = receivedSheet.getRange(row, 4).getValue().toString().trim();
      if (status !== 'Awaiting Testing') {
        return _respond({ success: false, error: 'PO#' + po + ' status is "' + status + '", expected "Awaiting Testing"' }, callback);
      }

      var now = new Date();
      var ts = _ts(now);
      receivedSheet.getRange(row, 4).setValue('Awaiting QC');
      receivedSheet.getRange(row, 8).setValue(ts); // Test Results Date = col H
      historySheet.appendRow([po, ts, 'Test Results APP', 'Awaiting Testing', 'Awaiting QC', user]);
      return _respond({ success: true, po: po, testDate: ts }, callback);
    }

    // ---- QC PASS ----
    if (action === 'qcpass') {
      if (!historySheet) return _respond({ success: false, error: 'Sheet "Transaction History" not found' }, callback);
      var po = (e.parameter.po || '').toString().trim();
      if (!po) return _respond({ success: false, error: 'PO Number is required' }, callback);

      var idx = _findPO(receivedSheet, po);
      if (idx === -1) return _respond({ success: false, error: 'PO#' + po + ' not found' }, callback);

      var row = idx + 2;
      var status = receivedSheet.getRange(row, 4).getValue().toString().trim();
      if (status !== 'Awaiting QC') {
        return _respond({ success: false, error: 'PO#' + po + ' status is "' + status + '", expected "Awaiting QC"' }, callback);
      }

      var now = new Date();
      var ts = _ts(now);
      receivedSheet.getRange(row, 4).setValue('QC Passed');
      receivedSheet.getRange(row, 9).setValue(ts); // QC Result Date = col I
      historySheet.appendRow([po, ts, 'QC Result APP', 'Awaiting QC', 'QC Passed', user]);
      return _respond({ success: true, po: po, qcDate: ts }, callback);
    }

    // ---- QC FAIL → IN INFIRMARY ----
    if (action === 'qcfail') {
      if (!historySheet) return _respond({ success: false, error: 'Sheet "Transaction History" not found' }, callback);
      var po = (e.parameter.po || '').toString().trim();
      var note = (e.parameter.note || '').toString().trim();
      if (!po) return _respond({ success: false, error: 'PO Number is required' }, callback);

      var idx = _findPO(receivedSheet, po);
      if (idx === -1) return _respond({ success: false, error: 'PO#' + po + ' not found' }, callback);

      var row = idx + 2;
      var status = receivedSheet.getRange(row, 4).getValue().toString().trim();
      if (status !== 'Awaiting QC') {
        return _respond({ success: false, error: 'PO#' + po + ' status is "' + status + '", expected "Awaiting QC"' }, callback);
      }

      var now = new Date();
      var ts = _ts(now);
      receivedSheet.getRange(row, 4).setValue('In Infirmary');
      receivedSheet.getRange(row, 9).setValue(ts); // QC Result Date = col I
      historySheet.appendRow([po, ts, 'QC Result APP', 'Awaiting QC', 'In Infirmary', user]);

      // Write to Infirmary Notes
      var infNotesSheet = ss.getSheetByName('Infirmary Notes');
      if (infNotesSheet && note) {
        infNotesSheet.appendRow([po, ts, 'QC Fail: ' + note]);
      }

      return _respond({ success: true, po: po, qcDate: ts }, callback);
    }

    // ---- MATERIAL SEARCH (by PO or serial number) ----
    if (action === 'materialsearch') {
      var matSheet = ss.getSheetByName('Material Info');
      if (!matSheet) return _respond({ success: false, error: 'Sheet "Material Info" not found' }, callback);

      var query = (e.parameter.query || '').toString().trim();
      if (!query) return _respond({ success: false, error: 'Search query is required' }, callback);
      var searchType = (e.parameter.searchType || '').toString().trim(); // 'po' or 'serial'

      var data = matSheet.getDataRange().getValues();
      if (data.length <= 1) return _respond({ success: true, data: [], query: query }, callback);

      var headers = data[0];
      var rows = [];
      for (var i = 1; i < data.length; i++) {
        var po = data[i][0].toString().trim();
        var component = data[i][1].toString().trim();
        var serial = data[i][2].toString().trim();
        var match = false;
        if (searchType === 'po') {
          match = (po !== '' && po === query);
        } else if (searchType === 'serial') {
          match = (serial !== '' && serial.indexOf(query) >= 0);
        } else {
          match = (po !== '' && po.indexOf(query) >= 0) || (serial !== '' && serial.indexOf(query) >= 0);
        }
        if (match) {
          var row = {};
          headers.forEach(function(h, j) { row[h] = data[i][j]; });
          rows.push(row);
        }
      }

      // Also get PO details from Received Orders for matched POs
      var poDetails = {};
      if (rows.length > 0) {
        var roData = receivedSheet.getDataRange().getValues();
        var roHeaders = roData[0];
        for (var i = 1; i < roData.length; i++) {
          var roPO = roData[i][0].toString().trim();
          for (var j = 0; j < rows.length; j++) {
            if (rows[j]['PO Number'].toString().trim() === roPO && !poDetails[roPO]) {
              var detail = {};
              roHeaders.forEach(function(h, k) { detail[h] = roData[i][k]; });
              poDetails[roPO] = detail;
              break;
            }
          }
        }
      }

      return _respond({ success: true, data: rows, poDetails: poDetails, query: query }, callback);
    }

    // ---- MATERIAL INFO ----
    if (action === 'materialinfo') {
      var matSheet = ss.getSheetByName('Material Info');
      if (!matSheet) return _respond({ success: false, error: 'Sheet "Material Info" not found' }, callback);

      var po = (e.parameter.po || '').toString().trim();
      var component = (e.parameter.component || '').toString().trim();
      var serial = (e.parameter.serial || '').toString().trim();
      if (!po) return _respond({ success: false, error: 'PO Number is required' }, callback);
      if (!component) return _respond({ success: false, error: 'Component is required' }, callback);

      var now = new Date();
      var ts = _ts(now);
      matSheet.appendRow([po, component, serial, ts]);
      return _respond({ success: true, po: po, component: component, serial: serial }, callback);
    }

    // ---- UPDATE MATERIAL SERIAL ----
    if (action === 'updatematerial') {
      var matSheet = ss.getSheetByName('Material Info');
      if (!matSheet) return _respond({ success: false, error: 'Sheet "Material Info" not found' }, callback);
      var rowNum = parseInt(e.parameter.row || '0');
      var newSerial = (e.parameter.serial || '').toString().trim();
      if (!rowNum || rowNum < 2) return _respond({ success: false, error: 'Invalid row' }, callback);
      if (rowNum > matSheet.getLastRow()) return _respond({ success: false, error: 'Row not found' }, callback);
      matSheet.getRange(rowNum, 3).setValue(newSerial);
      return _respond({ success: true }, callback);
    }

    // ---- DELETE MATERIAL ENTRY ----
    if (action === 'deletematerial') {
      var matSheet = ss.getSheetByName('Material Info');
      if (!matSheet) return _respond({ success: false, error: 'Sheet "Material Info" not found' }, callback);
      var rowNum = parseInt(e.parameter.row || '0');
      if (!rowNum || rowNum < 2) return _respond({ success: false, error: 'Invalid row' }, callback);
      if (rowNum > matSheet.getLastRow()) return _respond({ success: false, error: 'Row not found' }, callback);
      matSheet.deleteRow(rowNum);
      return _respond({ success: true }, callback);
    }

    // ---- MATERIAL INFO READ (get entries for a PO) ----
    if (action === 'materialread') {
      var matSheet = ss.getSheetByName('Material Info');
      if (!matSheet) return _respond({ success: false, error: 'Sheet "Material Info" not found' }, callback);

      var po = (e.parameter.po || '').toString().trim();
      var data = matSheet.getDataRange().getValues();
      if (data.length <= 1) return _respond({ success: true, data: [] }, callback);

      var headers = data[0];
      var rows = [];
      for (var i = 1; i < data.length; i++) {
        var row = {};
        headers.forEach(function(h, j) { row[h] = data[i][j]; });
        row['_row'] = i + 1; // sheet row number (1-indexed, +1 for header)
        if (!po || row['PO Number'].toString().trim() === po) rows.push(row);
      }
      return _respond({ success: true, data: rows }, callback);
    }

    // ---- SHIP OUT ----
    if (action === 'shipout') {
      if (!historySheet) return _respond({ success: false, error: 'Sheet "Transaction History" not found' }, callback);
      var po = (e.parameter.po || '').toString().trim();
      if (!po) return _respond({ success: false, error: 'PO Number is required' }, callback);

      var idx = _findPO(receivedSheet, po);
      if (idx === -1) return _respond({ success: false, error: 'PO#' + po + ' not found' }, callback);

      var row = idx + 2;
      var status = receivedSheet.getRange(row, 4).getValue().toString().trim();
      if (status !== 'QC Passed') {
        return _respond({ success: false, error: 'PO#' + po + ' status is "' + status + '", expected "QC Passed"' }, callback);
      }

      var now = new Date();
      var ts = _ts(now);
      receivedSheet.getRange(row, 4).setValue('Tender');
      receivedSheet.getRange(row, 10).setValue(ts); // Transfer Out Date = col J = index 10
      historySheet.appendRow([po, ts, 'Ship Out APP', 'QC Passed', 'Tender', user]);
      return _respond({ success: true, po: po, shipDate: ts }, callback);
    }

    // ---- SEND TO INFIRMARY ----
    if (action === 'toinfirmary') {
      if (!historySheet) return _respond({ success: false, error: 'Sheet "Transaction History" not found' }, callback);
      var po = (e.parameter.po || '').toString().trim();
      var note = (e.parameter.note || '').toString().trim();
      if (!po) return _respond({ success: false, error: 'PO Number is required' }, callback);
      if (!note) return _respond({ success: false, error: 'A note is required' }, callback);

      var idx = _findPO(receivedSheet, po);
      if (idx === -1) return _respond({ success: false, error: 'PO#' + po + ' not found' }, callback);

      var row = idx + 2;
      var currentStatus = receivedSheet.getRange(row, 4).getValue().toString().trim();
      if (currentStatus === 'Tender') {
        return _respond({ success: false, error: 'PO#' + po + ' is already Tendered' }, callback);
      }

      var now = new Date();
      var ts = _ts(now);
      receivedSheet.getRange(row, 4).setValue('In Infirmary');
      historySheet.appendRow([po, ts, 'Infirmary APP', currentStatus, 'In Infirmary', user]);

      var infNotesSheet = ss.getSheetByName('Infirmary Notes');
      if (infNotesSheet) {
        infNotesSheet.appendRow([po, ts, note]);
      }

      return _respond({ success: true, po: po, fromStatus: currentStatus }, callback);
    }

    // ---- ADD INFIRMARY NOTE ----
    if (action === 'addinfirmarynote') {
      var po = (e.parameter.po || '').toString().trim();
      var note = (e.parameter.note || '').toString().trim();
      if (!po) return _respond({ success: false, error: 'PO Number is required' }, callback);
      if (!note) return _respond({ success: false, error: 'A note is required' }, callback);

      var infNotesSheet = ss.getSheetByName('Infirmary Notes');
      if (!infNotesSheet) return _respond({ success: false, error: 'Sheet "Infirmary Notes" not found' }, callback);

      var now = new Date();
      var ts = _ts(now);
      infNotesSheet.appendRow([po, ts, note]);
      return _respond({ success: true, po: po }, callback);
    }

    // ---- READ INFIRMARY NOTES ----
    if (action === 'readinfirmarynotes') {
      var infNotesSheet = ss.getSheetByName('Infirmary Notes');
      if (!infNotesSheet) return _respond({ success: false, error: 'Sheet "Infirmary Notes" not found' }, callback);

      var po = (e.parameter.po || '').toString().trim();
      var data = infNotesSheet.getDataRange().getValues();
      if (data.length <= 1) return _respond({ success: true, data: [] }, callback);

      var headers = data[0];
      var rows = [];
      for (var i = 1; i < data.length; i++) {
        var row = {};
        headers.forEach(function(h, j) { row[h] = data[i][j]; });
        if (!po || row['PO Number'].toString().trim() === po) rows.push(row);
      }
      return _respond({ success: true, data: rows }, callback);
    }

    // ---- MOVE FROM INFIRMARY ----
    if (action === 'movefrominf') {
      if (!historySheet) return _respond({ success: false, error: 'Sheet "Transaction History" not found' }, callback);
      var po = (e.parameter.po || '').toString().trim();
      var newStatus = (e.parameter.newstatus || '').toString().trim();
      var note = (e.parameter.note || '').toString().trim();
      if (!po) return _respond({ success: false, error: 'PO Number is required' }, callback);
      if (!newStatus) return _respond({ success: false, error: 'New status is required' }, callback);

      var idx = _findPO(receivedSheet, po);
      if (idx === -1) return _respond({ success: false, error: 'PO#' + po + ' not found' }, callback);

      var row = idx + 2;
      var currentStatus = receivedSheet.getRange(row, 4).getValue().toString().trim();
      if (currentStatus !== 'In Infirmary') {
        return _respond({ success: false, error: 'PO#' + po + ' is not In Infirmary' }, callback);
      }

      var now = new Date();
      var ts = _ts(now);
      receivedSheet.getRange(row, 4).setValue(newStatus);
      historySheet.appendRow([po, ts, 'Infirmary APP', 'In Infirmary', newStatus, user]);

      if (note) {
        var infNotesSheet = ss.getSheetByName('Infirmary Notes');
        if (infNotesSheet) {
          infNotesSheet.appendRow([po, ts, 'Moved to ' + newStatus + ': ' + note]);
        }
      }

      return _respond({ success: true, po: po, newStatus: newStatus }, callback);
    }

    // ---- PO HISTORY (read transaction history for a PO) ----
    if (action === 'pohistory') {
      if (!historySheet) return _respond({ success: false, error: 'Sheet "Transaction History" not found' }, callback);
      var po = (e.parameter.po || '').toString().trim();
      if (!po) return _respond({ success: false, error: 'PO Number is required' }, callback);

      var data = historySheet.getDataRange().getValues();
      if (data.length <= 1) return _respond({ success: true, data: [] }, callback);

      var headers = data[0];
      var rows = [];
      for (var i = 1; i < data.length; i++) {
        if (data[i][0].toString().trim() === po) {
          var row = {};
          headers.forEach(function(h, j) { row[h] = data[i][j]; });
          rows.push(row);
        }
      }
      return _respond({ success: true, data: rows }, callback);
    }

    // ---- READ CONFIGS ----
    if (action === 'readconfigs') {
      var cfgSheet = ss.getSheetByName('Cabinet Configs');
      if (!cfgSheet) return _respond({ success: false, error: 'Sheet "Cabinet Configs" not found' }, callback);
      var data = cfgSheet.getDataRange().getValues();
      if (data.length <= 1) return _respond({ success: true, data: [] }, callback);
      var headers = data[0];
      var rows = [];
      for (var i = 1; i < data.length; i++) {
        var row = {};
        headers.forEach(function(h, j) { row[h] = data[i][j]; });
        rows.push(row);
      }
      return _respond({ success: true, data: rows }, callback);
    }

    // ---- ADD CONFIG ----
    if (action === 'addconfig') {
      var cfgSheet = ss.getSheetByName('Cabinet Configs');
      if (!cfgSheet) return _respond({ success: false, error: 'Sheet "Cabinet Configs" not found' }, callback);
      var name = (e.parameter.name || '').toString().trim();
      if (!name) return _respond({ success: false, error: 'Configuration name is required' }, callback);

      // Duplicate check
      var lastRow = cfgSheet.getLastRow();
      if (lastRow > 1) {
        var names = cfgSheet.getRange('A2:A' + lastRow).getValues().flat();
        if (names.some(function(v) { return v.toString().trim().toLowerCase() === name.toLowerCase(); })) {
          return _respond({ success: false, error: 'Configuration "' + name + '" already exists' }, callback);
        }
      }

      var rowData = [name];
      for (var i = 1; i <= 15; i++) {
        rowData.push((e.parameter['mat' + i] || '').toString().trim());
      }
      cfgSheet.appendRow(rowData);
      return _respond({ success: true, name: name }, callback);
    }

    // ---- UPDATE CONFIG ----
    if (action === 'updateconfig') {
      var cfgSheet = ss.getSheetByName('Cabinet Configs');
      if (!cfgSheet) return _respond({ success: false, error: 'Sheet "Cabinet Configs" not found' }, callback);
      var origName = (e.parameter.origname || '').toString().trim();
      var name = (e.parameter.name || '').toString().trim();
      if (!origName || !name) return _respond({ success: false, error: 'Configuration name is required' }, callback);

      var lastRow = cfgSheet.getLastRow();
      if (lastRow < 2) return _respond({ success: false, error: 'Configuration not found' }, callback);
      var names = cfgSheet.getRange('A2:A' + lastRow).getValues().flat();
      var rowIdx = -1;
      for (var i = 0; i < names.length; i++) {
        if (names[i].toString().trim().toLowerCase() === origName.toLowerCase()) { rowIdx = i; break; }
      }
      if (rowIdx === -1) return _respond({ success: false, error: 'Configuration "' + origName + '" not found' }, callback);

      var sheetRow = rowIdx + 2;
      cfgSheet.getRange(sheetRow, 1).setValue(name);
      for (var i = 1; i <= 15; i++) {
        cfgSheet.getRange(sheetRow, i + 1).setValue((e.parameter['mat' + i] || '').toString().trim());
      }
      return _respond({ success: true, name: name }, callback);
    }

    // ---- DELETE CONFIG ----
    if (action === 'deleteconfig') {
      var cfgSheet = ss.getSheetByName('Cabinet Configs');
      if (!cfgSheet) return _respond({ success: false, error: 'Sheet "Cabinet Configs" not found' }, callback);
      var name = (e.parameter.name || '').toString().trim();
      if (!name) return _respond({ success: false, error: 'Configuration name is required' }, callback);

      var lastRow = cfgSheet.getLastRow();
      if (lastRow < 2) return _respond({ success: false, error: 'Configuration not found' }, callback);
      var names = cfgSheet.getRange('A2:A' + lastRow).getValues().flat();
      var rowIdx = -1;
      for (var i = 0; i < names.length; i++) {
        if (names[i].toString().trim().toLowerCase() === name.toLowerCase()) { rowIdx = i; break; }
      }
      if (rowIdx === -1) return _respond({ success: false, error: 'Configuration "' + name + '" not found' }, callback);

      cfgSheet.deleteRow(rowIdx + 2);
      return _respond({ success: true, name: name }, callback);
    }

    // ---- PHOTO TEST (verify Drive access) ----
    if (action === 'phototest') {
      var parentFolder = PHOTO_FOLDER_ID ? DriveApp.getFolderById(PHOTO_FOLDER_ID) : DriveApp.getRootFolder();
      return _respond({ success: true, folderName: parentFolder.getName(), folderId: PHOTO_FOLDER_ID }, callback);
    }

    // ---- PHOTO CHUNK (receive a piece of base64 image via JSONP) ----
    if (action === 'photochunk') {
      var uploadId = (e.parameter.uploadId || '').toString();
      var chunkIndex = parseInt(e.parameter.chunkIndex || '0');
      var totalChunks = parseInt(e.parameter.totalChunks || '0');
      var chunk = (e.parameter.chunk || '').toString();
      // Decode if it was double-encoded
      try { chunk = decodeURIComponent(chunk); } catch(ex) {}
      if (!uploadId) return _respond({ success: false, error: 'Missing uploadId' }, callback);

      var cache = CacheService.getScriptCache();
      cache.put(uploadId + '_' + chunkIndex, chunk, 300);
      cache.put(uploadId + '_total', totalChunks.toString(), 300);
      return _respond({ success: true, chunk: chunkIndex, of: totalChunks }, callback);
    }

    // ---- PHOTO ASSEMBLE (combine chunks and save to Drive) ----
    if (action === 'photoassemble') {
      var uploadId = (e.parameter.uploadId || '').toString();
      var po = (e.parameter.po || '').toString().trim();
      var label = (e.parameter.label || 'photo').toString().trim();
      var totalChunks = parseInt(e.parameter.totalChunks || '0');
      if (!uploadId || !po || !totalChunks) return _respond({ success: false, error: 'Missing parameters' }, callback);

      var cache = CacheService.getScriptCache();
      var imageData = '';
      for (var i = 0; i < totalChunks; i++) {
        var chunk = cache.get(uploadId + '_' + i);
        if (!chunk) return _respond({ success: false, error: 'Chunk ' + i + ' expired or missing. Try again.' }, callback);
        imageData += chunk;
      }

      var parentFolder = PHOTO_FOLDER_ID ? DriveApp.getFolderById(PHOTO_FOLDER_ID) : DriveApp.getRootFolder();
      var poFolders = parentFolder.getFoldersByName(po);
      var poFolder = poFolders.hasNext() ? poFolders.next() : parentFolder.createFolder(po);

      var parts = imageData.split(',');
      var mimeMatch = parts[0].match(/data:(.*?);/);
      var mime = mimeMatch ? mimeMatch[1] : 'image/jpeg';
      var ext = mime === 'image/png' ? '.png' : '.jpg';
      var fileName = po + ' - ' + label + ext;
      var blob = Utilities.newBlob(Utilities.base64Decode(parts[1]), mime, fileName);

      var file = poFolder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

      for (var i = 0; i < totalChunks; i++) { cache.remove(uploadId + '_' + i); }
      cache.remove(uploadId + '_total');

      return _respond({ success: true, fileId: file.getId(), fileUrl: file.getUrl(), fileName: fileName, folderUrl: poFolder.getUrl() }, callback);
    }

    // ---- LIST PHOTOS (JSONP) ----
    if (action === 'listphotos') {
      var po = (e.parameter.po || '').toString().trim();
      if (!po) return _respond({ success: false, error: 'PO Number is required' }, callback);

      var parentFolder = PHOTO_FOLDER_ID ? DriveApp.getFolderById(PHOTO_FOLDER_ID) : DriveApp.getRootFolder();
      var poFolders = parentFolder.getFoldersByName(po);
      if (!poFolders.hasNext()) return _respond({ success: true, photos: [] }, callback);

      var poFolder = poFolders.next();
      var files = poFolder.getFiles();
      var photos = [];
      while (files.hasNext()) {
        var f = files.next();
        photos.push({
          name: f.getName(),
          url: 'https://drive.google.com/thumbnail?id=' + f.getId() + '&sz=w400',
          fullUrl: f.getUrl(),
          date: Utilities.formatDate(f.getDateCreated(), Session.getScriptTimeZone(), TS_FORMAT)
        });
      }
      return _respond({ success: true, photos: photos, folderUrl: poFolder.getUrl() }, callback);
    }

    // ---- Default: READ all orders ----
    var data = receivedSheet.getDataRange().getValues();
    if (data.length <= 1) return _respond({ success: true, data: [] }, callback);

    var headers = data[0];
    var rows = [];
    for (var i = 1; i < data.length; i++) {
      var row = {};
      headers.forEach(function(h, j) { row[h] = data[i][j]; });
      rows.push(row);
    }
    return _respond({ success: true, data: rows }, callback);

  } catch (err) {
    return _respond({ success: false, error: err.toString() }, callback);
  }
}

function doPost(e) {
  try {
    var data;
    // Handle form POST (from iframe) or JSON POST
    if (e.parameter && e.parameter.postData) {
      data = JSON.parse(e.parameter.postData);
    } else {
      data = JSON.parse(e.postData.contents);
    }

    // ---- PHOTO UPLOAD (POST only) ----
    if (data.action === 'uploadphoto') {
      var po = (data.po || '').toString().trim();
      var imageData = (data.image || '').toString();
      var label = (data.label || '').toString().trim() || 'photo';
      if (!po) return _respond({ success: false, error: 'PO Number is required' });
      if (!imageData) return _respond({ success: false, error: 'No image data' });

      // Get or create PO subfolder
      var parentFolder;
      if (PHOTO_FOLDER_ID) {
        parentFolder = DriveApp.getFolderById(PHOTO_FOLDER_ID);
      } else {
        parentFolder = DriveApp.getRootFolder();
      }

      var poFolders = parentFolder.getFoldersByName(po);
      var poFolder = poFolders.hasNext() ? poFolders.next() : parentFolder.createFolder(po);

      // Decode base64 image
      var parts = imageData.split(',');
      var mimeMatch = parts[0].match(/data:(.*?);/);
      var mime = mimeMatch ? mimeMatch[1] : 'image/jpeg';
      var ext = mime === 'image/png' ? '.png' : '.jpg';
      var blob = Utilities.newBlob(Utilities.base64Decode(parts[1]), mime, po + '_' + label + '_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss') + ext);

      var file = poFolder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

      return _respond({ success: true, fileId: file.getId(), fileUrl: file.getUrl(), fileName: file.getName() });
    }

    // ---- LIST PHOTOS (POST for consistency) ----
    if (data.action === 'listphotos') {
      var po = (data.po || '').toString().trim();
      if (!po) return _respond({ success: false, error: 'PO Number is required' });

      var parentFolder;
      if (PHOTO_FOLDER_ID) {
        parentFolder = DriveApp.getFolderById(PHOTO_FOLDER_ID);
      } else {
        parentFolder = DriveApp.getRootFolder();
      }

      var poFolders = parentFolder.getFoldersByName(po);
      if (!poFolders.hasNext()) return _respond({ success: true, photos: [] });

      var poFolder = poFolders.next();
      var files = poFolder.getFiles();
      var photos = [];
      while (files.hasNext()) {
        var f = files.next();
        photos.push({
          name: f.getName(),
          url: 'https://drive.google.com/thumbnail?id=' + f.getId() + '&sz=w400',
          fullUrl: f.getUrl(),
          date: Utilities.formatDate(f.getDateCreated(), Session.getScriptTimeZone(), TS_FORMAT)
        });
      }
      return _respond({ success: true, photos: photos });
    }

    // Fallback to GET handler for other POST actions
    var fakeE = { parameter: data };
    if (!data.action) fakeE.parameter.action = 'receive';
    return doGet(fakeE);
  } catch (err) {
    return _respond({ success: false, error: err.toString() });
  }
}
