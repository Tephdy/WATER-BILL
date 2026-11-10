// ================= CONFIGURATION =================
const ROOT_FOLDER_ID = '1SLd83WomxtEH3UPuzo3GpSBbUwiaWPHE';

// ================= GET SPREADSHEET (container-bound) =================
function getSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Script is not bound to a spreadsheet. Open it via Extensions > Apps Script.');
  return ss;
}

// ================= MAIN HANDLER =================
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return createResponse(false, null, 'No POST data received');
    }
    const requestData = JSON.parse(e.postData.contents);
    const action = requestData.action;

    switch (action) {
      case 'load':       return handleLoad();
      case 'save':       return handleSave(requestData);
      case 'update':     return handleUpdate(requestData);
      case 'delete':     return handleDelete(requestData);
      case 'clearTab':   return handleClearTab(requestData);
      case 'uploadPdf':  return handleUploadPdf(requestData);
      case 'getTabs':    return handleGetTabs();
      default:
        return createResponse(false, null, 'Unknown action: ' + action);
    }
  } catch (error) {
    return createResponse(false, null, 'Error: ' + error.toString());
  }
}

// ================= UPLOAD PDF -> Main / Property / Month / Tenant =================
function handleUploadPdf(data) {
  try {
    if (!data.base64 || !data.fileName) {
      return createResponse(false, null, 'Missing base64 or fileName');
    }
    const decoded = Utilities.base64Decode(data.base64);
    const blob = Utilities.newBlob(decoded, 'application/pdf', data.fileName);

    let mainFolder;
    try {
      mainFolder = DriveApp.getFolderById(data.rootFolderId || ROOT_FOLDER_ID);
    } catch (e) {
      return createResponse(false, null, 'Cannot open main folder: ' + e.toString());
    }

    const propertyName = (data.propertyName || 'Unknown Property').trim();
    const propertyFolder = getOrCreateFolder(mainFolder, propertyName);

    const monthName = (data.monthName || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMMM yyyy')).trim();
    const monthFolder = getOrCreateFolder(propertyFolder, monthName);

    const tenantName = (data.tenantName || 'Unknown Tenant').trim();
    const tenantFolder = getOrCreateFolder(monthFolder, tenantName);

    const file = tenantFolder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    const fileUrl = file.getUrl();
    if (!fileUrl) return createResponse(false, null, 'File created but no URL returned');

    return createResponse(true, {
      fileUrl: fileUrl,
      fileId: file.getId(),
      fileName: file.getName(),
      folderPath: [propertyName, monthName, tenantName].join(' / ')
    }, null);

  } catch (error) {
    return createResponse(false, null, 'Upload error: ' + error.toString());
  }
}

function getOrCreateFolder(parentFolder, folderName) {
  const safeName = folderName.replace(/[\\/:*?"<>|]/g, '_');
  const folders = parentFolder.getFoldersByName(safeName);
  if (folders.hasNext()) return folders.next();
  return parentFolder.createFolder(safeName);
}

// ================= LOAD =================
function handleLoad() {
  try {
    const ss = getSpreadsheet();
    const sheets = ss.getSheets();
    const data = {};

    sheets.forEach(sheet => {
      const sheetName = sheet.getName();
      const values = sheet.getDataRange().getValues();
      if (values.length <= 1) { data[sheetName] = []; return; }

      const headers = values[0].map(h => String(h).trim().toUpperCase());

      const col = {
        tenantName:     findColumn(headers, ['TENANT NAME','NAME','TENANT']),
        unitNo:         findColumn(headers, ['UNIT NO.','UNIT NO','UNIT','UNIT NUMBER']),
        prevReading:    findColumn(headers, ['PREVIOUS','METER READING PREVIOUS','PREV']),
        presentReading: findColumn(headers, ['CURRENT','METER PRESENT READING','PRESENT']),
        billingDate:    findColumn(headers, ['BILLING DATE','DATE']),
        consumption:    findColumn(headers, ['CONSUMPTION','CONSUMPTION (CU.M)','CU.M']),
        rate:           findColumn(headers, ['RATE','AMOUNT PER CUBIC METER','RATE PER CU.M']),
        total:          findColumn(headers, ['TOTAL','TOTAL AMOUNT']),
        status:         findColumn(headers, ['STATUS']),
        pdfUrl:         findColumn(headers, ['PDF','PDF URL','URL']),
        dateSent:       findColumn(headers, ['DATE SENT','SENT DATE']),
        balance:        findColumn(headers, ['BALANCE']),
        remarks:        findColumn(headers, ['REMARKS','REMARK'])
      };

      const rows = [];
      for (let i = 1; i < values.length; i++) {
        const row = values[i];
        if (row.every(c => c === '' || c === null || c === undefined)) continue;

        const rateVal = col.rate >= 0 ? parseFloat(row[col.rate]) : 0;

        const bill = {
          tenantName:     col.tenantName >= 0 ? String(row[col.tenantName] || '').trim() : '',
          unitNo:         col.unitNo >= 0 ? String(row[col.unitNo] || '').trim() : '',
          prevReading:    col.prevReading >= 0 ? Number(row[col.prevReading]) || 0 : 0,
          presentReading: col.presentReading >= 0 ? Number(row[col.presentReading]) || 0 : 0,
          billingDate:    col.billingDate >= 0 ? formatDateCell(row[col.billingDate]) : '',
          consumption:    col.consumption >= 0 ? Number(row[col.consumption]) || 0 : 0,
          rate:           isNaN(rateVal) ? 0 : rateVal,
          total:          col.total >= 0 ? Number(row[col.total]) || 0 : 0,
          status:         col.status >= 0 ? String(row[col.status] || 'Pending').trim() : 'Pending',
          pdfUrl:         col.pdfUrl >= 0 ? String(row[col.pdfUrl] || '').trim() : '',
          dateSent:       col.dateSent >= 0 ? formatDateCell(row[col.dateSent]) : '',
          balance:        col.balance >= 0 ? Number(row[col.balance]) || 0 : 0,
          remarks:        col.remarks >= 0 ? String(row[col.remarks] || '') : ''
        };

        // Force accurate total calculation on load
        bill.total = (bill.consumption * bill.rate) + bill.balance;

        rows.push(bill);
      }
      data[sheetName] = rows;
    });

    return createResponse(true, data, null);
  } catch (error) {
    return createResponse(false, null, 'Load error: ' + error.toString());
  }
}

// ================= SAVE =================
function handleSave(data) {
  try {
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName(data.tabName);

    if (!sheet) {
      sheet = ss.insertSheet(data.tabName);
      sheet.appendRow(['TENANT NAME', 'UNIT NO.', 'PREVIOUS', 'CURRENT', 'BILLING DATE', 'CONSUMPTION', 'RATE', 'TOTAL', 'STATUS', 'PDF', 'DATE SENT', 'REMARKS']);
    }
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['TENANT NAME', 'UNIT NO.', 'PREVIOUS', 'CURRENT', 'BILLING DATE', 'CONSUMPTION', 'RATE', 'TOTAL', 'STATUS', 'PDF', 'DATE SENT', 'REMARKS']);
    }
    ensureRemarksColumn(sheet);

    const rateVal = parseFloat(data.rate);
    const rate = isNaN(rateVal) ? 0 : rateVal;
    const consumption = Number(data.consumption) || 0;
    const balance = Number(data.balance) || 0;
    const total = (consumption * rate) + balance;

    sheet.appendRow([
      data.tenantName || '',
      data.unitNo || '',
      data.prevReading || 0,
      data.presentReading || 0,
      data.billingDate || '',
      consumption,
      rate,
      total,
      data.status || 'Pending',
      data.pdfUrl || '',
      data.dateSent || '',
      data.remarks || ''
    ]);
    return createResponse(true, { message: 'Saved' }, null);
  } catch (error) {
    return createResponse(false, null, 'Save error: ' + error.toString());
  }
}

// ================= UPDATE =================
function handleUpdate(data) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(data.tabName);
    if (!sheet) return createResponse(false, null, 'Sheet not found');

    const values = sheet.getDataRange().getValues();
    if (values.length <= 1) return createResponse(false, null, 'No data rows');

    const headers = values[0].map(h => String(h).trim().toUpperCase());
    const tCol = findColumn(headers, ['TENANT NAME','NAME','TENANT']);
    if (tCol < 0) return createResponse(false, null, 'Tenant column not found');

    let rowIdx = -1;
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][tCol] || '').trim() === data.tenantName.trim()) { rowIdx = i + 1; break; }
    }
    if (rowIdx === -1) return createResponse(false, null, 'Tenant not found');

    const col = {
      tenantName:     findColumn(headers, ['TENANT NAME','NAME','TENANT']),
      unitNo:         findColumn(headers, ['UNIT NO.','UNIT NO','UNIT','UNIT NUMBER']),
      prevReading:    findColumn(headers, ['PREVIOUS','METER READING PREVIOUS','PREV']),
      presentReading: findColumn(headers, ['CURRENT','METER PRESENT READING','PRESENT']),
      billingDate:    findColumn(headers, ['BILLING DATE','DATE']),
      consumption:    findColumn(headers, ['CONSUMPTION','CONSUMPTION (CU.M)','CU.M']),
      rate:           findColumn(headers, ['RATE','AMOUNT PER CUBIC METER','RATE PER CU.M']),
      total:          findColumn(headers, ['TOTAL','TOTAL AMOUNT']),
      status:         findColumn(headers, ['STATUS']),
      pdfUrl:         findColumn(headers, ['PDF','PDF URL','URL']),
      dateSent:       findColumn(headers, ['DATE SENT','SENT DATE']),
      balance:        findColumn(headers, ['BALANCE'])
    };
    col.remarks = ensureRemarksColumn(sheet);

    const rateVal = parseFloat(data.rate);
    const rate = isNaN(rateVal) ? 0 : rateVal;
    const consumption = Number(data.consumption) || 0;
    const balance = Number(data.balance) || 0;
    const total = (consumption * rate) + balance;

    if (col.tenantName >= 0) sheet.getRange(rowIdx, col.tenantName + 1).setValue(data.tenantName);
    if (col.unitNo >= 0) sheet.getRange(rowIdx, col.unitNo + 1).setValue(data.unitNo);
    if (col.prevReading >= 0) sheet.getRange(rowIdx, col.prevReading + 1).setValue(data.prevReading);
    if (col.presentReading >= 0) sheet.getRange(rowIdx, col.presentReading + 1).setValue(data.presentReading);
    if (col.billingDate >= 0) sheet.getRange(rowIdx, col.billingDate + 1).setValue(data.billingDate);
    if (col.consumption >= 0) sheet.getRange(rowIdx, col.consumption + 1).setValue(consumption);
    if (col.rate >= 0) sheet.getRange(rowIdx, col.rate + 1).setValue(rate);
    if (col.total >= 0) sheet.getRange(rowIdx, col.total + 1).setValue(total);
    if (col.status >= 0) sheet.getRange(rowIdx, col.status + 1).setValue(data.status);
    if (col.pdfUrl >= 0) sheet.getRange(rowIdx, col.pdfUrl + 1).setValue(data.pdfUrl || '');
    if (col.dateSent >= 0) sheet.getRange(rowIdx, col.dateSent + 1).setValue(data.dateSent || '');
    if (col.remarks >= 0) sheet.getRange(rowIdx, col.remarks + 1).setValue(data.remarks || '');

    return createResponse(true, { message: 'Updated' }, null);
  } catch (error) {
    return createResponse(false, null, 'Update error: ' + error.toString());
  }
}

// ================= DELETE =================
function handleDelete(data) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(data.tabName);
    if (!sheet) return createResponse(false, null, 'Sheet not found');

    const values = sheet.getDataRange().getValues();
    if (values.length <= 1) return createResponse(false, null, 'No data rows');

    const headers = values[0].map(h => String(h).trim().toUpperCase());
    const tCol = findColumn(headers, ['TENANT NAME','NAME','TENANT']);
    if (tCol < 0) return createResponse(false, null, 'Tenant column not found');

    let rowIdx = -1;
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][tCol] || '').trim() === data.tenantName.trim()) { rowIdx = i + 1; break; }
    }
    if (rowIdx === -1) return createResponse(false, null, 'Tenant not found');

    sheet.deleteRow(rowIdx);
    return createResponse(true, { message: 'Deleted' }, null);
  } catch (error) {
    return createResponse(false, null, 'Delete error: ' + error.toString());
  }
}

// ================= CLEAR TAB =================
function handleClearTab(data) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(data.tabName);
    if (!sheet) return createResponse(false, null, 'Sheet not found');
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);
    return createResponse(true, { message: 'Cleared' }, null);
  } catch (error) {
    return createResponse(false, null, 'Clear error: ' + error.toString());
  }
}

// ================= HELPERS =================
function createResponse(success, data, error) {
  return ContentService.createTextOutput(JSON.stringify({
    success: success, data: data || null, error: error || null
  })).setMimeType(ContentService.MimeType.JSON);
}

function findColumn(headers, names) {
  for (let i = 0; i < headers.length; i++) {
    const h = String(headers[i]).trim().toUpperCase();
    for (const n of names) {
      if (h === n.toUpperCase()) return i;
    }
  }
  return -1;
}

function ensureRemarksColumn(sheet) {
  const lastColumn = sheet.getLastColumn();
  if (lastColumn === 0) return -1;

  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(h => String(h).trim().toUpperCase());
  const remarksColumn = findColumn(headers, ['REMARKS', 'REMARK']);
  if (remarksColumn >= 0) return remarksColumn;

  sheet.getRange(1, lastColumn + 1).setValue('REMARKS');
  return lastColumn;
}

function formatDateCell(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const s = String(v).trim();
  if (!s) return '';
  const d = new Date(s);
  if (!isNaN(d.getTime())) return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return s;
}

// ================= SETUP =================
function setupSpreadsheet() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName('Sheet1') || ss.insertSheet('Sheet1');
  sheet.clear();
  sheet.appendRow(['TENANT NAME', 'UNIT NO.', 'PREVIOUS', 'CURRENT', 'BILLING DATE', 'CONSUMPTION', 'RATE', 'TOTAL', 'STATUS', 'PDF', 'DATE SENT', 'REMARKS']);
  sheet.getRange(1, 1, 1, 12).setFontWeight('bold').setBackground('#217346').setFontColor('#ffffff');
  sheet.setFrozenRows(1);
  return 'Setup complete';
}

function testConfiguration() {
  const ss = getSpreadsheet();
  const folder = DriveApp.getFolderById(ROOT_FOLDER_ID);
  return 'Spreadsheet: ' + ss.getName() + ' | Drive folder: ' + folder.getName();
}

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({
    success: true, message: 'Water Bill API running', timestamp: new Date().toISOString()
  })).setMimeType(ContentService.MimeType.JSON);
}
