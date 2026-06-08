// ========================================
// CULAC Invoice System - Google Apps Script Backend
// Version: 2.1 (Fixed queryString fallback)
// Updated: June 2026
// ========================================

const SHEET_NAME = 'Invoices';
const PAYMENT_SHEET = 'Payments';

// ========================================
// Parse parameters with queryString fallback
// ========================================
function getParams(e) {
  if (e && e.parameter && Object.keys(e.parameter).length > 0) {
    return e.parameter;
  }
  // Fallback: parse queryString manually
  if (e && e.queryString) {
    const params = {};
    e.queryString.split('&').forEach(function(pair) {
      const parts = pair.split('=');
      if (parts.length >= 1 && parts[0]) {
        params[decodeURIComponent(parts[0])] = decodeURIComponent(parts[1] || '');
      }
    });
    return params;
  }
  return {};
}

// ========================================
// doGet
// ========================================
function doGet(e) {
  try {
    const params = getParams(e);
    const action = params.action;

    if (!action) {
      return jsonOut({ success: false, error: 'Action parameter is required', debug: { queryString: e ? e.queryString : null, parameter: e ? e.parameter : null } });
    }

    switch (action) {
      case 'test':
        return testConnection();
      case 'list':
        return getInvoiceList();
      case 'get':
        if (!params.id) return jsonOut({ success: false, error: 'ID parameter is required' });
        return getInvoice(params.id);
      case 'getNextNumber':
        return getNextInvoiceNumber();
      default:
        return jsonOut({ success: false, error: 'Invalid action: ' + action });
    }
  } catch (error) {
    Logger.log('Error in doGet: ' + error.toString());
    return jsonOut({ success: false, error: 'Server error: ' + error.toString() });
  }
}

// ========================================
// doPost
// ========================================
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonOut({ success: false, error: 'No POST data provided' });
    }

    const data = JSON.parse(e.postData.contents);
    const action = data.action;

    if (!action) return jsonOut({ success: false, error: 'Action is required in POST data' });

    switch (action) {
      case 'create':
        return createInvoice(data);
      case 'update':
        return updateInvoice(data);
      case 'delete':
        return deleteInvoice(data);
      case 'recordPayment':
        return recordPayment(data);
      default:
        return jsonOut({ success: false, error: 'Invalid action: ' + action });
    }
  } catch (error) {
    Logger.log('Error in doPost: ' + error.toString());
    return jsonOut({ success: false, error: 'Server error: ' + error.toString() });
  }
}

// ========================================
// Helper
// ========================================
function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ========================================
// Test Connection
// ========================================
function testConnection() {
  return jsonOut({
    success: true,
    message: 'Connection successful!',
    timestamp: new Date().toISOString(),
    version: '2.1'
  });
}

// ========================================
// Invoice Functions
// ========================================
function createInvoice(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    setupInvoiceSheet(sheet);
  }

  const id = Utilities.getUuid();
  const timestamp = new Date().toISOString();

  sheet.appendRow([
    id,
    data.invoiceNo,
    data.date,
    data.customer,
    JSON.stringify(data.items),
    data.total,
    data.note || '',
    data.bankName,
    data.bankBranch,
    data.bankAcc,
    data.bankUser,
    'unpaid',
    '',
    '',
    timestamp
  ]);

  return jsonOut({ success: true, id: id, invoiceNo: data.invoiceNo });
}

function getInvoiceList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet || sheet.getLastRow() <= 1) {
    return jsonOut({ success: true, invoices: [] });
  }

  const data = sheet.getDataRange().getValues();
  const invoices = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[0]) {
      invoices.push({
        id: row[0],
        invoiceNo: row[1],
        date: row[2],
        customer: row[3],
        items: row[4],
        total: row[5],
        note: row[6] || '',
        bankName: row[7],
        bankBranch: row[8],
        bankAccount: row[9],
        bankUser: row[10],
        paymentStatus: row[11] || 'unpaid',
        paymentDate: row[12] || '',
        paymentDocNo: row[13] || '',
        createdAt: row[14]
      });
    }
  }

  invoices.sort(function(a, b) { return new Date(b.date) - new Date(a.date); });

  return jsonOut({ success: true, invoices: invoices });
}

function getInvoice(id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) return jsonOut({ success: false, error: 'Sheet not found' });

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      const row = data[i];
      return jsonOut({
        success: true,
        invoice: {
          id: row[0], invoiceNo: row[1], date: row[2], customer: row[3],
          items: row[4], total: row[5], note: row[6] || '',
          bankName: row[7], bankBranch: row[8], bankAccount: row[9], bankUser: row[10],
          paymentStatus: row[11] || 'unpaid', paymentDate: row[12] || '',
          paymentDocNo: row[13] || '', createdAt: row[14]
        }
      });
    }
  }

  return jsonOut({ success: false, error: 'Invoice not found' });
}

function updateInvoice(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) return jsonOut({ success: false, error: 'Sheet not found' });

  const sheetData = sheet.getDataRange().getValues();

  for (let i = 1; i < sheetData.length; i++) {
    if (sheetData[i][0] === data.id) {
      const row = i + 1;
      sheet.getRange(row, 2).setValue(data.invoiceNo);
      sheet.getRange(row, 3).setValue(data.date);
      sheet.getRange(row, 4).setValue(data.customer);
      sheet.getRange(row, 5).setValue(JSON.stringify(data.items));
      sheet.getRange(row, 6).setValue(data.total);
      sheet.getRange(row, 7).setValue(data.note || '');
      sheet.getRange(row, 8).setValue(data.bankName);
      sheet.getRange(row, 9).setValue(data.bankBranch);
      sheet.getRange(row, 10).setValue(data.bankAcc);
      sheet.getRange(row, 11).setValue(data.bankUser);
      return jsonOut({ success: true, id: data.id });
    }
  }

  return jsonOut({ success: false, error: 'Invoice not found' });
}

// ========================================
// Delete Invoice
// ========================================
function deleteInvoice(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) return jsonOut({ success: false, error: 'Sheet not found' });

  const sheetData = sheet.getDataRange().getValues();

  for (let i = 1; i < sheetData.length; i++) {
    if (sheetData[i][0] === data.id) {
      sheet.deleteRow(i + 1);
      return jsonOut({ success: true, id: data.id });
    }
  }

  return jsonOut({ success: false, error: 'Invoice not found' });
}

// ========================================
// Payment Functions
// ========================================
function recordPayment(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) return jsonOut({ success: false, error: 'Sheet not found' });

  const sheetData = sheet.getDataRange().getValues();

  for (let i = 1; i < sheetData.length; i++) {
    if (sheetData[i][0] === data.id) {
      const row = i + 1;
      sheet.getRange(row, 12).setValue('paid');
      sheet.getRange(row, 13).setValue(data.paymentDate);
      sheet.getRange(row, 14).setValue(data.paymentDocNo);

      logPayment({
        id: data.id,
        invoiceNo: sheetData[i][1],
        paymentDate: data.paymentDate,
        paymentDocNo: data.paymentDocNo,
        amount: sheetData[i][5]
      });

      return jsonOut({ success: true, id: data.id });
    }
  }

  return jsonOut({ success: false, error: 'Invoice not found' });
}

function logPayment(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(PAYMENT_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(PAYMENT_SHEET);
    setupPaymentSheet(sheet);
  }

  const timestamp = new Date().toISOString();
  const user = Session.getActiveUser().getEmail();

  sheet.appendRow([
    data.id, data.invoiceNo, data.paymentDate,
    data.paymentDocNo, data.amount, timestamp, user
  ]);
}

// ========================================
// Invoice Number Generator
// ========================================
function getNextInvoiceNumber() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);

  const now = new Date();
  const month = String(now.getMonth() + 1).padStart(2, '0');
  const year = String(now.getFullYear() + 543).slice(-2);
  const currentMonthYear = month + year;

  if (!sheet || sheet.getLastRow() <= 1) {
    return jsonOut({ success: true, nextNumber: 1, invoiceNo: '01' + currentMonthYear });
  }

  const data = sheet.getDataRange().getValues();
  let maxSequence = 0;

  for (let i = 1; i < data.length; i++) {
    const invoiceNo = String(data[i][1]);
    if (invoiceNo && invoiceNo.length >= 6) {
      const invoiceMonthYear = invoiceNo.slice(-4);
      if (invoiceMonthYear === currentMonthYear) {
        const sequence = parseInt(invoiceNo.substring(0, 2));
        if (!isNaN(sequence) && sequence > maxSequence) {
          maxSequence = sequence;
        }
      }
    }
  }

  const nextSequence = maxSequence + 1;
  const nextInvoiceNo = String(nextSequence).padStart(2, '0') + currentMonthYear;

  return jsonOut({ success: true, nextNumber: nextSequence, invoiceNo: nextInvoiceNo });
}

// ========================================
// Setup Functions
// ========================================
function setupInvoiceSheet(sheet) {
  sheet.appendRow([
    'ID', 'Invoice No', 'Date', 'Customer', 'Items (JSON)',
    'Total', 'Note', 'Bank Name', 'Bank Branch', 'Bank Account',
    'Bank User', 'Payment Status', 'Payment Date', 'Payment Doc No', 'Created At'
  ]);
  const headerRange = sheet.getRange('A1:O1');
  headerRange.setBackground('#d81b60');
  headerRange.setFontColor('white');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(2, 100);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 200);
  sheet.setColumnWidth(5, 300);
  sheet.setColumnWidth(6, 100);
  sheet.setColumnWidth(7, 200);
  sheet.setColumnWidth(12, 100);
  sheet.setColumnWidth(13, 100);
  sheet.setColumnWidth(14, 120);
}

function setupPaymentSheet(sheet) {
  sheet.appendRow([
    'Invoice ID', 'Invoice No', 'Payment Date',
    'Payment Doc No', 'Amount', 'Recorded At', 'Recorded By'
  ]);
  const headerRange = sheet.getRange('A1:G1');
  headerRange.setBackground('#d81b60');
  headerRange.setFontColor('white');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(2, 100);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 120);
  sheet.setColumnWidth(5, 100);
  sheet.setColumnWidth(6, 150);
  sheet.setColumnWidth(7, 200);
}

function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let invoiceSheet = ss.getSheetByName(SHEET_NAME);
  if (!invoiceSheet) {
    invoiceSheet = ss.insertSheet(SHEET_NAME);
    setupInvoiceSheet(invoiceSheet);
    Logger.log('✅ Invoices sheet created');
  } else {
    Logger.log('ℹ️ Invoices sheet already exists');
  }

  let paymentSheet = ss.getSheetByName(PAYMENT_SHEET);
  if (!paymentSheet) {
    paymentSheet = ss.insertSheet(PAYMENT_SHEET);
    setupPaymentSheet(paymentSheet);
    Logger.log('✅ Payments sheet created');
  } else {
    Logger.log('ℹ️ Payments sheet already exists');
  }

  Logger.log('🎉 Setup completed successfully!');

  SpreadsheetApp.getUi().alert(
    'Setup Complete',
    'CULAC Invoice System sheets have been set up successfully!\n\n' +
    '✅ Invoices sheet\n✅ Payments sheet\n\n' +
    'You can now deploy this as a Web App.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}
