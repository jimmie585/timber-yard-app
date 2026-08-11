/**
 * TIMBER YARD BACKEND — Google Apps Script (v2)
 * ------------------------------------------------
 * Uses ONE Google Sheet with multiple tabs:
 *   Stock, Movements, Customers, Users, SalesLines, SalesSummary, Payments
 *
 * SETUP:
 * 1. Create one Google Sheet (any name, e.g. "Timber Yard Data").
 * 2. Copy its ID from the URL: docs.google.com/spreadsheets/d/COPY_THIS_PART/edit
 * 3. Paste that ID into SHEET_ID below.
 * 4. Go to https://script.google.com -> New project. Delete starter code, paste this whole file in.
 *    (This must be a STANDALONE script, not bound to a sheet, since it opens the sheet by ID.)
 * 5. Edit BUSINESS_NAME / LOCATION / PHONE below.
 * 6. For SMS: sign up free at https://africastalking.com, grab your sandbox API key,
 *    paste into AT_API_KEY. Leave AT_USERNAME as 'sandbox' until you go live.
 * 7. Deploy > New deployment > Web app. Execute as: Me. Who has access: Anyone.
 * 8. Copy the Web app URL into APPS_SCRIPT_URL in index.html, employees.html, and boss.html.
 * 9. Any time you edit this file again: Deploy > Manage deployments > Edit > New version.
 */

// ====== CONFIGURATION — fill these in ======
var SHEET_ID = '1Ff50HG4QWgLG25BY6vU8lajQSQSrtsT9RQb5VrjgXlI'; // your one Google Sheet — all tabs live inside it

var BUSINESS_NAME = "Mathenge's Timberyard";
var BUSINESS_LOCATION = 'Nanyuki';
var BUSINESS_PHONE = '0724112335';

// Africa's Talking SMS credentials — https://africastalking.com (free sandbox to start)
var AT_USERNAME = 'sandbox';
var AT_API_KEY = 'PASTE_YOUR_AFRICASTALKING_API_KEY_HERE';
var AT_SENDER_ID = ''; // optional shortcode/sender ID, leave blank if you don't have one
var DEFAULT_COUNTRY_CODE = '254'; // Kenya. Change if needed.

// ====== ROUTES ======
function doGet(e) {
  ensureHeaders();
  var action = e.parameter.action || 'getData';
  if (action === 'getData') return respond(getAllData());
  return respond({ error: 'Unknown action' });
}

function doPost(e) {
  ensureHeaders();
  var body = JSON.parse(e.postData.contents);
  try {
    switch (body.action) {
      case 'addStockItem': return respond(addStockItem(body));
      case 'receiveStock': return respond(receiveStock(body));
      case 'adjustStock': return respond(adjustStock(body));
      case 'addSale': return respond(addSale(body));
      case 'addPayment': return respond(addPayment(body));
      case 'addCustomer': return respond(addCustomer(body));
      case 'emailReceipt': return respond(emailReceipt(body));
      case 'login': return respond(login(body));
      case 'addUser': return respond(addUser(body));
      default: return respond({ error: 'Unknown action: ' + body.action });
    }
  } catch (err) {
    return respond({ error: err.message });
  }
}

function respond(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ====== SHEET HELPERS ======
function stockSS() { return SpreadsheetApp.openById(SHEET_ID); }
function customerSS() { return SpreadsheetApp.openById(SHEET_ID); }

function stockSheet(name) {
  var s = stockSS().getSheetByName(name);
  if (!s) s = stockSS().insertSheet(name);
  return s;
}
function custSheet(name) {
  var s = customerSS().getSheetByName(name);
  if (!s) s = customerSS().insertSheet(name);
  return s;
}

function ensureHeaders() {
  var stockDefs = {
    Stock: ['ID', 'Name', 'Unit', 'Qty', 'Price', 'LowThreshold'],
    Movements: ['ID', 'StockName', 'Type', 'Qty', 'BalanceAfter', 'Date', 'Note']
  };
  Object.keys(stockDefs).forEach(function (name) {
    var s = stockSheet(name);
    if (s.getRange(1, 1).getValue() === '') s.appendRow(stockDefs[name]);
  });
  var custDefs = {
    Customers: ['ID', 'Name', 'Phone', 'Email'],
    Users: ['ID', 'Username', 'Password', 'Role', 'Name'],
    SalesLines: ['SaleID', 'ReceiptNo', 'Date', 'CustomerName', 'Phone', 'Item', 'Qty', 'UnitPrice', 'Subtotal'],
    SalesSummary: ['SaleID', 'ReceiptNo', 'Date', 'CustomerName', 'Phone', 'Email', 'Total', 'Paid', 'Balance', 'Status'],
    Payments: ['ID', 'SaleID', 'ReceiptNo', 'Amount', 'Date', 'Method']
  };
  Object.keys(custDefs).forEach(function (name) {
    var s = custSheet(name);
    if (s.getRange(1, 1).getValue() === '') s.appendRow(custDefs[name]);
  });
}

function sheetToObjects(sheetObj) {
  var rows = sheetObj.getDataRange().getValues();
  if (rows.length < 2) return [];
  var headers = rows[0];
  var out = [];
  for (var i = 1; i < rows.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) obj[headers[j]] = rows[i][j];
    out.push(obj);
  }
  return out;
}

function getAllData() {
  var users = sheetToObjects(custSheet('Users')).map(function (u) {
    return { ID: u.ID, Username: u.Username, Role: u.Role, Name: u.Name }; // never send passwords to the client
  });
  return {
    stock: sheetToObjects(stockSheet('Stock')),
    movements: sheetToObjects(stockSheet('Movements')),
    customers: sheetToObjects(custSheet('Customers')),
    users: users,
    salesSummary: sheetToObjects(custSheet('SalesSummary')),
    salesLines: sheetToObjects(custSheet('SalesLines')),
    payments: sheetToObjects(custSheet('Payments'))
  };
}

function login(body) {
  var rows = custSheet('Users').getDataRange().getValues();
  var uname = String(body.username || '').trim().toLowerCase();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][1]).toLowerCase() === uname && String(rows[i][2]) === String(body.password)) {
      return { success: true, role: rows[i][3], name: rows[i][4], username: rows[i][1] };
    }
  }
  return { success: false, error: 'Incorrect username or password' };
}

function addUser(body) {
  var rows = custSheet('Users').getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][1]).toLowerCase() === String(body.username).toLowerCase()) {
      return { success: false, error: 'That username already exists' };
    }
  }
  custSheet('Users').appendRow([nextId('usr'), body.username, body.password, body.role || 'employee', body.name || body.username]);
  return { success: true, data: getAllData() };
}

function nextId(prefix) { return prefix + '_' + new Date().getTime() + '_' + Math.floor(Math.random() * 1000); }

function logMovement(stockName, type, qty, balanceAfter, note) {
  stockSheet('Movements').appendRow([nextId('mv'), stockName, type, qty, balanceAfter, new Date(), note || '']);
}

// ====== STOCK ======
function addStockItem(body) {
  var id = nextId('stk');
  var qty = Number(body.qty) || 0;
  stockSheet('Stock').appendRow([id, body.name, body.unit, qty, Number(body.price) || 0, Number(body.lowThreshold) || 0]);
  if (qty > 0) logMovement(body.name, 'in', qty, qty, 'New item added');
  return getAllData();
}

function receiveStock(body) {
  var s = stockSheet('Stock');
  var rows = s.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === body.stockId) {
      var newQty = Number(rows[i][3]) + Number(body.qty);
      s.getRange(i + 1, 4).setValue(newQty);
      logMovement(rows[i][1], 'in', Number(body.qty), newQty, body.note || 'Stock received');
      break;
    }
  }
  return getAllData();
}

function adjustStock(body) {
  var s = stockSheet('Stock');
  var rows = s.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === body.stockId) {
      var oldQty = Number(rows[i][3]);
      var newQty = Number(body.qty);
      var diff = newQty - oldQty;
      s.getRange(i + 1, 4).setValue(newQty);
      s.getRange(i + 1, 5).setValue(Number(body.price));
      s.getRange(i + 1, 6).setValue(Number(body.lowThreshold));
      if (diff !== 0) logMovement(rows[i][1], diff > 0 ? 'in' : 'out', Math.abs(diff), newQty, body.note || 'Manual adjustment');
      break;
    }
  }
  return getAllData();
}

// ====== CUSTOMERS ======
function findOrCreateCustomer(name, phone, email) {
  var s = custSheet('Customers');
  var rows = s.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][1]).toLowerCase() === String(name).toLowerCase()) {
      if (email && !rows[i][3]) s.getRange(i + 1, 4).setValue(email);
      return rows[i][0];
    }
  }
  var id = nextId('cus');
  s.appendRow([id, name, phone || '', email || '']);
  return id;
}

function addCustomer(body) {
  custSheet('Customers').appendRow([nextId('cus'), body.name, body.phone || '', body.email || '']);
  return getAllData();
}

// ====== SALES ======
function getNextReceiptNo() {
  var rows = custSheet('SalesSummary').getDataRange().getValues();
  var max = 1000;
  for (var i = 1; i < rows.length; i++) { var n = Number(rows[i][1]); if (n > max) max = n; }
  return max + 1;
}

function addSale(body) {
  findOrCreateCustomer(body.customerName, body.phone, body.email);
  var receiptNo = getNextReceiptNo();
  var saleId = nextId('sale');
  var total = 0;
  var emailItems = [];
  var stockData = stockSheet('Stock');

  body.items.forEach(function (li) {
    var subtotal = Number(li.qty) * Number(li.price);
    total += subtotal;
    emailItems.push({ name: li.name, qty: li.qty, subtotal: subtotal });
    custSheet('SalesLines').appendRow([saleId, receiptNo, new Date(), body.customerName, body.phone || '', li.name, li.qty, li.price, subtotal]);
    var rows = stockData.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0] === li.stockId) {
        var newQty = Number(rows[i][3]) - Number(li.qty);
        stockData.getRange(i + 1, 4).setValue(newQty);
        logMovement(li.name, 'out', li.qty, newQty, 'Sale #' + receiptNo);
        break;
      }
    }
  });

  var paid = Math.min(Number(body.paid) || 0, total);
  var balance = total - paid;
  var status = balance <= 0 ? 'paid' : (paid > 0 ? 'partial' : 'unpaid');
  custSheet('SalesSummary').appendRow([saleId, receiptNo, new Date(), body.customerName, body.phone || '', body.email || '', total, paid, balance, status]);
  if (paid > 0) {
    custSheet('Payments').appendRow([nextId('pay'), saleId, receiptNo, paid, new Date(), 'Initial payment']);
    if (body.phone) {
      sendSms(body.phone, BUSINESS_NAME + ': Payment of KES ' + Math.round(paid).toLocaleString() + ' received. Receipt #' + receiptNo + '. Balance: KES ' + Math.round(balance).toLocaleString() + '. Thank you.');
    }
  }

  var emailSent = false;
  if (body.email) emailSent = sendReceiptEmail(body.email, receiptNo, body.customerName, emailItems, total, paid, balance, status);

  return { data: getAllData(), saleId: saleId, receiptNo: receiptNo, emailSent: emailSent };
}

function addPayment(body) {
  var s = custSheet('SalesSummary');
  var rows = s.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === body.saleId) {
      var newPaid = Number(rows[i][7]) + Number(body.amount);
      var total = Number(rows[i][6]);
      var newBalance = total - newPaid;
      var phone = rows[i][4];
      var receiptNo = rows[i][1];
      s.getRange(i + 1, 8).setValue(newPaid);
      s.getRange(i + 1, 9).setValue(newBalance);
      s.getRange(i + 1, 10).setValue(newBalance <= 0 ? 'paid' : (newPaid > 0 ? 'partial' : 'unpaid'));
      custSheet('Payments').appendRow([nextId('pay'), body.saleId, receiptNo, body.amount, new Date(), body.method || '']);
      if (phone) {
        sendSms(phone, BUSINESS_NAME + ': Payment of KES ' + Math.round(Number(body.amount)).toLocaleString() + ' received for receipt #' + receiptNo + '. New balance: KES ' + Math.round(newBalance).toLocaleString() + '. Thank you.');
      }
      break;
    }
  }
  return getAllData();
}

// ====== EMAIL RECEIPT ======
function emailReceipt(body) {
  var summaryRows = custSheet('SalesSummary').getDataRange().getValues();
  var sale = null;
  for (var i = 1; i < summaryRows.length; i++) {
    if (summaryRows[i][0] === body.saleId) {
      sale = { receiptNo: summaryRows[i][1], customerName: summaryRows[i][3], total: summaryRows[i][6], paid: summaryRows[i][7], balance: summaryRows[i][8], status: summaryRows[i][9] };
      break;
    }
  }
  if (!sale) return { error: 'Sale not found' };
  var lineRows = custSheet('SalesLines').getDataRange().getValues();
  var items = [];
  for (var j = 1; j < lineRows.length; j++) {
    if (lineRows[j][0] === body.saleId) items.push({ name: lineRows[j][5], qty: lineRows[j][6], subtotal: lineRows[j][8] });
  }
  sendReceiptEmail(body.email, sale.receiptNo, sale.customerName, items, sale.total, sale.paid, sale.balance, sale.status);
  return { success: true };
}

function sendReceiptEmail(toEmail, receiptNo, customerName, items, total, paid, balance, status) {
  if (!toEmail) return false;
  var subject = BUSINESS_NAME + ' \u2014 Receipt #' + receiptNo;
  var rowsHtml = items.map(function (li) {
    return '<tr><td style="padding:6px 10px;border-bottom:1px solid #eee;">' + li.name + ' &times; ' + li.qty + '</td>' +
      '<td style="padding:6px 10px;border-bottom:1px solid #eee;text-align:right;">KES ' + Number(li.subtotal).toLocaleString() + '</td></tr>';
  }).join('');
  var statusColor = status === 'paid' ? '#43604B' : '#B5482C';
  var html =
    '<div style="font-family:Georgia,serif;max-width:420px;margin:0 auto;border:1px solid #E3D6C1;border-radius:8px;padding:24px;background:#FBF7EF;color:#2B1B12;">' +
    '<h2 style="text-align:center;margin:0 0 4px;">' + BUSINESS_NAME + '</h2>' +
    '<p style="text-align:center;color:#7C6A58;font-size:12px;margin:0 0 18px;">' + BUSINESS_LOCATION + (BUSINESS_PHONE ? ' &middot; ' + BUSINESS_PHONE : '') + '</p>' +
    '<p style="font-size:13px;color:#7C6A58;">Receipt #' + receiptNo + ' &middot; ' + new Date().toLocaleDateString() + '</p>' +
    '<p style="font-size:14px;">Customer: <strong>' + customerName + '</strong></p>' +
    '<table style="width:100%;border-collapse:collapse;font-size:13px;margin-top:10px;">' + rowsHtml + '</table>' +
    '<div style="margin-top:14px;padding-top:10px;border-top:1px dashed #E3D6C1;font-size:14px;">' +
    '<p style="font-weight:bold;">Total: KES ' + Number(total).toLocaleString() + '</p>' +
    '<p>Paid: KES ' + Number(paid).toLocaleString() + '</p>' +
    '<p style="color:' + statusColor + ';font-weight:bold;">Balance: KES ' + Number(balance).toLocaleString() + ' (' + status.toUpperCase() + ')</p>' +
    '</div>' +
    '<p style="text-align:center;font-size:11px;color:#7C6A58;margin-top:20px;">Thank you for your business.</p>' +
    '</div>';
  MailApp.sendEmail({ to: toEmail, subject: subject, htmlBody: html });
  return true;
}

// ====== SMS (Africa's Talking) ======
function normalizePhone(phone) {
  var p = String(phone).replace(/[^0-9+]/g, '');
  if (p.indexOf('+') === 0) return p;
  if (p.indexOf('0') === 0) return '+' + DEFAULT_COUNTRY_CODE + p.substring(1);
  if (p.indexOf(DEFAULT_COUNTRY_CODE) === 0) return '+' + p;
  return '+' + p;
}

function sendSms(phone, message) {
  if (!phone || !AT_API_KEY || AT_API_KEY.indexOf('PASTE_YOUR') === 0) return false;
  try {
    var url = (AT_USERNAME === 'sandbox')
      ? 'https://api.sandbox.africastalking.com/version1/messaging'
      : 'https://api.africastalking.com/version1/messaging';
    var payload = { username: AT_USERNAME, to: normalizePhone(phone), message: message };
    if (AT_SENDER_ID) payload.from = AT_SENDER_ID;
    var options = {
      method: 'post',
      payload: payload,
      headers: { apiKey: AT_API_KEY, Accept: 'application/json' },
      muteHttpExceptions: true
    };
    UrlFetchApp.fetch(url, options);
    return true;
  } catch (e) {
    return false;
  }
}
