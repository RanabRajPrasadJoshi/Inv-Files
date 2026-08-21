/**
 * WOOLLY & CO. — Google Apps Script backend
 * ------------------------------------------------------------
 * This file is the ONLY backend the site has. It reads/writes the
 * Google Sheet directly and is the single source of truth for
 * price and stock. The frontend never trusts its own numbers —
 * every order is re-validated against the sheet in here.
 *
 * DEPLOY AS: Extensions > Apps Script > Deploy > New deployment
 *            Type: Web app
 *            Execute as: Me
 *            Who has access: Anyone
 * Copy the resulting /exec URL into script.js -> CONFIG.API_URL
 * ------------------------------------------------------------
 */

// ====== CONFIG ==============================================
const SHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
const PRODUCTS_SHEET_NAME = 'Products';
const ORDERS_SHEET_NAME = 'Orders';
const OWNER_EMAIL = 'ranabjoshi20@gmail.com'; 
const STORE_NAME = 'Woolly & Co.';

// Column order for the Products sheet (row 1 = headers, must match exactly)
const PRODUCT_COLS = [
  'ProductID', 'ProductName', 'Category', 'Price', 'Quantity',
  'ImageURL', 'Description', 'Featured', 'New'
];

// Column order for the Orders sheet
const ORDER_COLS = [
  'Timestamp', 'OrderID', 'ProductID', 'ProductName', 'Quantity',
  'Price', 'Total', 'CustomerName', 'Email', 'Phone', 'Address',
  'Note', 'OrderStatus'
];

// ====== ENTRY POINTS =========================================

function doGet(e) {
  try {
    const action = e.parameter.action;
    if (action === 'products') {
      return jsonOut({ ok: true, products: getAllProducts() });
    }
    return jsonOut({ ok: false, error: 'Unknown action.' });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    // Body is sent as text/plain JSON from the browser to dodge the
    // CORS preflight that Apps Script web apps can't answer.
    const body = JSON.parse(e.postData.contents || '{}');
    const action = body.action || e.parameter.action;

    if (action === 'createOrder') {
      return jsonOut(createOrder(body));
    }
    return jsonOut({ ok: false, error: 'Unknown action.' });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ====== PRODUCTS ==============================================

function getProductsSheet_() {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(PRODUCTS_SHEET_NAME);
}

function getOrdersSheet_() {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(ORDERS_SHEET_NAME);
}

/** Reads every product row into an array of plain objects. */
function getAllProducts() {
  const sheet = getProductsSheet_();
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(h => String(h).trim());
  const rows = values.slice(1);

  return rows
    .filter(r => r.some(cell => cell !== '' && cell !== null)) // skip blank rows
    .map(r => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = r[i]);
      return {
        id: String(obj.ProductID),
        name: String(obj.ProductName || ''),
        category: String(obj.Category || ''),
        price: Number(obj.Price) || 0,
        quantity: Math.max(0, Math.floor(Number(obj.Quantity) || 0)),
        image: String(obj.ImageURL || ''),
        description: String(obj.Description || ''),
        featured: truthy_(obj.Featured),
        isNew: truthy_(obj.New)
      };
    });
}

function truthy_(v) {
  const s = String(v).trim().toLowerCase();
  return s === 'true' || s === 'yes' || s === '1' || s === 'y';
}

/** Finds a product's row index (1-based, includes header) by ProductID. */
function findProductRow_(sheet, productId) {
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(h => String(h).trim());
  const idCol = headers.indexOf('ProductID');
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idCol]) === String(productId)) {
      return { rowIndex: i + 1, headers: headers };
    }
  }
  return null;
}

// ====== ORDERS ==================================================

/**
 * Validates stock/price straight from the sheet, reserves the stock,
 * writes the order row, emails the owner, and returns the result.
 * Wrapped in LockService so two simultaneous buyers can't both win
 * the last unit.
 */
function createOrder(payload) {
  const productId = String(payload.productId || '').trim();
  const qtyRequested = Math.floor(Number(payload.quantity)) || 0;
  const customer = payload.customer || {};

  if (!productId) return { ok: false, error: 'Missing product.' };
  if (qtyRequested < 1) return { ok: false, error: 'Quantity must be at least 1.' };
  if (!customer.name || !customer.email || !customer.phone || !customer.address) {
    return { ok: false, error: 'Please fill in all required fields.' };
  }

  const lock = LockService.getScriptLock();
  const gotLock = lock.tryLock(10000); // wait up to 10s for other orders to finish
  if (!gotLock) {
    return { ok: false, error: 'The store is busy right now. Please try again in a moment.' };
  }

  try {
    const sheet = getProductsSheet_();
    const found = findProductRow_(sheet, productId);
    if (!found) return { ok: false, error: 'That product could not be found.' };

    const { rowIndex, headers } = found;
    const priceCol = headers.indexOf('Price') + 1;
    const qtyCol = headers.indexOf('Quantity') + 1;
    const nameCol = headers.indexOf('ProductName') + 1;

    const currentPrice = Number(sheet.getRange(rowIndex, priceCol).getValue()) || 0;
    const currentQty = Math.floor(Number(sheet.getRange(rowIndex, qtyCol).getValue())) || 0;
    const productName = String(sheet.getRange(rowIndex, nameCol).getValue());

    if (currentQty <= 0) {
      return { ok: false, error: 'This product is now out of stock.' };
    }
    if (qtyRequested > currentQty) {
      return { ok: false, error: `Only ${currentQty} item(s) are available.` };
    }

    // ---- authoritative numbers, computed here, never from the browser ----
    const total = round2_(currentPrice * qtyRequested);
    const remaining = currentQty - qtyRequested;
    const orderId = generateOrderId_();

    // 1. reduce stock
    sheet.getRange(rowIndex, qtyCol).setValue(remaining);

    // 2. log the order
    const ordersSheet = getOrdersSheet_();
    ordersSheet.appendRow(ORDER_COLS.map(col => {
      switch (col) {
        case 'Timestamp': return new Date();
        case 'OrderID': return orderId;
        case 'ProductID': return productId;
        case 'ProductName': return productName;
        case 'Quantity': return qtyRequested;
        case 'Price': return currentPrice;
        case 'Total': return total;
        case 'CustomerName': return customer.name;
        case 'Email': return customer.email;
        case 'Phone': return customer.phone;
        case 'Address': return customer.address;
        case 'Note': return customer.note || '';
        case 'OrderStatus': return 'Pending';
        default: return '';
      }
    }));

    // 3. notify the owner (best-effort — don't fail the order over email trouble)
    try {
      sendOwnerEmail_({
        orderId, productId, productName, qtyRequested, currentPrice, total,
        remaining, customer
      });
    } catch (mailErr) {
      Logger.log('Email failed: ' + mailErr);
    }

    return {
      ok: true,
      orderId: orderId,
      product: productName,
      quantity: qtyRequested,
      price: currentPrice,
      total: total,
      remainingStock: remaining
    };
  } finally {
    lock.releaseLock();
  }
}

function round2_(n) {
  return Math.round(n * 100) / 100;
}

function generateOrderId_() {
  const now = new Date();
  const y = now.getFullYear();
  const m = String(now.getMonth() + 1).padStart(2, '0');
  const d = String(now.getDate()).padStart(2, '0');
  // A short random suffix keeps IDs unique even if two orders land in
  // the same second; the lock above already prevents stock races.
  const suffix = Math.floor(100 + Math.random() * 900);
  return `ORD-${y}${m}${d}-${suffix}`;
}

function sendOwnerEmail_(o) {
  const subject = `New order ${o.orderId} — ${STORE_NAME}`;
  const body = [
    `You've got a new order on ${STORE_NAME}.`,
    ``,
    `Order ID: ${o.orderId}`,
    `Product: ${o.productName} (${o.productId})`,
    `Quantity: ${o.qtyRequested}`,
    `Price: Rs. ${o.currentPrice}`,
    `Total: Rs. ${o.total}`,
    `Remaining stock: ${o.remaining}`,
    `Order status: Pending`,
    ``,
    `Customer: ${o.customer.name}`,
    `Email: ${o.customer.email}`,
    `Phone: ${o.customer.phone}`,
    `Address: ${o.customer.address}`,
    `Note: ${o.customer.note || '(none)'}`
  ].join('\n');

  MailApp.sendEmail({
    to: OWNER_EMAIL,
    replyTo: o.customer.email,
    subject: subject,
    body: body
  });
}

// ====== ONE-TIME SETUP HELPER =====================================
// Run this once from the Apps Script editor (select it in the function
// dropdown and press Run) to create both sheets with correct headers
// if they don't already exist.
function setupSheets() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  if (!ss.getSheetByName(PRODUCTS_SHEET_NAME)) {
    const p = ss.insertSheet(PRODUCTS_SHEET_NAME);
    p.appendRow(PRODUCT_COLS);
    p.appendRow(['P001', 'Felted Wool Tote', 'Bags', 2500, 12,
      'https://images.example.com/tote.jpg',
      'Hand-felted merino wool tote bag, lined with cotton.', 'TRUE', 'TRUE']);
    p.setFrozenRows(1);
  }

  if (!ss.getSheetByName(ORDERS_SHEET_NAME)) {
    const o = ss.insertSheet(ORDERS_SHEET_NAME);
    o.appendRow(ORDER_COLS);
    o.setFrozenRows(1);
  }
}
