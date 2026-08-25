/**
 * WOOLLY & CO. — Google Apps Script backend
 * ------------------------------------------------------------
 * This file is the ONLY backend the site has. It reads/writes the
 * Google Sheet directly and is the single source of truth for
 * price and stock. The frontend never trusts its own numbers —
 * every order is re-validated against the sheet in here.
 *
 * SHEET LAYOUT THIS EXPECTS (Products tab, row 1 = these exact headers):
 *   Image | Code | Categories and ITEMS name | Rate | Stock QTY | Price | Import QTY | Description | Featured
 *
 *   - Code                          -> the product's unique ID (e.g. F-DC-001)
 *   - Categories and ITEMS name     -> "Category | Item name", e.g. "Felted Deco | Buddha"
 *   - Rate                          -> the selling price shown on the site
 *   - Stock QTY                     -> live stock count
 *   - Price                         -> Rate × Stock QTY, your own inventory-value column (site ignores this)
 *   - Import QTY                    -> restock quantity; any value > 0 marks the item as "New" on the site
 *   - Image                         -> a direct image URL (NOT an inserted/pasted picture — see README)
 *   - Description / Featured        -> optional. Safe to leave blank.
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
const OWNER_EMAIL = 'owner@example.com'; // <-- change to the store owner's inbox
const STORE_NAME = 'Woolly & Co.';

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

/** Splits "Category | Item name" into { category, name }. Tolerates a
 *  missing "|" by treating the whole string as the name. */
function parseCategoryAndName_(raw) {
  const str = String(raw || '').trim();
  const idx = str.indexOf('|');
  if (idx === -1) return { category: '', name: str };
  return {
    category: str.slice(0, idx).trim(),
    name: str.slice(idx + 1).trim()
  };
}

function truthy_(v) {
  const s = String(v).trim().toLowerCase();
  return s === 'true' || s === 'yes' || s === '1' || s === 'y';
}

/** Reads every product row into an array of plain objects the frontend understands. */
function getAllProducts() {
  const sheet = getProductsSheet_();
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(h => String(h).trim());
  const rows = values.slice(1);

  const col = name => headers.indexOf(name); // -1 if the column doesn't exist (optional columns)

  return rows
    .filter(r => r.some(cell => cell !== '' && cell !== null)) // skip blank rows
    .map(r => {
      const get = name => { const i = col(name); return i === -1 ? '' : r[i]; };
      const { category, name } = parseCategoryAndName_(get('Categories and ITEMS name'));
      const importQty = Math.max(0, Math.floor(Number(get('Import QTY')) || 0));

      return {
        id: String(get('Code')),
        name: name,
        category: category,
        price: Number(get('Rate')) || 0,
        quantity: Math.max(0, Math.floor(Number(get('Stock QTY')) || 0)),
        image: String(get('Image') || ''),
        description: String(get('Description') || ''),
        featured: truthy_(get('Featured')),
        isNew: importQty > 0
      };
    });
}

/** Finds a product's row index (1-based, includes header) by Code. */
function findProductRow_(sheet, productId) {
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(h => String(h).trim());
  const idCol = headers.indexOf('Code');
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
    const rateCol = headers.indexOf('Rate') + 1;
    const stockCol = headers.indexOf('Stock QTY') + 1;
    const nameRawCol = headers.indexOf('Categories and ITEMS name') + 1;

    const currentPrice = Number(sheet.getRange(rowIndex, rateCol).getValue()) || 0;
    const currentQty = Math.floor(Number(sheet.getRange(rowIndex, stockCol).getValue())) || 0;
    const { name: productName } = parseCategoryAndName_(sheet.getRange(rowIndex, nameRawCol).getValue());

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

    // 1. reduce stock (Stock QTY column — the sheet's own Price column,
    //    being Rate × Stock QTY, is left for you to recalculate manually
    //    or with a formula, since it's your inventory-value figure, not
    //    something the storefront depends on)
    sheet.getRange(rowIndex, stockCol).setValue(remaining);

    // 2. log the order
    const ordersSheet = getOrdersSheet_();
    ordersSheet.appendRow(ORDER_COLS.map(colName => {
      switch (colName) {
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
// if they don't already exist. If you're importing the provided
// Woolly-and-Co-Product-Database.xlsx instead, you can skip this —
// it already has the Products/Orders tabs and your real product rows.
function setupSheets() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  if (!ss.getSheetByName(PRODUCTS_SHEET_NAME)) {
    const p = ss.insertSheet(PRODUCTS_SHEET_NAME);
    p.appendRow(['Image', 'Code', 'Categories and ITEMS name', 'Rate', 'Stock QTY', 'Price', 'Import QTY', 'Description', 'Featured']);
    p.appendRow(['', 'F-DC-001', 'Felted Deco | Bean M-with Soap', 580, 67, 38860, 100, 'Hand-felted decorative piece.', 'TRUE']);
    p.setFrozenRows(1);
  }

  if (!ss.getSheetByName(ORDERS_SHEET_NAME)) {
    const o = ss.insertSheet(ORDERS_SHEET_NAME);
    o.appendRow(ORDER_COLS);
    o.setFrozenRows(1);
  }
}
