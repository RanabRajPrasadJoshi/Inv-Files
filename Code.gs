/**
 * WOOLLY & CO. — Google Apps Script backend
 * ------------------------------------------------------------
 * This file is the ONLY backend the site has. It reads/writes the
 * Google Sheet directly and is the single source of truth for
 * price and stock. The frontend never trusts its own numbers —
 * every order is re-validated against the sheet in here.
 *
 * SHEET LAYOUT THIS EXPECTS (Products tab, row 1 = these headers —
 * matching ignores case, spacing and punctuation, so "Stock QTY",
 * "stock qty", "Stock-Qty" etc. all resolve the same):
 *   Image | Code | Categories and ITEMS name | Rate | Stock QTY | Price | Import QTY | Description | Featured
 *
 *   - Code                          -> unique product ID / SKU (e.g. F-DC-001)
 *   - Categories and ITEMS name     -> "Category | Item name", e.g. "Felted Deco | Buddha"
 *                                      Shown on the site as "Category — Item name" so it's unique
 *                                      even when two categories share an item name (e.g. "Buddha").
 *   - Rate                          -> THE SELLING PRICE shown on the site.
 *   - Stock QTY                     -> THE LIVE STOCK COUNT shown on the site.
 *   - Price                         -> Rate × Stock QTY, your own inventory-value column. The site
 *                                      never reads this column — it's for your bookkeeping only.
 *   - Import QTY                    -> restock quantity; any value > 0 marks the item "New" on the site.
 *   - Image                         -> a picture uploaded / dragged into the cell
 *                                      (Insert > Image > Insert image IN cell), OR a direct
 *                                      image link, OR a Google Drive share link. Uploaded pictures
 *                                      have no public link, so the website requests them from this
 *                                      script (?action=image&code=...) - see "PICTURES STORED
 *                                      INSIDE THE SHEET". Pictures floating OVER cells can't be read.
 *   - Description / Featured        -> optional. Safe to leave blank.
 *
 * ORDERS SHEET — expected headers (any order, matching is equally lenient):
 *   Timestamp | OrderID | ProductID | ProductName | Quantity | Price | Total |
 *   CustomerName | Email | Phone | Address | Note | OrderStatus
 *   Every value is written to whichever column actually has that header —
 *   NOT to a fixed position — so reordering or adding your own columns is safe.
 *   If a column for one of these is missing entirely, that value is silently
 *   skipped and logged; run debugOrderHeaderMapping() to check.
 *
 * MAIN INVENTORY LINK (Baccabiu Main Inventory > "Warehouse" tab, table 1 only):
 *   The Products tab here is a LIVE COPY of that table. STOCK_SOURCE decides who copies it:
 *   'formula' (default) = an IMPORTRANGE formula in the Products tab; the script never writes there.
 *   'script' = this script copies it (on every site load, at most every 20 seconds, and on
 *   every order; needs SOURCE_SPREADSHEET_ID). Mapping:
 *       Baccabiu SKU  -> Code            Description  -> Categories and ITEMS name
 *       Rate          -> Rate (price)    2026 Stock   -> Stock QTY (what the site shows)
 *       2026 Import   -> Import QTY      Image        -> Image / HasPicture
 *   Products removed from table 1 disappear from the site; new ones appear.
 *   Description / Featured in the Products tab are yours - the sync never overwrites them.
 *
 * ORDER STATUS & STOCK BEHAVIOUR (stock is DERIVED, never overwritten):
 *   shown on site = 2026 Stock (main inventory) - quantity in open web orders
 *   - A new order is "Pending" and immediately counts as open, so the stock shown drops
 *     (this is what stops two customers buying the last unit).
 *   - Status "Cancelled" releases it - the stock goes straight back, nothing to undo.
 *   - Status "Delivered" also releases it (see RELEASING_STATUSES): record that sale in
 *     the main inventory and its 2026 Stock falls by the same amount, so nothing is
 *     counted twice. Change RELEASING_STATUSES to ['cancelled'] if you do NOT record web
 *     sales in the main inventory.
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
const STORE_NAME = "Baccabi'u";

// ---- Link to the main inventory (the master copy of products + stock) ----
// Open "Baccabiu Main Inventory" in the browser; the ID is the long text in its address
// between /d/ and /edit  ->  https://docs.google.com/spreadsheets/d/<THIS PART>/edit
const SOURCE_SPREADSHEET_ID = '1tPVxpzCpfuZewfreZvQre7-241GSN5-Z6LAME3rxtUQ';
// How the Products tab gets its data from the main inventory:
//   'formula' -> an IMPORTRANGE formula typed into the Products tab does the copying (no script
//                sync; the script only READS the tab and never writes to it). Simplest.
//   'script'  -> this script copies the data itself (also brings pictures across).
const STOCK_SOURCE = 'formula';
const SOURCE_SHEET_NAME = 'Warehouse';
// Column headings in that tab (matching ignores case, spaces and line breaks):
const SOURCE_CODE_HEADER = 'Baccabiu SKU';
const SOURCE_NAME_HEADER = 'Description';   // "Felt Brooch | Mr. Crocodile"
const SOURCE_RATE_HEADER = 'Rate';
const SOURCE_STOCK_HEADER = '2026 Stock';   // <- the quantity the website shows
const SOURCE_IMPORT_HEADER = '2026 Import'; // > 0 marks a product "New"
const SOURCE_IMAGE_HEADER = 'Image';         // column C: picture inside the cell
const SOURCE_IMAGE_URL_HEADER = 'Image-URL'; // column L: a web link to the picture (used first)
const SOURCE_EXPECTED_ROWS = 81;            // table 1 size; only used to warn if it differs
const SOURCE_MAX_SCAN_ROWS = 1000;
const SYNC_MIN_INTERVAL_SECONDS = 20;       // how often site loads may refresh the copy
// Order statuses that no longer hold stock back (compared case-insensitively):
const RELEASING_STATUSES = ['cancelled', 'delivered'];

// Canonical Orders columns this script knows how to fill in.
// (Order here is just for readability — actual writing is by header match, not position.)
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
    if (action === 'image') {
      return jsonOut(getSheetImage_(e.parameter.code));
    }
    if (action === 'sync') {
      return jsonOut(runSyncCheck_());
    }
    if (action === 'imagecheck') {
      return jsonOut(imageCheck_());
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

// ====== HEADER MATCHING (case/spacing/punctuation-insensitive) ==========

/** Strips everything except letters and numbers and lowercases what's left,
 *  so "Stock QTY", "stock qty", "Stock-Qty", "STOCK_QTY" all normalize to
 *  the same "stockqty" key. This is deliberately very lenient, since sheet
 *  header formatting drifts easily and silently breaking price/stock
 *  reading (or order logging) is worse than being a bit permissive. */
function normalizeHeader_(h) {
  return String(h || '').toLowerCase().replace(/[^a-z0-9]/g, '');
}

/** Builds { canonicalName: columnIndex } from a Products header row. */
function buildHeaderMap_(headerRow) {
  // Each field accepts several heading names, so the Products tab works whether it
  // uses its own headings or the Warehouse headings copied in by IMPORTRANGE.
  const expected = {
    image: ['Image', 'Picture', 'Photo'],
    imageurl: ['Image URL', 'ImageURL', 'Image Link'],
    haspicture: ['HasPicture'],
    code: ['Code', 'Baccabiu SKU', 'SKU', 'ProductID'],
    categoryitem: ['English Name', 'Categories and ITEMS name', 'ProductName', 'Item', 'Name'],
    mandarin: ['Mandarin Name', 'Mandarin', 'Chinese Name'],
    rate: ['Rate', 'Selling Price'],
    stock: ['Stock QTY', '2026 Stock', 'Stock', 'Quantity'],
    price: ['Price'],
    importqty: ['Import QTY', '2026 Import'],
    description: ['Description'],
    featured: ['Featured']
  };
  const normalizedHeaders = headerRow.map(normalizeHeader_);
  const map = {};
  Object.keys(expected).forEach(key => {
    map[key] = -1;
    for (const name of expected[key]) {
      const idx = normalizedHeaders.indexOf(normalizeHeader_(name));
      if (idx !== -1) { map[key] = idx; break; }
    }
  });
  // Warehouse-style tab: "Description" holds "Category | Item", not a sales blurb.
  if (map.categoryitem === -1 && map.description !== -1 &&
      normalizedHeaders.indexOf(normalizeHeader_('Baccabiu SKU')) !== -1) {
    map.categoryitem = map.description;
    map.description = -1;
  }
  return map;
}

/** Builds { 'OrderID': columnIndex, ... } from an Orders header row,
 *  keyed by the canonical names in ORDER_COLS. -1 means that column
 *  doesn't exist in the sheet. */
function buildOrderHeaderMap_(headerRow) {
  const normalizedHeaders = headerRow.map(normalizeHeader_);
  const map = {};
  ORDER_COLS.forEach(name => {
    map[name] = normalizedHeaders.indexOf(normalizeHeader_(name));
  });
  return map;
}

// ====== PRODUCTS ==============================================

function getProductsSheet_() {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(PRODUCTS_SHEET_NAME);
}

function getOrdersSheet_() {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(ORDERS_SHEET_NAME);
}

/** Splits "Category | Item name" into { category, item }. Tolerates a
 *  missing "|" by treating the whole string as the item name. */
function parseCategoryAndItem_(raw) {
  const str = String(raw || '').trim();
  const idx = str.indexOf('|');
  if (idx === -1) return { category: '', item: str };
  return {
    category: str.slice(0, idx).trim(),
    item: str.slice(idx + 1).trim()
  };
}

function truthy_(v) {
  const s = String(v).trim().toLowerCase();
  return s === 'true' || s === 'yes' || s === '1' || s === 'y';
}

/**
 * Pulls a usable image URL out of a cell, handling three cases:
 *  1. A cell image inserted via "Insert image in cell > By URL" — Apps
 *     Script exposes this as an object with a getUrl() method.
 *  2. Plain text: a normal https://... URL typed/pasted into the cell.
 *  3. A Google Drive "share" link, which isn't directly viewable as an
 *     <img> src — this converts it to the direct-view form.
 *  Returns '' if nothing usable is found (e.g. an uploaded/pasted picture,
 *  which Apps Script has no way to read a URL for).
 */
function extractImageUrl_(cellValue) {
  if (cellValue && typeof cellValue.getContentUrl === 'function') return publicLinkOf_(cellValue);
  if (cellValue && typeof cellValue.getUrl === 'function') {
    return cellValue.getUrl() || '';
  }
  const str = String(cellValue || '').trim();
  if (!str) return '';

  const driveFileMatch = str.match(/drive\.google\.com\/file\/d\/([^/]+)/);
  if (driveFileMatch) return `https://drive.google.com/thumbnail?id=${driveFileMatch[1]}&sz=w1000`;

  const driveIdMatch = str.match(/drive\.google\.com\/.*[?&]id=([^&]+)/);
  if (driveIdMatch) return `https://drive.google.com/thumbnail?id=${driveIdMatch[1]}&sz=w1000`;

  return str;
}

/**
 * Reads the Products tab into { list:[...], byCode:{code: item} }.
 * Used for the extras you keep here (Description, Featured, pictures) and as a
 * fallback product list if the main inventory can't be reached.
 */
function readProductsTab_() {
  const out = { list: [], byCode: {} };
  const sheet = getProductsSheet_();
  if (!sheet) return out;
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return out;
  const headerMap = buildHeaderMap_(values[0]);

  values.slice(1).forEach(r => {
    if (!r.some(cell => cell !== '' && cell !== null)) return;
    const get = key => { const i = headerMap[key]; return (i === undefined || i === -1) ? '' : r[i]; };
    const code = String(get('code') || '').trim();
    const rawImage = get('image');
    const cellPic = isCellPicture_(rawImage);
    const link = (cellPic ? publicLinkOf_(rawImage) : extractImageUrl_(rawImage)) || extractImageUrl_(get('imageurl'));
    const item = {
      code: code,
      name: String(get('categoryitem') || '').trim(),
      mandarin: String(get('mandarin') || '').trim(),
      rate: Number(get('rate')) || 0,
      stock: Math.floor(Number(get('stock'))) || 0,
      importQty: Math.floor(Number(get('importqty'))) || 0,
      link: link,
      hasPicture: cellPic || (!link && truthy_(get('haspicture'))),
      description: String(get('description') || '').trim(),
      featured: truthy_(get('featured'))
    };
    out.list.push(item);
    if (code && !out.byCode[code]) out.byCode[code] = item;
  });
  return out;
}

/** Main inventory table 1 as plain data, cached briefly so every visitor doesn't re-read it. */
function getMainItemsCached_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('mainItems_v2');
  if (hit) { try { return JSON.parse(hit); } catch (ignore) { /* re-read */ } }
  const items = readMainTable_().items.map(it => ({
    code: it.code, name: it.name, mandarin: it.mandarin, rate: it.rate, stock: it.stock,
    importQty: it.importQty, link: it.link, hasPicture: it.hasPicture
  }));
  try { cache.put('mainItems_v2', JSON.stringify(items), SYNC_MIN_INTERVAL_SECONDS); } catch (ignore) {}
  return items;
}

/** Turns one merged product into the object the website understands. */
function toSiteProduct_(base, extra, reserved) {
  const { category, item } = parseCategoryAndItem_(base.name);
  const code = base.code;
  // Category + item together so two categories sharing an item name
  // (e.g. "Buddha" under both Felted Deco and Felted Keychain) stay distinguishable.
  const displayName = category && item ? `${category} — ${item}` : (item || category || code || 'Untitled product');

  // Picture: the Image URL link (Products tab, else Warehouse Image-URL column).
  let image = '', sheetImage = false;
  if (extra && extra.link) image = extra.link;
  else if (base.link) image = extractImageUrl_(base.link);

  let description = extra ? extra.description : '';
  if (description && description === base.name) description = ''; // copied heading, not a blurb

  const stockInMain = Math.max(0, Math.floor(Number(base.stock)) || 0);
  const held = reserved[code] || 0;
  return {
    id: code,
    sku: code,
    name: displayName,
    mandarin: String(base.mandarin || (extra && extra.mandarin) || '').trim(),
    category: category,
    price: Number(base.rate) || 0,                // Rate = selling price
    stock: stockInMain,                           // 2026 Stock in the main inventory
    reserved: held,                               // held by open web orders
    quantity: Math.max(0, stockInMain - held),    // what a customer can still buy
    importQty: Math.max(0, Math.floor(Number(base.importQty)) || 0),
    image: image,
    sheetImage: sheetImage,
    description: description,
    featured: !!(extra && extra.featured),
    isNew: (Number(base.importQty) || 0) > 0
  };
}

/** Every product for the website. The main inventory decides WHICH products exist and
 *  their name, rate and stock; the Products tab adds Description / Featured / pictures. */
function getAllProducts() {
  syncFromMainIfStale_();             // script mode only: keep the Products tab copy fresh
  const reserved = reservedByCode_(); // quantities held by open web orders
  const tab = readProductsTab_();

  // The Products tab (live IMPORTRANGE copy of the Warehouse top table) is what the site shows.
  const fromTab = tab.list.filter(it => it.code);
  if (fromTab.length) return fromTab.map(it => toSiteProduct_(it, it, reserved));

  // Products tab empty or still loading: read the Warehouse directly instead.
  if (isMainConfigured_()) {
    try { return getMainItemsCached_().map(it => toSiteProduct_(it, null, reserved)); }
    catch (err) { Logger.log('getAllProducts: main inventory unreadable: ' + err); }
  }
  return [];
}

/** Live (uncached) lookup of one product, used when an order is placed. */
function findProductLive_(code) {
  if (isMainConfigured_()) {
    const live = readMainTable_().items.filter(it => it.code === code)[0];
    if (live) return live;
  }
  return readProductsTab_().byCode[code] || null;
}

/** Finds a product's row index (1-based, includes header) by Code. */
function findProductRow_(sheet, productId) {
  const values = sheet.getDataRange().getValues();
  const headerMap = buildHeaderMap_(values[0]);
  const idCol = headerMap.code;
  if (idCol === -1) return null;
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idCol]) === String(productId)) {
      return { rowIndex: i + 1, headerMap: headerMap };
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
    // 0. Pull the latest numbers from the main inventory BEFORE deciding anything,
    //    so we never sell against a stale copy.
    if (STOCK_SOURCE === 'script') {
      if (!isMainConfigured_()) {
        return { ok: false, error: 'The shop is not connected to the main inventory yet. Please try again later.' };
      }
      try {
        syncFromMain_();
      } catch (syncErr) {
        Logger.log('createOrder: Products tab copy not refreshed (order still checked live): ' + syncErr);
      }
    }

    // Read THIS product's price and stock live from the main inventory (falls back to
    // the Products tab only if the main inventory isn't configured).
    let live;
    try {
      live = findProductLive_(productId);
    } catch (liveErr) {
      Logger.log('createOrder: could not read main inventory: ' + liveErr);
      return { ok: false, error: 'Could not check current stock right now. Please try again in a moment.' };
    }
    if (!live) return { ok: false, error: 'That product could not be found.' };

    const currentPrice = Number(live.rate) || 0;
    const stockInMain = Math.max(0, Math.floor(Number(live.stock)) || 0);
    const heldByOpenOrders = reservedByCode_()[productId] || 0;
    const currentQty = Math.max(0, stockInMain - heldByOpenOrders); // what a customer can still buy
    const { category, item } = parseCategoryAndItem_(live.name);
    const productName = category && item ? `${category} — ${item}` : (item || category || productId);

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

    // 1. No stock cell is changed. The order row written below (status "Pending")
    //    IS the hold: it lowers the stock shown on the site until it is Cancelled
    //    (stock returns) or Delivered (the main inventory takes over).

    // 2. log the order — written by matching each value to whichever
    //    column in YOUR Orders sheet actually has that header, not to an
    //    assumed fixed position. This is the fix for order data (OrderID,
    //    ProductID, ProductName, Quantity, Price) not showing up: the old
    //    version wrote values as a plain array assuming a fixed column
    //    order, which breaks the moment your sheet's columns are reordered
    //    or worded slightly differently.
    writeOrderRow_({
      Timestamp: new Date(),
      OrderID: orderId,
      ProductID: productId,
      ProductName: productName,
      Quantity: qtyRequested,
      Price: currentPrice,
      Total: total,
      CustomerName: customer.name,
      Email: customer.email,
      Phone: customer.phone,
      Address: customer.address,
      Note: customer.note || '',
      OrderStatus: 'Pending'
    });

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

/**
 * Appends one row to the Orders sheet, placing each value in whichever
 * column actually carries that header — by name, not by position. Any
 * canonical field with no matching column is skipped (and logged) rather
 * than silently shifting every other value over by one, which is what
 * caused OrderID/ProductID/ProductName/Quantity/Price to go missing before.
 */
function writeOrderRow_(valuesByCol) {
  const sheet = getOrdersSheet_();
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const headerMap = buildOrderHeaderMap_(headerRow);

  const rowWidth = Math.max(lastCol, ORDER_COLS.length);
  const rowValues = new Array(rowWidth).fill('');
  const missing = [];

  Object.keys(valuesByCol).forEach(key => {
    const idx = headerMap[key];
    if (idx !== undefined && idx !== -1) {
      rowValues[idx] = valuesByCol[key];
    } else {
      missing.push(key);
    }
  });

  if (missing.length) {
    Logger.log('writeOrderRow_: Orders sheet has no column for: ' + missing.join(', ') +
      ' — that data was not written for this order. Run debugOrderHeaderMapping() to check your headers.');
  }

  sheet.appendRow(rowValues);
}

function round2_(n) {
  return Math.round(n * 100) / 100;
}

function generateOrderId_() {
  const now = new Date();
  const y = now.getFullYear();
  const m = String(now.getMonth() + 1).padStart(2, '0');
  const d = String(now.getDate()).padStart(2, '0');
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

// ====== PICTURES STORED INSIDE THE SHEET ===========================
/**
 * A picture uploaded / dragged into an Image cell has no public link, so the
 * website can't load it directly. Instead the site asks this script for it:
 *     GET ?action=image&code=F-DC-002
 * The script reads that row's picture from the sheet and returns it as base64
 * text, which the site turns into an <img>. No Drive folder and no Drive
 * permission is involved - it only needs "connect to external service"
 * permission (the picture is downloaded from Google's own image server using
 * your sheet access). Results are cached for a few hours so repeat views are fast.
 */
const IMAGE_MAX_WIDTH = 800;       // ask Google for a downsized copy if possible (px)
const IMAGE_CACHE_SECONDS = 21600; // 6 hours (the longest CacheService allows)
const IMAGE_CHUNK = 90000;         // CacheService values are limited to 100 KB each
const IMAGE_MAX_CHUNKS = 40;       // skip caching absurdly large pictures

function imgKey_(code) { return String(code).replace(/[^A-Za-z0-9_-]/g, '_'); }

/** Returns { ok, mime, data(base64) } for the picture in a product's Image cell. */
function getSheetImage_(code) {
  code = String(code || '').trim();
  if (!code) return { ok: false, error: 'Missing product code.' };

  const cached = readImageCache_(code);
  if (cached) return cached;

  const isPic = c => !!c && typeof c.getContentUrl === 'function';
  const sheet = getProductsSheet_();
  const found = sheet ? findProductRow_(sheet, code) : null;

  // 1st choice: a picture in this sheet's Image cell (Products tab)
  let cell = (found && found.headerMap.image !== -1)
    ? sheet.getRange(found.rowIndex, found.headerMap.image + 1).getValue() : null;

  // 2nd choice: the picture in the main inventory's Image cell
  let mainItem = null;
  if (!isPic(cell) && isMainConfigured_()) {
    try {
      mainItem = readMainTable_().items.filter(it => it.code === code)[0] || null;
      if (mainItem && isPic(mainItem.imageCell)) cell = mainItem.imageCell;
    } catch (mainErr) {
      Logger.log('getSheetImage_: could not read main inventory: ' + mainErr);
    }
  }

  if (!isPic(cell)) {
    if (!found && !mainItem) return { ok: false, error: 'Product not found: ' + code };
    // Is there a picture floating OVER this row's cells? Scripts can't read those.
    let floatingHere = 0;
    try {
      if (found) floatingHere = sheet.getImages().filter(img => img.getAnchorCell().getRow() === found.rowIndex).length;
    } catch (ignore) { /* keep going with the generic message */ }
    if (floatingHere) {
      return { ok: false, error: 'The picture in row ' + found.rowIndex + ' (' + code + ') is floating OVER the cells, ' +
        'which a script cannot read. Delete it and use Insert > Image > Insert image IN cell.' };
    }
    return { ok: false, error: 'No in-cell picture found for ' + code + ' in the Products tab or the main inventory. ' +
      'Put the picture INSIDE the Image cell of that product\'s row (Insert > Image > Insert image in cell).' };
  }

  const blob = fetchCellImageBlob_(cell, IMAGE_MAX_WIDTH);
  const out = {
    ok: true,
    mime: blob.getContentType() || 'image/png',
    data: Utilities.base64Encode(blob.getBytes())
  };
  writeImageCache_(code, out);
  return out;
}

/** Downloads the bytes of an in-cell picture using the script's own access. */
function fetchCellImageBlob_(cellImage, maxWidth) {
  const contentUrl = cellImage.getContentUrl();
  if (!contentUrl) throw new Error('no content URL available for this picture');
  const attempts = [
    { muteHttpExceptions: true },  // content URLs are already signed for this account
    { muteHttpExceptions: true, headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } }
  ];
  const urls = [];
  // Try a smaller copy first (Google image servers accept a "=w800" suffix).
  if (maxWidth && contentUrl.indexOf('?') === -1) {
    urls.push(contentUrl.replace(/=[swh]\d+[^=\/]*$/i, '') + '=w' + maxWidth);
  }
  urls.push(contentUrl);

  let lastCode = 0;
  for (const url of urls) {
    for (const opts of attempts) {
      try {
        const r = UrlFetchApp.fetch(url, opts);
        lastCode = r.getResponseCode();
        if (lastCode !== 200) continue;
        const b = r.getBlob();
        if (/^image\//i.test(b.getContentType() || '')) return b;
      } catch (ignore) { /* try the next way */ }
    }
  }
  throw new Error('download failed (HTTP ' + lastCode + ')');
}

function readImageCache_(code) {
  try {
    const cache = CacheService.getScriptCache();
    const k = imgKey_(code);
    const metaRaw = cache.get('imgmeta_' + k);
    if (!metaRaw) return null;
    const meta = JSON.parse(metaRaw);
    const keys = [];
    for (let i = 0; i < meta.n; i++) keys.push('img_' + k + '_' + i);
    const parts = cache.getAll(keys);
    let data = '';
    for (let i = 0; i < meta.n; i++) {
      const part = parts['img_' + k + '_' + i];
      if (part === null || part === undefined) return null; // expired / evicted
      data += part;
    }
    return { ok: true, mime: meta.mime, data: data };
  } catch (err) {
    return null;
  }
}

function writeImageCache_(code, out) {
  try {
    const k = imgKey_(code);
    const put = {};
    let n = 0;
    for (let i = 0; i < out.data.length; i += IMAGE_CHUNK) {
      put['img_' + k + '_' + n] = out.data.slice(i, i + IMAGE_CHUNK);
      n++;
    }
    if (n > IMAGE_MAX_CHUNKS) return; // too big to be worth caching
    put['imgmeta_' + k] = JSON.stringify({ mime: out.mime, n: n });
    CacheService.getScriptCache().putAll(put, IMAGE_CACHE_SECONDS);
  } catch (err) {
    Logger.log('picture cache write skipped: ' + err);
  }
}

/** Menu action: forget cached pictures (use after replacing a picture in the sheet). */
function clearPictureCache() {
  const sheet = getProductsSheet_();
  const values = sheet.getDataRange().getValues();
  const map = buildHeaderMap_(values[0]);
  const keys = [];
  if (map.code !== -1) {
    values.slice(1).forEach(r => {
      const code = String(r[map.code] || '').trim();
      if (!code) return;
      const k = imgKey_(code);
      keys.push('imgmeta_' + k);
      for (let i = 0; i < IMAGE_MAX_CHUNKS; i++) keys.push('img_' + k + '_' + i);
    });
  }
  const cache = CacheService.getScriptCache();
  cache.remove('mainItems_v2');
  for (let i = 0; i < keys.length; i += 50) cache.removeAll(keys.slice(i, i + 50));

  const msg = 'Picture cache cleared. A new or replaced picture will show on the site within a minute ' +
    '(a visitor\'s browser can keep the old copy for up to 30 minutes).';
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { Logger.log(msg); }
}

/** Adds a "Woolly Shop" menu to the spreadsheet (runs automatically on open). */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Woolly Shop')
    .addItem('Sync from main inventory now', 'syncFromMainNow')
    .addItem('Check main inventory link', 'debugMainLink')
    .addItem('Keep copy updated every 5 min (optional)', 'installMainSyncTrigger')
    .addSeparator()
    .addItem('Check images (diagnostic)', 'debugImages')
    .addItem('Clear picture cache', 'clearPictureCache')
    .addToUi();
}

/** Shows, for every product, where its picture will come from on the website. */
function debugImages() {
  const lines = [];
  let products = [];
  try { CacheService.getScriptCache().remove('mainItems_v2'); products = getAllProducts(); }
  catch (err) { lines.push('Could not read products: ' + err); }
  let missing = 0;
  products.forEach(p => {
    let kind;
    if (p.image) kind = 'link - will show: ' + String(p.image).slice(0, 60);
    else if (p.sheetImage) kind = 'picture inside a cell - will show (loaded via the script)';
    else { kind = 'NO PICTURE (site will try images/' + p.sku + '.jpg, then a placeholder)'; missing++; }
    lines.push(p.sku + ' -> ' + kind);
  });
  lines.unshift(products.length + ' products, ' + missing + ' without a picture.\n');
  try {
    const floating = getProductsSheet_().getImages().length;
    if (floating) lines.push('\n' + floating + ' picture(s) in the Products tab FLOAT OVER cells - these cannot be read. ' +
      'Delete them and use Insert > Image > Insert image IN cell.');
  } catch (ignore) {}
  if (isMainConfigured_()) {
    try {
      const sh = SpreadsheetApp.openById(SOURCE_SPREADSHEET_ID).getSheetByName(SOURCE_SHEET_NAME);
      const fl = sh ? sh.getImages().length : 0;
      if (fl) lines.push('\n' + fl + ' picture(s) in the main inventory "' + SOURCE_SHEET_NAME + '" tab FLOAT OVER cells - ' +
        'these cannot be read. Select each one > three-dot menu > "Put image in selected cell".');
    } catch (ignore) {}
  }
  Logger.log(lines.join('\n'));
  try { SpreadsheetApp.getUi().alert(lines.join('\n').slice(0, 8000)); } catch (e) { /* run from editor: see log */ }
}

/**
 * Open  <your /exec URL>?action=imagecheck  in a browser to see exactly what the script
 * finds in the Warehouse Image column, and whether it can download each picture.
 */
function imageCheck_() {
  const products = getAllProducts();
  const noPic = products.filter(p => !p.image).map(p => p.sku);
  return {
    ok: true,
    codeVersion: 'products-tab-v1',
    source: readProductsTab_().list.length ? 'Products tab' : 'Warehouse (Products tab empty)',
    products: products.length,
    withPicture: products.length - noPic.length,
    withoutPicture: noPic,
    sample: products.slice(0, 3)
  };
}

// ====== MAIN INVENTORY LINK =========================================
/**
 * The Products tab is a live copy of table 1 of the main inventory's Warehouse
 * tab. Table 1 = the first block of rows under the heading row, ending at the
 * first blank row (or a repeated heading row). Anything below or beside it is ignored.
 */
function isMainConfigured_() {
  return !!SOURCE_SPREADSHEET_ID && SOURCE_SPREADSHEET_ID.indexOf('PASTE') === -1;
}

function isCellPicture_(v) { return !!v && typeof v.getContentUrl === 'function'; }

/** The picture's own web link, if it was inserted from one (uploaded pictures have none
 *  we can use - their Google address only works for the sheet owner). */
function publicLinkOf_(cellImage) {
  let u = '';
  try { u = cellImage.getUrl() || ''; } catch (ignore) {}
  return (/^https?:/i.test(u) && !/googleusercontent\.com/i.test(u)) ? extractImageUrl_(u) : '';
}

function columnLetter_(n) {
  let s = '';
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

/** Reads table 1 from the main inventory. Returns { items, headerRowNumber, skippedNoCode, fileName }. */
function readMainTable_() {
  if (!isMainConfigured_()) throw new Error('SOURCE_SPREADSHEET_ID has not been set in Code.gs yet.');

  let ss;
  try {
    ss = SpreadsheetApp.openById(SOURCE_SPREADSHEET_ID);
  } catch (err) {
    throw new Error('Could not open the main inventory (check the ID and that your Google account can open it): ' + err);
  }
  const sh = ss.getSheetByName(SOURCE_SHEET_NAME);
  if (!sh) {
    let tabs = '';
    try { tabs = ' Tabs found: ' + ss.getSheets().map(t => '"' + t.getName() + '"').join(', ') + '.'; } catch (ignore) {}
    throw new Error('The main inventory has no tab named "' + SOURCE_SHEET_NAME + '".' + tabs +
      ' (Set SOURCE_SHEET_NAME in Code.gs to the exact tab name.)');
  }

  const lastRow = Math.min(sh.getLastRow(), SOURCE_MAX_SCAN_ROWS);
  const lastCol = sh.getLastColumn();
  if (lastRow < 2) throw new Error('The "' + SOURCE_SHEET_NAME + '" tab is empty.');
  const values = sh.getRange(1, 1, lastRow, lastCol).getValues(); // in-cell pictures come back as objects

  const want = {
    code: normalizeHeader_(SOURCE_CODE_HEADER),
    name: normalizeHeader_(SOURCE_NAME_HEADER),
    rate: normalizeHeader_(SOURCE_RATE_HEADER),
    stock: normalizeHeader_(SOURCE_STOCK_HEADER),
    imp: normalizeHeader_(SOURCE_IMPORT_HEADER),
    image: normalizeHeader_(SOURCE_IMAGE_HEADER)
  };
  const labels = {
    code: SOURCE_CODE_HEADER, name: SOURCE_NAME_HEADER, rate: SOURCE_RATE_HEADER,
    stock: SOURCE_STOCK_HEADER, imp: SOURCE_IMPORT_HEADER, image: SOURCE_IMAGE_HEADER
  };

  // heading row = first row (within the top 20) holding both the code and the stock headings
  let headerIdx = -1;
  for (let i = 0; i < Math.min(values.length, 20); i++) {
    const keys = values[i].map(normalizeHeader_);
    if (keys.indexOf(want.code) !== -1 && keys.indexOf(want.stock) !== -1) { headerIdx = i; break; }
  }
  if (headerIdx === -1) {
    const firstRow = values[0].slice(0, 14).map(v => String(v).replace(/\s+/g, ' ').trim()).join(' | ');
    throw new Error('Could not find a heading row containing "' + SOURCE_CODE_HEADER + '" and "' +
      SOURCE_STOCK_HEADER + '" in the top 20 rows of "' + SOURCE_SHEET_NAME + '". Row 1 reads: ' + firstRow +
      '. (Fix SOURCE_CODE_HEADER / SOURCE_STOCK_HEADER in Code.gs to match.)');
  }

  const keys = values[headerIdx].map(normalizeHeader_);
  const col = {};
  Object.keys(want).forEach(k => { col[k] = keys.indexOf(want[k]); });
  ['code', 'name', 'rate', 'stock'].forEach(k => {
    if (col[k] === -1) throw new Error('The main inventory has no "' + labels[k] + '" column.');
  });
  col.man = keys.findIndex(k => /^(mandarin|chinese)(name)?$/.test(k));
  // Image-URL column (a web link to the picture)
  col.imageUrl = keys.indexOf(normalizeHeader_(SOURCE_IMAGE_URL_HEADER));
  if (col.imageUrl === -1) col.imageUrl = keys.findIndex(k => /^(image|photo|picture|img)(url|link)$/.test(k));
  // Image column: exact heading first, otherwise any heading mentioning image/photo/picture
  if (col.image === -1) col.image = keys.findIndex((k, j) => j !== col.imageUrl && /image|photo|picture|pic|img/.test(k));

  // links in the Image-URL column (a cell can show text but link somewhere else)
  let urlLinks = [];
  if (col.imageUrl !== -1) {
    try {
      urlLinks = sh.getRange(1, col.imageUrl + 1, lastRow, 1).getRichTextValues().map(r => {
        const rt = r[0];
        if (!rt) return '';
        let u = '';
        try { u = rt.getLinkUrl() || ''; } catch (ignore) {}
        if (!u) { try { rt.getRuns().some(run => (u = run.getLinkUrl() || '')); } catch (ignore) {} }
        return u;
      });
    } catch (ignore) {}
  }

  // =IMAGE("https://...") formulas: the value reads as blank, so read the formula instead
  let imageFormulas = [];
  if (col.image !== -1) {
    try { imageFormulas = sh.getRange(1, col.image + 1, lastRow, 1).getFormulas().map(r => r[0]); } catch (ignore) {}
  }
  // Pictures floating OVER the cells: usable only if they were inserted from a web link
  const floatingByRow = {};
  try {
    sh.getImages().forEach(im => {
      const r = im.getAnchorCell().getRow();
      let u = '';
      try { u = im.getUrl() || ''; } catch (ignore) {}
      if (!floatingByRow[r]) floatingByRow[r] = { url: u };
    });
  } catch (ignore) {}

  const items = [];
  let skippedNoCode = 0;
  for (let i = headerIdx + 1; i < values.length; i++) {
    const row = values[i];
    const sku = String(row[col.code] === null || row[col.code] === undefined ? '' : row[col.code]).trim();
    const nameRaw = String(row[col.name] === null || row[col.name] === undefined ? '' : row[col.name]).trim();

    if (!sku && !nameRaw) break;                    // blank row = end of table 1
    if (normalizeHeader_(sku) === want.code) break; // a second heading row = another table starts
    if (!sku) { skippedNoCode++; continue; }        // has a name but no code: can't be sold

    const img = col.image === -1 ? '' : row[col.image];
    let link = '';
    let hasPicture = false;
    let imageSource = 'none';
    if (isCellPicture_(img)) {
      link = publicLinkOf_(img);   // only when it was inserted from a normal web link
      hasPicture = true;           // the script can always fetch it, as a backup
      imageSource = 'in-cell picture';
    } else if (String(img || '').trim()) {
      link = extractImageUrl_(img);
      imageSource = 'link in cell';
    } else if (imageFormulas[i] && /=\s*IMAGE\s*\(/i.test(imageFormulas[i])) {
      const m = imageFormulas[i].match(/IMAGE\s*\(\s*"([^"]+)"/i);
      if (m) { link = extractImageUrl_(m[1]); imageSource = 'IMAGE() formula'; }
    }
    // Image-URL column wins: a web link loads fastest and needs no script work
    if (col.imageUrl !== -1) {
      const typed = String(row[col.imageUrl] === null || row[col.imageUrl] === undefined ? '' : row[col.imageUrl]).trim();
      const u = /^https?:/i.test(typed) ? typed : (urlLinks[i] || '');
      if (/^https?:/i.test(u)) {
        link = extractImageUrl_(u);
        imageSource = hasPicture ? 'Image-URL link (+ in-cell picture as backup)' : 'Image-URL link';
      }
    }
    const floating = floatingByRow[i + 1];
    if (!link && !hasPicture && floating) {
      if (floating.url) { link = extractImageUrl_(floating.url); imageSource = 'floating picture (from a link)'; }
      else imageSource = 'FLOATING over cells - cannot be read';
    }

    items.push({
      code: sku,
      name: nameRaw,
      mandarin: col.man === -1 ? '' : String(row[col.man] || '').trim(),
      rate: Number(row[col.rate]) || 0,
      stock: Math.floor(Number(row[col.stock])) || 0,
      importQty: col.imp === -1 ? 0 : (Math.floor(Number(row[col.imp])) || 0),
      link: link,
      hasPicture: hasPicture,
      imageCell: hasPicture ? img : null,
      imageSource: imageSource,
      row: i + 1
    });
  }

  return { items: items, headerRowNumber: headerIdx + 1, skippedNoCode: skippedNoCode, fileName: ss.getName(),
    imageColumn: col.image === -1 ? '' : String(values[headerIdx][col.image]).trim(),
    imageColumnLetter: col.image === -1 ? '' : columnLetter_(col.image + 1),
    imageUrlColumn: col.imageUrl === -1 ? '' : columnLetter_(col.imageUrl + 1) + ' "' + String(values[headerIdx][col.imageUrl]).trim() + '"',
    floatingCount: Object.keys(floatingByRow).length };
}

/**
 * Copies table 1 into the Products tab: updates changed values, adds new products,
 * removes products no longer in the table. Never touches Description / Featured or
 * pictures you placed in this sheet. The caller should hold the script lock.
 * Returns { count, added, removed }.
 */
function syncFromMain_() {
  if (STOCK_SOURCE !== 'script') {
    throw new Error('Script sync is switched off because STOCK_SOURCE is "' + STOCK_SOURCE +
      '" (the Products tab is filled by a formula). Set STOCK_SOURCE to \'script\' in Code.gs to use it.');
  }
  const main = readMainTable_();
  const items = main.items;
  if (!items.length) {
    // never wipe the shop because of a read problem
    throw new Error('No products were found in table 1 of the main inventory, so nothing was changed.');
  }

  const sheet = getProductsSheet_();

  // 0. make sure every column we write to exists
  const needed = ['Image', 'Code', 'Categories and ITEMS name', 'Rate', 'Stock QTY', 'Price',
    'Import QTY', 'Description', 'Featured', 'HasPicture'];
  const startCols = sheet.getLastColumn();
  const header = startCols ? sheet.getRange(1, 1, 1, startCols).getValues()[0] : [];
  const normalized = header.map(normalizeHeader_);
  let changedHeader = false;
  needed.forEach(name => {
    if (normalized.indexOf(normalizeHeader_(name)) === -1) {
      header.push(name);
      normalized.push(normalizeHeader_(name));
      changedHeader = true;
    }
  });
  if (changedHeader) sheet.getRange(1, 1, 1, header.length).setValues([header]);
  const colIndex = name => normalized.indexOf(normalizeHeader_(name)) + 1; // 1-based

  const codeCol = colIndex('Code');
  const wanted = {};
  items.forEach(it => { wanted[it.code] = true; });

  // 1. remove products that are no longer in table 1 (bottom-up keeps row numbers valid)
  let removed = 0;
  let lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const codes = sheet.getRange(2, codeCol, lastRow - 1, 1).getValues().map(r => String(r[0] || '').trim());
    for (let k = codes.length - 1; k >= 0; k--) {
      if (codes[k] && !wanted[codes[k]]) {
        try { sheet.deleteRow(k + 2); } catch (e) { sheet.getRange(k + 2, 1, 1, header.length).clearContent(); }
        removed++;
      }
    }
  }

  // 2. work out where each product lives; append the new ones
  lastRow = sheet.getLastRow();
  const rowOf = {};
  if (lastRow >= 2) {
    sheet.getRange(2, codeCol, lastRow - 1, 1).getValues().forEach((r, idx) => {
      const c = String(r[0] || '').trim();
      if (c) rowOf[c] = idx + 2;
    });
  }
  const toAdd = items.filter(it => !rowOf[it.code]);
  if (toAdd.length) {
    const startRow = Math.max(lastRow, 1) + 1;
    sheet.getRange(startRow, codeCol, toAdd.length, 1).setValues(toAdd.map(it => [it.code]));
    toAdd.forEach((it, k) => { rowOf[it.code] = startRow + k; });
  }

  // 3. refresh the copied columns, one column at a time (other columns are untouched)
  const maxRow = Math.max.apply(null, items.map(it => rowOf[it.code]));
  const n = maxRow - 1;
  const writeCol = (name, fn) => {
    const range = sheet.getRange(2, colIndex(name), n, 1);
    const vals = range.getValues();
    items.forEach(it => { vals[rowOf[it.code] - 2][0] = fn(it); });
    range.setValues(vals);
  };
  writeCol('Categories and ITEMS name', it => it.name);
  writeCol('Rate', it => it.rate);
  writeCol('Stock QTY', it => it.stock);
  writeCol('Price', it => round2_(it.rate * it.stock)); // inventory value, for your own reference
  writeCol('Import QTY', it => it.importQty);
  writeCol('HasPicture', it => (it.hasPicture ? 'TRUE' : ''));

  // 4. image LINKS from the main inventory (cell by cell, only where different, so
  //    pictures you placed in this sheet are never rewritten)
  const withLink = items.filter(it => it.link);
  if (withLink.length) {
    const imgCol = colIndex('Image');
    const shown = sheet.getRange(2, imgCol, n, 1).getDisplayValues();
    withLink.forEach(it => {
      const idx = rowOf[it.code] - 2;
      if (String(shown[idx][0]).trim() !== it.link) sheet.getRange(rowOf[it.code], imgCol).setValue(it.link);
    });
  }

  return { count: items.length, added: toAdd.length, removed: removed };
}

/** Called on every product request; copies at most once per SYNC_MIN_INTERVAL_SECONDS. */
function syncFromMainIfStale_() {
  if (STOCK_SOURCE !== 'script' || !isMainConfigured_()) return;
  const cache = CacheService.getScriptCache();
  if (cache.get('mainSyncFresh')) return;

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(0)) return; // an order or another sync is running - use what's there
  try {
    if (cache.get('mainSyncFresh')) return;
    try {
      syncFromMain_();
      cache.put('mainSyncFresh', '1', SYNC_MIN_INTERVAL_SECONDS);
    } catch (err) {
      Logger.log('Sync from main inventory failed: ' + err);
      cache.put('mainSyncFresh', '1', 10); // don't hammer a broken link on every request
    }
  } finally {
    lock.releaseLock();
  }
}

/**
 * Orders that are still "open" hold stock back. Returns { code: quantityHeld }.
 * An order stops holding stock once its status is in RELEASING_STATUSES.
 */
function reservedByCode_() {
  const out = {};
  const sheet = getOrdersSheet_();
  if (!sheet || sheet.getLastRow() < 2) return out;
  const values = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
  const h = buildOrderHeaderMap_(values[0]);
  if (h['ProductID'] === -1 || h['Quantity'] === -1) return out;

  for (let i = 1; i < values.length; i++) {
    const code = String(values[i][h['ProductID']] || '').trim();
    if (!code) continue;
    const status = h['OrderStatus'] === -1 ? '' : String(values[i][h['OrderStatus']] || '').trim().toLowerCase();
    if (RELEASING_STATUSES.indexOf(status) !== -1) continue;
    out[code] = (out[code] || 0) + (Math.floor(Number(values[i][h['Quantity']])) || 0);
  }
  return out;
}

/**
 * Runs the copy right now and reports the result as plain data, so you can test the
 * link from a browser:  <your exec URL>?action=sync
 * Returns { ok:true, linked, added, removed } or { ok:false, error }.
 */
function runSyncCheck_() {
  if (!isMainConfigured_()) {
    return { ok: false, error: 'SOURCE_SPREADSHEET_ID has not been set in Code.gs, or the new version has not been deployed yet.' };
  }
  const cache = CacheService.getScriptCache();
  if (cache.get('syncCheckBusy')) return { ok: false, error: 'Please wait about 10 seconds and try again.' };
  cache.put('syncCheckBusy', '1', 10);

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return { ok: false, error: 'The shop is busy right now - try again in a moment.' };
  try {
    const r = syncFromMain_();
    cache.put('mainSyncFresh', '1', SYNC_MIN_INTERVAL_SECONDS);
    return { ok: true, linked: r.count, added: r.added, removed: r.removed, expected: SOURCE_EXPECTED_ROWS };
  } catch (err) {
    return { ok: false, error: String(err).split(SOURCE_SPREADSHEET_ID).join('<ID>') };
  } finally {
    lock.releaseLock();
  }
}

/** Menu: copy from the main inventory right now and report what happened. */
function syncFromMainNow() {
  let msg;
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    try {
      const r = syncFromMain_();
      CacheService.getScriptCache().put('mainSyncFresh', '1', SYNC_MIN_INTERVAL_SECONDS);
      msg = 'Linked ' + r.count + ' products from the "' + SOURCE_SHEET_NAME + '" tab (table 1).' +
        (r.added ? '\n' + r.added + ' new product(s) added.' : '') +
        (r.removed ? '\n' + r.removed + ' product(s) no longer in the main inventory were removed.' : '');
      if (SOURCE_EXPECTED_ROWS && r.count !== SOURCE_EXPECTED_ROWS) {
        msg += '\n\nHeads-up: you expected ' + SOURCE_EXPECTED_ROWS + ' products but ' + r.count +
          ' were found. Use "Check main inventory link" to see where table 1 starts and ends.';
      }
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    msg = 'Sync failed: ' + err;
  }
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { /* run from the editor: see the log */ }
}

/** Menu: shows exactly which rows of the main inventory are being linked. */
function debugMainLink() {
  let msg;
  try {
    const m = readMainTable_();
    const first = m.items.slice(0, 3).map(it => it.code);
    const last = m.items.slice(-3).map(it => it.code);
    msg = 'Main inventory: "' + m.fileName + '" > tab "' + SOURCE_SHEET_NAME + '"\n' +
      'Heading row found on row ' + m.headerRowNumber + '\n' +
      'Products in table 1: ' + m.items.length + (SOURCE_EXPECTED_ROWS ? '  (expected ' + SOURCE_EXPECTED_ROWS + ')' : '') + '\n' +
      'First: ' + first.join(', ') + '\nLast: ' + last.join(', ') +
      (m.skippedNoCode ? '\n\n' + m.skippedNoCode + ' row(s) skipped because they have no ' + SOURCE_CODE_HEADER + '.' : '');
    if (SOURCE_EXPECTED_ROWS && m.items.length !== SOURCE_EXPECTED_ROWS) {
      msg += '\n\nThe count differs from ' + SOURCE_EXPECTED_ROWS + '. Table 1 ends at the first completely blank row, ' +
        'so check for a blank row inside the table or extra rows right below it.';
    }
  } catch (err) {
    msg = 'Link problem: ' + err;
  }
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { /* run from the editor: see the log */ }
}

/** Optional: keeps the Products copy fresh every 5 minutes even when nobody visits the site. */
function installMainSyncTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    const fn = t.getHandlerFunction();
    // also clears the old "order status" trigger, which is no longer used
    if (fn === 'mainSyncTrigger' || fn === 'onOrdersEdit') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('mainSyncTrigger').timeBased().everyMinutes(5).create();
  const msg = 'Done - the Products copy will now refresh from the main inventory every 5 minutes.';
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { /* run from the editor */ }
}

function mainSyncTrigger() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return;
  try {
    syncFromMain_();
  } catch (err) {
    Logger.log('Scheduled sync failed: ' + err);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Left in place only so an old "order status" trigger (if you installed one earlier)
 * doesn't throw errors. Stock no longer changes when you edit a status - open orders
 * simply hold stock until they're Cancelled or Delivered.
 */
function onOrdersEdit(e) { /* intentionally empty */ }

// ====== ONE-TIME PERMISSION SETUP ==================================
/**
 * Run this ONCE from the Apps Script editor (pick authorizeAll in the
 * function dropdown, click Run). It touches every service this project uses
 * so Google shows ONE consent screen covering all of them: external requests
 * (needed to read pictures stored in the sheet), email, triggers and menus.
 * Drive is NOT used anywhere. It deliberately has no try/catch - if a
 * permission is still missing, the error appears in the log.
 * After approving, go to Deploy > Manage deployments > pencil > New version.
 */
function authorizeAll() {
  UrlFetchApp.fetch('https://www.google.com', { muteHttpExceptions: true });
  MailApp.getRemainingDailyQuota();
  ScriptApp.getProjectTriggers();
  LockService.getDocumentLock();
  CacheService.getScriptCache();
  SpreadsheetApp.getActiveSpreadsheet().getName();
  // also prove this account can open the main inventory (throws a clear error if not)
  if (isMainConfigured_()) Logger.log('Main inventory opened OK: ' + SpreadsheetApp.openById(SOURCE_SPREADSHEET_ID).getName());
  Logger.log('All permissions are granted. Now: Deploy > Manage deployments > pencil > New version > Deploy.');
}

// ====== DIAGNOSTIC HELPERS =====================================
// Run these from the Apps Script editor (select in the function dropdown,
// press Run, then View > Logs) whenever something looks wrong on the site
// or in the sheet — they print exactly which column each field resolved
// to (-1 = not found, meaning that header text doesn't match anything
// Code.gs recognizes).

function debugHeaderMapping() {
  const sheet = getProductsSheet_();
  const headerRow = sheet.getDataRange().getValues()[0];
  Logger.log('Raw Products headers: ' + JSON.stringify(headerRow));
  Logger.log('Resolved Products column map: ' + JSON.stringify(buildHeaderMap_(headerRow)));
}

function debugOrderHeaderMapping() {
  const sheet = getOrdersSheet_();
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  Logger.log('Raw Orders headers: ' + JSON.stringify(headerRow));
  Logger.log('Resolved Orders column map (-1 = missing): ' + JSON.stringify(buildOrderHeaderMap_(headerRow)));
}

// ====== ONE-TIME SETUP HELPER =====================================
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