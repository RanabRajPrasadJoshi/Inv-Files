/**
 * WOOLLY & CO. — frontend logic
 * ------------------------------------------------------------
 * No framework, no bundler. Talks to Google Apps Script for
 * everything product- and order-related. Google Sheets stays the
 * only source of truth for price/stock — this file never invents
 * or caches those numbers past the next refresh.
 * ------------------------------------------------------------
 */

// ====== CONFIG ================================================
const CONFIG = {
  // Paste your Apps Script Web App /exec URL here after deploying Code.gs
  API_URL: 'https://script.google.com/macros/s/AKfycbz5yOMD8mvsn8j8guj56Tn-Iq5BKlpBIW9vao7N7hGIlsQCFU2n0i_g1QmjARu36c5r/exec',
  REFRESH_MS: 30000,      // background product refresh interval
  CURRENCY_PREFIX: 'Rs. ',
  LOW_STOCK_THRESHOLD: 5
};

// ====== STATE ==================================================
const state = {
  products: [],          // latest products from the sheet
  productsById: new Map(),
  lastFetch: 0,
  cart: loadCart(),       // [{id, qty}]
  filters: { category: 'all', inStockOnly: false, featuredOnly: false, newOnly: false, sort: 'default', q: '' },
  inventorySort: { key: 'category', dir: 'asc' },
  inventoryFilter: { q: '', category: 'all', status: 'all' },
  dataSignature: '',
  route: { view: 'home', params: {} }
};

// ====== DOM SHORTCUTS ==========================================
const $ = (sel, root = document) => root.querySelector(sel);
const $$ = (sel, root = document) => Array.from(root.querySelectorAll(sel));

// ====== INIT ====================================================
document.addEventListener('DOMContentLoaded', () => {
  $('#year').textContent = new Date().getFullYear();
  bindGlobalUI();
  watchSheetImages();
  window.addEventListener('hashchange', route);
  route();               // handle whatever hash is already in the URL
  loadProducts(true);
  setInterval(() => loadProducts(false), CONFIG.REFRESH_MS);
});

// ====== DATA LOADING ============================================
async function loadProducts(isFirstLoad) {
  try {
    const res = await fetch(`${CONFIG.API_URL}?action=products`, { cache: 'no-store' });
    const data = await res.json();
    if (!data.ok) throw new Error(data.error || 'Could not load products.');

    const products = (data.products || []).filter(p => p && (p.id || p.name));
    const signature = JSON.stringify(products);
    const wasEmpty = state.lastFetch === 0;
    state.lastFetch = Date.now();
    updateInventoryStamp();

    // Background refresh with nothing changed: leave the page alone (no flicker,
    // no lost quantity you were typing on a product page).
    if (!wasEmpty && signature === state.dataSignature) return;
    state.dataSignature = signature;

    state.products = products;
    state.productsById = new Map(products.map(p => [String(p.id), p]));

    reconcileCartWithStock();
    renderCurrentView();
    renderCategoryFilters();
    renderCategoryStrip();
    renderInventoryCategoryOptions();
  } catch (err) {
    console.error(err);
    if (isFirstLoad) {
      showToast("Couldn't load products. Check the API_URL in script.js.");
      renderFetchError();
    }
  }
}

function renderFetchError() {
  const msg = `Products couldn't be loaded. Confirm CONFIG.API_URL in script.js points at your deployed Apps Script web app.`;
  ['#featuredGrid', '#newGrid', '#shopGrid'].forEach(sel => {
    const el = $(sel);
    if (el) el.innerHTML = `<p class="empty-state">${msg}</p>`;
  });
  const shopCount = $('#shopResultCount');
  if (shopCount) shopCount.textContent = 'Could not load products';
  const invBody = $('#inventoryBody');
  if (invBody) invBody.innerHTML = `<tr><td colspan="8" class="inv-loading">${msg}</td></tr>`;
}

/** If stock drops below what's in someone's cart while they're browsing, trim it. */
function reconcileCartWithStock() {
  let changed = false;
  state.cart = state.cart.filter(line => {
    const p = state.productsById.get(String(line.id));
    if (!p || p.quantity <= 0) { changed = true; return false; }
    if (line.qty > p.quantity) { line.qty = p.quantity; changed = true; }
    return true;
  });
  if (changed) { saveCart(); renderCartBadge(); }
}

// ====== ROUTER ===================================================
function route() {
  const hash = location.hash.replace(/^#/, '') || '/home';
  const [pathPart, queryPart] = hash.split('?');
  const parts = pathPart.split('/').filter(Boolean);
  const view = parts[0] || 'home';
  const param = parts[1] || null;
  const params = Object.fromEntries(new URLSearchParams(queryPart || ''));

  state.route = { view, param, params };

  $$('.view').forEach(v => v.classList.remove('active'));
  const target = $(`#view-${view}`) || $('#view-home');
  target.classList.add('active');

  $$('.main-nav a, .mobile-menu a').forEach(a => {
    a.classList.toggle('active', a.dataset.route === `/${view}`);
  });

  closeMobileMenu();
  closeCartDrawer();
  window.scrollTo({ top: 0, behavior: 'instant' in window ? 'instant' : 'auto' });

  if (view === 'shop') {
    state.filters.category = params.cat || 'all';
    state.filters.featuredOnly = params.featured === '1';
    state.filters.newOnly = params.new === '1';
    state.filters.inStockOnly = false;
    syncFilterUI();
  }

  renderCurrentView();
}

function renderCurrentView() {
  if (!state.products.length && state.lastFetch === 0) return; // still loading
  switch (state.route.view) {
    case 'home': renderHome(); break;
    case 'shop': renderShop(); break;
    case 'product': renderProductDetail(state.route.param); break;
    case 'cart': renderCartView(); break;
    case 'inventory': renderInventory(); break;
    default: break;
  }
}

function navigate(hash) { location.hash = hash; }

// ====== RENDER: HOME =============================================
function renderHome() {
  const featured = state.products.filter(p => p.featured).slice(0, 8);
  const fresh = state.products.filter(p => p.isNew).slice(0, 8);
  $('#featuredGrid').innerHTML = (featured.length ? featured : state.products.slice(0, 8))
    .map(productCard).join('') || emptyGridMsg();
  $('#newGrid').innerHTML = (fresh.length ? fresh : state.products.slice(-8))
    .map(productCard).join('') || emptyGridMsg();
}

function emptyGridMsg() {
  return `<p class="empty-state">No products yet — add rows to the Products sheet.</p>`;
}

function renderCategoryStrip() {
  const cats = uniqueCategories();
  $('#categoryStrip').innerHTML = cats.map(c =>
    `<a class="cat-chip" href="#/shop?cat=${encodeURIComponent(c)}" data-route="/shop">${escapeHtml(c)}</a>`
  ).join('');
  bindRouteLinks($('#categoryStrip'));
}

function uniqueCategories() {
  return Array.from(new Set(state.products.map(p => p.category).filter(Boolean))).sort();
}

// ====== RENDER: SHOP ==============================================
function renderShop() {
  let list = applyFilters(state.products);
  $('#shopResultCount').textContent = state.products.length
    ? `${list.length} piece${list.length === 1 ? '' : 's'}`
    : 'Loading products…';

  $('#shopGrid').innerHTML = list.map(productCard).join('');
  $('#shopEmpty').hidden = list.length !== 0;
}

function applyFilters(products) {
  const f = state.filters;
  let list = products.slice();

  if (f.q) {
    const q = f.q.toLowerCase();
    list = list.filter(p =>
      p.name.toLowerCase().includes(q) ||
      String(p.id).toLowerCase().includes(q) ||
      String(p.mandarin || '').toLowerCase().includes(q) ||
      String(p.description || '').toLowerCase().includes(q) ||
      String(p.category || '').toLowerCase().includes(q)
    );
  }
  if (f.category && f.category !== 'all') {
    list = list.filter(p => p.category === f.category);
  }
  if (f.inStockOnly) list = list.filter(p => p.quantity > 0);
  if (f.featuredOnly) list = list.filter(p => p.featured);
  if (f.newOnly) list = list.filter(p => p.isNew);

  switch (f.sort) {
    case 'price-asc': list.sort((a, b) => a.price - b.price); break;
    case 'price-desc': list.sort((a, b) => b.price - a.price); break;
    case 'name-asc': list.sort((a, b) => a.name.localeCompare(b.name)); break;
    default:
      list.sort((a, b) => (b.featured - a.featured) || (b.isNew - a.isNew));
  }
  return list;
}

function renderCategoryFilters() {
  const wrap = $('#filterCategory');
  if (!wrap) return;
  const cats = uniqueCategories();
  wrap.innerHTML = [
    `<label class="check"><input type="radio" name="cat" value="all" ${state.filters.category === 'all' ? 'checked' : ''}> All</label>`,
    ...cats.map(c => `<label class="check"><input type="radio" name="cat" value="${escapeHtml(c)}" ${state.filters.category === c ? 'checked' : ''}> ${escapeHtml(c)}</label>`)
  ].join('');
  $$('input[name="cat"]', wrap).forEach(r => r.addEventListener('change', e => {
    state.filters.category = e.target.value;
    renderShop();
  }));
}

function syncFilterUI() {
  renderCategoryFilters();
  $('#filterInStock').checked = state.filters.inStockOnly;
  $('#filterFeatured').checked = state.filters.featuredOnly;
  $('#sortSelect').value = state.filters.sort;
}

// ====== PRODUCT CARD ===============================================
function productCard(p) {
  const status = stockStatus(p);
  const disabled = p.quantity <= 0 ? 'disabled' : '';
  return `
  <article class="product-card">
    <div class="product-media">
      <a href="#/product/${encodeURIComponent(p.id)}" data-route-product="${p.id}">
        <img src="${escapeAttr(imageFor(p))}" ${sheetAttr(p)} alt="${escapeAttr(p.name)}" loading="lazy"
             onerror="imgFallback(this)">
      </a>
      <span class="stock-badge ${status.cls}">${status.label}</span>
      <span class="price-tag">${formatPrice(p.price)}</span>
    </div>
    <div class="product-info">
      <div class="product-meta-row">
        <span class="product-cat">${escapeHtml(p.category || 'Handmade')}</span>
        <span class="sku-tag">${escapeHtml(p.sku || p.id)}</span>
      </div>
      <h3 class="product-name"><a href="#/product/${encodeURIComponent(p.id)}" data-route-product="${p.id}">${escapeHtml(p.name)}</a></h3>
      ${p.mandarin ? `<p class="product-mandarin">${escapeHtml(p.mandarin)}</p>` : ''}
      <div class="product-actions">
        <button class="btn btn-ghost" ${disabled} data-add-cart="${p.id}">Add to cart</button>
        <button class="btn btn-primary" ${disabled} data-buy-now="${p.id}">Buy now</button>
      </div>
    </div>
  </article>`;
}

function stockStatus(p) {
  if (p.quantity <= 0) return { label: 'Out of stock', cls: 'out' };
  if (p.quantity <= CONFIG.LOW_STOCK_THRESHOLD) return { label: `Only ${p.quantity} left`, cls: 'low' };
  return { label: 'In stock', cls: 'in' };
}

// ====== PICTURES STORED IN THE SHEET ==========================================
/**
 * A picture uploaded straight into the sheet's Image cell has no public link,
 * so the products feed marks it (sheetImage: true) and we fetch it from the
 * Apps Script on demand:  ?action=image&code=F-DC-002  ->  { mime, data(base64) }.
 * Each picture is requested once, kept in memory, and also saved in the
 * browser's Cache Storage for 30 minutes so repeat visits are instant.
 */
const SHEET_IMG = { urls: new Map(), pending: new Map(), queue: [], active: 0, MAX_PARALLEL: 3 };
const SHEET_IMG_CACHE = 'woolly-sheet-images-v1';
const SHEET_IMG_TTL_MS = 30 * 60 * 1000;

/** Attribute that marks an <img> as "fill me from the sheet". */
function sheetAttr(p) {
  if (!p.sheetImage) return '';
  const code = escapeAttr(p.sku || p.id);
  if (p.image) return `data-sheet-backup="${code}"`;          // link first, script copy if it breaks
  if (SHEET_IMG.urls.has(sheetKey(p))) return '';   // already showing the real picture
  return `data-sheet-code="${code}" data-sheet-ver="${escapeAttr(p.imgv || '')}"`;
}

/** One cache entry per photo version: a replaced photo gets a new key, so it is fetched again. */
function sheetKey(p) { return String(p.sku || p.id) + '|' + (p.imgv || ''); }

function runSheetImageQueue() {
  while (SHEET_IMG.active < SHEET_IMG.MAX_PARALLEL && SHEET_IMG.queue.length) {
    const job = SHEET_IMG.queue.shift();
    SHEET_IMG.active++;
    job.fn().then(job.resolve, job.reject).finally(() => {
      SHEET_IMG.active--;
      runSheetImageQueue();
    });
  }
}
function queueSheetImageJob(fn) {
  return new Promise((resolve, reject) => {
    SHEET_IMG.queue.push({ fn, resolve, reject });
    runSheetImageQueue();
  });
}

async function downloadSheetImage(code, ver) {
  const res = await fetch(`${CONFIG.API_URL}?action=image&code=${encodeURIComponent(code)}&v=${encodeURIComponent(ver || '')}`, { cache: 'no-store' });
  const data = await res.json();
  if (!data.ok) throw new Error(data.error || 'picture unavailable');
  const bin = atob(data.data);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return new Blob([bytes], { type: data.mime || 'image/png' });
}

async function loadSheetImage(code, ver) {
  const key = 'sheet-image/' + encodeURIComponent(code) + '/' + encodeURIComponent(ver || '');
  const canCache = typeof caches !== 'undefined';

  if (canCache) {
    try {
      const cache = await caches.open(SHEET_IMG_CACHE);
      const hit = await cache.match(key);
      if (hit && (ver || Date.now() - Number(hit.headers.get('X-Fetched-At') || 0) < SHEET_IMG_TTL_MS)) {
        return URL.createObjectURL(await hit.blob());
      }
    } catch (e) { /* cache unavailable - just download */ }
  }

  const blob = await queueSheetImageJob(() => downloadSheetImage(code, ver));

  if (canCache) {
    try {
      const cache = await caches.open(SHEET_IMG_CACHE);
      await cache.put(key, new Response(blob, {
        headers: { 'Content-Type': blob.type, 'X-Fetched-At': String(Date.now()) }
      }));
    } catch (e) { /* ignore */ }
  }
  return URL.createObjectURL(blob);
}

function getSheetImageUrl(code, ver) {
  const key = code + '|' + (ver || '');
  if (SHEET_IMG.urls.has(key)) return Promise.resolve(SHEET_IMG.urls.get(key));
  if (SHEET_IMG.pending.has(key)) return SHEET_IMG.pending.get(key);
  const pr = loadSheetImage(code, ver)
    .then(url => { SHEET_IMG.urls.set(key, url); return url; })
    .finally(() => SHEET_IMG.pending.delete(key));
  SHEET_IMG.pending.set(key, pr);
  return pr;
}

function fillSheetImage(img) {
  const code = img.dataset.sheetCode;
  if (!code || img.dataset.sheetLoading === '1') return;
  img.dataset.sheetLoading = '1';
  getSheetImageUrl(code, img.dataset.sheetVer || '')
    .then(url => { if (url) img.src = url; })
    .catch(err => {
      console.warn('Picture for ' + code + ' could not be loaded:', err.message);
      img.removeAttribute('data-sheet-code'); // keep the placeholder, don't retry in a loop
    })
    .finally(() => { img.dataset.sheetLoading = ''; });
}

/** Watches the page: any product <img> that appears with data-sheet-code gets filled in. */
function watchSheetImages() {
  const scan = node => {
    if (node.nodeType !== 1) return;
    if (node.matches && node.matches('img[data-sheet-code]')) fillSheetImage(node);
    if (node.querySelectorAll) node.querySelectorAll('img[data-sheet-code]').forEach(fillSheetImage);
  };
  new MutationObserver(muts => muts.forEach(m => m.addedNodes.forEach(scan)))
    .observe(document.body, { childList: true, subtree: true });
  scan(document.body);
}

/**
 * Where a product's picture comes from. In order:
 *  1. a full link in the sheet's Image/ImageURL column (http..., data:...)
 *  2. a bare file name in that column (e.g. "tote.jpg") -> looked up in /images
 *  3. a picture uploaded INTO the sheet cell -> placeholder first, then the real
 *     picture is fetched from the Apps Script (see "PICTURES STORED IN THE SHEET")
 *  4. nothing in the sheet -> /images/<Code>.jpg  (e.g. images/F-DC-001.jpg)
 */
function imageFor(p) {
  const v = String(p.image || '').trim();
  if (/^(https?:|data:|\/\/|\/)/i.test(v)) return v;
  if (v) return 'images/' + v;
  if (p.sheetImage) {
    const ready = SHEET_IMG.urls.get(sheetKey(p));
    return ready || placeholderImg(); // real picture is swapped in by fillSheetImage()
  }
  // no link: look for a photo named after the SKU in the website's images folder,
  // e.g. images/F003-EG-MR.jpg (other file types are tried automatically)
  return 'images/' + encodeURIComponent(p.sku || p.id) + '.jpg';
}

/** If a picture fails to load, try the other common file types, then the placeholder. */
const IMAGE_TYPES = ['jpg', 'jpeg', 'png', 'webp', 'JPG', 'JPEG', 'PNG'];
function imgFallback(img) {
  // a picture link that failed, but the script can still send the picture from the sheet
  if (img.dataset.sheetBackup && !img.dataset.sheetCode) {
    img.dataset.sheetCode = img.dataset.sheetBackup;
    img.src = placeholderImg();
    fillSheetImage(img);
    return;
  }
  const m = (img.getAttribute('src') || '').match(/^images\/(.+)\.([A-Za-z]+)$/);
  const tried = Number(img.dataset.tried || 0);
  if (m && tried < IMAGE_TYPES.length - 1) {
    img.dataset.tried = String(tried + 1);
    img.src = 'images/' + m[1] + '.' + IMAGE_TYPES[tried + 1];
    return;
  }
  img.onerror = null;
  img.src = placeholderImg();
}

function placeholderImg() {
  return 'data:image/svg+xml;utf8,' + encodeURIComponent(
    `<svg xmlns="http://www.w3.org/2000/svg" width="400" height="400"><rect width="400" height="400" fill="#F1E9D8"/><text x="50%" y="50%" font-family="sans-serif" font-size="18" fill="#8A7F6D" text-anchor="middle">Baccabi'u</text></svg>`
  );
}

function formatPrice(n) {
  return CONFIG.CURRENCY_PREFIX + Number(n).toLocaleString('en-IN', { maximumFractionDigits: 2 });
}

// ====== RENDER: PRODUCT DETAIL ======================================
function renderProductDetail(id) {
  const p = state.productsById.get(String(id));
  const wrap = $('#productDetail');
  if (!p) {
    wrap.innerHTML = state.products.length
      ? `<p class="empty-state">That product isn't available anymore.</p><p><a class="text-link" href="#/shop" data-route="/shop">← Back to shop</a></p>`
      : `<p class="empty-state">Loading…</p>`;
    bindRouteLinks(wrap);
    return;
  }
  const status = stockStatus(p);
  const max = Math.max(p.quantity, 0);

  wrap.innerHTML = `
    <div class="pd-grid">
      <div class="pd-media">
        <img src="${escapeAttr(imageFor(p))}" ${sheetAttr(p)} alt="${escapeAttr(p.name)}" onerror="imgFallback(this)">
      </div>
      <div class="pd-copy">
        <p class="pd-crumb"><a href="#/shop" data-route="/shop">Shop</a> / <a href="#/shop?cat=${encodeURIComponent(p.category)}" data-route="/shop">${escapeHtml(p.category)}</a></p>
        <div class="pd-sku">SKU: ${escapeHtml(p.sku || p.id)}</div>
        <h1>${escapeHtml(p.name)}</h1>
        ${p.mandarin ? `<p class="pd-mandarin">${escapeHtml(p.mandarin)}</p>` : ''}
        <div class="pd-price">${formatPrice(p.price)}</div>
        <p class="pd-desc">${escapeHtml(p.description)}</p>
        <span class="pd-stock stock-badge ${status.cls}">${status.label}</span>
        <div class="qty-row">
          <div class="qty-stepper">
            <button type="button" data-qty-down>&minus;</button>
            <input type="number" id="pdQty" value="1" min="1" max="${max || 1}" ${max ? '' : 'disabled'}>
            <button type="button" data-qty-up>+</button>
          </div>
          <span class="muted">${max} available</span>
        </div>
        <div class="pd-actions">
          <button class="btn btn-ghost" id="pdAddCart" ${max ? '' : 'disabled'}>Add to cart</button>
          <button class="btn btn-primary" id="pdBuyNow" ${max ? '' : 'disabled'}>Buy now</button>
        </div>
      </div>
    </div>`;

  bindRouteLinks(wrap);

  const qtyInput = $('#pdQty');
  $('[data-qty-down]', wrap).addEventListener('click', () => {
    qtyInput.value = Math.max(1, Number(qtyInput.value) - 1);
  });
  $('[data-qty-up]', wrap).addEventListener('click', () => {
    qtyInput.value = Math.min(max, Number(qtyInput.value) + 1);
  });
  qtyInput.addEventListener('change', () => {
    qtyInput.value = clamp(Number(qtyInput.value) || 1, 1, max);
  });

  $('#pdAddCart').addEventListener('click', () => {
    addToCart(p.id, Number(qtyInput.value) || 1);
  });
  $('#pdBuyNow').addEventListener('click', () => {
    openCheckout(p.id, Number(qtyInput.value) || 1);
  });
}

function clamp(n, min, max) { return Math.min(Math.max(n, min), max); }

// ====== CART ==========================================================
function loadCart() {
  try { return JSON.parse(localStorage.getItem('woolly_cart') || '[]'); }
  catch { return []; }
}
function saveCart() {
  localStorage.setItem('woolly_cart', JSON.stringify(state.cart));
  renderCartBadge();
}

function addToCart(id, qty) {
  const p = state.productsById.get(String(id));
  if (!p || p.quantity <= 0) return showToast('That piece is out of stock.');
  const line = state.cart.find(l => String(l.id) === String(id));
  const wanted = (line ? line.qty : 0) + qty;
  const capped = Math.min(wanted, p.quantity);
  if (capped < wanted) showToast(`Only ${p.quantity} available — cart adjusted.`);

  if (line) line.qty = capped;
  else state.cart.push({ id: String(id), qty: capped });

  saveCart();
  showToast(`${p.name} added to cart.`);
  renderDrawer();
  if (state.route.view === 'cart') renderCartView();
}

function updateCartQty(id, qty) {
  const p = state.productsById.get(String(id));
  const line = state.cart.find(l => String(l.id) === String(id));
  if (!line) return;
  const max = p ? p.quantity : 999;
  line.qty = clamp(qty, 1, max);
  saveCart();
  renderDrawer();
  if (state.route.view === 'cart') renderCartView();
}

function removeFromCart(id) {
  state.cart = state.cart.filter(l => String(l.id) !== String(id));
  saveCart();
  renderDrawer();
  if (state.route.view === 'cart') renderCartView();
}

function cartLines() {
  return state.cart
    .map(l => ({ line: l, product: state.productsById.get(String(l.id)) }))
    .filter(x => x.product);
}
function cartTotal() {
  return cartLines().reduce((sum, x) => sum + x.product.price * x.line.qty, 0);
}
function cartCount() {
  return state.cart.reduce((sum, l) => sum + l.qty, 0);
}

function renderCartBadge() {
  $('#cartCount').textContent = cartCount();
}

function cartLineHtml(product, line, { removable = true } = {}) {
  return `
  <div class="cart-line" data-cart-line="${product.id}">
    <img src="${escapeAttr(imageFor(product))}" ${sheetAttr(product)} alt="${escapeAttr(product.name)}" onerror="imgFallback(this)">
    <div>
      <div class="cart-line-name">${escapeHtml(product.name)}</div>
      <div class="cart-line-meta">${formatPrice(product.price)} each</div>
      <div class="cart-line-qty">
        <button type="button" data-cart-dec="${product.id}">&minus;</button>
        <span>${line.qty}</span>
        <button type="button" data-cart-inc="${product.id}">+</button>
      </div>
      ${removable ? `<button class="cart-line-remove" data-cart-remove="${product.id}">Remove</button>` : ''}
    </div>
    <div class="cart-line-total">${formatPrice(product.price * line.qty)}</div>
  </div>`;
}

function renderDrawer() {
  const body = $('#drawerBody');
  const foot = $('#drawerFoot');
  const lines = cartLines();

  if (!lines.length) {
    body.innerHTML = `<p class="cart-empty">Your cart is empty.</p>`;
    foot.innerHTML = `<a class="btn btn-primary full" href="#/shop" data-route="/shop">Browse the shop</a>`;
    bindRouteLinks(foot);
    return;
  }

  body.innerHTML = lines.map(({ product, line }) => cartLineHtml(product, line)).join('');
  foot.innerHTML = `
    <div class="cart-summary-row total"><span>Total</span><span>${formatPrice(cartTotal())}</span></div>
    <a class="btn btn-primary full" href="#/cart" data-route="/cart" style="margin-top:14px;">View cart &amp; checkout</a>`;
  bindRouteLinks(foot);
  bindCartLineControls(body);
}

function renderCartView() {
  const el = $('#cartView');
  const lines = cartLines();
  if (!lines.length) {
    el.innerHTML = `<p class="cart-empty">Your cart is empty. <a class="text-link" href="#/shop" data-route="/shop">Go find something you like →</a></p>`;
    bindRouteLinks(el);
    return;
  }
  el.innerHTML = `
    <div class="cart-table-wrap">
      ${lines.map(({ product, line }) => cartLineHtml(product, line)).join('')}
      <div class="cart-page-summary">
        <div class="cart-summary-row"><span>Subtotal</span><span>${formatPrice(cartTotal())}</span></div>
        <div class="cart-summary-row total"><span>Total</span><span>${formatPrice(cartTotal())}</span></div>
        <button class="btn btn-primary full" id="cartCheckoutBtn" style="margin-top:16px;">Checkout</button>
      </div>
    </div>`;
  bindCartLineControls(el);
  $('#cartCheckoutBtn').addEventListener('click', () => openCheckoutFromCart());
}

function bindCartLineControls(root) {
  $$('[data-cart-inc]', root).forEach(b => b.addEventListener('click', () => {
    const id = b.dataset.cartInc;
    const line = state.cart.find(l => String(l.id) === String(id));
    updateCartQty(id, line.qty + 1);
  }));
  $$('[data-cart-dec]', root).forEach(b => b.addEventListener('click', () => {
    const id = b.dataset.cartDec;
    const line = state.cart.find(l => String(l.id) === String(id));
    if (line.qty <= 1) removeFromCart(id); else updateCartQty(id, line.qty - 1);
  }));
  $$('[data-cart-remove]', root).forEach(b => b.addEventListener('click', () => removeFromCart(b.dataset.cartRemove)));
}

// ====== INVENTORY =========================================================
/**
 * Every product in the main inventory (Warehouse tab, table 1), with its
 * picture, stock in the inventory, what open web orders are holding, and
 * what's still available. Refreshed on the same 30-second cycle as the shop.
 */
function inventoryStatus(p) {
  return p.quantity <= 0 ? 'out' : (p.quantity <= CONFIG.LOW_STOCK_THRESHOLD ? 'low' : 'in');
}

function inventoryRows() {
  const f = state.inventoryFilter;
  const q = f.q.toLowerCase();
  let rows = state.products.map(p => {
    const stock = Number.isFinite(p.stock) ? p.stock : p.quantity; // older Code.gs has no "stock"
    return {
      ...p,
      stock,
      reserved: Number(p.reserved) || 0,
      value: p.price * stock,
      status: inventoryStatus(p)
    };
  });
  if (q) rows = rows.filter(r =>
    String(r.sku || r.id).toLowerCase().includes(q) ||
    String(r.name).toLowerCase().includes(q) ||
    String(r.mandarin || '').toLowerCase().includes(q) ||
    String(r.category || '').toLowerCase().includes(q));
  if (f.category !== 'all') rows = rows.filter(r => r.category === f.category);
  if (f.status !== 'all') rows = rows.filter(r => r.status === f.status);

  const { key, dir } = state.inventorySort;
  const mul = dir === 'asc' ? 1 : -1;
  const statusOrder = { out: 0, low: 1, in: 2 };
  rows.sort((a, b) => {
    let av = a[key], bv = b[key];
    if (key === 'status') { av = statusOrder[av]; bv = statusOrder[bv]; }
    if (key === 'image') { av = a.image || a.sheetImage ? 1 : 0; bv = b.image || b.sheetImage ? 1 : 0; }
    if (typeof av === 'string' || typeof bv === 'string') {
      const c = String(av || '').localeCompare(String(bv || ''), undefined, { numeric: true, sensitivity: 'base' });
      return c * mul || String(a.sku).localeCompare(String(b.sku), undefined, { numeric: true });
    }
    if (av < bv) return -1 * mul;
    if (av > bv) return 1 * mul;
    return String(a.sku).localeCompare(String(b.sku), undefined, { numeric: true });
  });
  return rows;
}

function renderInventoryCategoryOptions() {
  const sel = $('#invCategory');
  if (!sel) return;
  const current = state.inventoryFilter.category;
  sel.innerHTML = `<option value="all">All categories</option>` +
    uniqueCategories().map(c => `<option value="${escapeAttr(c)}">${escapeHtml(c)}</option>`).join('');
  sel.value = uniqueCategories().includes(current) ? current : 'all';
  state.inventoryFilter.category = sel.value;
}

function updateInventoryStamp() {
  const el = $('#inventoryUpdated');
  if (el && state.lastFetch) {
    el.textContent = 'Live from the main inventory · updated ' +
      new Date(state.lastFetch).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', second: '2-digit' });
  }
}

function renderInventory() {
  const body = $('#inventoryBody');
  if (!state.lastFetch) {
    body.innerHTML = `<tr><td colspan="8" class="inv-loading">Loading inventory…</td></tr>`;
    return;
  }
  const all = state.products;
  const rows = inventoryRows();
  const filtered = rows.length !== all.length;

  $('#inventoryEmpty').hidden = all.length !== 0;
  $('#inventoryTable').style.display = all.length ? '' : 'none';

  // ---- summary cards (always for the whole inventory, not just the filtered rows) ----
  const sum = (list, f) => list.reduce((s, r) => s + f(r), 0);
  const stockOf = p => (Number.isFinite(p.stock) ? p.stock : p.quantity);
  const totalUnits = sum(all, stockOf);
  const totalValue = sum(all, p => p.price * stockOf(p));
  const outCount = all.filter(p => inventoryStatus(p) === 'out').length;
  const lowCount = all.filter(p => inventoryStatus(p) === 'low').length;

  $('#inventorySummary').innerHTML = `
    <div class="summary-card"><span class="num">${all.length}</span><span class="label">Products</span></div>
    <div class="summary-card"><span class="num">${totalUnits.toLocaleString('en-IN')}</span><span class="label">Units in stock</span></div>
    <div class="summary-card"><span class="num">${formatPrice(totalValue)}</span><span class="label">Stock value</span></div>
    <div class="summary-card ${lowCount ? 'warn' : ''}"><span class="num">${lowCount}</span><span class="label">Low stock (≤ ${CONFIG.LOW_STOCK_THRESHOLD})</span></div>
    <div class="summary-card ${outCount ? 'warn' : ''}"><span class="num">${outCount}</span><span class="label">Out of stock</span></div>`;

  $('#inventoryCount').textContent = filtered
    ? `Showing ${rows.length} of ${all.length} products`
    : `Showing all ${all.length} products`;

  // ---- table body ----
  body.innerHTML = rows.length ? rows.map(r => {
    const status = { in: 'In stock', low: 'Low stock', out: 'Out of stock' }[r.status];
    const link = `#/product/${encodeURIComponent(r.id)}`;
    return `
    <tr>
      <td class="col-image" data-label="Image"><a href="${link}"><img src="${escapeAttr(imageFor(r))}" ${sheetAttr(r)} alt="${escapeAttr(r.name)}" loading="lazy" onerror="imgFallback(this)"></a></td>
      <td class="col-code" data-label="Code">${escapeHtml(r.sku || r.id)}</td>
      <td data-label="Category">${escapeHtml(r.category || '—')}</td>
      <td class="col-name" data-label="Item"><a href="${link}">${escapeHtml(r.name)}</a>${r.isNew ? ' <span class="inv-new">New</span>' : ''}${r.mandarin ? `<div class="inv-mandarin">${escapeHtml(r.mandarin)}</div>` : ''}</td>
      <td class="col-num" data-label="Rate">${formatPrice(r.price)}</td>
      <td class="col-num strong" data-label="Stock QTY">${r.stock.toLocaleString('en-IN')}</td>
      <td class="col-num" data-label="Value">${formatPrice(r.value)}</td>
      <td data-label="Status"><span class="inv-status ${r.status}">${status}</span></td>
    </tr>`;
  }).join('') : `<tr><td colspan="8" class="inv-loading">No products match these filters.</td></tr>`;

  const fUnits = sum(rows, r => r.stock), fValue = sum(rows, r => r.value);
  $('#inventoryFoot').innerHTML = rows.length ? `
    <tr>
      <td colspan="5" data-label="">${filtered ? 'Totals (filtered)' : 'Totals'}</td>
      <td class="col-num" data-label="Stock QTY">${fUnits.toLocaleString('en-IN')}</td>
      <td class="col-num" data-label="Value">${formatPrice(fValue)}</td>
      <td></td>
    </tr>` : '';

  // reflect current sort in header
  $$('#inventoryTable thead th').forEach(th => {
    th.classList.toggle('sorted', th.dataset.sort === state.inventorySort.key);
    th.classList.toggle('asc', th.dataset.sort === state.inventorySort.key && state.inventorySort.dir === 'asc');
  });
  updateInventoryStamp();
}

function exportInventoryCsv() {
  const rows = inventoryRows();
  const header = ['Code', 'Category', 'Item', 'Mandarin', 'Rate', 'Stock QTY', 'Value', 'Status', 'Image'];
  const lines = [header.join(',')];
  rows.forEach(r => {
    const cells = [r.sku || r.id, r.category, r.name, r.mandarin || '', r.price, r.stock, r.value, r.status,
      /^https?:/i.test(r.image || '') ? r.image : (r.sheetImage ? 'picture in sheet' : '')]
      .map(v => `"${String(v).replace(/"/g, '""')}"`);
    lines.push(cells.join(','));
  });
  const blob = new Blob(['﻿' + lines.join('\n')], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `inventory-${new Date().toISOString().slice(0, 10)}.csv`;
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

// ====== CHECKOUT ========================================================
let checkoutContext = null; // { mode: 'single'|'cart', productId?, quantity? }

function openCheckout(productId, quantity) {
  const p = state.productsById.get(String(productId));
  if (!p || p.quantity < quantity) return showToast('Not enough stock for that quantity.');
  checkoutContext = { mode: 'single', productId, quantity };
  renderCheckoutSummary();
  showModal();
}

function openCheckoutFromCart() {
  if (!cartLines().length) return;
  checkoutContext = { mode: 'cart' };
  renderCheckoutSummary();
  showModal();
}

function renderCheckoutSummary() {
  const el = $('#checkoutSummary');
  if (checkoutContext.mode === 'single') {
    const p = state.productsById.get(String(checkoutContext.productId));
    el.innerHTML = `
      <img src="${escapeAttr(imageFor(p))}" ${sheetAttr(p)} alt="" onerror="imgFallback(this)">
      <div class="checkout-summary-meta">
        <strong>${escapeHtml(p.name)}</strong>
        <span>${formatPrice(p.price)} × ${checkoutContext.quantity}</span><br>
        <span class="muted">${p.quantity} in stock</span>
      </div>
      <div class="checkout-summary-total">${formatPrice(p.price * checkoutContext.quantity)}</div>`;
  } else {
    const lines = cartLines();
    el.innerHTML = `
      <div class="checkout-summary-meta">
        <strong>${lines.length} item${lines.length === 1 ? '' : 's'} in cart</strong>
        ${lines.map(({ product, line }) => `<div>${escapeHtml(product.name)} × ${line.qty}</div>`).join('')}
      </div>
      <div class="checkout-summary-total">${formatPrice(cartTotal())}</div>`;
  }
}

function showModal() {
  $('#checkoutForm').hidden = false;
  $('#checkoutForm').reset();
  $('#checkoutError').hidden = true;
  $('#checkoutSuccess').hidden = true;
  $('#checkoutSummary').style.display = '';
  $('#checkoutModal').classList.add('open');
  $('#modalScrim').classList.add('open');
  document.body.style.overflow = 'hidden';
}
function hideModal() {
  $('#checkoutModal').classList.remove('open');
  $('#modalScrim').classList.remove('open');
  document.body.style.overflow = '';
}

async function submitCheckout(e) {
  e.preventDefault();
  const btn = $('#checkoutSubmit');
  const errEl = $('#checkoutError');
  errEl.hidden = true;

  const customer = {
    name: $('#cName').value.trim(),
    email: $('#cEmail').value.trim(),
    phone: $('#cPhone').value.trim(),
    address: $('#cAddress').value.trim(),
    note: $('#cNote').value.trim()
  };
  if (!customer.name || !customer.email || !customer.phone || !customer.address) {
    errEl.textContent = 'Please fill in all required fields.';
    errEl.hidden = false;
    return;
  }

  const orderLines = checkoutContext.mode === 'single'
    ? [{ productId: checkoutContext.productId, quantity: checkoutContext.quantity }]
    : cartLines().map(({ product, line }) => ({ productId: product.id, quantity: line.qty }));

  btn.disabled = true;
  btn.textContent = 'Placing order…';

  try {
    const results = [];
    // Orders are created one product at a time against Code.gs, which is
    // the simplest shape for the sheet's one-product-per-row Orders log.
    for (const ol of orderLines) {
      const res = await postOrder({ ...ol, customer });
      if (!res.ok) throw new Error(res.error || 'Order failed.');
      results.push(res);
    }

    // remove the ordered lines from the cart, then pull fresh stock
    // numbers for everything else in the background
    orderLines.forEach(ol => removeFromCart(ol.productId));

    showSuccess(results);
    loadProducts(false);
  } catch (err) {
    errEl.textContent = err.message || 'Something went wrong placing your order.';
    errEl.hidden = false;
  } finally {
    btn.disabled = false;
    btn.textContent = 'Place order';
  }
}

async function postOrder(payload) {
  // Sent as text/plain so the browser skips the CORS preflight that
  // Apps Script web apps can't respond to.
  const res = await fetch(CONFIG.API_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify({ action: 'createOrder', ...payload })
  });
  return res.json();
}

function showSuccess(results) {
  $('#checkoutForm').hidden = true;
  $('#checkoutSummary').style.display = 'none';
  const last = results[results.length - 1];
  const idsText = results.map(r => r.orderId).join(', ');
  $('#successMessage').textContent = results.length > 1
    ? `Your order has been placed successfully. Order IDs: ${idsText}.`
    : `Your order has been placed successfully. Order ID: ${last.orderId}. Total: ${formatPrice(last.total)}. Remaining stock: ${last.remainingStock}.`;
  $('#checkoutSuccess').hidden = false;
}

// ====== GLOBAL UI BINDINGS ===============================================
function bindGlobalUI() {
  document.addEventListener('click', globalClickRouter);

  $('#hamburgerBtn').addEventListener('click', toggleMobileMenu);
  $('#scrim').addEventListener('click', closeMobileMenu);

  $('#searchToggle').addEventListener('click', toggleSearchBar);
  $('#searchClose').addEventListener('click', closeSearchBar);
  $('#searchInput').addEventListener('input', debounce(e => {
    state.filters.q = e.target.value.trim();
    if (state.route.view !== 'shop') navigate('/shop');
    else renderShop();
  }, 200));

  $('#cartToggle').addEventListener('click', openCartDrawer);
  $('#cartDrawerClose').addEventListener('click', closeCartDrawer);
  $('#drawerScrim').addEventListener('click', closeCartDrawer);

  $('#filterInStock').addEventListener('change', e => { state.filters.inStockOnly = e.target.checked; renderShop(); });
  $('#filterFeatured').addEventListener('change', e => { state.filters.featuredOnly = e.target.checked; renderShop(); });
  $('#sortSelect').addEventListener('change', e => { state.filters.sort = e.target.value; renderShop(); });
  $('#clearFilters').addEventListener('click', () => {
    state.filters = { category: 'all', inStockOnly: false, featuredOnly: false, newOnly: false, sort: 'default', q: '' };
    $('#searchInput').value = '';
    syncFilterUI();
    renderShop();
  });

  $('#checkoutClose').addEventListener('click', hideModal);
  $('#modalScrim').addEventListener('click', hideModal);
  $('#checkoutForm').addEventListener('submit', submitCheckout);
  $('#successClose').addEventListener('click', () => { hideModal(); navigate('/shop'); });

  $$('#inventoryTable thead th[data-sort]').forEach(th => {
    th.addEventListener('click', () => {
      const key = th.dataset.sort;
      if (state.inventorySort.key === key) {
        state.inventorySort.dir = state.inventorySort.dir === 'asc' ? 'desc' : 'asc';
      } else {
        state.inventorySort = { key, dir: 'asc' };
      }
      renderInventory();
    });
  });
  $('#inventoryRefresh').addEventListener('click', async () => {
    showToast('Refreshing inventory…');
    await loadProducts(false);
    renderInventory();
  });
  $('#invSearch').addEventListener('input', debounce(e => {
    state.inventoryFilter.q = e.target.value.trim();
    renderInventory();
  }, 150));
  $('#invCategory').addEventListener('change', e => { state.inventoryFilter.category = e.target.value; renderInventory(); });
  $('#invStatus').addEventListener('change', e => { state.inventoryFilter.status = e.target.value; renderInventory(); });
  $('#inventoryExport').addEventListener('click', exportInventoryCsv);

  renderCartBadge();
}

/** Delegated handler for anything rendered dynamically (product cards, etc). */
function globalClickRouter(e) {
  const routeLink = e.target.closest('[data-route]');
  if (routeLink) { /* normal anchor navigation via hash, nothing extra needed */ }

  const prodLink = e.target.closest('[data-route-product]');
  if (prodLink) { /* anchor handles navigation */ }

  const addBtn = e.target.closest('[data-add-cart]');
  if (addBtn && !addBtn.disabled) { addToCart(addBtn.dataset.addCart, 1); }

  const buyBtn = e.target.closest('[data-buy-now]');
  if (buyBtn && !buyBtn.disabled) { openCheckout(buyBtn.dataset.buyNow, 1); }
}

function bindRouteLinks() { /* no-op: navigation is plain <a href="#/..."> now */ }

function toggleMobileMenu() {
  const open = $('#mobileMenu').classList.toggle('open');
  $('#scrim').classList.toggle('open', open);
  $('#hamburgerBtn').setAttribute('aria-expanded', String(open));
}
function closeMobileMenu() {
  $('#mobileMenu').classList.remove('open');
  $('#scrim').classList.remove('open');
  $('#hamburgerBtn').setAttribute('aria-expanded', 'false');
}

function toggleSearchBar() {
  const bar = $('#searchBar');
  bar.classList.toggle('open');
  if (bar.classList.contains('open')) $('#searchInput').focus();
}
function closeSearchBar() { $('#searchBar').classList.remove('open'); }

function openCartDrawer() {
  renderDrawer();
  $('#cartDrawer').classList.add('open');
  $('#drawerScrim').classList.add('open');
}
function closeCartDrawer() {
  $('#cartDrawer').classList.remove('open');
  $('#drawerScrim').classList.remove('open');
}

// ====== HELPERS ===========================================================
function showToast(msg) {
  const t = $('#toast');
  t.textContent = msg;
  t.classList.add('show');
  clearTimeout(showToast._t);
  showToast._t = setTimeout(() => t.classList.remove('show'), 2600);
}

function debounce(fn, ms) {
  let h;
  return (...args) => { clearTimeout(h); h = setTimeout(() => fn(...args), ms); };
}

function escapeHtml(str) {
  return String(str).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}
function escapeAttr(str) { return escapeHtml(str); }