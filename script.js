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
  API_URL: 'https://script.google.com/macros/s/AKfycbz3DV7wqny6CHZcwCRydmNUuRxD56b70uePMoW5Vp2v184mPnbF2O2pSBZ8qxYfmmBJ/exec',
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
  filters: { category: 'all', inStockOnly: false, featuredOnly: false, sort: 'default', q: '' },
  inventorySort: { key: 'category', dir: 'asc' },
  route: { view: 'home', params: {} }
};

// ====== DOM SHORTCUTS ==========================================
const $ = (sel, root = document) => root.querySelector(sel);
const $$ = (sel, root = document) => Array.from(root.querySelectorAll(sel));

// ====== INIT ====================================================
document.addEventListener('DOMContentLoaded', () => {
  $('#year').textContent = new Date().getFullYear();
  bindGlobalUI();
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

    state.products = data.products;
    state.productsById = new Map(data.products.map(p => [String(p.id), p]));
    state.lastFetch = Date.now();

    reconcileCartWithStock();
    renderCurrentView();
    renderCategoryFilters();
    renderCategoryStrip();

    if (!isFirstLoad) {
      // quiet refresh — only nudge the user if something they're looking
      // at changed, rather than interrupting with a toast every 30s
    }
  } catch (err) {
    console.error(err);
    if (isFirstLoad) {
      showToast("Couldn't load products. Check the API_URL in script.js.");
      renderFetchError();
    }
  }
}

function renderFetchError() {
  const grids = ['#featuredGrid', '#newGrid', '#shopGrid'];
  grids.forEach(sel => {
    const el = $(sel);
    if (el) el.innerHTML = `<p class="empty-state">Products couldn't be loaded. Confirm CONFIG.API_URL in script.js points at your deployed Apps Script web app.</p>`;
  });
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
      p.description.toLowerCase().includes(q) ||
      p.category.toLowerCase().includes(q)
    );
  }
  if (f.category && f.category !== 'all') {
    list = list.filter(p => p.category === f.category);
  }
  if (f.inStockOnly) list = list.filter(p => p.quantity > 0);
  if (f.featuredOnly) list = list.filter(p => p.featured);

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
        <img src="${escapeAttr(p.image || placeholderImg())}" alt="${escapeAttr(p.name)}" loading="lazy"
             onerror="this.src='${placeholderImg()}'">
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

function placeholderImg() {
  return 'data:image/svg+xml;utf8,' + encodeURIComponent(
    `<svg xmlns="http://www.w3.org/2000/svg" width="400" height="400"><rect width="400" height="400" fill="#F1E9D8"/><text x="50%" y="50%" font-family="sans-serif" font-size="18" fill="#8A7F6D" text-anchor="middle">Woolly &amp; Co.</text></svg>`
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
        <img src="${escapeAttr(p.image || placeholderImg())}" alt="${escapeAttr(p.name)}" onerror="this.src='${placeholderImg()}'">
      </div>
      <div class="pd-copy">
        <p class="pd-crumb"><a href="#/shop" data-route="/shop">Shop</a> / <a href="#/shop?cat=${encodeURIComponent(p.category)}" data-route="/shop">${escapeHtml(p.category)}</a></p>
        <div class="pd-sku">SKU: ${escapeHtml(p.sku || p.id)}</div>
        <h1>${escapeHtml(p.name)}</h1>
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
    <img src="${escapeAttr(product.image || placeholderImg())}" alt="${escapeAttr(product.name)}" onerror="this.src='${placeholderImg()}'">
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
 * A plain, sortable table of everything in the Products sheet — meant for
 * the store owner to eyeball total stock, not a customer-facing page.
 * Reads straight from state.products, which is refreshed on the same
 * 30-second cycle as the shop.
 */
function inventoryRows() {
  const rows = state.products.map(p => ({
    ...p,
    value: p.price * p.quantity,
    status: p.quantity <= 0 ? 'out' : (p.quantity <= CONFIG.LOW_STOCK_THRESHOLD ? 'low' : 'in')
  }));
  const { key, dir } = state.inventorySort;
  const mul = dir === 'asc' ? 1 : -1;
  rows.sort((a, b) => {
    let av = a[key], bv = b[key];
    if (typeof av === 'string') { av = av.toLowerCase(); bv = String(bv).toLowerCase(); }
    if (av < bv) return -1 * mul;
    if (av > bv) return 1 * mul;
    return 0;
  });
  return rows;
}

function renderInventory() {
  const rows = inventoryRows();
  const el = $('#inventoryTable');
  $('#inventoryEmpty').hidden = rows.length !== 0 || state.lastFetch === 0;
  el.style.display = rows.length || state.lastFetch === 0 ? '' : 'none';

  // ---- summary cards ----
  const totalUnits = rows.reduce((s, r) => s + r.quantity, 0);
  const totalValue = rows.reduce((s, r) => s + r.value, 0);
  const outCount = rows.filter(r => r.status === 'out').length;
  const lowCount = rows.filter(r => r.status === 'low').length;

  $('#inventorySummary').innerHTML = `
    <div class="summary-card"><span class="num">${rows.length}</span><span class="label">Products</span></div>
    <div class="summary-card"><span class="num">${totalUnits.toLocaleString('en-IN')}</span><span class="label">Total units in stock</span></div>
    <div class="summary-card"><span class="num">${formatPrice(totalValue)}</span><span class="label">Total stock value</span></div>
    <div class="summary-card ${lowCount ? 'warn' : ''}"><span class="num">${lowCount}</span><span class="label">Low stock (≤ ${CONFIG.LOW_STOCK_THRESHOLD})</span></div>
    <div class="summary-card ${outCount ? 'warn' : ''}"><span class="num">${outCount}</span><span class="label">Out of stock</span></div>`;

  // ---- table body ----
  $('#inventoryBody').innerHTML = rows.map(r => {
    const status = { in: 'In stock', low: 'Low stock', out: 'Out of stock' }[r.status];
    return `
    <tr>
      <td class="col-image"><img src="${escapeAttr(r.image || placeholderImg())}" alt="" onerror="this.src='${placeholderImg()}'"></td>
      <td>${escapeHtml(r.sku || r.id)}</td>
      <td>${escapeHtml(r.category || '—')}</td>
      <td>${escapeHtml(r.name)}</td>
      <td class="col-num">${formatPrice(r.price)}</td>
      <td class="col-num">${r.quantity}</td>
      <td class="col-num">${formatPrice(r.value)}</td>
      <td><span class="inv-status ${r.status}">${status}</span></td>
    </tr>`;
  }).join('');

  $('#inventoryFoot').innerHTML = rows.length ? `
    <tr>
      <td colspan="5">Totals</td>
      <td class="col-num">${totalUnits.toLocaleString('en-IN')}</td>
      <td class="col-num">${formatPrice(totalValue)}</td>
      <td></td>
    </tr>` : '';

  // reflect current sort in header
  $$('#inventoryTable thead th').forEach(th => {
    th.classList.toggle('sorted', th.dataset.sort === state.inventorySort.key);
    th.classList.toggle('asc', th.dataset.sort === state.inventorySort.key && state.inventorySort.dir === 'asc');
  });
}

function exportInventoryCsv() {
  const rows = inventoryRows();
  const header = ['Code', 'Category', 'Item', 'Rate', 'Stock Qty', 'Value', 'Status'];
  const lines = [header.join(',')];
  rows.forEach(r => {
    const cells = [r.sku || r.id, r.category, r.name, r.price, r.quantity, r.value, r.status]
      .map(v => `"${String(v).replace(/"/g, '""')}"`);
    lines.push(cells.join(','));
  });
  const blob = new Blob([lines.join('\n')], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `inventory-${new Date().toISOString().slice(0, 10)}.csv`;
  a.click();
  URL.revokeObjectURL(url);
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
      <img src="${escapeAttr(p.image || placeholderImg())}" alt="" onerror="this.src='${placeholderImg()}'">
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
    state.filters = { category: 'all', inStockOnly: false, featuredOnly: false, sort: 'default', q: '' };
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
  $('#inventoryRefresh').addEventListener('click', () => { loadProducts(false); showToast('Inventory refreshed.'); });
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