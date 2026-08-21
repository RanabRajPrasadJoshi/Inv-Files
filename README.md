# Woolly & Co. — static shop backed by Google Sheets

A complete e-commerce frontend (HTML/CSS/JS, no framework) with Google Sheets
as the product database and Google Apps Script as the API layer. No traditional
backend, no database server.

```
Google Sheet  →  Google Apps Script (Code.gs)  →  Website (index.html/style.css/script.js)
```

## Files

| File | Purpose |
|---|---|
| `index.html` | Page shell — header, all views (home/shop/product/cart/about/contact), cart drawer, checkout modal |
| `style.css` | Full visual design |
| `script.js` | Loads products from Apps Script, client-side router, search/filters, cart (localStorage), checkout |
| `Code.gs` | Apps Script — serves products as JSON, validates + creates orders, emails the owner |

## 1. Create the Google Sheet

Create a new Google Sheet (any name). It needs two tabs, exactly named:

### `Products` sheet — columns (row 1 = headers)

| ProductID | ProductName | Category | Price | Quantity | ImageURL | Description | Featured | New |
|---|---|---|---|---|---|---|---|---|
| P001 | Felted Wool Tote | Bags | 2500 | 12 | https://... | Hand-felted merino wool tote... | TRUE | TRUE |
| P002 | Wool Trivet | Home | 800 | 0 | https://... | Thick felted trivet... | FALSE | FALSE |

- `Featured` / `New`: `TRUE`/`FALSE`, `Yes`/`No`, or `1`/`0` all work.
- `Quantity` is the live stock count — editing it here updates the site on the
  next refresh (page load or the 30-second auto-refresh).
- `ImageURL` should be a direct, publicly viewable image link (e.g. an image
  hosted on Imgur, a public Google Drive "anyone with the link" image
  converted to a direct URL, Unsplash, your own hosting, etc).

### `Orders` sheet — columns (row 1 = headers)

`Timestamp | OrderID | ProductID | ProductName | Quantity | Price | Total | CustomerName | Email | Phone | Address | Note | OrderStatus`

You don't need to type these headers by hand — see step 2's `setupSheets()`
helper, which creates both tabs with correct headers (and one sample product)
for you.

## 2. Add the Apps Script

1. In the Sheet, go to **Extensions → Apps Script**.
2. Delete any starter code and paste in the contents of `Code.gs`.
3. At the top of `Code.gs`, set:
   ```js
   const OWNER_EMAIL = 'your-real-email@example.com';
   ```
4. In the function dropdown at the top of the editor, select `setupSheets`
   and click **Run** once. Approve the permissions prompt. This creates the
   `Products` and `Orders` tabs if they don't already exist.
5. Replace the sample product / add your real products directly in the
   `Products` sheet.

## 3. Deploy the Apps Script as a Web App

1. In the Apps Script editor: **Deploy → New deployment**.
2. Click the gear icon next to "Select type" → choose **Web app**.
3. Settings:
   - **Execute as:** Me
   - **Who has access:** Anyone
4. Click **Deploy**, authorize the requested permissions, then copy the
   **Web app URL** (ends in `/exec`).

Every time you edit `Code.gs` after this, use **Deploy → Manage deployments →
edit (pencil) → New version → Deploy** so the live URL picks up your changes.

## 4. Point the website at your Apps Script

Open `script.js` and set:

```js
const CONFIG = {
  API_URL: 'https://script.google.com/macros/s/AKfycb.../exec', // your URL from step 3
  ...
};
```

That's the only code change required to connect the site to your own sheet.

## 5. Host the static site

Any static host works — no build step needed. Upload `index.html`,
`style.css`, and `script.js` (as siblings, same folder) to:

- **Netlify** — drag-and-drop the folder onto app.netlify.com/drop, or connect
  a Git repo containing these three files.
- **GitHub Pages** — push the three files to a repo, then enable Pages on the
  `main` branch in Settings → Pages.
- **Cloudflare Pages** — connect the repo, leave the build command empty,
  set the output directory to the repo root.

## How live updates work

- On page load, the site calls `?action=products` on your Apps Script URL and
  renders whatever is currently in the `Products` sheet.
- It repeats that call every 30 seconds in the background (`CONFIG.REFRESH_MS`)
  so price/stock/image edits in the Sheet show up without anyone reloading.
- Nothing about price or stock is ever stored authoritatively in the browser —
  the cart only remembers *which product IDs and quantities* were chosen;
  the price shown is always whatever the last product fetch returned.

## How an order is protected against overselling

1. Browser calls `POST ?action=createOrder` with `{ productId, quantity, customer }`.
2. `Code.gs` takes a `LockService` script lock so only one order can touch the
   sheet at a time.
3. It re-reads the **current** price and quantity directly from the `Products`
   sheet — the numbers the browser sent are never trusted for money math.
4. If there isn't enough stock, it returns `{ ok: false, error: "..." }` and
   nothing is written.
5. If there's enough stock, it reduces `Quantity` in place, appends a row to
   `Orders` with status `Pending`, emails `OWNER_EMAIL`, and returns
   `{ ok: true, orderId, remainingStock, ... }`.
6. The lock releases; the next queued order sees the just-updated quantity.

## Customizing

- **Colors/fonts**: edit the CSS custom properties at the top of `style.css`
  (`:root { --bg, --ink, --rust, --moss, ... }`).
- **Copy**: hero headline, about/contact text, footer links all live directly
  in `index.html`.
- **Low-stock threshold**: `CONFIG.LOW_STOCK_THRESHOLD` in `script.js`
  (default 5) controls when a card switches from "In stock" to "Only X left".
- **Currency**: `CONFIG.CURRENCY_PREFIX` in `script.js` (default `'Rs. '`).

## Notes & limits

- Apps Script web apps can't answer CORS preflight requests, so order POSTs
  are sent with `Content-Type: text/plain` (containing JSON) specifically to
  avoid triggering one. Don't change that header when editing `script.js`.
- Cart/checkout data lives in `localStorage`, per browser/device — there's no
  cross-device cart sync, by design (no accounts, no database beyond Sheets).
- Multi-item cart checkouts create one order row per product (looping the
  same `createOrder` call), so every row in `Orders` stays single-product,
  matching the columns requested in the sheet.
