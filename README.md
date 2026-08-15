# GST Ledger

A small web app for tracking purchase and sale invoices with GST — built to
replace a manually-maintained Google Sheet. Runs entirely in your browser,
no server or account needed.

## What it does

- Separate **Purchase** and **Sale** tabs, each with an entry form and a table
- **Vendors** and **Clients** master lists — save a party once, then just start
  typing their name on an invoice and their GSTIN autofills
- **Auto in-state / out-of-state detection** — type a city (yours is set to Indore,
  Madhya Pradesh by default) and the app looks up its state, compares it to your
  business's state, and classifies the invoice. A GSTIN also auto-fills the
  same way if you'd rather use that. City names that aren't in the app's
  built-in list simply leave the state blank for you to pick by hand — nothing
  gets silently misclassified.
- **Auto GST split** — computes CGST + SGST for in-state invoices, or IGST for
  out-of-state, from the taxable amount and rate you enter
- **GSTIN format validation** on every entry
- **Duplicate bill-number warning** per vendor/client
- **Summary tab** — total purchases/sales, input vs output GST, and net GST
  payable, filterable by month range
- **Excel export/import** — export the full ledger (or just one register) as
  an `.xlsx` file matching your original columns; import a previously
  exported file to restore or move data to another device

## How your data is stored

This is a static site — there's no database and no login. Every entry you
add is saved to your browser's local storage automatically, so it's still
there the next time you open the page **on the same browser, on the same
device**. It does **not** sync across devices on its own.

To back up your data or move it to another computer, use **Export full
backup (Excel)** on the Summary tab, and **Import backup** to load it back
in. Since browser storage can occasionally be cleared (private browsing,
clearing site data, reinstalling the browser), it's worth exporting a backup
every so often — e.g. at the end of each month, or before filing.

## Deploying to GitHub Pages

1. Create a new repository on GitHub (e.g. `gst-ledger`).
2. Add these three files to the repository root: `index.html`, `style.css`,
   `app.js`.
   - Easiest way: on the repo page, click **Add file → Upload files**, drag
     in the three files, and commit.
3. Go to the repo's **Settings → Pages**.
4. Under **Build and deployment**, set **Source** to `Deploy from a branch`,
   branch `main`, folder `/ (root)`. Save.
5. GitHub will give you a URL like
   `https://<your-username>.github.io/gst-ledger/` — that's your app. It can
   take a minute to go live after the first deploy.
6. Open it, click **Business settings** in the top right, and fill in your
   business name, GSTIN, and state — this is what the app compares every
   invoice against to decide in-state vs out-of-state.

Because it's a static site, anyone with the link can open it — but each
visitor only sees their *own* browser's local data, never yours. If you'd
rather keep the link private, you can leave the repository private and use
GitHub Pages' access controls (available on paid GitHub plans), or just not
share the URL.

## Notes on the GST logic

- Freight is added to the total **after** tax, as a separate line — matching
  your current sheet. If a vendor bills freight as part of a taxable
  composite supply, you may want to fold it into the taxable amount instead
  before entering it.
- GST rate defaults to 18% with a dropdown for the other standard slabs
  (0%, 0.25%, 3%, 5%, 12%, 28%) plus a custom option.
- State codes follow the official CBIC GST state code list (2-digit prefix
  of every GSTIN).

## Local development

No build step — just open `index.html` in a browser, or serve the folder
with any static file server (e.g. `python3 -m http.server`) if you want
GSTIN autocapitalize and datalist suggestions to behave exactly as they will
once deployed.
