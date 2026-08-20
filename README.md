# Procurement Toolkit

A password-protected Streamlit app for Harb Electric tendering. It turns flat transaction exports
into **one row per item number** — price history laid out newest → oldest, or a warehouse receipt
checked against what was ordered.

Three modes, switched from a pill bar across the top of the page (no sidebar):

| Page | Source document | What it does | Sample |
|---|---|---|---|
| **Selling Prices** | sales lines | price history, newest first (OC Net Price) | `sample_data.xlsx` |
| **Purchase Prices** | vendor purchase lines | price history, newest first (Unit Cost) | `sample_purchases.xlsx` |
| **Warehouse Receipt** | warehouse receipt lines | combines repeated items and checks received vs ordered | `sample_receipt.xlsx` |

## Input

Any `.xlsx` / `.xlsm` / `.xls` / `.csv`. Extra columns are ignored, and the header does not have to
be the first row — the app scans the first 20 rows for it.

**Selling Prices** expects:

| No. | Description | Document No. | Customer Name | Quantity | Pre-Discount Price | Discount % | Net Price | OC Net Price | Currency | Posting Date |
|---|---|---|---|---|---|---|---|---|---|---|

**Purchase Prices** expects:

| Document No | Vendor | Date | Item No | Description | Unit Cost | QTY. | Discount Amount | UOM | Discount | Amount | Amount Including VAT | Currency | Shipment No |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|

The item column may repeat — each repetition is one transaction for that item.

## Output

```
No. | R  Date  Qty  Price  Curr. | R  Date  Qty  Price  Curr. | ...
```

- **First column** — each item number once (`No.` / `Item No`), in the order it first appears.
- The `R, Date, Qty, Price, Curr.` block repeats as many times as the item with the **most**
  transactions requires; shorter items leave the trailing blocks blank.
- **R** — 1, 2, 3 … the occurrence number.
- **Date** — Posting Date (sales) or Date (purchases). Block 1 is the most recent transaction,
  then progressively older. Undated lines sort last; same-date lines keep their original file order.
- **Qty** — Quantity / QTY.
- **Price** — OC Net Price (sales; falls back to Net Price if there is no OC column), or Unit Cost
  (purchases; falls back to Amount if the file has no Unit Cost column). Unit Cost is per-unit, so
  prices stay comparable across lines bought in different quantities. The purchases page also
  offers **Amount** (the line total) and **Amount Including VAT** in a dropdown.
- **Curr.** — Currency.

Filled **R** cells are shaded brand blue (blank padding is left alone), on screen and in the
download. Change `BRAND_BLUE` in `theme.py` to use a different colour.

A Description column can be added next to the item number with the checkbox. Results are previewed
in the page and downloadable as a formatted `.xlsx`.

## Warehouse Receipt

A different transform: instead of spreading transactions sideways, it **collapses** them.

Expected columns:

| Source Document | Source No. | Item No. | Description | Bin Code | Quantity | Qty. to Receive | Over-Receipt Quantity |
|---|---|---|---|---|---|---|---|

Every line for the same `Item No.` is combined into one row, summing `Quantity` and
`Qty. to Receive`, and a **Difference** column is added:

```
Item No. | Description | Quantity | Qty. to Receive | Difference

Difference = Qty. to Receive − Quantity
```

- **Negative** — the warehouse is receiving *less* than the order expects; the cell is filled
  solid red (`#D32F2F`) with white text.
- **Positive** — the warehouse is receiving *more* than the order expects; solid green
  (`#157A40`) with white text.
- Zero is left unshaded.

The sign is written into the number itself (`-2`, `+20`), so the meaning does not depend on
colour alone. Highlighting applies to both the on-screen table and the Excel download; the export
also carries an auto-filter and a frozen header.

Two checkboxes: **Include Source No. and Bin Code** lists every source document and bin an item
appears in, and **Only show mismatches** hides items where received matches ordered.

The line count and over-receipt totals stay out of the table — the line count is reported in the
summary cards above it instead.

## Adding another source document

Both pages are the same code with a different `Profile` (header aliases + labels). To support a
third export, add a `Profile` in `core.py` and one `st.Page` entry in `streamlit_app.py` — the top
pill bar picks it up automatically.

## The password

The app asks for a password before showing anything. It is read from Streamlit secrets — it is
**not** stored in this repository.

**On Streamlit Community Cloud:** open the app → **Settings → Secrets** and paste:

```toml
app_password = "your-password-here"
```

**Locally:** copy `.streamlit/secrets.toml.example` to `.streamlit/secrets.toml` and set the same
key (that file is gitignored).

Change the password by editing that secret — no code change and no redeploy needed. Note this is a
single shared password, not per-user accounts; anyone with the link still needs it to get in.

## Deploying to Streamlit Community Cloud

1. Push this repo to GitHub.
2. Go to [share.streamlit.io](https://share.streamlit.io) → **Create app** → **Deploy a public app
   from GitHub**.
3. Repository `hayssamchebli-chh/Farah-app`, branch `main`, main file `streamlit_app.py`.
4. Before clicking Deploy, open **Advanced settings → Secrets** and add the `app_password` line
   above (or add it right after, under Settings → Secrets).
5. Deploy. Every push to `main` redeploys automatically.

## Running locally

```bash
pip install -r requirements.txt
```

```bash
streamlit run streamlit_app.py
```

## Offline version

`standalone/index.html` is the same tool as a single self-contained web page — open it directly in
a browser, no Python and no network needed. It covers the **selling prices** workflow only, and has
**no password gate**; it is meant for local use.
Its on-screen R cells are blue too, but its `.xlsx` download is unstyled — the free build of SheetJS
cannot write cell fills. Use the Streamlit app when you need the colour in the file.

## Branding

The interface follows Harb Electric's identity, sampled from
[harbelectric.com](https://harbelectric.com):

| Token | Value | Used for |
|---|---|---|
| Brand blue | `#005AA7` | buttons, accents, R cells, focus rings |
| Brand blue (dark) | `#00447E` | hover/pressed states |
| Ink | `#16171E` | headings and body text |
| Grey | `#626974` | secondary text, labels |
| Canvas / surface | `#F4F6F9` / `#FFFFFF` | page background / cards |
| Type | Teko (display) + Barlow (UI) | the site's own fonts |

The app name lives in `APP_TITLE` at the top of `theme.py` (and the browser tab title in
`streamlit_app.py`); the mode pill bar is centred above the masthead.

All foreground/background pairs meet WCAG AA (white on brand blue is 6.95:1).
The tokens live at the top of `streamlit_app.py` and in `.streamlit/config.toml`.

The header uses a generic bolt mark rather than the official logo file — replace `LOGO_SVG` in
`streamlit_app.py` with the real asset if you want the exact company logo.

## Files

- `streamlit_app.py` — password gate, page navigation, and the shared page layout
- `core.py` — the pivot engine: header detection, parsing, pivot, Excel export, per-source profiles
- `theme.py` — Harb Electric palette, CSS and page chrome
- `requirements.txt` — Python dependencies
- `.streamlit/config.toml` — brand theme (committed)
- `.streamlit/secrets.toml.example` — template for the password secret
- `sample_data.xlsx`, `sample_purchases.xlsx`, `sample_receipt.xlsx` — example inputs
- `standalone/` — offline browser-only version (SheetJS 0.18.5 vendored)
