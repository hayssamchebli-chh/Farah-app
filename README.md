# Price History Pivot

A single-page browser tool that turns a flat sales-lines Excel export into a pivot with
**one row per item number** and every transaction for that item laid out newest → oldest.

No install, no server, no upload — the file is parsed entirely in your browser.

## Input

Any `.xlsx` / `.xls` / `.csv` whose header row contains these columns (extra columns are ignored,
and the header does not have to be the first row):

| No. | Description | Document No. | Customer Name | Quantity | Pre-Discount Price | Discount % | Net Price | OC Net Price | Currency | Posting Date |
|---|---|---|---|---|---|---|---|---|---|---|

`No.` may repeat — each repetition is one transaction for that item.

## Output

```
No. | R  Date  Qty  Price  Curr. | R  Date  Qty  Price  Curr. | ...
```

- **No.** — each item number once, in the order it first appears in the file.
- The `R, Date, Qty, Price, Curr.` block repeats as many times as the item with the **most**
  transactions requires; shorter items leave the trailing blocks blank.
- **R** — 1, 2, 3 … the occurrence number.
- **Date** — Posting Date. Block 1 is the most recent transaction, then progressively older.
  Lines with no date sort last; same-date lines keep their original file order.
- **Qty** — Quantity.
- **Price** — OC Net Price (falls back to Net Price if there is no OC column).
- **Curr.** — Currency.

A "Description" column can be added next to `No.` with the checkbox.

Preview the first 100 rows in the page, then **Download Excel** for the full result.

## Running it

Open `index.html` — that's it. Or serve the folder:

```bash
npx http-server . -p 8899
```

Hosted version (GitHub Pages): enable Pages on this repo with source **GitHub Actions**; the
included workflow publishes on every push to `main`.

## Files

- `index.html` — markup
- `styles.css` — styling (light + dark)
- `app.js` — parsing, pivot, and Excel export
- `vendor/xlsx.full.min.js` — SheetJS 0.18.5, vendored so the tool works offline
