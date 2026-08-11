# Price History Pivot

A password-protected Streamlit app that turns a flat sales-lines Excel export into a pivot with
**one row per item number** and every transaction for that item laid out newest → oldest.

## Input

Any `.xlsx` / `.xlsm` / `.xls` / `.csv` whose header row contains these columns (extra columns are
ignored, and the header does not have to be the first row):

| No. | Description | Document No. | Customer Name | Quantity | Pre-Discount Price | Discount % | Net Price | OC Net Price | Currency | Posting Date |
|---|---|---|---|---|---|---|---|---|---|---|

`No.` may repeat — each repetition is one transaction for that item.
`sample_data.xlsx` in this repo is a small example.

## Output

```
No. | R  Date  Qty  Price  Curr. | R  Date  Qty  Price  Curr. | ...
```

- **No.** — each item number once, in the order it first appears in the file.
- The `R, Date, Qty, Price, Curr.` block repeats as many times as the item with the **most**
  transactions requires; shorter items leave the trailing blocks blank.
- **R** — 1, 2, 3 … the occurrence number.
- **Date** — Posting Date. Block 1 is the most recent transaction, then progressively older.
  Undated lines sort last; same-date lines keep their original file order.
- **Qty** — Quantity.
- **Price** — OC Net Price (falls back to Net Price if there is no OC column).
- **Curr.** — Currency.

Filled **R** cells are shaded blue (blank padding is left alone), on screen and in the download.
Change `R_BLUE` at the top of `streamlit_app.py` to use a different colour.

A Description column can be added next to `No.` with the checkbox. Results are previewed in the
page and downloadable as a formatted `.xlsx`.

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
a browser, no Python and no network needed. It has **no password gate**; it is meant for local use.
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

All foreground/background pairs meet WCAG AA (white on brand blue is 6.95:1).
The tokens live at the top of `streamlit_app.py` and in `.streamlit/config.toml`.

The header uses a generic bolt mark rather than the official logo file — replace `LOGO_SVG` in
`streamlit_app.py` with the real asset if you want the exact company logo.

## Files

- `streamlit_app.py` — the app: parsing, pivot, Excel export, password gate
- `requirements.txt` — Python dependencies
- `.streamlit/config.toml` — brand theme (committed)
- `.streamlit/secrets.toml.example` — template for the password secret
- `sample_data.xlsx` — example input
- `standalone/` — offline browser-only version (SheetJS 0.18.5 vendored)
