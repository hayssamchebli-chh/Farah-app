"""Price History Pivot — Streamlit app.

Upload a flat sales-lines Excel export and get one row per item No., with every
transaction laid out newest -> oldest as repeated [R, Date, Qty, Price, Curr.] blocks.

Access is gated by a password stored in Streamlit secrets (see README).
"""

from __future__ import annotations

import hmac
import io
import re
from datetime import date, datetime, timedelta

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

BLOCK = ["R", "Date", "Qty", "Price", "Curr."]

# Fill for the R cells that actually hold an occurrence number. Change here to
# restyle both the Excel export and the on-screen preview.
R_BLUE = "1F6FEB"

# Header aliases, matched after normalising (lowercase, alphanumerics only).
FIELDS = {
    "no": ["no", "nos", "itemno", "item", "itemnumber", "partno", "partnumber"],
    "desc": ["description", "itemdescription", "desc"],
    "qty": ["quantity", "qty"],
    "price": ["ocnetprice", "netpriceoc", "ocnet", "priceoc"],
    "curr": ["currency", "curr", "currencycode"],
    "date": ["postingdate", "postingdt", "date", "documentdate"],
}
LABELS = {
    "no": "No.",
    "qty": "Quantity",
    "price": "OC Net Price",
    "curr": "Currency",
    "date": "Posting Date",
}

st.set_page_config(page_title="Price History Pivot", page_icon="📊", layout="wide")


# --------------------------------------------------------------------------- auth


def check_password() -> bool:
    """Show a password prompt until the correct password is entered."""
    if st.session_state.get("authenticated"):
        return True

    try:
        expected = st.secrets.get("app_password")
    except Exception:  # no secrets.toml at all
        expected = None
    if not expected:
        st.error(
            "No app password is configured. Add `app_password = \"...\"` to the app's "
            "secrets (Streamlit Cloud → Settings → Secrets, or `.streamlit/secrets.toml` "
            "when running locally)."
        )
        return False

    st.title("Price History Pivot")
    st.caption("This tool is private. Enter the password to continue.")

    with st.form("login"):
        pwd = st.text_input("Password", type="password")
        submitted = st.form_submit_button("Enter")

    if submitted:
        if hmac.compare_digest(pwd, str(expected)):
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("Incorrect password.")
    return False


# ---------------------------------------------------------------------- parsing


def norm(value) -> str:
    return re.sub(r"[^a-z0-9]", "", str("" if value is None else value).lower())


def map_headers(header_row: list) -> dict:
    """Map column positions to our field names."""
    mapping: dict[str, int] = {}
    for i, cell in enumerate(header_row):
        n = norm(cell)
        if not n:
            continue
        for field, aliases in FIELDS.items():
            if field not in mapping and n in aliases:
                mapping[field] = i
    # Fall back to a plain "Net Price" column if "OC Net Price" is absent.
    if "price" not in mapping:
        for i, cell in enumerate(header_row):
            if norm(cell) == "netprice":
                mapping["price"] = i
                break
    return mapping


def find_header(rows: list[list]) -> tuple[int, dict] | None:
    """The first row (within the first 20) giving both a No. and a Posting Date column."""
    for r in range(min(len(rows), 20)):
        mapping = map_headers(rows[r])
        if "no" in mapping and "date" in mapping:
            return r, mapping
    return None


def to_date(value):
    """Excel serials, datetimes and day-first strings -> date. Otherwise None."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime().date()
    if isinstance(value, (int, float)):
        serial = float(value)
        if serial < 1 or serial > 2958465:
            return None
        # 1900 date system, including Excel's phantom 1900-02-29.
        days = int(serial) - (0 if serial < 61 else 1)
        return date(1899, 12, 31) + timedelta(days=days)

    text = str(value).strip()
    if not text:
        return None
    m = re.match(r"^(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{2,4})$", text)
    if m:
        d, mth, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if y < 100:
            y += 2000 if y < 70 else 1900
        try:
            return date(y, mth, d)
        except ValueError:
            return None
    parsed = pd.to_datetime(text, dayfirst=True, errors="coerce")
    return None if pd.isna(parsed) else parsed.date()


def to_num(value):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = re.sub(r"[\s,']", "", str(value).strip())
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def build_pivot(rows: list[list], include_desc: bool) -> dict:
    found = find_header(rows)
    if not found:
        raise ValueError('Could not find a header row containing both "No." and "Posting Date".')
    header_row, mapping = found

    missing = [f for f in ("no", "qty", "price", "curr", "date") if f not in mapping]
    if missing:
        raise ValueError("Missing column(s): " + ", ".join(LABELS[f] for f in missing))

    groups: dict[str, dict] = {}
    order: list[str] = []
    skipped = 0

    def cell(row, idx):
        return row[idx] if idx is not None and idx < len(row) else None

    for r in range(header_row + 1, len(rows)):
        row = rows[r]
        raw_no = cell(row, mapping["no"])
        no = "" if raw_no is None or pd.isna(raw_no) else str(raw_no).strip()
        if no.endswith(".0") and str(raw_no).replace(".", "", 1).isdigit():
            no = no[:-2]  # numeric item codes read back as floats
        if not no:
            if any(c is not None and not pd.isna(c) and str(c).strip() for c in row):
                skipped += 1
            continue

        key = no.upper()
        if key not in groups:
            groups[key] = {"no": no, "desc": "", "entries": []}
            order.append(key)
        group = groups[key]
        if include_desc and not group["desc"] and "desc" in mapping:
            d = cell(row, mapping["desc"])
            group["desc"] = "" if d is None or pd.isna(d) else str(d).strip()

        curr = cell(row, mapping["curr"])
        group["entries"].append(
            {
                "date": to_date(cell(row, mapping["date"])),
                "qty": to_num(cell(row, mapping["qty"])),
                "price": to_num(cell(row, mapping["price"])),
                "curr": "" if curr is None or pd.isna(curr) else str(curr).strip(),
                "seq": r,
            }
        )

    if not order:
        raise ValueError("No data rows found under the header.")

    # Newest first; undated lines sink to the bottom, ties keep original file order.
    max_dup = 0
    for key in order:
        entries = groups[key]["entries"]
        entries.sort(key=lambda e: (e["date"] is not None, e["date"] or date.min, -e["seq"]), reverse=True)
        max_dup = max(max_dup, len(entries))

    header = ["No."] + (["Description"] if include_desc else [])
    for _ in range(max_dup):
        header += BLOCK

    body = []
    for key in order:
        group = groups[key]
        out = [group["no"]] + ([group["desc"]] if include_desc else [])
        for i in range(max_dup):
            if i < len(group["entries"]):
                e = group["entries"][i]
                out += [i + 1, e["date"], e["qty"], e["price"], e["curr"]]
            else:
                out += [None] * 5
        body.append(out)

    return {
        "header": header,
        "body": body,
        "max_dup": max_dup,
        "items": len(order),
        "lines": sum(len(groups[k]["entries"]) for k in order),
        "skipped": skipped,
        "include_desc": include_desc,
    }


# ----------------------------------------------------------------------- export


def to_excel(pivot: dict) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Price History"

    ws.append(pivot["header"])
    for c in range(1, len(pivot["header"]) + 1):
        ws.cell(row=1, column=c).font = Font(bold=True)
        ws.cell(row=1, column=c).alignment = Alignment(horizontal="center")

    for row in pivot["body"]:
        ws.append(row)

    first_block = 2 if pivot["include_desc"] else 1  # 0-based count of leading columns
    widths = [18] + ([40] if pivot["include_desc"] else [])
    widths += [4, 12, 9, 12, 7] * pivot["max_dup"]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    blue_fill = PatternFill("solid", start_color=R_BLUE, end_color=R_BLUE)
    blue_font = Font(bold=True, color="FFFFFF")

    for b in range(pivot["max_dup"]):
        r_col = first_block + b * 5 + 1     # 1-based R column of this block
        date_col = r_col + 1
        for r in range(2, len(pivot["body"]) + 2):
            ws.cell(row=r, column=date_col).number_format = "DD/MM/YYYY"
            r_cell = ws.cell(row=r, column=r_col)
            if r_cell.value is not None:    # only filled R cells, not the padding
                r_cell.fill = blue_fill
                r_cell.font = blue_font
                r_cell.alignment = Alignment(horizontal="center")

    ws.freeze_panes = ws.cell(row=2, column=first_block + 1)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def to_frame(pivot: dict) -> pd.DataFrame:
    """Preview frame with de-duplicated column labels and dates as dd/mm/yyyy text."""
    labels = []
    counts: dict[str, int] = {}
    for name in pivot["header"]:
        counts[name] = counts.get(name, 0) + 1
        labels.append(name if counts[name] == 1 else f"{name}.{counts[name]}")

    rows = [
        [v.strftime("%d/%m/%Y") if isinstance(v, date) else ("" if v is None else v) for v in row]
        for row in pivot["body"]
    ]
    return pd.DataFrame(rows, columns=labels)


def style_frame(df: pd.DataFrame, pivot: dict):
    """Blue-fill the R cells that hold an occurrence number, matching the export."""
    first_block = 2 if pivot["include_desc"] else 1
    r_cols = [df.columns[first_block + b * 5] for b in range(pivot["max_dup"])]
    css = f"background-color: #{R_BLUE}; color: white; font-weight: 600;"

    def paint(col: pd.Series):
        return [css if v != "" else "" for v in col]

    return df.style.apply(paint, subset=r_cols)


# -------------------------------------------------------------------------- app


def main() -> None:
    st.title("📊 Price History Pivot")
    st.caption(
        "Upload a sales-lines Excel export and get one row per item, "
        "with every transaction laid out newest → oldest."
    )

    uploaded = st.file_uploader(
        "Excel or CSV file",
        type=["xlsx", "xlsm", "xls", "csv"],
        help="Expected columns: No., Description, Document No., Customer Name, Quantity, "
        "Pre-Discount Price, Discount %, Net Price, OC Net Price, Currency, Posting Date",
    )
    if not uploaded:
        st.info("Choose a file to begin. Nothing is stored — the file is processed per session.")
        return

    data = uploaded.getvalue()
    try:
        if uploaded.name.lower().endswith(".csv"):
            sheet_rows = {"CSV": pd.read_csv(io.BytesIO(data), header=None, dtype=object)}
        else:
            book = pd.read_excel(io.BytesIO(data), sheet_name=None, header=None, dtype=object)
            sheet_rows = book
    except Exception as exc:  # noqa: BLE001 — surface any reader failure to the user
        st.error(f"Could not read that file: {exc}")
        return

    col1, col2 = st.columns([2, 1])
    with col1:
        sheet = st.selectbox("Sheet", list(sheet_rows.keys()))
    with col2:
        include_desc = st.checkbox("Include a Description column")

    rows = sheet_rows[sheet].where(pd.notna(sheet_rows[sheet]), None).values.tolist()

    try:
        pivot = build_pivot(rows, include_desc)
    except ValueError as exc:
        st.error(str(exc))
        return

    st.success(
        f"{pivot['items']} unique item{'' if pivot['items'] == 1 else 's'} · "
        f"{pivot['lines']} lines · up to {pivot['max_dup']} "
        f"transaction{'' if pivot['max_dup'] == 1 else 's'} per item"
    )
    if pivot["skipped"]:
        st.warning(f"{pivot['skipped']} row(s) were skipped for having no No. value.")

    st.dataframe(style_frame(to_frame(pivot), pivot), use_container_width=True, hide_index=True)

    st.download_button(
        "⬇️ Download Excel",
        data=to_excel(pivot),
        file_name=f"price-history-pivot-{date.today():%Y-%m-%d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )


if __name__ == "__main__" and check_password():
    main()
