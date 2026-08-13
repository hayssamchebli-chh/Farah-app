"""Pivot engine shared by every page of the app.

Reads a flat transaction export and returns one row per item, with each
transaction laid out newest -> oldest as repeated [R, Date, Qty, Price, Curr.]
blocks. Column names differ per source document, so each page passes a
`Profile` describing the header aliases to look for.
"""

from __future__ import annotations

import io
import re
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

BLOCK = ["R", "Date", "Qty", "Price", "Curr."]
REQUIRED = ("no", "qty", "price", "curr", "date")


@dataclass(frozen=True)
class Profile:
    """Everything that differs between one source document and another."""

    key: str
    title: str
    blurb: str
    # field -> accepted header names, matched after normalising
    fields: dict[str, list[str]]
    # field -> the name shown in error messages
    labels: dict[str, str]
    expected: str                                  # help text listing the columns
    file_stem: str
    # optional: let the user pick which column supplies Price
    price_choices: dict[str, list[str]] = field(default_factory=dict)
    price_fallback: list[str] = field(default_factory=list)

    def with_price(self, aliases: list[str]) -> Profile:
        """Copy of this profile with Price read from `aliases` instead."""
        return Profile(**{**self.__dict__, "fields": {**self.fields, "price": aliases}})


# --------------------------------------------------------------------------- cells


def norm(value) -> str:
    return re.sub(r"[^a-z0-9]", "", str("" if value is None else value).lower())


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
    if text.startswith("(") and text.endswith(")"):       # (1,234) = negative
        text = "-" + text[1:-1]
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


# -------------------------------------------------------------------------- header


def map_headers(header_row: list, profile: Profile) -> dict:
    """Map column positions to field names."""
    mapping: dict[str, int] = {}
    for i, cell in enumerate(header_row):
        n = norm(cell)
        if not n:
            continue
        for name, aliases in profile.fields.items():
            if name not in mapping and n in aliases:
                mapping[name] = i
    if "price" not in mapping and profile.price_fallback:
        for i, cell in enumerate(header_row):
            if norm(cell) in profile.price_fallback:
                mapping["price"] = i
                break
    return mapping


def find_header(rows: list[list], profile: Profile) -> tuple[int, dict] | None:
    """First row (within the first 20) giving both an item column and a date column."""
    for r in range(min(len(rows), 20)):
        mapping = map_headers(rows[r], profile)
        if "no" in mapping and "date" in mapping:
            return r, mapping
    return None


# --------------------------------------------------------------------------- pivot


def build_pivot(rows: list[list], profile: Profile, include_desc: bool) -> dict:
    found = find_header(rows, profile)
    if not found:
        raise ValueError(
            f'Could not find a header row containing both "{profile.labels["no"]}" '
            f'and "{profile.labels["date"]}".'
        )
    header_row, mapping = found

    missing = [f for f in REQUIRED if f not in mapping]
    if missing:
        raise ValueError("Missing column(s): " + ", ".join(profile.labels[f] for f in missing))

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
        entries.sort(
            key=lambda e: (e["date"] is not None, e["date"] or date.min, -e["seq"]),
            reverse=True,
        )
        max_dup = max(max_dup, len(entries))

    header = [profile.labels["no"]] + (["Description"] if include_desc else [])
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


# -------------------------------------------------------------------------- output


def to_excel(pivot: dict, accent: str, sheet_title: str = "Price History") -> bytes:
    accent = accent.lstrip("#").upper()
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_title[:31]

    ws.append(pivot["header"])
    for c in range(1, len(pivot["header"]) + 1):
        ws.cell(row=1, column=c).font = Font(bold=True)
        ws.cell(row=1, column=c).alignment = Alignment(horizontal="center")

    for row in pivot["body"]:
        ws.append(row)

    first_block = 2 if pivot["include_desc"] else 1  # count of leading columns
    widths = [18] + ([40] if pivot["include_desc"] else [])
    widths += [4, 12, 9, 12, 7] * pivot["max_dup"]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    fill = PatternFill("solid", start_color=accent, end_color=accent)
    white_bold = Font(bold=True, color="FFFFFF")

    for b in range(pivot["max_dup"]):
        r_col = first_block + b * 5 + 1     # 1-based R column of this block
        date_col = r_col + 1
        for r in range(2, len(pivot["body"]) + 2):
            ws.cell(row=r, column=date_col).number_format = "DD/MM/YYYY"
            r_cell = ws.cell(row=r, column=r_col)
            if r_cell.value is not None:    # only filled R cells, not the padding
                r_cell.fill = fill
                r_cell.font = white_bold
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


def style_frame(df: pd.DataFrame, pivot: dict, accent: str):
    """Shade the R cells that hold an occurrence number, matching the export."""
    first_block = 2 if pivot["include_desc"] else 1
    r_cols = [df.columns[first_block + b * 5] for b in range(pivot["max_dup"])]
    css = f"background-color: {accent}; color: white; font-weight: 600;"

    def paint(col: pd.Series):
        return [css if v != "" else "" for v in col]

    return df.style.apply(paint, subset=r_cols)


# ------------------------------------------------------------------------ profiles

SALES = Profile(
    key="sales",
    title="Selling Prices",
    blurb="Sales lines — the last price each item sold at, newest first.",
    fields={
        "no": ["no", "nos", "itemno", "item", "itemnumber", "partno", "partnumber"],
        "desc": ["description", "itemdescription", "desc"],
        "qty": ["quantity", "qty"],
        "price": ["ocnetprice", "netpriceoc", "ocnet", "priceoc"],
        "curr": ["currency", "curr", "currencycode"],
        "date": ["postingdate", "postingdt", "date", "documentdate"],
    },
    labels={
        "no": "No.",
        "qty": "Quantity",
        "price": "OC Net Price",
        "curr": "Currency",
        "date": "Posting Date",
    },
    expected=(
        "No., Description, Document No., Customer Name, Quantity, Pre-Discount Price, "
        "Discount %, Net Price, OC Net Price, Currency, Posting Date"
    ),
    file_stem="selling-price-history",
    price_fallback=["netprice"],
)

PURCHASES = Profile(
    key="purchases",
    title="Purchase Prices",
    blurb="Vendor purchase lines — what each item last cost, newest first.",
    fields={
        "no": ["itemno", "itemnumber", "item", "no", "partno", "partnumber"],
        "desc": ["description", "itemdescription", "desc"],
        "qty": ["qty", "quantity", "qtys"],
        "price": ["amount"],
        "curr": ["currency", "curr", "currencycode"],
        "date": ["date", "documentdate", "postingdate", "postingdt"],
    },
    labels={
        "no": "Item No",
        "qty": "QTY.",
        "price": "Amount",
        "curr": "Currency",
        "date": "Date",
    },
    expected=(
        "Document No, Vendor, Date, Item No, Description, Unit Cost, QTY., "
        "Discount Amount, UOM, Discount, Amount, Amount Including VAT, Currency, Shipment No"
    ),
    file_stem="purchase-price-history",
    price_choices={
        "Amount": ["amount"],
        "Unit Cost": ["unitcost", "unitprice", "cost"],
        "Amount Including VAT": ["amountincludingvat", "amountinclvat", "amountinclvat"],
    },
)
