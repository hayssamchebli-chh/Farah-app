"""Price History Pivot — Harb Electric.

Three modes over one engine:
  * Selling Prices    — sales lines, price history newest first (OC Net Price)
  * Purchase Prices   — vendor lines, price history newest first (Amount)
  * Warehouse Receipt — receipt lines collapsed per item, received vs ordered

Access is gated by a password stored in Streamlit secrets.
"""

from __future__ import annotations

import hmac
import importlib
import io
from datetime import date

import pandas as pd
import streamlit as st

import core
import theme

# Streamlit Cloud reloads this entrypoint after a deploy but can keep an
# already-imported module in sys.modules, so a new page here would call into an
# old core.py and blow up with AttributeError. Re-executing both modules costs
# microseconds and keeps them in step with the entrypoint.
for _module in (core, theme):
    try:
        importlib.reload(_module)
    except Exception:  # noqa: BLE001 — never let a reload hiccup take the app down
        pass

st.set_page_config(
    page_title="Procurement Toolkit · Harb Electric",
    page_icon="📊",
    layout="wide",
)


# --------------------------------------------------------------------------- auth


def check_password() -> bool:
    """Show a password prompt until the correct password is entered."""
    if st.session_state.get("authenticated"):
        return True

    try:
        expected = st.secrets.get("app_password")
    except Exception:  # no secrets.toml at all
        expected = None

    theme.inject_css()
    if not expected:
        st.error(
            "No app password is configured. Add `app_password = \"...\"` to the app's "
            "secrets (Streamlit Cloud → Settings → Secrets, or `.streamlit/secrets.toml` "
            "when running locally)."
        )
        return False

    _, mid, _ = st.columns([1, 1.1, 1])
    with mid:
        st.markdown(
            f'<div class="hb-login">{theme.LOGO_SVG}'
            f"<h2>{theme.APP_TITLE}</h2>"
            "<p>Harb Electric internal tool. Enter the access password to continue.</p>",
            unsafe_allow_html=True,
        )
        with st.form("login"):
            pwd = st.text_input(
                "Password", type="password", autocomplete="current-password",
                placeholder="Access password",
            )
            submitted = st.form_submit_button(
                "Sign in", use_container_width=True, type="primary"
            )
        st.markdown("</div>", unsafe_allow_html=True)

        if submitted:
            if hmac.compare_digest(pwd, str(expected)):
                st.session_state["authenticated"] = True
                st.rerun()
            else:
                st.error("Incorrect password. Please try again.")
    return False


# ---------------------------------------------------------------------------- page


def read_book(uploaded) -> dict[str, pd.DataFrame] | None:
    """Read every sheet as raw cells, or show the failure and return None."""
    data = uploaded.getvalue()
    try:
        if uploaded.name.lower().endswith(".csv"):
            return {"CSV": pd.read_csv(io.BytesIO(data), header=None, dtype=object)}
        return pd.read_excel(io.BytesIO(data), sheet_name=None, header=None, dtype=object)
    except Exception as exc:  # noqa: BLE001 — surface any reader failure to the user
        st.error(f"Could not read that file: {exc}")
        return None


def render(profile: core.Profile) -> None:
    """The whole tool for one source document type."""
    st.caption(profile.blurb)

    theme.step(1, "Upload the export")
    uploaded = st.file_uploader(
        "Excel or CSV file",
        type=["xlsx", "xlsm", "xls", "csv"],
        label_visibility="collapsed",
        key=f"upload_{profile.key}",
        help=f"Expected columns: {profile.expected}",
    )
    if not uploaded:
        st.caption(
            "Accepted: .xlsx, .xlsm, .xls, .csv — the header row does not have to be the "
            "first row. Files are processed in this session only and are never stored."
        )
        theme.footer()
        return

    book = read_book(uploaded)
    if book is None:
        theme.footer()
        return

    theme.step(2, "Choose the sheet")
    cols = st.columns([2, 1.2, 1.2] if profile.price_choices else [2, 1])
    with cols[0]:
        sheet = st.selectbox(
            "Sheet", list(book.keys()), key=f"sheet_{profile.key}",
            help="Pick the tab holding the transaction lines.",
        )
    if profile.price_choices:
        with cols[1]:
            choice = st.selectbox(
                "Price column", list(profile.price_choices.keys()),
                key=f"price_{profile.key}",
                help="Which column to report as Price.",
            )
            profile = profile.with_price(profile.price_choices[choice])
    with cols[-1]:
        include_desc = st.checkbox(
            "Include a Description column", key=f"desc_{profile.key}",
            help="Adds the item description beside the item number.",
        )

    frame = book[sheet]
    rows = frame.where(pd.notna(frame), None).values.tolist()

    try:
        pivot = core.build_pivot(rows, profile, include_desc)
    except ValueError as exc:
        st.error(str(exc))
        theme.footer()
        return

    theme.step(3, "Review and export")
    theme.stat_cards(
        [
            (pivot["items"], "Unique items"),
            (pivot["lines"], "Source lines"),
            (pivot["max_dup"], "Max transactions / item"),
        ]
    )
    if pivot["skipped"]:
        st.warning(
            f"{pivot['skipped']} row(s) were skipped because they had no "
            f"{profile.labels['no']} value. Check for subtotal or note rows in the source file."
        )

    st.dataframe(
        core.style_frame(core.to_frame(pivot), pivot, theme.BRAND_BLUE),
        use_container_width=True,
        hide_index=True,
    )
    st.caption(
        "Each block reads newest → oldest. Blue **R** cells number the transactions per item; "
        "blank blocks mean that item has fewer transactions."
    )

    st.download_button(
        "Download Excel",
        data=core.to_excel(pivot, theme.BRAND_BLUE, profile.title),
        file_name=f"{profile.file_stem}-{date.today():%Y-%m-%d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        key=f"dl_{profile.key}",
    )
    theme.footer()


def render_receipt(profile: core.Profile) -> None:
    """Warehouse receipt check: one row per item, warehouse qty vs ordered qty."""
    st.caption(profile.blurb)

    theme.step(1, "Upload the warehouse receipt")
    uploaded = st.file_uploader(
        "Excel or CSV file",
        type=["xlsx", "xlsm", "xls", "csv"],
        label_visibility="collapsed",
        key=f"upload_{profile.key}",
        help=f"Expected columns: {profile.expected}",
    )
    if not uploaded:
        st.caption(
            "Accepted: .xlsx, .xlsm, .xls, .csv — the header row does not have to be the "
            "first row. Files are processed in this session only and are never stored."
        )
        theme.footer()
        return

    book = read_book(uploaded)
    if book is None:
        theme.footer()
        return

    theme.step(2, "Choose the sheet")
    col1, col2, col3 = st.columns([2, 1.2, 1.2])
    with col1:
        sheet = st.selectbox(
            "Sheet", list(book.keys()), key=f"sheet_{profile.key}",
            help="Pick the tab holding the receipt lines.",
        )
    with col2:
        extras = st.checkbox(
            "Include Source No. and Bin Code", key=f"extras_{profile.key}",
            help="Lists every source document and bin the item appears in.",
        )
    with col3:
        only_diff = st.checkbox(
            "Only show mismatches", key=f"onlydiff_{profile.key}",
            help="Hides items where the warehouse quantity matches the order.",
        )

    frame = book[sheet]
    rows = frame.where(pd.notna(frame), None).values.tolist()

    try:
        receipt = core.build_receipt(rows, profile, extras)
    except ValueError as exc:
        st.error(str(exc))
        theme.footer()
        return

    theme.step(3, "Review and export")
    theme.stat_cards(
        [
            (receipt["items"], "Unique items"),
            (receipt["lines"], "Source lines"),
            (receipt["combined"], "Items combined"),
            (receipt["short"], "Short (negative)", "neg"),
            (receipt["over"], "Over (positive)", "pos"),
        ]
    )
    if receipt["skipped"]:
        st.warning(
            f"{receipt['skipped']} row(s) were skipped because they had no "
            f"{profile.labels['no']} value."
        )

    shown = dict(receipt)
    if only_diff:
        diff_col = receipt["diff_col"]
        shown["body"] = [r for r in receipt["body"] if r[diff_col] != 0]
        if not shown["body"]:
            st.success("Every item matches — no shortfalls and no over-receipts.")

    if shown["body"]:
        st.dataframe(
            core.style_receipt(core.receipt_frame(shown), shown, theme.NEG, theme.POS),
            use_container_width=True,
            hide_index=True,
        )
    st.caption(
        f"**Difference = {profile.labels['recv']} − {profile.labels['qty']}.** "
        "Negative (red) means the warehouse is receiving less than the order expects; "
        "positive (green) means more."
    )

    st.download_button(
        "Download Excel",
        data=core.receipt_to_excel(shown, theme.NEG, theme.POS),
        file_name=f"{profile.file_stem}-{date.today():%Y-%m-%d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        key=f"dl_{profile.key}",
    )
    theme.footer()


def selling_prices() -> None:
    render(core.SALES)


def purchase_prices() -> None:
    render(core.PURCHASES)


def warehouse_receipt() -> None:
    render_receipt(core.RECEIPT)


def app() -> None:
    pages = [
        st.Page(
            selling_prices, title="Selling Prices", icon=":material/sell:",
            url_path="selling_prices", default=True,
        ),
        st.Page(
            purchase_prices, title="Purchase Prices", icon=":material/local_shipping:",
            url_path="purchase_prices",
        ),
        st.Page(
            warehouse_receipt, title="Warehouse Receipt", icon=":material/inventory_2:",
            url_path="warehouse_receipt",
        ),
    ]
    # Streamlit's own nav is hidden: the pill bar below is the navigation.
    nav = st.navigation(pages, position="hidden")
    theme.app_header()                   # app title on top
    theme.nav_bar(pages, nav.title)      # then the mode pills
    theme.masthead(nav.title)            # then the current mode's title bar
    nav.run()


if __name__ == "__main__" and check_password():
    app()
