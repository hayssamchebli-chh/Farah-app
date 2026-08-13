"""Harb Electric look and feel: palette, page chrome and the shared CSS.

Colours and type sampled from harbelectric.com.
"""

from __future__ import annotations

import streamlit as st

BRAND_BLUE = "#005AA7"
BRAND_BLUE_DARK = "#00447E"
BRAND_INK = "#16171E"
BRAND_GREY = "#626974"
BRAND_LINE = "#E3E7EC"
BRAND_SURFACE = "#FFFFFF"
BRAND_CANVAS = "#F4F6F9"

# Difference highlights: (fill, text on that fill). Solid, saturated fills —
# both pairs clear WCAG AA (5.0:1 and 5.4:1).
NEG = ("#B22222", "#FFFFFF")   # short — warehouse is receiving less than ordered
POS = ("#32cd32", "#FFFFFF")   # over  — warehouse is receiving more than ordered
# The same meaning as text on a white surface (stat cards, borders).
NEG_INK = "#B3261E"
POS_INK = "#0F6B36"

# Inline mark: ascending bars in a rounded square, matching the tab icon.
# Swap in the official Harb Electric logo file here to use the real asset.
LOGO_SVG = (
    '<svg width="38" height="38" viewBox="0 0 32 32" role="img" aria-label="Price History Pivot">'
    f'<rect width="32" height="32" rx="7" fill="{BRAND_BLUE}"/>'
    '<rect x="7.5" y="18" width="4.5" height="7.5" rx="1.2" fill="#fff" opacity=".72"/>'
    '<rect x="13.75" y="13" width="4.5" height="12.5" rx="1.2" fill="#fff" opacity=".86"/>'
    '<rect x="20" y="7" width="4.5" height="18.5" rx="1.2" fill="#fff"/></svg>'
)

CSS = f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Barlow:wght@400;500;600;700&family=Teko:wght@500;600&display=swap');

:root {{
  --brand: {BRAND_BLUE};
  --brand-dark: {BRAND_BLUE_DARK};
  --ink: {BRAND_INK};
  --grey: {BRAND_GREY};
  --line: {BRAND_LINE};
  --surface: {BRAND_SURFACE};
  --canvas: {BRAND_CANVAS};
}}

html, body, .stApp, [data-testid="stAppViewContainer"], .stMarkdown,
h1, h2, h3, h4, p, label, span, div, input, button, select, textarea {{
  font-family: 'Barlow', -apple-system, 'Segoe UI', Roboto, sans-serif;
}}
/* ...but never the icon spans: Streamlit renders Material ligatures, so an
   overridden font shows the raw word ("upload", "visibility") instead. */
[data-testid="stIconMaterial"], .material-icons, .material-icons-outlined,
span[class*="material-symbols"], [data-testid="stExpanderIcon"] {{
  font-family: 'Material Symbols Rounded' !important;
}}
[data-testid="stAppViewContainer"] {{ background: var(--canvas); }}
[data-testid="stHeader"] {{ background: transparent; }}
/* clear Streamlit's floating toolbar (Share / star / edit) so the masthead
   never sits underneath it */
.block-container {{ padding-top: 4.75rem; max-width: 1280px; }}
[data-testid="stToolbar"] {{ z-index: 100; }}

/* ---- top navigation (pills) ---- */
/* the injected <style> lands in its own block and would otherwise reserve a
   full row of vertical gap */
[data-testid="stElementContainer"]:has(style) {{ display: none !important; }}
.st-key-hbnav {{ margin: 0 !important; gap: .7rem; flex-wrap: wrap; }}
[data-testid="stPageLink"] a {{
  display: inline-flex; align-items: center; justify-content: center; gap: .6rem;
  padding: .8rem 2.4rem; min-height: 56px;
  border: 1px solid var(--line); border-radius: 999px;
  background: var(--surface); color: var(--grey) !important;
  font-weight: 600; font-size: 1.12rem; letter-spacing: .01em;
  text-decoration: none !important;
  transition: background 160ms ease, border-color 160ms ease, color 160ms ease;
}}
[data-testid="stPageLink"] a [data-testid="stIconMaterial"] {{
  font-size: 1.35rem !important; width: 1.35rem; height: 1.35rem;
}}
[data-testid="stPageLink"] a:hover {{
  border-color: var(--brand); color: var(--brand) !important;
  background: rgba(0,90,167,.05);
}}
[data-testid="stPageLink"] a:focus-visible {{
  outline: 3px solid rgba(0,90,167,.45); outline-offset: 2px;
}}
[data-testid="stPageLink"] a p {{ margin: 0; font-weight: 600; font-size: 1.12rem; }}
/* the current page: a solid brand pill */
[class*="st-key-hbnav_on"] [data-testid="stPageLink"] a,
[class*="st-key-hbnav_on"] [data-testid="stPageLink"] a:hover {{
  background: var(--brand); border-color: var(--brand);
  box-shadow: 0 2px 8px rgba(0,90,167,.22);
}}
[class*="st-key-hbnav_on"] [data-testid="stPageLink"] a p,
[class*="st-key-hbnav_on"] [data-testid="stPageLink"] a span {{ color: #fff !important; }}

/* ---- masthead ---- */
.hb-head {{
  display: flex; align-items: center; gap: .85rem;
  padding: 1rem 1.35rem; margin-bottom: 1.4rem;
  background: var(--surface); border: 1px solid var(--line);
  border-radius: 12px; border-top: 3px solid var(--brand);
  box-shadow: 0 1px 2px rgba(22,23,30,.05);
}}
.hb-head .hb-title {{
  font-family: 'Teko', 'Barlow', sans-serif; font-weight: 600;
  font-size: 1.9rem; line-height: 1; letter-spacing: .02em;
  color: var(--ink); margin: 0;
}}
.hb-head .hb-sub {{
  font-size: .82rem; color: var(--grey); margin: .15rem 0 0;
  letter-spacing: .04em; text-transform: uppercase;
}}
.hb-head .hb-spacer {{ flex: 1 1 auto; }}
.hb-badge {{
  font-size: .72rem; font-weight: 600; letter-spacing: .06em; text-transform: uppercase;
  color: var(--brand); background: rgba(0,90,167,.08);
  border: 1px solid rgba(0,90,167,.18); border-radius: 999px; padding: .3rem .7rem;
}}

/* ---- section labels ---- */
.hb-step {{
  display: flex; align-items: center; gap: .55rem;
  font-size: .78rem; font-weight: 700; letter-spacing: .1em; text-transform: uppercase;
  color: var(--grey); margin: .2rem 0 .6rem;
}}
.hb-step span.n {{
  display: inline-grid; place-items: center; width: 1.35rem; height: 1.35rem;
  border-radius: 50%; background: var(--brand); color: #fff; font-size: .72rem;
}}

/* ---- stat cards ---- */
.hb-stats {{ display: flex; flex-wrap: wrap; gap: .9rem; margin: .2rem 0 1.1rem; }}
.hb-stat {{
  flex: 1 1 170px; background: var(--surface); border: 1px solid var(--line);
  border-radius: 10px; padding: .85rem 1rem; border-left: 3px solid var(--brand);
}}
.hb-stat .v {{
  font-family: 'Teko','Barlow',sans-serif; font-size: 2rem; line-height: 1.05;
  color: var(--ink); font-weight: 600; font-variant-numeric: tabular-nums;
}}
.hb-stat .k {{
  font-size: .74rem; letter-spacing: .08em; text-transform: uppercase; color: var(--grey);
}}
.hb-stat.neg {{ border-left-color: {NEG_INK}; }}
.hb-stat.neg .v {{ color: {NEG_INK}; }}
.hb-stat.pos {{ border-left-color: {POS_INK}; }}
.hb-stat.pos .v {{ color: {POS_INK}; }}

/* ---- panels ---- */
[data-testid="stFileUploader"], [data-testid="stDataFrame"] {{
  background: var(--surface); border: 1px solid var(--line); border-radius: 10px;
}}
[data-testid="stFileUploader"] {{ padding: .5rem .75rem; }}
[data-testid="stFileUploader"] section {{ border-radius: 8px; }}
[data-testid="stFileUploaderDropzone"] {{ background: transparent; }}

/* ---- controls ---- */
[data-testid="stFormSubmitButton"] button,
[data-testid="stDownloadButton"] button,
[data-testid="stBaseButton-primary"] {{
  background: var(--brand) !important; color: #fff !important;
  border: 1px solid var(--brand) !important;
  border-radius: 8px; font-weight: 600; letter-spacing: .02em;
  padding: .5rem 1.15rem; min-height: 44px; cursor: pointer;
  transition: background 180ms ease, box-shadow 180ms ease;
}}
[data-testid="stFormSubmitButton"] button:hover,
[data-testid="stDownloadButton"] button:hover,
[data-testid="stBaseButton-primary"]:hover {{
  background: var(--brand-dark) !important; border-color: var(--brand-dark) !important;
  box-shadow: 0 2px 8px rgba(0,90,167,.25);
}}
[data-testid="stFormSubmitButton"] button:focus-visible,
[data-testid="stDownloadButton"] button:focus-visible,
[data-testid="stBaseButton-primary"]:focus-visible,
input:focus-visible {{
  outline: 3px solid rgba(0,90,167,.45); outline-offset: 2px;
}}
/* the uploader's own "Browse files" button stays secondary */
[data-testid="stFileUploader"] button {{
  border-radius: 8px; font-weight: 600; cursor: pointer; min-height: 44px;
  border-color: var(--brand); color: var(--brand);
}}
[data-testid="stFileUploader"] button:hover {{
  background: rgba(0,90,167,.06); border-color: var(--brand-dark); color: var(--brand-dark);
}}
[data-testid="stTextInput"] input {{ border-radius: 8px; }}
[data-testid="stTextInput"] input:focus {{ border-color: var(--brand); }}

/* ---- login card ---- */
.hb-login {{
  background: var(--surface); border: 1px solid var(--line); border-top: 3px solid var(--brand);
  border-radius: 14px; padding: 2rem 2rem 1.2rem;
  box-shadow: 0 10px 30px rgba(22,23,30,.07);
}}
.hb-login h2 {{
  font-family: 'Teko','Barlow',sans-serif; font-weight: 600; font-size: 1.75rem;
  color: var(--ink); margin: .9rem 0 .1rem; line-height: 1.1;
}}
.hb-login p {{ color: var(--grey); font-size: .9rem; margin: 0 0 .4rem; }}

/* ---- footer ---- */
.hb-foot {{
  margin-top: 2.2rem; padding-top: 1rem; border-top: 1px solid var(--line);
  color: var(--grey); font-size: .78rem; display: flex; gap: .5rem; flex-wrap: wrap;
}}
.hb-foot b {{ color: var(--ink); font-weight: 600; }}

@media (prefers-reduced-motion: reduce) {{
  * {{ transition: none !important; animation: none !important; }}
}}
@media (max-width: 640px) {{
  .hb-head {{ flex-wrap: wrap; }}
  .hb-head .hb-title {{ font-size: 1.6rem; }}
}}
</style>
"""


def inject_css() -> None:
    st.markdown(CSS, unsafe_allow_html=True)


def masthead(title: str, badge: str = "Internal tool") -> None:
    # NB: keep this HTML flush left — indented lines are parsed as a code block.
    inject_css()
    st.markdown(
        '<div class="hb-head">'
        + LOGO_SVG
        + f'<div><p class="hb-title">{title}</p>'
        '<p class="hb-sub">Harb Electric &middot; Tendering</p></div>'
        '<div class="hb-spacer"></div>'
        f'<span class="hb-badge">{badge}</span>'
        "</div>",
        unsafe_allow_html=True,
    )


def nav_bar(pages, active_title: str) -> None:
    """A row of pill links across the top of the page, current page filled in.

    Each link sits in a keyed container so Streamlit stamps a `st-key-…` class
    on it, which the CSS uses to fill in the active pill.
    """
    with st.container(horizontal=True, gap="small", key="hbnav"):
        for i, page in enumerate(pages):
            state = "on" if page.title == active_title else "off"
            with st.container(key=f"hbnav_{state}_{i}", width="content"):
                st.page_link(page, width="content")


def step(number: int, label: str) -> None:
    st.markdown(
        f'<div class="hb-step"><span class="n">{number}</span>{label}</div>',
        unsafe_allow_html=True,
    )


def stat_cards(cards) -> None:
    """Cards of (value, label) or (value, label, tone) where tone is neg/pos."""
    html = ['<div class="hb-stats">']
    for card in cards:
        value, label = card[0], card[1]
        tone = f" {card[2]}" if len(card) > 2 else ""
        html.append(
            f'<div class="hb-stat{tone}"><div class="v">{value:,}</div>'
            f'<div class="k">{label}</div></div>'
        )
    st.markdown("".join(html) + "</div>", unsafe_allow_html=True)


def footer() -> None:
    st.markdown(
        '<div class="hb-foot"><b>Harb Electric</b><span>&middot;</span>'
        "<span>Files are processed in-session and never stored.</span></div>",
        unsafe_allow_html=True,
    )
