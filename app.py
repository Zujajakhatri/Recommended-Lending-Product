"""
Recommended Lending Products for a Digital Retail Bank (DRB) — Streamlit dashboard
Personal loans · vehicle finance · home finance

Run with:  streamlit run app.py

Keep these files in the same folder as app.py:
    Personal_Loans_Estimation.xlsx   (data: sheets "Data", "Inflation rate ", "Year wise product data")
    FT-Logo NEW.png                  (logo)
    index.html                       (written dashboard insights)
    Consumer_Finance_Analysis.ipynb  (notebook, for the download button)
    optional: PDF report, for the download button
"""

import base64
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
import streamlit.components.v1 as components
from streamlit import config as _st_config

# ── Monochrome theme (also set in .streamlit/config.toml) ────────────────────
_THEME = {
    "theme.base": "light",
    "theme.primaryColor": "#000000",
    "theme.backgroundColor": "#FFFFFF",
    "theme.secondaryBackgroundColor": "#F3F3F3",
    "theme.textColor": "#1A1A1A",
}
_needs_rerun = False
for _k, _v in _THEME.items():
    try:
        if _st_config.get_option(_k) != _v:
            _st_config.set_option(_k, _v)
            _needs_rerun = True
    except Exception:
        pass

st.set_page_config(
    page_title="Recommended Lending Products · DRB",
    page_icon="▪",
    layout="wide",
    initial_sidebar_state="expanded",
)
if _needs_rerun:
    st.rerun()

# ════════════════════════════════════════════════════════════════════════════
# CONSTANTS
# ════════════════════════════════════════════════════════════════════════════
DATA_FILE = "Personal_Loans_Estimation.xlsx"
LOGO_FILE = "FT-Logo NEW.png"
REPORT_PDF = "DRB_lending_products_report.pdf"
NOTEBOOK_FILE = "Consumer_Finance_Analysis.ipynb"
PREPARED_BY = "Prepared by: Zujaja Khatri – Data and Research Analyst"
DISCLAIMER = ("<b style='color:#555;font-weight:600'>Disclaimer:</b> Figures are compiled from publicly available "
              "sources and include estimates; actual values may differ.")

ADV = "Advances (PKR Bn)"
ASSETS = "Total Assets (PKR Bn)"
DEP = "Total Deposits (PKR Bn)"
CF = "Consumer Finance (PKR Bn)"

PRODUCTS = ["Personal Loans", "Auto Finance", "Home Finance", "Credit Card Financing"]
CORE = ["Personal Loans", "Auto Finance", "Home Finance"]        # the three products in scope

SHORT_TYPE = {
    "Conventional + Islamic Banking": "Conventional + Islamic",
    "Full Islamic": "Full Islamic",
    "Foreign Commercial": "Foreign",
    "Microfinance Bank": "Microfinance",
    "Digital Bank": "Digital",
    "Islamic Digital Bank": "Islamic Digital",
}

FONT = "IBM Plex Sans, Helvetica Neue, Arial, sans-serif"
GREYS = ["#000000", "#3a3a3a", "#666666", "#8f8f8f", "#b8b8b8", "#dadada"]
HILITE, MUTED = "#000000", "#c4c4c4"

ST_VER = tuple(int(p) for p in st.__version__.split(".")[:2])
HAS_PILLS = hasattr(st, "pills")

# ════════════════════════════════════════════════════════════════════════════
# STYLING
# ════════════════════════════════════════════════════════════════════════════
st.markdown(
    """
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@300;400;500;600;700&display=swap');

:root{
  --ink:#000000; --ink-2:#2b2b2b; --mid:#6b6b6b; --soft:#9a9a9a;
  --line:#e2e2e2; --panel:#f5f5f5; --paper:#ffffff;
}
html, body, .stApp, [class*="css"], .stMarkdown, button, input, select, textarea{
  font-family:'IBM Plex Sans', 'Helvetica Neue', Arial, sans-serif !important;
}
.stApp{ background:var(--paper); color:var(--ink-2); }
#MainMenu, footer, [data-testid="stDecoration"]{ display:none !important; }
[data-testid="stHeader"]{ background:transparent; }
.block-container{ padding-top:1.4rem; padding-bottom:3rem; max-width:1500px; }
h1,h2,h3,h4,p,li,label,span,div{ color:inherit; }

/* Sidebar */
[data-testid="stSidebar"]{ background:var(--panel); border-right:1px solid var(--line); }
.sb-logo{ display:flex; justify-content:center; padding:0 0 12px 0; margin-top:-6px; }
.sb-logo img{ width:140px; max-width:100%; height:auto; display:block; }
[data-testid="stSidebarHeader"]{ height:2.25rem !important; min-height:0 !important; padding-top:0.5rem !important;
                                  padding-bottom:0 !important; }
[data-testid="stSidebarUserContent"]{ padding-top:0 !important; }
.sb-brand{ text-align:center; border-bottom:1px solid #d0d0d0; padding:0 0 14px 0; margin-bottom:6px; }
.sb-brand .t1{ font-size:17px; font-weight:700; color:#000; line-height:1.2; letter-spacing:-0.01em; }
.sb-brand .t2{ font-size:12px; color:var(--mid); margin-top:4px; }
.sb-h{ font-size:12px; font-weight:600; color:#000; margin:18px 0 2px 0; }
.sb-ref{ list-style:none; padding:0; margin:6px 0 0 0; }
.sb-ref li{ font-size:11.5px; line-height:1.45; color:#1a1a1a; margin:0 0 9px 0; padding-left:14px; text-indent:-14px;
            word-break:break-word; }
.sb-ref li i{ color:#000; }
.sb-sec{ font-size:11px; font-weight:600; letter-spacing:.08em; text-transform:uppercase; color:var(--mid);
         margin:22px 0 4px 0; padding-top:14px; border-top:1px solid #d9d9d9; }
[data-testid="stSidebar"] [data-testid="stExpander"] details{ border:1px solid #dcdcdc; border-radius:8px; background:#fff; }
[data-testid="stSidebar"] [data-testid="stExpander"] summary p{ font-size:13px; font-weight:600; }
[data-testid="stSidebar"] button[kind="secondary"]{ border-radius:8px; border-color:#cfcfcf; }

/* Navigation radio → list (only the page navigator, keyed "nav") */
.st-key-nav [role="radiogroup"]{ gap:2px; width:100%; }
.st-key-nav [role="radiogroup"] > div{ width:100%; }
.st-key-nav [role="radiogroup"] label{ padding:9px 12px !important; border-radius:6px; margin:0 !important; width:100%;
                   border:1px solid transparent; cursor:pointer; box-sizing:border-box; display:flex !important; }
.st-key-nav [role="radiogroup"] label:hover{ background:#e9e9e9; }
.st-key-nav [role="radiogroup"] label[data-selected="true"], .st-key-nav [role="radiogroup"] label:has(input:checked){ background:#000 !important; border-color:#000; }
.st-key-nav [role="radiogroup"] label[data-selected="true"] *, .st-key-nav [role="radiogroup"] label:has(input:checked) *{ color:#fff !important; }
.st-key-nav div:has(> div[data-testid="stMarkdownContainer"]) > div:first-child:not([data-testid]),
.st-key-nav [role="radiogroup"] > label > div:first-child{ display:none !important; }
.st-key-nav p{ font-size:14px !important; font-weight:500; }

/* Widget accents: neutralise Streamlit red */
span[data-baseweb="tag"]{ background:#1f1f1f !important; border-radius:4px !important; }
span[data-baseweb="tag"], span[data-baseweb="tag"] *{ color:#fff !important; }
div[data-baseweb="select"] > div{ border-color:#cfcfcf !important; border-radius:6px !important; background:#fff !important; }
div[data-baseweb="slider"] div[role="slider"]{ background:#000 !important; border-color:#000 !important; box-shadow:none !important; }
div[data-testid="stSliderThumbValue"], div[data-testid="stTickBarMin"], div[data-testid="stTickBarMax"]{ color:#000 !important; }
div[data-baseweb="slider"] > div > div > div{ background:#000; }
button[data-baseweb="tab"][aria-selected="true"]{ color:#000 !important; }
div[data-baseweb="tab-highlight"]{ background:#000 !important; }
div[data-baseweb="tab-border"]{ background:var(--line) !important; }
button[data-baseweb="tab"] p{ font-weight:500; }

/* Pills: selected = black, unselected = white (old + new Streamlit markup) */
button[kind="pills"], button[data-variant="pills"]{ background:#fff !important; border:1px solid #cfcfcf !important;
  border-radius:999px !important; }
button[kind="pills"] *, button[data-variant="pills"] *{ color:#1a1a1a !important; }
button[kind="pillsActive"], button[data-variant="pills"][aria-pressed="true"]{ background:#000 !important;
  border-color:#000 !important; }
button[kind="pillsActive"] *, button[data-variant="pills"][aria-pressed="true"] *{ color:#fff !important; }

/* Masthead */
.mast{ background:#000; color:#fff; border-radius:10px; padding:22px 28px 20px 28px;
       display:flex; justify-content:space-between; align-items:flex-end; gap:24px; flex-wrap:wrap; margin-bottom:22px; }
.mast .ttl{ font-size:28px; font-weight:700; letter-spacing:-0.02em; color:#fff; line-height:1.1; }
.mast .sub{ font-size:13.5px; color:#bdbdbd; margin-top:6px; }
.mast .meta{ text-align:right; font-size:13px; color:#fff; font-weight:500; letter-spacing:.01em;
             white-space:nowrap; align-self:center; }
.mast .meta .lbl{ color:#9a9a9a; font-weight:400; margin-right:4px; }
.mast .meta .sep{ color:#5c5c5c; margin:0 10px; }

/* Section header */
.sec{ margin:26px 0 12px 0; padding-bottom:8px; border-bottom:2px solid #000; display:flex;
      justify-content:space-between; align-items:baseline; gap:12px; flex-wrap:wrap; }
.sec .h{ font-size:19px; font-weight:700; color:#000; letter-spacing:-0.01em; }
.sec .s{ font-size:12.5px; color:var(--mid); }

/* KPI cards */
.kpi{ background:#fff; border:1px solid var(--line); border-top:3px solid #000; border-radius:4px 4px 8px 8px;
      padding:14px 16px 13px 16px; height:100%; min-height:122px; }
.kpi .l{ font-size:12.5px; color:var(--mid); font-weight:500; line-height:1.3; min-height:16px; }
.kpi .v{ font-size:27px; font-weight:600; color:#000; letter-spacing:-0.02em; margin-top:6px;
         font-variant-numeric:tabular-nums; line-height:1.15; word-break:break-word; }
.kpi .v.sm{ font-size:19px; letter-spacing:-0.01em; }
.kpi .d{ font-size:12.5px; margin-top:7px; font-variant-numeric:tabular-nums; color:var(--mid); }
.kpi .d .up{ color:#000; font-weight:600; }
.kpi .d .dn{ color:#000; font-weight:600; background:#e9e9e9; padding:1px 5px; border-radius:3px; }
.kpi.dark{ background:#000; border-color:#000; }
.kpi.dark .l, .kpi.dark .d{ color:#bdbdbd; }
.kpi.dark .v, .kpi.dark .d .up{ color:#fff; }

/* Chart cards */
.ct{ font-size:15px; font-weight:600; color:#000; letter-spacing:-0.005em; }
.cs{ font-size:12.5px; color:var(--mid); margin-top:2px; margin-bottom:2px; }
div[data-testid="stVerticalBlockBorderWrapper"]{ border-color:var(--line) !important; border-radius:8px !important; background:#fff; }

/* Insights */
.ins{ list-style:none; padding:0; margin:0; }
.ins li{ border-left:3px solid #000; background:#f7f7f7; padding:11px 14px; margin:0 0 8px 0;
         font-size:14px; line-height:1.55; color:#1a1a1a; border-radius:0 6px 6px 0; }
.ins li b{ color:#000; }

.note{ font-size:12.5px; color:var(--mid); background:#f7f7f7; border:1px solid var(--line);
       border-radius:6px; padding:9px 12px; margin:4px 0 10px 0; line-height:1.5; }
.note b{ color:#000; }

/* Roadmap cards */
.phase{ border:1px solid var(--line); border-radius:8px; padding:14px 16px; background:#fff; height:100%;
        min-height:205px; box-sizing:border-box; }
.phase.dark{ background:#000; border-color:#000; }
.phase .n{ display:inline-block; background:#000; color:#fff; font-size:12px; font-weight:600;
           border-radius:4px; padding:1px 7px; margin-bottom:6px; }
.phase.dark .n{ background:#fff; color:#000; }
.phase .w{ font-size:12px; color:var(--mid); }
.phase h4{ margin:2px 0 4px; font-size:16px; font-weight:600; color:#000; }
.phase p{ margin:0; font-size:13.5px; color:#333; line-height:1.5; }
.phase.dark h4, .phase.dark p{ color:#fff; }
.phase.dark .w{ color:#bdbdbd; }

[data-testid="stDataFrame"]{ border:1px solid var(--line); border-radius:6px; }
@media (max-width: 900px){ .mast .meta{ text-align:left; white-space:normal; } .kpi .v{ font-size:22px; } }
</style>
""",
    unsafe_allow_html=True,
)


# ════════════════════════════════════════════════════════════════════════════
# DATA LOADING
# ════════════════════════════════════════════════════════════════════════════
def _find(name):
    here = Path(__file__).resolve().parent
    for p in (here / name, here / "data" / name, Path.cwd() / name, Path.cwd() / "data" / name):
        if p.exists():
            return p
    return None


def _find_any(suffix, contains=""):
    """First file in the app folder with this extension (and text in its name)."""
    here = Path(__file__).resolve().parent
    hits = sorted(p for p in here.glob(f"*{suffix}") if contains.lower() in p.name.lower())
    return hits[0] if hits else None


def logo_data_uri():
    path = _find(LOGO_FILE) or _find_any(".png", "logo")
    if path is None:
        return None
    return f"data:image/png;base64,{base64.b64encode(path.read_bytes()).decode()}"


@st.cache_data(show_spinner=False)
def load_data(path_str):
    bank = pd.read_excel(path_str, sheet_name="Data")
    infl = pd.read_excel(path_str, sheet_name="Inflation rate ")
    prod = pd.read_excel(path_str, sheet_name="Year wise product data")

    # clean column names (fix spelling in the source file)
    bank.columns = bank.columns.str.strip()
    bank = bank.rename(columns={"Cunsumer Finance (PKR bn)": CF})
    bank["Bank"] = bank["Bank"].str.strip()
    infl.columns = infl.columns.str.strip()
    prod = prod.rename(columns={"Personal_Loans": "Personal Loans", "Home_Finance": "Home Finance",
                                "Auto_Financ": "Auto Finance", "Auto_Finance": "Auto Finance",
                                "Credit_Card_Financing": "Credit Card Financing"})

    # consumer finance growth per bank (0 → x gives inf, show as n/a)
    bank = bank.sort_values(["Bank", "Year"]).reset_index(drop=True)
    bank["CF Growth (%)"] = (bank.groupby("Bank")[CF].pct_change() * 100).replace([np.inf, -np.inf], np.nan)
    return bank, infl.set_index("Year")["Avg_Inflation_Rate"] * 100, prod.set_index("Year")[PRODUCTS]


def cagr(start, end, years):
    if years <= 0 or pd.isna(start) or pd.isna(end) or start <= 0:
        return np.nan
    return ((end / start) ** (1 / years) - 1) * 100


# ════════════════════════════════════════════════════════════════════════════
# FORMATTING & UI HELPERS
# ════════════════════════════════════════════════════════════════════════════
def f_bn(v, dec=1):
    if v is None or pd.isna(v):
        return "—"
    return f"PKR {v / 1000:,.2f} Tn" if abs(v) >= 1000 else f"PKR {v:,.{dec}f} Bn"


def f_pct(v, dec=1, sign=False):
    if v is None or pd.isna(v):
        return "—"
    v = v + (1e-9 if v >= 0 else -1e-9)       # so 3.55 rounds to 3.6, not 3.5
    return f"{v:+,.{dec}f}%" if sign else f"{v:,.{dec}f}%"


def delta(v, ref="", unit="%"):
    if v is None or pd.isna(v):
        return "<span>No prior-year comparison</span>"
    cls, arrow = ("up", "▲") if v >= 0 else ("dn", "▼")
    return f"<span class='{cls}'>{arrow} {abs(v):,.1f}{unit}</span> {ref}"


def kpi(label, value, foot="", small=False, dark=False):
    st.markdown(
        f"<div class='kpi{' dark' if dark else ''}'><div class='l'>{label}</div>"
        f"<div class='v{' sm' if small else ''}'>{value}</div><div class='d'>{foot}</div></div>",
        unsafe_allow_html=True,
    )


def section(title, sub=""):
    st.markdown(f"<div class='sec'><div class='h'>{title}</div><div class='s'>{sub}</div></div>",
                unsafe_allow_html=True)


def note(html):
    st.markdown(f"<div class='note'>{html}</div>", unsafe_allow_html=True)


def insights(items):
    st.markdown("<ul class='ins'>" + "".join(f"<li>{i}</li>" for i in items) + "</ul>", unsafe_allow_html=True)


def show_fig(fig, title, sub=""):
    with st.container(border=True):
        st.markdown(f"<div class='ct'>{title}</div>" + (f"<div class='cs'>{sub}</div>" if sub else ""),
                    unsafe_allow_html=True)
        cfg = {"displayModeBar": False, "responsive": True}
        if ST_VER >= (1, 46):
            st.plotly_chart(fig, width="stretch", config=cfg)
        else:
            st.plotly_chart(fig, use_container_width=True, config=cfg)


def show_table(df, column_config=None, height=None, hide_index=True):
    kw = dict(hide_index=hide_index, column_config=column_config or {})
    if height:
        kw["height"] = height
    if ST_VER >= (1, 46):
        st.dataframe(df, width="stretch", **kw)
    else:
        st.dataframe(df, use_container_width=True, **kw)


def style(fig, height=360, legend=True, ysuffix="", xsuffix="", ytitle=None, xtitle=None):
    fig.update_layout(
        template="none",
        height=height,
        font=dict(family=FONT, size=12, color="#1f1f1f"),
        paper_bgcolor="white",
        plot_bgcolor="white",
        margin=dict(l=8, r=16, t=36 if legend else 12, b=8),
        showlegend=legend,
        legend=dict(orientation="h", yanchor="bottom", y=1.0, x=0, xanchor="left",
                    title=None, font=dict(size=11.5), bgcolor="rgba(0,0,0,0)"),
        hoverlabel=dict(bgcolor="white", bordercolor="#000", font=dict(family=FONT, size=12, color="#000")),
        bargap=0.28,
    )
    fig.update_xaxes(showgrid=False, linecolor="#9e9e9e", tickcolor="#9e9e9e", ticks="outside",
                     ticklen=4, ticksuffix=xsuffix, title=xtitle, title_font=dict(size=12, color="#555"),
                     automargin=True)
    fig.update_yaxes(showgrid=True, gridcolor="#ececec", zeroline=False, linecolor="#9e9e9e",
                     ticksuffix=ysuffix, title=ytitle, title_font=dict(size=12, color="#555"), automargin=True)
    return fig


def hbar(df, x, y, height=None, fmt=",.1f", suffix="", highlight=None):
    """Horizontal bar, largest on top. Bars in `highlight` are black, the rest grey."""
    df = df.iloc[::-1]
    colors = [HILITE if (highlight is None or n in highlight) else MUTED for n in df[y]]
    fig = go.Figure(go.Bar(
        x=df[x], y=df[y], orientation="h", marker_color=colors,
        text=[f"{v:{fmt}}{suffix}" if pd.notna(v) else "" for v in df[x]],
        textposition="outside", cliponaxis=False, textfont=dict(size=11.5, color="#1f1f1f"),
        hovertemplate=f"<b>%{{y}}</b><br>%{{x:{fmt}}}{suffix}<extra></extra>",
    ))
    style(fig, height=height or max(260, 30 * len(df) + 60), legend=False, xsuffix=suffix)
    xmax = pd.to_numeric(df[x], errors="coerce").max()
    xmin = min(0, pd.to_numeric(df[x], errors="coerce").min())
    rng = [xmin * 1.25, xmax * 1.22] if pd.notna(xmax) and xmax > 0 else None
    fig.update_xaxes(showgrid=True, gridcolor="#f0f0f0", showticklabels=False, ticks="", range=rng)
    fig.update_yaxes(showgrid=False, ticks="", linecolor="#ffffff")
    return fig


# ════════════════════════════════════════════════════════════════════════════
# LOAD
# ════════════════════════════════════════════════════════════════════════════
data_path = _find(DATA_FILE) or _find_any(".xlsx", "personal_loans")
if data_path is None:
    st.error(f"Data file not found. Place {DATA_FILE} in the same folder as app.py.")
    st.stop()

BANK_ALL, INFL, PROD = load_data(str(data_path))
YEARS = sorted(BANK_ALL["Year"].unique().tolist())
ALL_TYPES = sorted(BANK_ALL["Bank Type"].unique().tolist())


# ════════════════════════════════════════════════════════════════════════════
# SIDEBAR — navigation + filters
# ════════════════════════════════════════════════════════════════════════════
PAGES = ["Executive overview", "Products & inflation", "Banks & competition", "Product fit & recommendation"]


def reset_filters():
    for k in [k for k in st.session_state if str(k).startswith("f_")]:
        del st.session_state[k]


with st.sidebar:
    LOGO_URI = logo_data_uri()
    if LOGO_URI:
        st.markdown(f"<div class='sb-logo'><img src='{LOGO_URI}' alt='Future Technologies'></div>",
                    unsafe_allow_html=True)
    st.markdown("<div class='sb-brand'><div class='t1'>Recommended Lending<br>Products</div>"
                "<div class='t2'>Personal · vehicle · home finance for a Digital Retail Bank</div></div>",
                unsafe_allow_html=True)
    page = st.radio("Dashboard", PAGES, label_visibility="collapsed", key="nav")

    st.markdown("<div class='sb-sec'>Filters</div>", unsafe_allow_html=True)

    st.markdown("<div class='sb-h'>Years</div>", unsafe_allow_html=True)
    if HAS_PILLS:
        sel_years = st.pills("Years", YEARS, selection_mode="multi", default=YEARS, key="f_years",
                             label_visibility="collapsed")
    else:
        sel_years = st.multiselect("Years", YEARS, default=YEARS, key="f_years", label_visibility="collapsed")
    SEL_YEARS = sorted(sel_years) if sel_years else [YEARS[-1]]
    year = SEL_YEARS[-1]                    # snapshot year (latest selected)
    p0, p1 = SEL_YEARS[0], SEL_YEARS[-1]    # growth period (first → last selected)
    st.caption(f"Snapshot: {year} · Trend: {p0}–{p1}" if p1 > p0 else f"Snapshot: {year}")

    with st.expander("Bank types", expanded=False):
        if HAS_PILLS:
            sel_types = st.pills("Bank types", ALL_TYPES, selection_mode="multi", default=ALL_TYPES,
                                 key="f_types", format_func=lambda v: SHORT_TYPE.get(v, v),
                                 label_visibility="collapsed")
        else:
            sel_types = st.multiselect("Bank types", ALL_TYPES, default=ALL_TYPES, key="f_types",
                                       format_func=lambda v: SHORT_TYPE.get(v, v), label_visibility="collapsed")
        sel_types = list(sel_types or [])
        st.caption(f"{len(sel_types)} of {len(ALL_TYPES)} selected · applies to bank pages")

    st.button("Reset filters", key="reset_btn", on_click=reset_filters,
              **({"width": "stretch"} if ST_VER >= (1, 46) else {"use_container_width": True}))

    # References — edit to match your sources
    REFERENCES = [
        ("State Bank of Pakistan", "2020–2025", "Banking statistics and consumer finance data", "SBP"),
        ("Pakistan Bureau of Statistics", "2020–2025", "Consumer Price Index (CPI) inflation", "PBS"),
        ("Individual banks", "2020–2025", "Annual reports and financial statements", "Bank websites"),
        ("State Bank of Pakistan", "2022", "Licensing and regulatory framework for digital banks", "SBP"),
    ]
    st.markdown("<div class='sb-sec'>References</div>", unsafe_allow_html=True)
    st.markdown(
        "<ul class='sb-ref'>"
        + "".join(f"<li>{a}. ({y}). <i>{t}</i>. {p}.</li>" for a, y, t, p in REFERENCES)
        + "</ul>",
        unsafe_allow_html=True,
    )

if not sel_types:
    st.warning("Select at least one bank type in the sidebar to populate the dashboard.")
    st.stop()

BANK = BANK_ALL[BANK_ALL["Bank Type"].isin(sel_types)].copy()
FILTERED = set(sel_types) != set(ALL_TYPES)
BY = BANK[BANK["Year"] == year]
N = p1 - p0

# Price index (start year = 1) and products in start-year prices
PRICE_INDEX = (1 + INFL / 100).cumprod()
PRICE_INDEX = PRICE_INDEX / PRICE_INDEX.loc[p0]
REAL = PROD.div(PRICE_INDEX, axis=0)
YOY = PROD.pct_change() * 100

PROD_CAGR = pd.Series({p: cagr(PROD.loc[p0, p], PROD.loc[p1, p], N) for p in PRODUCTS})
REAL_CAGR = pd.Series({p: cagr(REAL.loc[p0, p], REAL.loc[p1, p], N) for p in PRODUCTS})

# ════════════════════════════════════════════════════════════════════════════
# MASTHEAD
# ════════════════════════════════════════════════════════════════════════════
st.markdown(
    f"<div class='mast'><div><div class='ttl'>{page}</div>"
    f"<div class='sub'>Recommended lending products for a Digital Retail Bank · personal, vehicle and home finance</div></div>"
    f"<div class='meta'><span class='lbl'>Coverage:</span> {BY['Bank'].nunique()} Banks"
    f"<span class='sep'>|</span>{len(sel_types)} Bank {'Type' if len(sel_types) == 1 else 'Types'}"
    f"<span class='sep'>|</span>"
    f"{f'{p0}–{p1} Trend Analysis' if p1 > p0 else f'{year} Snapshot'}</div></div>",
    unsafe_allow_html=True,
)


# ════════════════════════════════════════════════════════════════════════════
# SHARED CHART BUILDERS
# ════════════════════════════════════════════════════════════════════════════
def product_trend_fig(height=360):
    d = PROD.loc[SEL_YEARS]
    fig = go.Figure()
    for i, p in enumerate(PRODUCTS):
        fig.add_bar(x=d.index, y=d[p], name=p, marker_color=GREYS[i],
                    hovertemplate=f"<b>{p}</b><br>%{{x}}: PKR %{{y:,.1f}} Bn<extra></extra>")
    style(fig, height=height, ytitle="PKR Bn")
    fig.update_layout(barmode="stack")
    fig.update_xaxes(dtick=1)
    return fig


def product_mix_fig(height=360):
    a, b = PROD.loc[p0], PROD.loc[p1]
    sa, sb = a / a.sum() * 100, b / b.sum() * 100
    fig = go.Figure()
    fig.add_bar(y=PRODUCTS, x=sa[PRODUCTS], name=str(p0), orientation="h", marker_color=MUTED,
                text=[f"{v:.1f}%" for v in sa[PRODUCTS]], textposition="outside", cliponaxis=False)
    if p1 > p0:
        fig.add_bar(y=PRODUCTS, x=sb[PRODUCTS], name=str(p1), orientation="h", marker_color=HILITE,
                    text=[f"{v:.1f}%" for v in sb[PRODUCTS]], textposition="outside", cliponaxis=False)
    style(fig, height=height, xsuffix="%")
    fig.update_yaxes(autorange="reversed", showgrid=False, ticks="")
    fig.update_xaxes(showgrid=True, gridcolor="#f0f0f0", range=[0, max(sa.max(), sb.max()) * 1.25])
    return fig


def product_lines_fig(height=400):
    """One line per product, nominal PKR Bn."""
    d = PROD.loc[SEL_YEARS]
    shades = ["#000000", "#555555", "#8f8f8f", "#b8b8b8"]
    fig = go.Figure()
    for p, c in zip(PRODUCTS, shades):
        fig.add_scatter(x=d.index, y=d[p], name=p, mode="lines+markers",
                        line=dict(color=c, width=2.5), marker=dict(size=8, color=c),
                        hovertemplate=f"<b>{p}</b><br>%{{x}}: PKR %{{y:,.1f}} Bn<extra></extra>")
    style(fig, height=height, ytitle="PKR Bn", xtitle="Year")
    fig.update_layout(legend=dict(orientation="h", yanchor="top", y=-0.18, x=0.5, xanchor="center"),
                      margin=dict(l=8, r=16, t=12, b=8))
    fig.update_xaxes(dtick=1)
    return fig


def inflation_fig(height=320):
    d = INFL.loc[SEL_YEARS]
    peak = d.idxmax()
    fig = go.Figure(go.Bar(
        x=d.index, y=d.values, marker_color=[HILITE if y == peak else MUTED for y in d.index],
        text=[f_pct(v) for v in d.values], textposition="outside", cliponaxis=False,
        hovertemplate="<b>%{x}</b><br>Inflation: %{y:.2f}%<extra></extra>",
    ))
    style(fig, height=height, legend=False, ysuffix="%")
    fig.update_xaxes(dtick=1)
    fig.update_yaxes(range=[0, d.max() * 1.2])
    return fig


def nominal_real_fig(product, height=340):
    fig = go.Figure()
    fig.add_scatter(x=SEL_YEARS, y=PROD.loc[SEL_YEARS, product], name="Nominal", mode="lines+markers",
                    line=dict(color="#000", width=2.5), marker=dict(size=8),
                    hovertemplate="<b>%{x}</b><br>Nominal: PKR %{y:,.1f} Bn<extra></extra>")
    fig.add_scatter(x=SEL_YEARS, y=REAL.loc[SEL_YEARS, product], name=f"Real ({p0} prices)",
                    mode="lines+markers", line=dict(color="#8f8f8f", width=2.5, dash="dash"), marker=dict(size=8),
                    hovertemplate="<b>%{x}</b><br>Real: PKR %{y:,.1f} Bn<extra></extra>")
    style(fig, height=height, ytitle="PKR Bn")
    fig.update_xaxes(dtick=1)
    return fig


def type_summary(b):
    """Bank type totals for the snapshot year: CF, CF as % of advances, market share."""
    t = b[b["Year"] == year].groupby("Bank Type")[[ADV, DEP, CF]].sum()
    t["CF % of advances"] = t[CF] / t[ADV].replace(0, np.nan) * 100
    t["Advances / deposits (%)"] = t[ADV] / t[DEP].replace(0, np.nan) * 100
    t["Market share (%)"] = t[CF] / t[CF].sum() * 100
    t0 = b[b["Year"] == p0].groupby("Bank Type")[CF].sum()
    t["CF CAGR (%)"] = [cagr(t0.get(i, np.nan), t.loc[i, CF], N) for i in t.index]
    t["Banks"] = b[b["Year"] == year].groupby("Bank Type")["Bank"].nunique()
    return t.reset_index()


# ════════════════════════════════════════════════════════════════════════════
# DOWNLOAD BUTTONS + INSIGHTS
# ════════════════════════════════════════════════════════════════════════════
def notebook_download_button():
    nb_path = _find(NOTEBOOK_FILE) or _find_any(".ipynb")
    if nb_path is None:
        st.button("Download data processing file (.ipynb)", key="nb_btn", type="primary", disabled=True,
                  help=f"Place {NOTEBOOK_FILE} in the same folder as app.py")
        return
    st.download_button("Download data processing file (.ipynb)", data=nb_path.read_bytes(),
                       file_name=NOTEBOOK_FILE, mime="application/x-ipynb+json", key="nb_btn", type="primary")


def pdf_download_button():
    pdf_path = _find(REPORT_PDF) or _find_any(".pdf")
    if pdf_path is None:
        st.button("Download PDF report", key="pdf_btn", type="primary", disabled=True,
                  help=f"Place {REPORT_PDF} in the same folder as app.py")
        return
    st.download_button("Download PDF report", data=pdf_path.read_bytes(), file_name=pdf_path.name,
                       mime="application/pdf", key="pdf_btn", type="primary")


def insights_button():
    """Shows / hides the written Dashboard Insights (index.html)."""
    if "show_insights" not in st.session_state:
        st.session_state["show_insights"] = False
    label = "Hide dashboard insights" if st.session_state["show_insights"] else "Dashboard insights"

    try:
        row = st.container(horizontal=True, gap="small")          # Streamlit 1.48+
        b1 = b2 = b3 = row
    except TypeError:
        b1, b2, b3, _ = st.columns([1, 1, 1.3, 4], gap="small")
    with b1:
        clicked = st.button(label, key="insights_btn", type="primary")
    with b2:
        pdf_download_button()
    with b3:
        notebook_download_button()
    if clicked:
        st.session_state["show_insights"] = not st.session_state["show_insights"]
        st.rerun()

    if st.session_state["show_insights"]:
        html_path = _find("index.html")
        if html_path is None:
            note("<b>index.html</b> not found. Place it in the same folder as app.py.")
        else:
            with st.container(border=True):
                html_text = html_path.read_text(encoding="utf-8")
                if LOGO_URI:                       # the iframe can't read local files, so embed the logo
                    html_text = html_text.replace(f'src="{LOGO_FILE}"', f'src="{LOGO_URI}"')
                if hasattr(st, "iframe"):
                    st.iframe(html_text, height=900)
                else:
                    components.html(html_text, height=900, scrolling=True)
        st.write("")


# ════════════════════════════════════════════════════════════════════════════
# PAGE 1 — EXECUTIVE OVERVIEW
# ════════════════════════════════════════════════════════════════════════════
def page_overview():
    insights_button()

    cf_y = BY[CF].sum()
    prev = BANK[BANK["Year"] == year - 1][CF].sum() if (year - 1) in YEARS else np.nan
    cf_yoy = (cf_y / prev - 1) * 100 if prev else np.nan
    cf_cagr = cagr(BANK[BANK["Year"] == p0][CF].sum(), cf_y, N)
    adv_y = BY[ADV].sum()
    cf_share = cf_y / adv_y * 100 if adv_y else np.nan
    ts = type_summary(BANK).set_index("Bank Type")
    digital = ts.loc["Digital Bank", "CF % of advances"] if "Digital Bank" in ts.index else np.nan

    c = st.columns(5, gap="small")
    with c[0]:
        kpi(f"Consumer finance {year}", f_bn(cf_y), delta(cf_yoy, f"vs {year - 1}"), dark=True)
    with c[1]:
        kpi(f"Consumer finance CAGR {p0}–{p1}", f_pct(cf_cagr) if N else "—", "Nominal, all products")
    with c[2]:
        kpi(f"Consumer finance share of advances", f_pct(cf_share), f"{year}, in-scope banks")
    with c[3]:
        kpi(f"Personal loans real CAGR", f_pct(REAL_CAGR["Personal Loans"]) if N else "—",
            f"After inflation, {p0}–{p1}")
    with c[4]:
        kpi("Digital bank: CF % of advances", f_pct(digital), f"vs {f_pct(ts.loc['Conventional + Islamic Banking', 'CF % of advances']) if 'Conventional + Islamic Banking' in ts.index else '—'} for conventional banks")

    section("Market at a glance", f"Outstanding consumer financing by product, {p0}–{p1}")
    l, r = st.columns([1.35, 1], gap="medium")
    with l:
        show_fig(product_trend_fig(), "Consumer financing by product", "PKR Bn, stacked")
    with r:
        show_fig(product_mix_fig(), "Product mix", f"Share of total, {p0} vs {p1}")

    section("Key findings")
    pl = "Personal Loans"
    insights([
        f"Personal loans fell from <b>{f_bn(PROD.loc[p0, pl])}</b> to <b>{f_bn(PROD.loc[p1, pl])}</b>. "
        f"After inflation they shrank <b>{f_pct(abs(REAL_CAGR[pl]))} a year</b>: banks cut supply when inflation peaked.",
        f"Only <b>home finance</b> ({f_pct(REAL_CAGR['Home Finance'], sign=True)}) and <b>credit cards</b> "
        f"({f_pct(REAL_CAGR['Credit Card Financing'], sign=True)}) grew faster than inflation.",
        f"Conventional banks put only <b>{f_pct(ts.loc['Conventional + Islamic Banking', 'CF % of advances']) if 'Conventional + Islamic Banking' in ts.index else '—'}</b> "
        f"of advances into consumer finance; the digital bank puts <b>{f_pct(digital)}</b>. The digital model for retail lending is proven.",
        f"Inflation fell to <b>{f_pct(INFL.loc[p1])}</b> in {p1} and personal loans grew "
        f"<b>{f_pct(YOY.loc[p1, pl], sign=True)}</b>: demand is returning.",
        "<b>Recommendation:</b> launch with <b>digital personal loans</b>, add vehicle finance in year 2 and home finance from year 3.",
    ])


# ════════════════════════════════════════════════════════════════════════════
# PAGE 2 — PRODUCTS & INFLATION
# ════════════════════════════════════════════════════════════════════════════
def page_products():
    c = st.columns(4, gap="small")
    for i, p in enumerate(PRODUCTS):
        with c[i]:
            kpi(p, f_bn(PROD.loc[p1, p]),
                f"Nominal {f_pct(PROD_CAGR[p], sign=True)} · Real {f_pct(REAL_CAGR[p], sign=True)} a year"
                if N else f"{p1} outstanding", dark=(p == "Personal Loans"))

    section("Consumer finance by product", f"{p0}–{p1}, PKR Bn")
    show_fig(product_lines_fig(), f"Consumer Finance by Product, {p0}–{p1} (PKR Bn)",
             "Outstanding balances, nominal")

    CARD_H = 470                       # all four cards on this page get the same size
    cfg = {"displayModeBar": False}

    def plot(fig):
        if ST_VER >= (1, 46):
            st.plotly_chart(fig, width="stretch", config=cfg)
        else:
            st.plotly_chart(fig, use_container_width=True, config=cfg)

    section("Inflation", "Nominal values hide how much the market shrank in 2022–2024")
    l, r = st.columns([1, 1], gap="medium")
    with l:
        with st.container(border=True, height=CARD_H):
            st.markdown("<div class='ct'>Average inflation rate</div><div class='cs'>% per year, peak year in black</div>",
                        unsafe_allow_html=True)
            plot(inflation_fig(height=300))
            st.markdown("<div class='cs'><b>Nominal</b> is the rupee value in the data. <b>Real</b> divides by the "
                        f"price index, so every year is shown in {p0} prices.</div>", unsafe_allow_html=True)
    with r:
        with st.container(border=True, height=CARD_H):
            prod = st.selectbox("Product", PRODUCTS, key="f_prod", label_visibility="collapsed")
            st.markdown(f"<div class='ct'>{prod}: nominal vs real</div><div class='cs'>PKR Bn</div>",
                        unsafe_allow_html=True)
            plot(nominal_real_fig(prod, height=310))

    section("Growth", f"{p0}–{p1}")
    l, r = st.columns([1, 1], gap="medium")
    with l:
        g = pd.DataFrame({"Product": PRODUCTS, "Real CAGR (%)": REAL_CAGR[PRODUCTS].values}) \
            .sort_values("Real CAGR (%)", ascending=False)
        with st.container(border=True, height=CARD_H):
            st.markdown("<div class='ct'>Real growth per year</div><div class='cs'>After inflation</div>",
                        unsafe_allow_html=True)
            plot(hbar(g, "Real CAGR (%)", "Product", suffix="%", highlight=["Personal Loans"], height=360))
    with r:
        y = YOY.loc[SEL_YEARS[1:]] if len(SEL_YEARS) > 1 else YOY.loc[[year]]
        t = y.T.round(1)
        t.columns = [str(cn) for cn in t.columns]
        t.insert(0, "Product", t.index)
        with st.container(border=True, height=CARD_H):
            st.markdown("<div class='ct'>Year-on-year growth</div><div class='cs'>% change, nominal</div>",
                        unsafe_allow_html=True)
            show_table(t, column_config={c: st.column_config.NumberColumn(format="%.1f")
                                         for c in t.columns if c != "Product"})
            vol = (YOY.loc[SEL_YEARS].std()).round(1)
            st.markdown(f"<div class='cs'>Volatility (std. dev. of yearly growth): "
                        + " · ".join(f"{p} <b>{vol[p]:.1f}%</b>" for p in PRODUCTS) + "</div>",
                        unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════════════════════
# PAGE 3 — BANKS & COMPETITION
# ════════════════════════════════════════════════════════════════════════════
def page_banks():
    ranked = BY.sort_values(CF, ascending=False).reset_index(drop=True)
    total = ranked[CF].sum()
    ranked["Market share (%)"] = ranked[CF] / total * 100 if total else np.nan
    top5 = ranked.head(5)["Market share (%)"].sum()
    hhi = (ranked["Market share (%)"] ** 2).sum()

    c = st.columns(4, gap="small")
    with c[0]:
        kpi(f"Largest lender {year}", ranked.iloc[0]["Bank"], f_bn(ranked.iloc[0][CF]), small=True, dark=True)
    with c[1]:
        kpi("Top 5 market share", f_pct(top5), "of consumer finance")
    with c[2]:
        kpi("Concentration (HHI)", f"{hhi:,.0f}", "Below 1,500 = not concentrated")
    with c[3]:
        kpi("Advances / deposits", f_pct(BY[ADV].sum() / BY[DEP].sum() * 100), "Share of deposits lent out")

    cfg = {"displayModeBar": False}

    def plot(fig):
        if ST_VER >= (1, 46):
            st.plotly_chart(fig, width="stretch", config=cfg)
        else:
            st.plotly_chart(fig, use_container_width=True, config=cfg)

    def card(title, sub, height):
        """Bordered card with a fixed height and a title."""
        box = st.container(border=True, height=height)
        box.markdown(f"<div class='ct'>{title}</div><div class='cs'>{sub}</div>", unsafe_allow_html=True)
        return box

    section("Who lends today", f"{year}")
    # slider in the same column width as the first chart
    sl, _ = st.columns([1, 1], gap="medium")
    with sl:
        n = st.slider("Banks to show", 5, 20, 10, key="f_topn")
    top = ranked.head(n)
    ts = type_summary(BANK).dropna(subset=["CF % of advances"]).sort_values("CF % of advances", ascending=False)
    h = max(300, 30 * n + 60)          # chart height grows with the number of banks
    row1_h = h + 90                    # same card height on both sides

    l, r = st.columns([1, 1], gap="medium")
    with l:
        with card(f"Top {n} banks by consumer finance", "PKR Bn, top 5 in black", row1_h):
            plot(hbar(top, CF, "Bank", highlight=top.head(5)["Bank"].tolist(), height=h))
    with r:
        with card("Consumer finance as % of advances",
                  "By bank type: how much each type focuses on retail lending", row1_h):
            plot(hbar(ts, "CF % of advances", "Bank Type", suffix="%", highlight=["Digital Bank"], height=h))

    section("Bank types", f"{p0}–{p1}")
    trend = BANK[BANK["Year"].isin(SEL_YEARS)].groupby(["Year", "Bank Type"])[CF].sum().unstack()
    trend = trend.loc[:, trend.sum() > 0]
    share = trend.div(trend.sum(axis=1), axis=0) * 100
    order = share.mean().sort_values(ascending=False).index.tolist()
    fig = go.Figure()
    for i, t in enumerate(order):
        fig.add_scatter(x=share.index, y=share[t], name=t, stackgroup="one", mode="lines",
                        line=dict(width=0.6, color="#ffffff"), fillcolor=GREYS[i % len(GREYS)],
                        hovertemplate=f"<b>{t}</b><br>%{{x}}: %{{y:.2f}}%<extra></extra>")
    style(fig, height=360, ysuffix="%")
    fig.update_xaxes(dtick=1)
    fig.update_yaxes(range=[0, 100])

    ROW2_H = 470
    l, r = st.columns([1, 1], gap="medium")
    with l:
        with card("Share of consumer finance by bank type", "% of total", ROW2_H):
            plot(fig)
    with r:
        tbl = type_summary(BANK).sort_values(CF, ascending=False)
        tbl["Bank Type"] = tbl["Bank Type"].map(lambda v: SHORT_TYPE.get(v, v))
        with card(f"Bank type summary, {year}", "CF = consumer finance", ROW2_H):
            # show "—" instead of "None" where a bank type has no lending
            view = pd.DataFrame({
                "Bank Type": tbl["Bank Type"],
                "Banks": tbl["Banks"],
                "CF (Bn)": tbl[CF].map(lambda v: f"{v:,.1f}"),
                "CF / adv.": tbl["CF % of advances"].map(f_pct),
                "Share": tbl["Market share (%)"].map(f_pct),
                "CAGR": tbl["CF CAGR (%)"].map(f_pct),
            })
            show_table(view)
            st.markdown("<div class='cs'><b>Islamic digital banks</b> (Raqami, Mashreq) have no lending yet, so a "
                        "Shariah-compliant digital loan faces little direct competition. — = no lending.</div>",
                        unsafe_allow_html=True)

    with st.expander("All banks, " + str(year)):
        show_table(ranked[["Bank", "Bank Type", CF, "Market share (%)", ADV, DEP, "CF Growth (%)"]].round(2),
                   height=420)


# ════════════════════════════════════════════════════════════════════════════
# PAGE 4 — PRODUCT FIT & RECOMMENDATION
# ════════════════════════════════════════════════════════════════════════════
def page_fit():
    section("Product comparison", f"{p0}–{p1}")
    vol = YOY.loc[SEL_YEARS, CORE].std()
    comp = pd.DataFrame({
        "Product": CORE,
        f"{p0} (PKR Bn)": PROD.loc[p0, CORE].values,
        f"{p1} (PKR Bn)": PROD.loc[p1, CORE].values,
        "Nominal CAGR (%)": PROD_CAGR[CORE].values,
        "Real CAGR (%)": REAL_CAGR[CORE].values,
        "Volatility (%)": vol[CORE].values,
        f"Growth {p1} (%)": YOY.loc[p1, CORE].values,
    }).round(1)
    show_table(comp, column_config={c: st.column_config.NumberColumn(format="%.1f")
                                    for c in comp.columns if c != "Product"})

    section("DRB fit score", "Scores from 1 (poor) to 5 (best), weighted")
    scores = pd.DataFrame({
        "Easy to offer digitally": [5, 3, 1],
        "Short loan tenor": [5, 3, 1],
        "Low risk": [4, 2, 2],
        "Market gap": [5, 3, 4],
        "Real growth": [2, 2, 4],
    }, index=CORE)

    # fixed weights (total = 100%)
    weights = {"Easy to offer digitally": 0.25, "Short loan tenor": 0.20, "Low risk": 0.20,
               "Market gap": 0.20, "Real growth": 0.15}
    scores["Total score"] = sum(scores[f] * w for f, w in weights.items())
    res = scores.sort_values("Total score", ascending=False).reset_index().rename(columns={"index": "Product"})

    FIT_H = 340                        # both cards get the same size
    cfg = {"displayModeBar": False}
    l, r = st.columns([1, 1], gap="medium")
    with l:
        with st.container(border=True, height=FIT_H):
            st.markdown("<div class='ct'>Product fit for a new digital bank</div>"
                        "<div class='cs'>Weighted score out of 5</div>", unsafe_allow_html=True)
            fig = hbar(res, "Total score", "Product", fmt=".2f", highlight=[res.iloc[0]["Product"]], height=250)
            fig.update_xaxes(range=[0, 5.8])
            if ST_VER >= (1, 46):
                st.plotly_chart(fig, width="stretch", config=cfg)
            else:
                st.plotly_chart(fig, use_container_width=True, config=cfg)
    with r:
        with st.container(border=True, height=FIT_H):
            st.markdown("<div class='ct'>Score breakdown</div><div class='cs'>1 = poor, 5 = best</div>",
                        unsafe_allow_html=True)
            short = res.rename(columns={"Easy to offer digitally": "Digital", "Short loan tenor": "Tenor",
                                        "Low risk": "Risk", "Market gap": "Gap", "Real growth": "Growth",
                                        "Total score": "Total"})
            show_table(short, column_config={"Total": st.column_config.NumberColumn(format="%.2f")})
            st.markdown("<div class='cs'><b>Weights:</b> " + " · ".join(f"{f} {w:.0%}" for f, w in weights.items())
                        + "<br>Scores are analyst assessment based on the data and DRB operating constraints.</div>",
                        unsafe_allow_html=True)

    section("Roadmap", "Recommended launch sequence")
    phases = [
        ("1", "Year 1", "Digital personal loans",
         "Salary advance and small-ticket loans fully in the app, with a Shariah-compliant option.", True),
        ("2", "Year 2", "Vehicle finance",
         "Car and motorbike finance through dealer partnerships, once rates and the deposit base are stable.", False),
        ("3", "Year 3+", "Home finance",
         "Home finance through partners or government schemes, when long-term funding is in place.", False),
    ]
    c = st.columns(3, gap="small")
    for col, (n, w, h, p, dark) in zip(c, phases):
        with col:
            st.markdown(f"<div class='phase{' dark' if dark else ''}'><span class='n'>{n}</span>"
                        f"<div class='w'>{w}</div><h4>{h}</h4><p>{p}</p></div>", unsafe_allow_html=True)

    section("Why this order")
    insights([
        f"<b>Personal loans</b> are the most stable product (volatility {vol['Personal Loans']:.1f}%), need no "
        "collateral or dealer network, and their short tenor matches a new bank's short-term deposits.",
        f"<b>Vehicle finance</b> recovered fastest in {p1} ({f_pct(YOY.loc[p1, 'Auto Finance'], sign=True)}) but is the "
        f"most volatile ({vol['Auto Finance']:.1f}%) and depends on dealers, car prices and interest rates.",
        f"<b>Home finance</b> has the best real growth ({f_pct(REAL_CAGR['Home Finance'], sign=True)} a year) but needs "
        "15–20 year funding, property checks and large capital.",
        f"<b>Credit cards</b> grew {f_pct(PROD_CAGR['Credit Card Financing'])} a year. A digital card or buy now, pay later "
        "product fits naturally next to personal loans.",
    ])


# ════════════════════════════════════════════════════════════════════════════
# ROUTER
# ════════════════════════════════════════════════════════════════════════════
if page == PAGES[0]:
    page_overview()
elif page == PAGES[1]:
    page_products()
elif page == PAGES[2]:
    page_banks()
else:
    page_fit()

st.markdown(
    "<div style='margin-top:36px;padding-top:12px;border-top:1px solid #e2e2e2;display:flex;"
    "justify-content:space-between;align-items:flex-start;gap:24px;flex-wrap:wrap'>"
    "<div style='font-size:11.5px;color:#8a8a8a'>"
    "Recommended Lending Products for a DRB · figures in PKR billions unless stated · real values use average inflation"
    f"<div style='margin-top:6px;font-size:12.5px;color:#1a1a1a;font-weight:500'>{PREPARED_BY}</div></div>"
    f"<div style='font-size:11.5px;color:#8a8a8a;max-width:460px;text-align:right'>{DISCLAIMER}</div></div>",
    unsafe_allow_html=True,
)