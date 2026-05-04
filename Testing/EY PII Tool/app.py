# app.py  — PII Detection Tool

import io
from io import BytesIO
import warnings
import pandas as pd
import streamlit as st
from pandas.errors import SettingWithCopyWarning

# Import the helper functions for PII detection, OCR, and redaction processing
from pi_detection_functions import *

# -----------------------------
# Global warning suppression
# -----------------------------
warnings.filterwarnings("ignore", category=SettingWithCopyWarning)
warnings.filterwarnings("ignore", category=FutureWarning)
warnings.filterwarnings("ignore", category=UserWarning)

def _preferred_excel_engine() -> str:
    """Return 'xlsxwriter' if available, otherwise 'openpyxl'."""
    try:
        import xlsxwriter  # noqa: F401
        return "xlsxwriter"
    except Exception:
        return "openpyxl"

# -----------------------------
# Page config & optional CSS
# -----------------------------
st.set_page_config(page_title="PII Detection Tool", page_icon="🕵️", layout="wide")

# ===== Optional styling block =====
# ===== Optional styling block (dark theme + #2e2e38 sidebar) =====
st.markdown(
    """
    <style>
      :root{
        --bg: #111217;         /* page canvas */
        --panel:#1A1C23;       /* panels/cards */
        --text:#E5E7EB;        /* primary text */
        --muted:#9CA3AF;       /* secondary text */
        --brand:#FFD400;       /* EY yellow */
        --brand-600:#FACC15;   /* hover yellow */
        --pill-bg:#2B2F3A;     /* chip bg */
        --pill-bd:#3A3F4B;     /* chip border */
        --table-panel:#1A1C23; /* dataframe background */
      }

      /* Page background */
      [data-testid="stAppViewContainer"]{
        background: var(--bg);
      }

      /* Sidebar (toolbar) */
      section[data-testid="stSidebar"]{
        background: #2e2e38 !important; /* requested color */
        border-right: 1px solid #23242c;
      }

      /* Page width */
      .block-container {max-width: 1100px;}

      /* Hero title */
      .pii-title{
        font-size: 2.6rem;
        font-weight: 700;
        margin: 0.25rem 0 0.25rem 0;
        color: var(--text);
      }
      .pii-subtitle{
        color: var(--muted);
        margin-bottom: 0.5rem;
      }

      /* Info pills */
      .pill{
        display:inline-block;
        padding:0.25rem 0.6rem;
        border-radius:999px;
        background: var(--pill-bg);
        color: var(--text);
        font-size:0.85rem;
        border:1px solid var(--pill-bd);
        margin-right:0.35rem;
      }

      /* Section titles */
      .section-title{
        font-size:1.15rem;
        font-weight:600;
        margin-top:1rem;
        margin-bottom:0.25rem;
        color: var(--text);
      }

      /* Subtle caption */
      .caption{
        color: var(--muted);
        font-size:0.9rem;
      }

      /* DataFrame background + text size */
      .stDataFrame {font-size: 0.92rem;}
      [data-testid="stDataFrame"] div[data-testid="stHorizontalBlock"]{
        background: var(--table-panel);
      }

      /* Buttons (optional: EY yellow) */
      .stButton>button{
        background: var(--brand);
        color:#000;
        border:1px solid #EAB308;
        border-radius:8px;
        font-weight:700;
      }
      .stButton>button:hover{ background: var(--brand-600); }
    
      
        /* Style only the final "Download PII results (Excel)" button */
        .dl-results-btn .stDownloadButton > button {
            background: #FFD400 !important;   /* EY yellow, like Scan */
            color: #000 !important;           /* <-- black text */
            border: 1px solid #EAB308 !important;
            border-radius: 8px !important;
            font-weight: 700 !important;
        }
        .dl-results-btn .stDownloadButton > button:hover {
            background: #FACC15 !important;   /* hover yellow */
        }

        /* Scope styling to a wrapper we add around the specific download button(s) */
        #dl-results-btn div[data-testid="stDownloadButton"] > button,
        #dl-results-btn button[data-baseweb="button"],                       /* fallback for some versions */
        #dl-results-btn div[data-testid="baseButton-secondary"] > button {   /* extra fallback */
            background: #FFD400 !important;   /* EY yellow - same as Scan */
            color: #000 !important;           /* <-- BLACK TEXT */
            border: 1px solid #EAB308 !important;
            border-radius: 8px !important;
            font-weight: 700 !important;
        }

        #dl-results-btn div[data-testid="stDownloadButton"] > button:hover,
        #dl-results-btn button[data-baseweb="button"]:hover,
        #dl-results-btn div[data-testid="baseButton-secondary"] > button:hover {
            background: #FACC15 !important;   /* hover yellow */
        }

        /* Style only the button whose label is exactly this string */
        div[data-testid="stDownloadButton"][aria-label="Download PII results (Excel)"] > button,
        button[aria-label="Download PII results (Excel)"],
        div[aria-label="Download PII results (Excel)"] > button,
        div[aria-label="Download PII results (Excel)"] button {
            background: #FFD400 !important;   /* EY yellow */
            color: #000 !important;           /* BLACK TEXT */
            border: 1px solid #EAB308 !important;
            border-radius: 8px !important;
            font-weight: 700 !important;
        }
        /* Hover state */
        div[data-testid="stDownloadButton"][aria-label="Download PII results (Excel)"] > button:hover,
        button[aria-label="Download PII results (Excel)"]:hover,
        div[aria-label="Download PII results (Excel)"] > button:hover,
        div[aria-label="Download PII results (Excel)"] button:hover {
            background: #FACC15 !important;
        }

    
        /* ===== Scoped styles for the Disclaimer expander ===== */
        /* Wrap target in a #disclaimer-box (see Python snippet below) to avoid affecting other expanders */

        /* Container, header button, and content get the same background */
        #disclaimer-box [data-testid="stExpander"] {
            background: #3A3F4B !important;
            border: 1px solid #3A3F4B !important;
            border-radius: 10px !important;
        }
        #disclaimer-box [data-testid="stExpander"] div[role="button"] {
            background: #3A3F4B !important;     /* header strip */
            color: var(--text) !important;       /* your light text */
            border-radius: 10px 10px 0 0 !important;
        }
        /* Expander content area */
        #disclaimer-box [data-testid="stExpander"] > div:not([role="button"]),
        #disclaimer-box [data-testid="stExpander"] section,
        #disclaimer-box [data-testid="stExpander"] .st-expander-content {
            background: #3A3F4B !important;
            color: var(--text) !important;
            border-radius: 0 0 10px 10px !important;
        }
        /* Toggle icon color */
        #disclaimer-box [data-testid="stExpander"] [data-testid="stExpanderToggleIcon"] svg {
            color: var(--text) !important;
        }
        /* Markdown text inside the expander */
        #disclaimer-box [data-testid="stMarkdownContainer"] p,
        #disclaimer-box [data-testid="stMarkdownContainer"] li,
        #disclaimer-box [data-testid="stMarkdownContainer"] strong,
        #disclaimer-box [data-testid="stMarkdownContainer"] em {
            color: var(--text) !important;
        }

        /* Older Streamlit builds fallback (class names may exist) */
        #disclaimer-box .streamlit-expanderHeader,
        #disclaimer-box .streamlit-expanderContent {
            background: #3A3F4B !important;
            color: var(--text) !important;
        }

        /* ===== Force expander background (header + content) to #3A3F4B ===== */

        /* Container (modern builds) */
        div[data-testid="stExpander"] {
            background: #3A3F4B !important;
            border: 1px solid #3A3F4B !important;
            border-radius: 10px !important;
            box-shadow: none !important;
        }

        /* Header bar */
        div[data-testid="stExpander"] > div[role="button"] {
            background: #3A3F4B !important;
            color: var(--text) !important;
            border-radius: 10px 10px 0 0 !important;
        }

        /* Content area (modern builds) */
        div[data-testid="stExpander"] .st-expander-content,
        div[data-testid="stExpander"] > div:not([role="button"]) {
            background: #3A3F4B !important;
            color: var(--text) !important;
            border-radius: 0 0 10px 10px !important;
        }

        /* Toggle chevron color */
        div[data-testid="stExpander"] [data-testid="stExpanderToggleIcon"] svg {
            color: var(--text) !important;
        }

        /* ===== Fallbacks for older Streamlit builds ===== */
        .streamlit-expanderHeader,
        .streamlit-expanderContent {
            background: #3A3F4B !important;
            color: var(--text) !important;
            box-shadow: none !important;
        }

        /* Keep the expander header ALWAYS #3A3F4B in every state */

        /* Modern Streamlit builds: the header is a button inside the expander container */
        div[data-testid="stExpander"] > div[role="button"],
        div[data-testid="stExpander"] > div[role="button"][aria-expanded="true"],
        div[data-testid="stExpander"] > div[role="button"][aria-expanded="false"],
        div[data-testid="stExpander"] > div[role="button"]:hover,
        div[data-testid="stExpander"] > div[role="button"]:focus,
        div[data-testid="stExpander"] > div[role="button"]:active {
            background: #3A3F4B !important;
            background-color: #3A3F4B !important;   /* some themes set background-color specifically */
            background-image: none !important;      /* kill gradients/images if any */
            color: var(--text) !important;
            border-color: #3A3F4B !important;
            box-shadow: none !important;
            outline: none !important;
        }

        /* Make sure any child text/icon inside the header also uses the light text color */
        div[data-testid="stExpander"] > div[role="button"] * {
            color: var(--text) !important;
            fill: var(--text) !important;
        }

        /* Keep the outer expander container the same color (prevents a frame showing through) */
        div[data-testid="stExpander"] {
            background: #3A3F4B !important;
            border: 1px solid #3A3F4B !important;
            border-radius: 10px !important;
            box-shadow: none !important;
        }

        /* Content area stays the same color */
        div[data-testid="stExpander"] .st-expander-content,
        div[data-testid="stExpander"] > div:not([role="button"]) {
            background: #3A3F4B !important;
            color: var(--text) !important;
            border-radius: 0 0 10px 10px !important;
        }

        /* Chevron color */
        div[data-testid="stExpander"] [data-testid="stExpanderToggleIcon"] svg {
            color: var(--text) !important;
            fill: var(--text) !important;
        }

        /* Fallbacks for older Streamlit builds */
        .streamlit-expanderHeader,
        .streamlit-expanderHeader:hover,
        .streamlit-expanderHeader:focus,
        .streamlit-expanderHeader:active {
            background: #3A3F4B !important;
            background-image: none !important;
            color: var(--text) !important;
            box-shadow: none !important;
        }
        .streamlit-expanderContent {
            background: #3A3F4B !important;
            color: var(--text) !important;
        }

        /* ===== Scope styles to ONLY the expander that immediately follows #disc-anchor ===== */

        /* Expander container */
        #disc-anchor + div [data-testid="stExpander"] {
            background: #3A3F4B !important;
            border: 1px solid #3A3F4B !important;
            border-radius: 10px !important;
            box-shadow: none !important;
        }

        /* Header button (all states) */
        #disc-anchor + div [data-testid="stExpander"] > div[role="button"],
        #disc-anchor + div [data-testid="stExpander"] > div[role="button"][aria-expanded="true"],
        #disc-anchor + div [data-testid="stExpander"] > div[role="button"][aria-expanded="false"],
        #disc-anchor + div [data-testid="stExpander"] > div[role="button"]:hover,
        #disc-anchor + div [data-testid="stExpander"] > div[role="button"]:focus,
        #disc-anchor + div [data-testid="stExpander"] > div[role="button"]:active {
            background: #3A3F4B !important;
            background-color: #3A3F4B !important;
            background-image: none !important;
            color: var(--text) !important;
            border-color: #3A3F4B !important;
            box-shadow: none !important;
            outline: none !important;
            border-radius: 10px 10px 0 0 !important;
        }

        /* Ensure header text & icon stay light */
        #disc-anchor + div [data-testid="stExpander"] > div[role="button"] *,
        #disc-anchor + div [data-testid="stExpander"] [data-testid="stExpanderToggleIcon"] svg {
            color: var(--text) !important;
            fill: var(--text) !important;
        }

        /* Content pane */
        #disc-anchor + div [data-testid="stExpander"] .st-expander-content,
        #disc-anchor + div [data-testid="stExpander"] > div:not([role="button"]) {
            background: #3A3F4B !important;
            color: var(--text) !important;
            border-radius: 0 0 10px 10px !important;
        }

        /* ==== Fallbacks for older Streamlit DOMs (class names may exist) ==== */
        #disc-anchor + div .streamlit-expanderHeader,
        #disc-anchor + div .streamlit-expanderHeader:hover,
        #disc-anchor + div .streamlit-expanderHeader:focus,
        #disc-anchor + div .streamlit-expanderHeader:active,
        #disc-anchor + div .streamlit-expanderContent {
            background: #3A3F4B !important;
            color: var(--text) !important;
            box-shadow: none !important;
        }

        /* --- Make ALL download buttons match the Scan button --- */
        /* Base selector (modern Streamlit) */
        div[data-testid="stDownloadButton"] > button {
            background: var(--brand) !important;        /* #FFD400 */
            color: #000 !important;                     /* <-- BLACK TEXT */
            border: 1px solid #EAB308 !important;
            border-radius: 8px !important;
            font-weight: 700 !important;
            box-shadow: none !important;
            outline: none !important;
        }
        /* Ensure descendants (text/icon) are black too */
        div[data-testid="stDownloadButton"] > button * {
            color: #000 !important;
            fill: #000 !important;
        }
        /* Hover */
        div[data-testid="stDownloadButton"] > button:hover {
            background: var(--brand-600) !important;    /* #FACC15 */
        }

        /* Fallbacks for older DOMs */
        button[data-baseweb="button"][kind],                /* some builds wrap as baseweb button */
        div[data-testid="baseButton-secondary"] > button {  /* extra variant */
            color: #000 !important;
        }

        /* ===== Scope to the multiselect right after #redact-anchor ===== */

        /* Selected chips inside the control (Base Web Tag) */
        #redact-anchor + div [data-baseweb="select"] [data-baseweb="tag"] {
            background: var(--brand) !important;   /* EY yellow */
            color: #000 !important;                /* BLACK text */
            border-color: var(--brand) !important;
        }
        /* Ensure chip text and close (×) icon are black */
        #redact-anchor + div [data-baseweb="select"] [data-baseweb="tag"] * {
            color: #000 !important;
            fill: #000 !important;
        }

        /* Caret / dropdown icon inside the select stays readable */
        #redact-anchor + div [data-baseweb="select"] svg {
            color: #E5E7EB !important;  /* light grey on dark bg */
            fill: #E5E7EB !important;
        }

        /* Selected row inside the dropdown menu (popover) → yellow + BLACK text */
        /* (The menu is rendered in a portal/popover, so this selector is global.) */
        div[data-baseweb="popover"] [role="option"][aria-selected="true"] {
            background: var(--brand) !important;
            color: #000 !important;
        }
        /* Make all text/icons in the selected row black */
        div[data-baseweb="popover"] [role="option"][aria-selected="true"] * {
            color: #000 !important;
            fill:  #000 !important;
        }

        /* Optional: on hover (non-selected) keep your current dark theme contrast */
        /* div[data-baseweb="popover"] [role="option"]:hover { background: var(--brand-600) !important; } */

        /* ===== Make selected chips (Tags) show BLACK text on EY yellow for this one multiselect ===== */

        /* The Tag background for selected values (the chips shown in the input) */
        #redact-anchor + div [data-baseweb="select"] [data-baseweb="tag"] {
            background: var(--brand) !important;      /* EY yellow */
            border-color: var(--brand) !important;
        }

        /* Force ALL inner text/icon of the Tag to black */
        #redact-anchor + div [data-baseweb="select"] [data-baseweb="tag"] *,
        #redact-anchor + div [data-baseweb="select"] [data-baseweb="tag"] svg,
        #redact-anchor + div [data-baseweb="select"] [data-baseweb="tag"] svg path,
        #redact-anchor + div [data-baseweb="select"] [data-baseweb="tag"] button {
            color: #000 !important;                   /* text nodes */
            fill:  #000 !important;                   /* svg icon fill */
            stroke:#000 !important;                   /* svg paths that use stroke */
            -webkit-text-fill-color:#000 !important;  /* Safari text rendering */
        }

        /* Optional: chip hover tone (slightly darker EY yellow) */
        #redact-anchor + div [data-baseweb="select"] [data-baseweb="tag"]:hover {
            background: var(--brand-600) !important;
        }

        /* Keep the caret/dropdown icon readable against your dark input field */
        #redact-anchor + div [data-baseweb="select"] svg {
            color: #E5E7EB !important;
            fill:  #E5E7EB !important;
        }

        /* ===== Keep selected row in the dropdown menu black-on-yellow (popover lives outside the control) ===== */
        div[data-baseweb="popover"] [role="option"][aria-selected="true"] {
            background: var(--brand) !important;
            color: #000 !important;
        }
        div[data-baseweb="popover"] [role="option"][aria-selected="true"] * {
            color: #000 !important;
            fill:  #000 !important;
            stroke:#000 !important;
        }

        /* (Optional) non-selected hover tone in the menu */
        /* div[data-baseweb="popover"] [role="option"]:hover { background: var(--brand-600) !important; } */

        /* ===== Force BLACK text + icon on selected chips (Base Web Tag) for the multiselect after #redact-anchor ===== */

        /* Chip container (the Tag) — EY yellow background */
        #redact-anchor + div [data-baseweb="select"] [data-baseweb="tag"] {
            background: var(--brand) !important;        /* EY yellow (#FFD400) */
            border-color: var(--brand) !important;
        }

        /* Text + icon inside the Tag — force BLACK across all descendants */
        #redact-anchor + div [data-baseweb="select"] [data-baseweb="tag"] *,
        #redact-anchor + div [data-baseweb="select"] [data-baseweb="tag"] span,
        #redact-anchor + div [data-baseweb="select"] [data-baseweb="tag"] button,
        #redact-anchor + div [data-baseweb="select"] [data-baseweb="tag"] button *,
        #redact-anchor + div [data-baseweb="select"] [data-baseweb="tag"] svg,
        #redact-anchor + div [data-baseweb="select"] [data-baseweb="tag"] svg *,
        #redact-anchor + div [data-baseweb="select"] [data-baseweb="tag"] svg path {
            color: #000 !important;                     /* text nodes + some icons */
            fill:  #000 !important;                     /* svg icon fill */
            stroke:#000 !important;                     /* svg paths that use stroke */
            -webkit-text-fill-color:#000 !important;    /* Safari text rendering */
        }

        /* Force black text on every element inside a selected pill/tag */
        [data-baseweb="tag"],
        [data-baseweb="tag"] *,
        [data-baseweb="tag"] span,
        [data-baseweb="tag"] > span,
        [data-baseweb="tag"] span span {
            background-color: var(--brand) !important;
            color: #000000 !important;
            -webkit-text-fill-color: #000000 !important;
            fill: #000000 !important;
        }
        /* Catch the inner text node specifically */
        span[data-baseweb="tag"] > span:first-child {
            color: #000000 !important;
            -webkit-text-fill-color: #000000 !important;
        }

        /* Hover/focus/active on the chip shouldn't flip colors */
        #redact-anchor + div [data-baseweb="select"] [data-baseweb="tag"]:hover,
        #redact-anchor + div [data-baseweb="select"] [data-baseweb="tag"]:focus,
        #redact-anchor + div [data-baseweb="select"] [data-baseweb="tag"]:active {
            background: var(--brand-600) !important;    /* slightly darker yellow (#FACC15) */
        }
        #redact-anchor + div [data-baseweb="select"] [data-baseweb="tag"]:hover *,
        #redact-anchor + div [data-baseweb="select"] [data-baseweb="tag"]:focus *,
        #redact-anchor + div [data-baseweb="select"] [data-baseweb="tag"]:active * {
            color:#000 !important; fill:#000 !important; stroke:#000 !important;
        }

        /* Ensure the caret/dropdown icon in the input itself stays legible on your dark field */
        #redact-anchor + div [data-baseweb="select"] svg:not([data-icon]) {
            color: #E5E7EB !important;
            fill:  #E5E7EB !important;
        }

        /* (You already have this, included for completeness)
            Selected row in the open dropdown menu should be black-on-yellow too */
        div[data-baseweb="popover"] [role="option"][aria-selected="true"] {
            background: var(--brand) !important;
            color: #000 !important;
        }
        div[data-baseweb="popover"] [role="option"][aria-selected="true"] * {
            color:#000 !important; fill:#000 !important; stroke:#000 !important;
        }

        /* Sidebar EY logo: top-left + spacing below */
        section[data-testid="stSidebar"] .ey-sb-logo {
        display: block;
        margin: 8px 12px 16px 12px;  /* left/right + a 16px gap below */
        }

        /***** Sidebar separator line *****/
        section[data-testid="stSidebar"] hr.ey-sep {
        border: 0;                      /* remove default border */
        height: 2px;                    /* line thickness */
        background: #3a3a47;            /* grey line */
        margin: 12px 0 14px 0;          /* spacing above/below the line */
        }

        /* Smaller pills ONLY in the file-types row */
        #file-pills .pill{
        padding:3px 8px;
        border-radius:12px;
        font-size:12px;
        line-height:1.1;
        margin-right:0.25rem;
        }

      /* Header bar transparent (optional) */
      [data-testid="stHeader"]{ background: transparent; }
    </style>
    """,
    unsafe_allow_html=True
)

# -----------------------------
# Session state initialization
# -----------------------------
DEFAULT_STATE = {
    "pi_df": None,
    "image_df": None,
    "redacted_files": None,
    "scan_completed": False,
    "redact_completed": False,
    "run_scan": False,
    "disclaimer_open": True,
}
for k, v in DEFAULT_STATE.items():
    if k not in st.session_state:
        st.session_state[k] = v

# --- Sidebar logo at absolute top (left-aligned by default) ---
st.sidebar.image("LOGO.png", width=50)  # adjust width as you like
# Spacer to add more room below the logo
st.sidebar.markdown("<div style='height:24px;'></div>", unsafe_allow_html=True)

# -----------------------------
# Header
# -----------------------------
st.markdown('<div class="pii-title">PII Detection Tool</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="pii-subtitle">Scan documents for potential personal identifiers and export redacted copies.</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div id="file-pills">'
    '<span class="pill">.xlsx</span>'
    '<span class="pill">.xls</span>'
    '<span class="pill">.docx</span>'
    '<span class="pill">.pdf</span>'
    '</div>',
    unsafe_allow_html=True
)

# -----------------------------
# Disclaimer
# -----------------------------
st.markdown("<div id='disc-anchor'></div>", unsafe_allow_html=True)   # ⬅️ start scope
with st.expander("Disclaimer", expanded=st.session_state.disclaimer_open):
    st.markdown(
        """
**Objective:** Automated assistance to find potential personally identifiable information (PII).  
**Supported File Types:** `.xlsx`, `.xls`, `.docx`, `.pdf`  
**Process:**
- Use the left sidebar to upload files.
- Click **Scan**.
- The tool scans for: Passport, SSN, Driver's License, Credit/Debit Card, Emails, Phone, Addresses, DOB, States, and more.
- In the results, choose the PII to redact from **Excel files** and export.

⚠️ This tool is not a substitute for manual review. Always validate results.
""",
        unsafe_allow_html=True,
    )
# st.markdown('</div>', unsafe_allow_html=True)                     # ⬅️ end scope

# -----------------------------
# Sidebar — Upload & Scan
# -----------------------------

with st.sidebar:

    # To draw a grey separator above the heading
    st.markdown('<hr class="ey-sep">', unsafe_allow_html=True)

    st.markdown("### Upload Section")
    st.caption("Upload multiple files to scan for sensitive information.")
    st.session_state.source_files = st.file_uploader(
        "Upload Files",
        type=["xlsx", "xls", "docx", "pdf"],
        accept_multiple_files=True,
    )

    # Status pill 
    status = "Ready"
    if st.session_state.get("run_scan"):
        status = "Scanning…"
    elif st.session_state.get("scan_completed"):
        status = "Scanned"

    status_color = {"Ready": "#f3f4f6", "Scanning…": "#fff7ed", "Scanned": "#ecfdf5"}[status]
    status_border = {"Ready": "#e5e7eb", "Scanning…": "#ffedd5", "Scanned": "#d1fae5"}[status]
    status_text = {"Ready": "#374151", "Scanning…": "#9a3412", "Scanned": "#065f46"}[status]

    st.markdown(
        f"""
        <div style="
            display:inline-block;
            padding:3px 8px;                 /* tighter than 0.25rem 0.6rem */
            border-radius:12px;               /* smaller pill radius */
            background:{status_color};
            border:1px solid {status_border};
            color:{status_text};
            font-weight:600;
            font-size:12px;                   /* ↓ smaller font */
            line-height:1.1;                  /* keep pill height tight */
            margin:0.6rem 0;                  /* less vertical space */
        ">
        Status: {status}
        </div>
        """,
        unsafe_allow_html=True
    )

    scan_clicked = st.button("Scan")
    if scan_clicked:
        st.session_state.run_scan = True
        st.session_state.scan_completed = False
        st.session_state.redact_completed = False

# -----------------------------
# Scan processing
# -----------------------------
center_placeholder = st.empty()

if st.session_state.run_scan and st.session_state.source_files:
    with center_placeholder, st.spinner("Scanning files for PII. Please wait..."):
        # 1) Detect PII from text-bearing parts (PDF pages, DOCX paragraphs/tables, Excel cells)
        text_pi_df = detect_pi(st.session_state.source_files)
        if isinstance(text_pi_df, pd.DataFrame) and not text_pi_df.empty:
            text_pi_df = text_pi_df[text_pi_df["Confidence (%)"] >= 40]
            st.session_state.pi_df = clean_pi_data(text_pi_df)
        else:
            st.session_state.pi_df = None

        # 2) OCR image-only regions from PDF/DOCX/XLSX, then detect PII in OCR text
        img_text_df = extract_image_texts(st.session_state.source_files, pdf_scale=2.0)
        if isinstance(img_text_df, pd.DataFrame) and not img_text_df.empty:
            image_pi_df = detect_pi_from_textframe(img_text_df)
            if isinstance(image_pi_df, pd.DataFrame) and not image_pi_df.empty:
                cleaned_image_df = clean_pi_data(image_pi_df)
                # Remove from OCR results any (File, PI Type, Detected Value) that already
                # appears in the text-extracted results — prevents duplicate display when
                # a text-layer PDF is also OCR-scanned
                if isinstance(st.session_state.pi_df, pd.DataFrame) and not st.session_state.pi_df.empty:
                    text_keys = set(
                        zip(
                            st.session_state.pi_df["File"].astype(str),
                            st.session_state.pi_df["PI Type"].astype(str),
                            st.session_state.pi_df["Detected Value"].astype(str).str.strip(),
                        )
                    )
                    # Also deduplicate on value alone per file (catches format variations)
                    text_file_vals = set(
                        zip(
                            st.session_state.pi_df["File"].astype(str),
                            st.session_state.pi_df["Detected Value"].astype(str).str.strip(),
                        )
                    )
                    def _is_ocr_dup(row):
                        fv = (str(row["File"]), str(row["Detected Value"]).strip())
                        ftp = (str(row["File"]), str(row["PI Type"]), str(row["Detected Value"]).strip())
                        return ftp in text_keys or fv in text_file_vals
                    mask = cleaned_image_df.apply(_is_ocr_dup, axis=1)
                    cleaned_image_df = cleaned_image_df[~mask].copy()
                st.session_state.image_df = cleaned_image_df if not cleaned_image_df.empty else None
            else:
                st.session_state.image_df = None
        else:
            st.session_state.image_df = None

        st.session_state.run_scan = False
        st.session_state.scan_completed = True
        st.session_state.redact_completed = False

st.markdown("---")
st.markdown('<div class="section-title">Actions</div>', unsafe_allow_html=True)
st.markdown('<div class="caption">Select PII types to redact and download cleaned files (Excel, DOCX, PDF supported).</div>', unsafe_allow_html=True)

# -----------------------------
# Helpers — download & redact
# -----------------------------
def download_single_sheet_report(df: pd.DataFrame, label: str = "Download PII Report", filename: str = "pii_report.xlsx"):
    if df is None or df.empty:
        return
    buf = BytesIO()
    df.to_excel(buf, index=False)
    buf.seek(0)
    st.download_button(
        label=label,
        data=buf,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# Helper to make Excel-safe sheet names
EXCEL_INVALID_SHEET_CHARS_RE = re.compile(r'[\\/*?:\[\]]')
MAX_SHEETNAME_LEN = 31

def _sanitize_sheet_name(name: str, existing: set | None = None, default: str = "Sheet1") -> str:
    """
    Make an Excel-safe sheet name and ensure uniqueness within `existing`.
    - Replace invalid chars: \ / * ? : [ ]
    - Remove control chars
    - Strip surrounding single quotes
    - Collapse whitespace
    - Truncate to 31 chars
    - De-duplicate with suffixes: " (1)", " (2)", ...
    """
    if not isinstance(name, str) or not name.strip():
        name = default

    # Replace invalid characters with hyphen
    name = EXCEL_INVALID_SHEET_CHARS_RE.sub('-', name)

    # Remove ASCII control characters
    name = re.sub(r'[\x00-\x1F]', '', name)

    # Trim spaces and leading/trailing single quote
    name = name.strip().strip("'")

    # Collapse multiple spaces
    name = re.sub(r'\s+', ' ', name).strip()

    # Fallback if empty after cleaning
    if not name:
        name = default

    # Truncate to Excel's limit
    base = name[:MAX_SHEETNAME_LEN]

    # If uniqueness tracking not provided, return base
    if existing is None:
        return base or default

    # Ensure unique within workbook
    candidate = base or default
    i = 1
    while candidate in existing:
        suffix = f" ({i})"
        # Ensure length remains <= 31 with suffix
        cut = MAX_SHEETNAME_LEN - len(suffix)
        candidate = (base[:cut] if len(base) > cut else base) + suffix
        i += 1

    existing.add(candidate)
    return candidate

def download_two_sheet_report(pi_df: pd.DataFrame, image_df: pd.DataFrame) -> bool:
    if (pi_df is None or pi_df.empty) and (image_df is None or image_df.empty):
        return False

    excel_bytes = io.BytesIO()
    engine = _preferred_excel_engine()

    with pd.ExcelWriter(excel_bytes, engine=engine) as writer:
        existing_names: set[str] = set()
        sheets = [("Text PII", pi_df), ("Image/OCR PII", image_df)]

        for raw_name, df in sheets:
            safe_name = _sanitize_sheet_name(raw_name, existing=existing_names)

            # Ensure we always have headers, even if empty
            if df is None or df.empty:
                df_to_write = pd.DataFrame(
                    columns=["File", "Page/Sheet", "Cell", "PI Type", "Detected Value", "Confidence (%)", "Occurrence"]
                )
            else:
                df_to_write = df

            # Write sheet (no per-cell mutation afterwards)
            df_to_write.to_excel(writer, index=False, sheet_name=safe_name)

            # Optional text formatting if xlsxwriter exists (column-wise, not per cell)
            if engine == "xlsxwriter":
                wb = writer.book
                ws = writer.sheets[safe_name]
                text_fmt = wb.add_format({"num_format": "@"})
                if len(df_to_write.columns) > 0:
                    ws.set_column(0, len(df_to_write.columns) - 1, 20, text_fmt)

    # Important: seek to start and pass the buffer itself
    excel_bytes.seek(0)

    # Styled download button (see Section B for CSS)
    st.download_button(
        label="Download PII results (Excel)",
        data=excel_bytes,  # <- pass the buffer object (more reliable across versions)
        file_name="pii_results.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        #use_container_width=True,
        type="primary",     # we’ll re-apply the yellow styling below
        key="dl_two_sheet_report"  # stable key avoids re-run confusion
    )
    return True

def redact_section(df_for_redaction: pd.DataFrame):
    """Render UI to select PI Types and perform Excel redaction."""
    if df_for_redaction is None or df_for_redaction.empty:
        return

    st.info("Redaction replaces matched text with [REDACTED] in Excel and DOCX files, and draws black rectangles over text in PDF files. Images embedded in documents are not redacted.")
    distinct_types = sorted(df_for_redaction['PI Type'].dropna().unique().tolist())

    st.markdown("<div id='redact-anchor'></div>", unsafe_allow_html=True)
    selected_types = st.multiselect("Select PII Types to redact:", options=distinct_types)

    redact_clicked = st.button(
        "Redact Files",
        type="primary",
        disabled=(df_for_redaction is None or st.session_state.redact_completed or not selected_types),
    )
    if redact_clicked:
        mapping = load_pi_mapping(df_for_redaction, selected_types)
        st.session_state.redacted_files = process_source_files(st.session_state.source_files, mapping)
        st.session_state.redact_completed = True
        st.success("Files redacted successfully.")

    if st.session_state.redact_completed and st.session_state.redacted_files:
        st.subheader("Download Redacted Files")
        for fname, filebytes in st.session_state.redacted_files.items():
            filebytes.seek(0)
            ext = fname.lower().split(".")[-1]
            if ext == "pdf":
                mime = "application/pdf"
            elif ext == "docx":
                mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            else:
                mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            st.download_button(
                label=f"Download redacted {fname}",
                data=filebytes,
                file_name=f"redacted_{fname}",
                mime=mime,
            )

# -----------------------------
# Results area
# -----------------------------
if st.session_state.scan_completed:
    pi_df = st.session_state.get("pi_df")
    image_df = st.session_state.get("image_df")

    has_text = isinstance(pi_df, pd.DataFrame) and not pi_df.empty
    has_image = isinstance(image_df, pd.DataFrame) and not image_df.empty

    if has_text and not has_image:
        st.subheader("Detected PII (Text)")
        st.dataframe(pi_df, use_container_width=True)
        download_single_sheet_report(pi_df)
        redact_section(pi_df)

    elif has_image and not has_text:
        st.subheader("Detected PII (Images / OCR)")
        st.dataframe(image_df, use_container_width=True)
        download_single_sheet_report(image_df, label="Download PII Report (OCR)", filename="pii_report_ocr.xlsx")
        st.info("Redaction is available only for text content in Excel files.")

    elif has_text and has_image:
        st.subheader("Detected PII (Combined)")
        combined = pd.concat([pi_df, image_df], ignore_index=True)
        st.dataframe(combined, use_container_width=True)
        download_two_sheet_report(pi_df, image_df)
        # Redact only text-based PII from Excel files
        redact_section(pi_df)

    else:
        st.info("No PII detected in the uploaded files yet. Try different documents.")

else:
    st.caption("Upload files in the sidebar and click **Scan** to begin.")