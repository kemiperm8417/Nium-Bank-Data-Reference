"""
Nium design-system styling for the Streamlit app.

Source of truth: ~/Documents/vscode/nium-design-system (tokens.css, DESIGN.md).
`.streamlit/config.toml` carries what Streamlit's theme engine can express
(colors, font, radius); this module injects the rest — the token variables,
the type scale, pill-styled chips, card-styled panels and button anatomy.
Light presentation only: Streamlit applies a custom theme regardless of the
OS setting, so a partial dark mode would mismatch its own widgets.
"""

import streamlit as st

# Light tokens, verbatim from tokens.css.
TOKENS = """
:root {
  --bg-page:#f5f5f7; --bg-surface:#ffffff; --bg-surface-2:#fbfbfd;
  --line:rgba(0,0,0,0.08); --line-strong:rgba(0,0,0,0.14);
  --text-1:#1d1d1f; --text-2:#424245; --text-3:#6e6e73; --text-4:#86868b;
  --accent:#7461D4; --accent-2:#24BAD6; --accent-3:#E43E6D;
  --positive:#2BB673; --negative:#E43E6D;
  --shadow-sm:0 1px 2px rgba(11,14,88,0.04),0 1px 1px rgba(11,14,88,0.03);
  --shadow-md:0 4px 16px rgba(11,14,88,0.06),0 1px 2px rgba(11,14,88,0.04);
  --radius-sm:8px; --radius-md:12px; --radius-lg:16px;
  --ease:cubic-bezier(0.25,0.1,0.25,1); --dur-fast:100ms; --dur:200ms;
}
"""

CSS = TOKENS + """
/* Page anatomy — dense, 1200px working width */
[data-testid="stMainBlockContainer"] { max-width: 1200px; padding-top: 2.2rem; }
.stApp { font-feature-settings: "ss01", "cv11"; -webkit-font-smoothing: antialiased; }

/* Type scale (DESIGN.md §2) */
.stApp h1 { font-size: 32px; font-weight: 600; letter-spacing: -0.025em; color: var(--text-1); margin-bottom: 4px; }
.stApp h2 { font-size: 20px; font-weight: 600; letter-spacing: -0.02em; color: var(--text-1); margin-top: 1.4rem; }
.stApp h3 { font-size: 15px; font-weight: 600; letter-spacing: -0.015em; color: var(--text-1); }
[data-testid="stCaptionContainer"] { color: var(--text-3); font-size: 13px; line-height: 1.5; }
.stApp .lede { color: var(--text-3); font-size: 15px; line-height: 1.55; margin: 0 0 18px; }

/* Numbers: tabular figures everywhere data appears */
[data-testid="stDataFrame"], [data-testid="stMetric"], .num {
  font-variant-numeric: tabular-nums; font-feature-settings: "tnum"; letter-spacing: -0.02em;
}

/* Cards — panels that hold content get a soft card (DESIGN.md §3–4) */
[data-testid="stExpander"], [data-testid="stStatusWidget"],
[data-testid="stAlertContainer"], [data-testid="stDataFrame"] {
  background: var(--bg-surface); border: 1px solid var(--line);
  border-radius: var(--radius-lg); box-shadow: var(--shadow-sm);
}
[data-testid="stAlertContainer"] { padding: 2px 4px; }
[data-testid="stExpander"] details { border: none; }

/* Buttons — .btn anatomy: secondary is quiet, primary is near-black, never purple surfaces */
.stButton > button, .stDownloadButton > button {
  border: 1px solid var(--line-strong); background: var(--bg-surface); color: var(--text-1);
  font-weight: 500; font-size: 13px; padding: 7px 14px; border-radius: var(--radius-md);
  box-shadow: var(--shadow-sm); transition: background var(--dur-fast) var(--ease), transform var(--dur-fast) var(--ease);
}
.stButton > button:hover, .stDownloadButton > button:hover {
  background: color-mix(in oklab, var(--text-1) 5%, var(--bg-surface)); transform: translateY(-1px);
  border-color: var(--line-strong); color: var(--text-1);
}
.stButton > button[kind="primary"], .stDownloadButton > button[kind="primary"] {
  background: var(--text-1); color: var(--bg-surface); border-color: transparent;
}
.stButton > button[kind="primary"]:hover, .stDownloadButton > button[kind="primary"]:hover {
  background: #000; color: var(--bg-surface);
}
.stButton > button:disabled { opacity: 0.55; transform: none; }

/* Multiselect chips → .pill.accent */
[data-baseweb="tag"] {
  background: color-mix(in oklab, var(--accent) 10%, transparent) !important;
  border: 1px solid color-mix(in oklab, var(--accent) 22%, transparent) !important;
  color: var(--accent) !important; border-radius: 999px !important;
  font-size: 12px; font-weight: 500;
}
[data-baseweb="tag"] span, [data-baseweb="tag"] svg { color: var(--accent) !important; fill: var(--accent) !important; }

/* Inputs */
[data-baseweb="select"] > div, [data-baseweb="input"] > div {
  background: var(--bg-surface); border-color: var(--line-strong); border-radius: var(--radius-md);
}

/* Progress bar → accent */
[data-testid="stProgress"] > div > div > div > div { background: var(--accent); }

/* Browser-build download link → .btn.accent */
a.nium-btn {
  display: inline-flex; align-items: center; gap: 8px; padding: 8px 16px;
  border-radius: var(--radius-md); background: var(--accent); color: #fff !important;
  font-weight: 600; font-size: 13px; text-decoration: none !important; box-shadow: var(--shadow-sm);
  transition: transform var(--dur-fast) var(--ease), filter var(--dur-fast) var(--ease);
}
a.nium-btn:hover { transform: translateY(-1px); filter: brightness(1.05); }

/* Status pills used in the summary table caption */
.pill { display:inline-flex; align-items:center; gap:6px; padding:3px 9px; border-radius:999px;
  font-size:11.5px; font-weight:500; border:1px solid var(--line); background:var(--bg-surface-2); color:var(--text-2); }
.pill.positive { color: var(--positive); background: color-mix(in oklab, var(--positive) 10%, transparent); border-color: color-mix(in oklab, var(--positive) 22%, transparent); }
.pill.negative { color: var(--negative); background: color-mix(in oklab, var(--negative) 10%, transparent); border-color: color-mix(in oklab, var(--negative) 22%, transparent); }

/* Hide Streamlit chrome that fights the design */
#MainMenu, footer { visibility: hidden; }
"""


def apply() -> None:
    """Inject the design-system CSS. Call once, right after set_page_config."""
    st.markdown("<style>%s</style>" % CSS, unsafe_allow_html=True)
