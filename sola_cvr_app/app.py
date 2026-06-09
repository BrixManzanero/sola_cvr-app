"""
app.py
SOLA CVR Dashboard — entry point.

Run with:
    streamlit run app.py

This file only handles page setup and wiring.
All UI logic lives in frontend/  and all data logic in backend/.
"""

import warnings
import streamlit as st

warnings.filterwarnings("ignore")

# ── Auto-install xlrd for Lazada .xls support ────────────────────────────────
try:
    import xlrd  # noqa: F401
except ImportError:
    import subprocess, sys
    subprocess.check_call([sys.executable, "-m", "pip", "install", "xlrd", "-q"])

# ── Page config (must be first Streamlit call) ────────────────────────────────
st.set_page_config(
    page_title="SOLA CVR Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ── Inject global CSS ─────────────────────────────────────────────────────────
from frontend.styles import APP_CSS
st.markdown(APP_CSS, unsafe_allow_html=True)

# ── Build tab bar ─────────────────────────────────────────────────────────────
from frontend.config import PLATFORMS
from frontend.components import render_header, render_tab

render_header()

tab_labels = [f"{v['tab_icon']}  {v['label']}" for v in PLATFORMS.values()]
tabs       = st.tabs(tab_labels)

for (platform, _), tab in zip(PLATFORMS.items(), tabs):
    render_tab(tab, platform)
