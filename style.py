"""CSS + small reusable UI helpers, kept out of app.py to keep pages tidy."""
import pandas as pd
import streamlit as st

# ---- palette -----------------------------------------------------------
GOOD = "#1DB874"    # updated / running / success
WARN = "#F5A623"    # pending / waiting / medium
BAD = "#E5484D"     # failed / not updated / high
NEUTRAL = "#8B90A0"  # disabled / unknown
INFO = "#4F7CFE"

STATUS_COLORS = {
    "Updated": GOOD, "Updated (Early)": GOOD, "Running": GOOD, "Completed": GOOD,
    "QMC Running": WARN, "Waiting": WARN, "Pending": WARN, "Dashboard Pending Refresh": WARN,
    "Updated (Late)": WARN, "Not Specified": NEUTRAL,
    "Not Updated": BAD, "Not updated": BAD, "Failed": BAD, "Error": BAD,
    "Disabled": NEUTRAL, "Unknown": NEUTRAL, "N/A": NEUTRAL, "—": NEUTRAL,
}
SEVERITY_COLORS = {"High": BAD, "Medium": WARN, "Low": INFO, "Unknown": NEUTRAL}


def color_for(value: str) -> str:
    return STATUS_COLORS.get(str(value), SEVERITY_COLORS.get(str(value), INFO))


def inject_css():
    st.markdown(f"""
    <style>
    .block-container {{ padding-top: 1.6rem; max-width: 1400px; }}
    #MainMenu, footer {{visibility: hidden;}}

    [data-testid="stMetric"] {{
        background: #F4F6FB; border: 1px solid #E7E9F3; border-radius: 12px;
        padding: 14px 16px 10px 16px;
    }}
    [data-testid="stMetricLabel"] {{ font-size: 0.82rem; color: #5B5F73; }}
    [data-testid="stMetricValue"] {{ font-size: 1.55rem; }}

    .section-card {{
        background: #FFFFFF; border: 1px solid #E7E9F3; border-radius: 14px;
        padding: 18px 20px; margin-bottom: 14px;
    }}
    .badge {{
        display:inline-block; padding: 3px 11px; border-radius: 999px;
        font-size: 0.78rem; font-weight: 600; color: white; white-space: nowrap;
    }}
    .pill-row span {{ margin-right: 6px; }}
    .app-title {{ font-size: 1.9rem; font-weight: 800; margin-bottom: 0; }}
    .app-sub {{ color:#6B7080; margin-top:-6px; margin-bottom: 1rem;}}
    .stTabs [data-baseweb="tab-list"] {{ gap: 4px; }}
    .stTabs [data-baseweb="tab"] {{
        background: #F4F6FB; border-radius: 10px 10px 0 0; padding: 8px 16px; font-weight: 600;
    }}
    .stTabs [aria-selected="true"] {{ background: {INFO}; color: white !important; }}
    </style>
    """, unsafe_allow_html=True)


def badge(text) -> str:
    if text is None or (isinstance(text, float) and pd.isna(text)):
        return ""
    return f'<span class="badge" style="background:{color_for(text)}">{text}</span>'


def badge_html_col(df: pd.DataFrame, col: str) -> pd.Series:
    return df[col].map(badge)


def render_pills(counter: dict, order: list[str] | None = None):
    keys = order or list(counter.keys())
    html = '<div class="pill-row">' + "".join(
        f'{badge(k)} <b>{counter[k]}</b>&nbsp;&nbsp;' for k in keys if k in counter
    ) + "</div>"
    st.markdown(html, unsafe_allow_html=True)


def fmt_dt(v, fmt="%d %b, %I:%M %p") -> str:
    ts = pd.to_datetime(v, errors="coerce")
    return "—" if pd.isna(ts) else ts.strftime(fmt).replace(" 0", " ")


def empty_state(msg: str):
    st.info(msg, icon="🗂️")
