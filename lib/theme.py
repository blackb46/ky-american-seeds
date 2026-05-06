"""KAS brand theme constants and CSS injection."""
from __future__ import annotations
import streamlit as st
from pathlib import Path

# KAS palette: deep agricultural green + warm gold + neutral cream.
KAS_GREEN = "#388E3C"
KAS_GREEN_LIGHT = "#66BB6A"
KAS_GREEN_DARK = "#2E7D32"
KAS_GOLD = "#C9A227"
KAS_GOLD_LIGHT = "#E8C547"
KAS_CREAM = "#FAFAF7"
KAS_TAN = "#F1EDE0"
KAS_INK = "#1A1A1A"
KAS_RED = "#B00020"

PLOTLY_COLORWAY = [
    KAS_GREEN, KAS_GOLD, KAS_GREEN_LIGHT, KAS_GOLD_LIGHT,
    "#5D4037", "#0277BD", "#6A1B9A", "#EF6C00",
]


def inject_css() -> None:
    st.markdown(f"""
    <style>
    :root {{
      --kas-green: {KAS_GREEN};
      --kas-green-dark: {KAS_GREEN_DARK};
      --kas-gold: {KAS_GOLD};
      --kas-cream: {KAS_CREAM};
      --kas-tan: {KAS_TAN};
      --kas-ink: {KAS_INK};
    }}

    .stApp {{ background: var(--kas-cream); }}

    /* Header bar with logo */
    .kas-header {{
      display: flex; align-items: center; gap: 18px;
      padding: 14px 18px; margin: -1rem -1rem 1rem -1rem;
      background: linear-gradient(90deg, var(--kas-green) 0%, var(--kas-green-dark) 100%);
      color: white;
      border-bottom: 4px solid var(--kas-gold);
      box-shadow: 0 2px 6px rgba(0,0,0,0.08);
    }}
    .kas-header h1 {{ color: white; margin: 0; font-size: 1.4rem; font-weight: 600;
                      letter-spacing: 0.3px; }}
    .kas-header .kas-sub {{ color: var(--kas-gold); font-size: 0.85rem; font-weight: 500; }}

    /* KPI cards */
    .kpi-card {{
      background: white; border-radius: 10px; padding: 18px 22px;
      box-shadow: 0 1px 4px rgba(0,0,0,0.06); border-left: 4px solid var(--kas-green);
      height: 100%;
    }}
    .kpi-card .label {{ color: #6b6b6b; font-size: 0.7rem; text-transform: uppercase;
                         letter-spacing: 0.6px; font-weight: 600;
                         white-space: nowrap; overflow: hidden;
                         text-overflow: ellipsis; }}
    .kpi-card .value {{ color: var(--kas-ink); font-size: 1.35rem; font-weight: 700;
                         margin-top: 6px; white-space: nowrap;
                         overflow: hidden; text-overflow: ellipsis;
                         font-variant-numeric: tabular-nums; }}
    .kpi-card .delta {{ color: #6b6b6b; font-size: 0.75rem; margin-top: 4px; }}
    .kpi-card.gold {{ border-left-color: var(--kas-gold); }}

    /* Confidence pills */
    .conf-high {{ color: #1B5E20; }}
    .conf-medium {{ color: #C9A227; }}
    .conf-low {{ color: #B00020; font-weight: 600; }}

    /* Issue rows */
    .issue-error {{ background: #FDECEA; border-left: 3px solid #B00020;
                    padding: 6px 10px; margin: 4px 0; border-radius: 4px; }}
    .issue-warning {{ background: #FFF8E1; border-left: 3px solid #C9A227;
                      padding: 6px 10px; margin: 4px 0; border-radius: 4px; }}
    .issue-info {{ background: #E3F2FD; border-left: 3px solid #0277BD;
                   padding: 6px 10px; margin: 4px 0; border-radius: 4px; }}

    /* Tabs polish */
    button[data-baseweb="tab"] {{ font-weight: 600; }}
    button[data-baseweb="tab"][aria-selected="true"] {{
      color: var(--kas-green) !important; border-bottom-color: var(--kas-gold) !important;
    }}

    /* Buttons */
    .stButton button[kind="primary"] {{
      background: var(--kas-green); border: none;
    }}
    .stButton button[kind="primary"]:hover {{ background: var(--kas-green-dark); }}

    /* Sidebar */
    section[data-testid="stSidebar"] {{ background: var(--kas-tan); }}

    /* Radio-as-tabs (used for top-level navigation) */
    div[role="radiogroup"][aria-label="Section"] {{
        gap: 8px;
        margin-bottom: 6px;
    }}
    div[role="radiogroup"][aria-label="Section"] > label {{
        background: white;
        border: 1px solid #d8d8d3;
        border-radius: 8px;
        padding: 8px 18px;
        margin: 0 !important;
        cursor: pointer;
        transition: background 0.15s, border-color 0.15s;
        font-weight: 600;
    }}
    div[role="radiogroup"][aria-label="Section"] > label:hover {{
        background: var(--kas-tan);
    }}
    div[role="radiogroup"][aria-label="Section"] > label > div:first-child {{
        display: none !important;     /* hide the radio circle */
    }}
    div[role="radiogroup"][aria-label="Section"] > label[data-checked="true"] {{
        background: var(--kas-green);
        border-color: var(--kas-green);
        color: white !important;
    }}
    div[role="radiogroup"][aria-label="Section"] > label[data-checked="true"] * {{
        color: white !important;
    }}
    /* Streamlit's older radio markup variant */
    div[role="radiogroup"][aria-label="Section"] > label:has(input:checked) {{
        background: var(--kas-green);
        border-color: var(--kas-green);
        color: white !important;
    }}
    div[role="radiogroup"][aria-label="Section"] > label:has(input:checked) * {{
        color: white !important;
    }}

    /* Test mode banner */
    .test-mode-banner {{
      background: #FFF8E1; border: 2px dashed var(--kas-gold);
      color: #6D4C00; padding: 10px 14px; border-radius: 8px; margin-bottom: 14px;
      font-weight: 600; text-align: center;
    }}
    </style>
    """, unsafe_allow_html=True)


def header(logo_path: str | Path = "logo.png", title: str = "Kentucky American Seeds — Transaction Manager") -> None:
    import base64
    p = Path(logo_path)
    img_html = ""
    if p.exists():
        b64 = base64.b64encode(p.read_bytes()).decode()
        img_html = f'<img src="data:image/png;base64,{b64}" style="height:84px;">'
    st.markdown(f"""
    <div class="kas-header">
        {img_html}
        <h1>{title}</h1>
    </div>
    """, unsafe_allow_html=True)


def kpi(col, label: str, value: str, delta: str | None = None, gold: bool = False) -> None:
    cls = "kpi-card gold" if gold else "kpi-card"
    delta_html = f'<div class="delta">{delta}</div>' if delta else ""
    col.markdown(
        f'<div class="{cls}"><div class="label">{label}</div>'
        f'<div class="value">{value}</div>{delta_html}</div>',
        unsafe_allow_html=True,
    )


def confidence_pill(level: str | None) -> str:
    if not level:
        return ""
    cls = {"high": "conf-high", "medium": "conf-medium", "low": "conf-low"}.get(level.lower(), "")
    icon = {"high": "🟢", "medium": "🟡", "low": "🔴"}.get(level.lower(), "⚪")
    return f'<span class="{cls}">{icon} {level}</span>'
