
from __future__ import annotations
import textwrap

import io
import shutil
from pathlib import Path
from datetime import datetime, date

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import plotly.io as pio
import streamlit as st

APP_DIR = Path(__file__).resolve().parent
DEFAULT_FILE = APP_DIR / "Chandra.xlsx"

MASTER_SALES = APP_DIR / "MASTER_SALES.csv"
MASTER_SERVICE = APP_DIR / "MASTER_SERVICE.csv"
IMPORT_LOG = APP_DIR / "IMPORT_LOG.csv"
SMP_FILE = APP_DIR / "SMP_DATA.xlsx"

def save_master_csv(df: pd.DataFrame, path: Path):
    df.to_csv(path, index=False, encoding="utf-8-sig")

@st.cache_data(show_spinner=False)
def read_master_csv_cached(path_str: str, modified_time: float):
    """Read a master CSV only when the file itself changes."""
    return pd.read_csv(path_str, low_memory=False)

def read_master_csv(path: Path):
    return read_master_csv_cached(str(path), path.stat().st_mtime)

def read_monthly_file(uploaded_file, expected_kind: str):
    """Read xlsx/xls/csv. For Excel, auto-detect a likely sales/service sheet."""
    name = uploaded_file.name.lower()
    data = uploaded_file.getvalue()

    if name.endswith(".csv"):
        return pd.read_csv(io.BytesIO(data), low_memory=False)

    if not (name.endswith(".xlsx") or name.endswith(".xls")):
        raise ValueError("Please upload Excel (.xlsx/.xls) or CSV.")

    xls = pd.ExcelFile(io.BytesIO(data), engine="openpyxl")
    sheets = xls.sheet_names

    # Prefer recognizable raw-data sheets.
    if expected_kind == "sales":
        preferred = ["SOLD_DATA", "SALES_DATA", "SALE_DATA", "SOLD DATA", "SALES DATA", "SALE DATA"]
        key_headers = {"VIN", "Sold Date"}
    else:
        preferred = ["SERVICE_DATA", "SERVICES_DATA", "SERVICE DATA", "SERVICES DATA"]
        key_headers = {"VIN", "Service Date"}

    for p in preferred:
        if p in sheets:
            # Try both row 1 and row 2 as header because the current workbook uses row 2.
            for header in [1, 0]:
                df = pd.read_excel(io.BytesIO(data), sheet_name=p, header=header, engine="openpyxl")
                cols = {str(c).strip() for c in df.columns}
                if key_headers.issubset(cols):
                    return df

    # Auto-detect sheet/header based on required columns.
    for sh in sheets:
        for header in [0, 1, 2]:
            try:
                df = pd.read_excel(io.BytesIO(data), sheet_name=sh, header=header, engine="openpyxl")
                cols = {str(c).strip() for c in df.columns}
                if key_headers.issubset(cols):
                    return df
            except Exception:
                pass

    raise ValueError(
        f"Could not find the required columns {sorted(key_headers)} in {uploaded_file.name}."
    )


def save_default_full_workbook(sales_df: pd.DataFrame, service_df: pd.DataFrame):
    """Keep Chandra.xlsx synchronized with the active full master data."""
    backup_dir = APP_DIR / "backups"
    backup_dir.mkdir(exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    if DEFAULT_FILE.exists():
        shutil.copy2(DEFAULT_FILE, backup_dir / f"Chandra_before_full_replace_{stamp}.xlsx")

    with pd.ExcelWriter(DEFAULT_FILE, engine="openpyxl") as writer:
        sales_df.to_excel(writer, sheet_name="SOLD_DATA", index=False, startrow=1)
        service_df.to_excel(writer, sheet_name="SERVICE_DATA", index=False, startrow=1)


def align_columns(old: pd.DataFrame, new: pd.DataFrame):
    """Union columns so future files may safely add columns."""
    cols = list(dict.fromkeys(list(old.columns) + list(new.columns)))
    return old.reindex(columns=cols), new.reindex(columns=cols)

def append_monthly_data(master: pd.DataFrame, new: pd.DataFrame, kind: str):
    old, fresh = align_columns(master, new)
    before = len(old)
    combined = pd.concat([old, fresh], ignore_index=True)

    # Remove exact duplicate rows only. Do not collapse legitimate multiple service lines/ROs.
    combined = combined.drop_duplicates().reset_index(drop=True)
    after = len(combined)
    exact_dupes_removed = before + len(fresh) - after

    # Also normalize obviously empty rows.
    if "VIN" in combined.columns:
        vin = combined["VIN"].fillna("").astype(str).str.strip()
        combined = combined[vin.ne("")].copy()

    return combined.reset_index(drop=True), exact_dupes_removed

def seed_master_from_chandra():
    if MASTER_SALES.exists() and MASTER_SERVICE.exists():
        return
    if not DEFAULT_FILE.exists():
        return
    sold_seed = pd.read_excel(DEFAULT_FILE, sheet_name="SOLD_DATA", header=1, engine="openpyxl")
    service_seed = pd.read_excel(DEFAULT_FILE, sheet_name="SERVICE_DATA", header=1, engine="openpyxl")
    if not MASTER_SALES.exists():
        save_master_csv(sold_seed, MASTER_SALES)
    if not MASTER_SERVICE.exists():
        save_master_csv(service_seed, MASTER_SERVICE)

seed_master_from_chandra()

NAVY = "#062D5D"
NAVY2 = "#062D5D"
BLUE = "#2F75B5"
GREEN = "#079447"
ORANGE = "#F28C28"
RED = "#D94A3A"
TEAL = "#19A6B3"
PURPLE = "#6447A5"
LIGHT = "#F4F7FB"
GRID = "#DDE5EF"
TEXT = "#152238"

APP_FONT = "Source Sans Pro"

def retention_color(v):
    p = float(v) * 100 if float(v) <= 1 else float(v)
    if p >= 80:
        return GREEN
    if p >= 50:
        return ORANGE
    return RED

def retention_band(v):
    p = float(v) * 100 if float(v) <= 1 else float(v)
    if p >= 80:
        return "Green — Strong retention"
    if p >= 50:
        return "Orange — Needs attention"
    return "Red — Critical retention"


st.set_page_config(
    page_title="Mody Skoda South — Sold Vehicle Cohort Retention Dashboard",
    page_icon="🚘",
    layout="wide",
    initial_sidebar_state="expanded",
)


# Global chart typography/theme — applied to every Plotly chart in the app.
pio.templates["mody_skoda"] = go.layout.Template(
    layout=go.Layout(
        font=dict(family=APP_FONT, color="#152238", size=12),
        title=dict(font=dict(family=APP_FONT, color="#2F75B5", size=15)),
        paper_bgcolor="#FFFFFF",
        plot_bgcolor="#FFFFFF",
        legend=dict(font=dict(family=APP_FONT, color="#39495E")),
        hoverlabel=dict(
            font=dict(family=APP_FONT, color="#FFFFFF"),
            bgcolor="#062D5D",
            bordercolor="#B6C6D8",
        ),
        xaxis=dict(
            tickfont=dict(family=APP_FONT, color="#39495E"),
            title=dict(font=dict(family=APP_FONT, color="#39495E")),
            gridcolor="#DDE5EF",
            linecolor="#B6C6D8",
        ),
        yaxis=dict(
            tickfont=dict(family=APP_FONT, color="#39495E"),
            title=dict(font=dict(family=APP_FONT, color="#39495E")),
            gridcolor="#DDE5EF",
            linecolor="#B6C6D8",
        ),
    )
)
pio.templates.default = "plotly+mody_skoda"


st.markdown(
    f"""
    <style>
      .stApp {{background: {LIGHT};}}
      .block-container {{padding-top: 0.7rem; padding-bottom: 2rem; max-width: 1800px;}}
      .main-title {{
          background:{NAVY}; color:white; padding:16px 22px; border:1px solid #2F75B5; border-radius:8px;
          font-size:28px; font-weight:800; margin-bottom:8px;
          box-shadow:0 2px 8px rgba(0,0,0,.12);
      }}
      .subtitle {{color:#506176; margin:2px 0 12px 2px; font-size:14px;}}
      .kpi {{
          background:#FFFFFF; border:1px solid #B6C6D8; border-radius:10px;
          padding:14px 14px 12px 14px; min-height:105px;
          box-shadow:0 2px 8px rgba(11,61,115,.07);
      }}
      .kpi-label {{font-size:13px; font-weight:700; color:#39495E;}}
      .kpi-value {{font-size:28px; font-weight:800; color:{TEXT}; line-height:1.15; margin-top:5px;}}
      .kpi-note {{font-size:11px; color:#718099; margin-top:5px;}}
      .panel-title {{
          background:{NAVY}; color:white; padding:8px 12px; border:1px solid #2F75B5; border-radius:7px 7px 0 0;
          font-size:14px; font-weight:800; text-align:center; margin-top:5px;
      }}
      div[data-testid="stMetric"] {{
          background:#FFFFFF; border:1px solid #B6C6D8; border-radius:10px; padding:10px 14px;
      }}
      div[data-testid="stDataFrame"] {{border:1px solid #B6C6D8; border-radius:8px; overflow:hidden;}}
      .small-note {{font-size:12px;color:#718099;}}
      .good {{color:{GREEN};font-weight:700}}
      .warn {{color:{ORANGE};font-weight:700}}
      .bad {{color:{RED};font-weight:700}}
      section[data-testid="stSidebar"] {{background:#EDF3FA;}}
      .stButton>button {{border-radius:7px; font-weight:700;}}
    
/* v3.6 icon KPI layout */
.kpi {{
    position: relative !important;
    display: flex !important;
    align-items: center !important;
    gap: 10px !important;
    min-height: 86px !important;
    height: 86px !important;
    padding: 9px 11px !important;
    background: #FFFFFF !important;
    border: 1px solid #B6C6D8 !important;
    border-radius: 9px !important;
    box-shadow: 0 2px 8px rgba(6,45,93,0.07) !important;
    overflow: hidden !important;
}}
.kpi-icon {{
    flex: 0 0 38px !important;
    width: 38px !important;
    text-align: center !important;
    font-size: 30px !important;
    line-height: 1 !important;
}}
.kpi-content {{
    flex: 1 1 auto !important;
    min-width: 0 !important;
}}
.kpi-label {{
    font-size: 10px !important;
    font-weight: 800 !important;
    color: #20324A !important;
    line-height: 1.08 !important;
    margin-bottom: 4px !important;
}}
.kpi-value {{
    font-size: 23px !important;
    font-weight: 800 !important;
    color: #071B3B !important;
    line-height: 1 !important;
    margin: 1px 0 4px 0 !important;
}}
.kpi-note {{
    font-size: 9px !important;
    color: #718099 !important;
    line-height: 1.05 !important;
}}


.kpi-icon svg {{
    width: 32px !important;
    height: 32px !important;
    display: block !important;
    margin: 0 auto !important;
}}
.service-category-icon svg {{
    display: block;
    margin: 0 auto;
}}


      /* v3.28 — full approved dark emerald + mint theme, top to bottom */
      .stApp {{ background:#F4F7FB !important; color:#152238 !important; }}
      .block-container {{ padding-top:3.6rem !important; }}
      header[data-testid="stHeader"] {{ background:#F4F7FB !important; }}
      section[data-testid="stSidebar"] {{ background:#EDF3FA !important; }}
      section[data-testid="stSidebar"] * {{ color:#152238 !important; }}

      html, body, [class*="css"], .stApp, button, input, select, textarea,
      [data-testid="stMarkdownContainer"], [data-testid="stWidgetLabel"] {{
          font-family:"Segoe UI","Arial",sans-serif !important;
      }}
      h1,h2,h3,h4,h5,h6 {{ color:#2F75B5 !important; font-weight:750 !important; }}

      .main-title {{
          background:#062D5D !important;
          color:#2F75B5 !important;
          border:1px solid #2F75B5 !important;
          font-size:28px !important;
          font-weight:800 !important;
      }}
      .subtitle, .small-note {{ color:#506176 !important; }}
      .panel-title {{
          background:#062D5D !important;
          color:#2F75B5 !important;
          border-color:#2F75B5 !important;
      }}
      .kpi, .filter-box, .lost-card, .lost-action-box {{
          background:#FFFFFF !important;
          border-color:#B6C6D8 !important;
          color:#152238 !important;
      }}
      .kpi-label, .lost-card-label {{ color:#39495E !important; }}
      .kpi-value {{ color:#152238 !important; }}
      .kpi-note, .lost-card-note {{ color:#718099 !important; }}

      div[data-testid="stDataFrame"] {{
          border:1px solid #B6C6D8 !important;
          background:#FFFFFF !important;
      }}
      div[data-testid="stMetric"] {{
          background:#FFFFFF !important;
          border:1px solid #B6C6D8 !important;
          border-radius:9px !important;
          padding:10px 12px !important;
      }}
      div[data-testid="stMetric"] label {{ color:#39495E !important; }}
      div[data-testid="stMetricValue"] {{ color:#152238 !important; }}

      div[data-baseweb="select"] > div,
      div[data-baseweb="input"] > div,
      input, textarea {{
          background:#FFFFFF !important;
          color:#152238 !important;
          border-color:#B6C6D8 !important;
      }}
      div[data-baseweb="popover"], div[data-baseweb="menu"] {{
          background:#FFFFFF !important;
          color:#152238 !important;
      }}
      button {{
          border-color:#B6C6D8 !important;
      }}
      div[data-testid="stTabs"] button {{
          color:#39495E !important;
          font-weight:600 !important;
      }}
      div[data-testid="stTabs"] button[aria-selected="true"] {{
          color:#2F75B5 !important;
          border-bottom-color:#2F75B5 !important;
      }}
      hr {{ border-color:#DDE5EF !important; }}


      /* v3.29 — fast tab-style page selector */
      div[role="radiogroup"] {{
          gap: 0.25rem !important;
          flex-wrap: wrap !important;
      }}
      div[role="radiogroup"] label {{
          background:#FFFFFF !important;
          border:1px solid #DDE5EF !important;
          border-radius:7px !important;
          padding:0.25rem 0.45rem !important;
      }}
      div[role="radiogroup"] label:has(input:checked) {{
          border-color:#2F75B5 !important;
          background:#062D5D !important;
      }}


      /* v3.31 — force the approved font consistently from top to bottom */
      .stApp, .stApp *,
      section[data-testid="stSidebar"], section[data-testid="stSidebar"] *,
      [data-testid="stMarkdownContainer"], [data-testid="stMarkdownContainer"] *,
      [data-testid="stWidgetLabel"], [data-testid="stWidgetLabel"] *,
      [data-testid="stMetric"], [data-testid="stMetric"] *,
      [data-testid="stDataFrame"], [data-testid="stDataFrame"] *,
      [data-baseweb="select"], [data-baseweb="select"] *,
      [data-baseweb="input"], [data-baseweb="input"] *,
      [data-baseweb="popover"], [data-baseweb="popover"] *,
      [data-baseweb="menu"], [data-baseweb="menu"] *,
      button, input, textarea, select, option, label, p, span, div, th, td,
      h1, h2, h3, h4, h5, h6 {{
          font-family: "Segoe UI", Arial, sans-serif !important;
      }}

      /* Plotly uses SVG text; force the same family there too. */
      .js-plotly-plot text,
      .plotly text,
      svg text {{
          font-family: "Segoe UI", Arial, sans-serif !important;
      }}

      /* Approved typography hierarchy */
      .main-title {{
          font-family: "Segoe UI", Arial, sans-serif !important;
          font-weight: 800 !important;
          letter-spacing: .15px !important;
          line-height: 1.25 !important;
      }}
      .panel-title, h1, h2, h3 {{
          font-family: "Segoe UI", Arial, sans-serif !important;
          font-weight: 700 !important;
      }}
      .kpi-label, [data-testid="stMetricLabel"] {{
          font-weight: 700 !important;
      }}
      .kpi-value, [data-testid="stMetricValue"] {{
          font-weight: 700 !important;
      }}


      /* v3.33 — visibility/accessibility override */
      div[role="radiogroup"] label {{
          background:#FFFFFF !important;
          border:1px solid #B6C6D8 !important;
      }}
      div[role="radiogroup"] label,
      div[role="radiogroup"] label *,
      div[role="radiogroup"] p,
      div[role="radiogroup"] span {{
          color:#152238 !important;
          opacity:1 !important;
          -webkit-text-fill-color:#152238 !important;
          font-weight:600 !important;
      }}
      div[role="radiogroup"] label:has(input:checked) {{
          background:#062D5D !important;
          border:2px solid #2F75B5 !important;
      }}
      div[role="radiogroup"] label:has(input:checked),
      div[role="radiogroup"] label:has(input:checked) * {{
          color:#FFFFFF !important;
          -webkit-text-fill-color:#FFFFFF !important;
      }}

      /* All Streamlit buttons, including CSV/Excel download buttons */
      .stDownloadButton button,
      [data-testid="stDownloadButton"] button,
      .stButton button,
      [data-testid="stBaseButton-secondary"],
      [data-testid="stBaseButton-primary"] {{
          background:#062D5D !important;
          color:#FFFFFF !important;
          border:1px solid #2F75B5 !important;
          opacity:1 !important;
      }}
      .stDownloadButton button *,
      [data-testid="stDownloadButton"] button *,
      .stButton button *,
      [data-testid="stBaseButton-secondary"] *,
      [data-testid="stBaseButton-primary"] * {{
          color:#FFFFFF !important;
          -webkit-text-fill-color:#FFFFFF !important;
          opacity:1 !important;
          font-weight:700 !important;
      }}
      .stDownloadButton button:hover,
      [data-testid="stDownloadButton"] button:hover,
      .stButton button:hover {{
          background:#2F75B5 !important;
          border-color:#D2E5F5 !important;
      }}

      /* File uploader and its Browse files button */
      [data-testid="stFileUploader"] {{
          background:#FFFFFF !important;
          border-radius:8px !important;
      }}
      [data-testid="stFileUploader"] *,
      [data-testid="stFileUploaderDropzone"] *,
      [data-testid="stFileUploader"] small {{
          color:#152238 !important;
          -webkit-text-fill-color:#152238 !important;
          opacity:1 !important;
      }}
      [data-testid="stFileUploader"] button {{
          background:#062D5D !important;
          color:#FFFFFF !important;
          border:1px solid #2F75B5 !important;
      }}
      [data-testid="stFileUploader"] button * {{
          color:#FFFFFF !important;
          -webkit-text-fill-color:#FFFFFF !important;
      }}

      /* Download/file labels, captions and helper text */
      [data-testid="stCaptionContainer"],
      [data-testid="stCaptionContainer"] *,
      .stDownloadButton,
      .stDownloadButton * {{
          opacity:1 !important;
      }}


      /* v3.34 — dataframe/table visibility fix */
      [data-testid="stDataFrame"],
      [data-testid="stDataFrame"] > div,
      [data-testid="stDataFrame"] iframe {{
          background:#FFFFFF !important;
          color:#152238 !important;
      }}

      /* Streamlit dataframe/grid cells */
      [data-testid="stDataFrame"] [role="grid"],
      [data-testid="stDataFrame"] [role="row"],
      [data-testid="stDataFrame"] [role="gridcell"],
      [data-testid="stDataFrame"] [role="columnheader"],
      [data-testid="stDataFrame"] div,
      [data-testid="stDataFrame"] span,
      [data-testid="stDataFrame"] p {{
          color:#152238 !important;
          -webkit-text-fill-color:#152238 !important;
          opacity:1 !important;
      }}

      [data-testid="stDataFrame"] [role="columnheader"] {{
          background:#062D5D !important;
          color:#2F75B5 !important;
          font-weight:700 !important;
      }}

      [data-testid="stDataFrame"] [role="gridcell"] {{
          background:#FFFFFF !important;
          border-color:#062D5D !important;
      }}

      /* HTML tables / pandas Styler output */
      table {{
          background:#FFFFFF !important;
          color:#152238 !important;
      }}
      table thead th {{
          background:#062D5D !important;
          color:#2F75B5 !important;
          border-color:#B6C6D8 !important;
      }}
      table tbody td {{
          background:#FFFFFF !important;
          color:#152238 !important;
          border-color:#062D5D !important;
      }}
      table tbody tr:nth-child(even) td {{
          background:#062D5D !important;
      }}

      /* Glide Data Grid / canvas-backed tables */
      .gdg-wmyidgi,
      .glideDataEditor {{
          color:#152238 !important;
      }}


      /* v3.36 — restore v3.28 font + text colour treatment */
      html, body, [class*="css"], .stApp, button, input, select, textarea,
      [data-testid="stMarkdownContainer"], [data-testid="stWidgetLabel"] {{
          font-family:"Segoe UI","Arial",sans-serif !important;
      }}

      .stApp {{ color:#152238 !important; }}
      section[data-testid="stSidebar"] * {{ color:#152238 !important; }}

      h1,h2,h3,h4,h5,h6 {{
          color:#2F75B5 !important;
          font-weight:750 !important;
      }}

      .main-title {{
          color:#2F75B5 !important;
          font-family:"Segoe UI","Arial",sans-serif !important;
          font-weight:800 !important;
      }}
      .subtitle, .small-note {{ color:#506176 !important; }}
      .panel-title {{ color:#2F75B5 !important; }}
      .kpi-label, .lost-card-label {{ color:#39495E !important; }}
      .kpi-value {{ color:#152238 !important; }}
      .kpi-note, .lost-card-note {{ color:#718099 !important; }}

      div[data-testid="stMetric"] label {{ color:#39495E !important; }}
      div[data-testid="stMetricValue"] {{ color:#152238 !important; }}

      /* v3.28-style filter text, with the v3.35 arrow visibility correction */
      [data-testid="stSelectbox"] label,
      [data-testid="stMultiSelect"] label,
      [data-testid="stSelectbox"] label *,
      [data-testid="stMultiSelect"] label * {{
          color:#152238 !important;
          -webkit-text-fill-color:#152238 !important;
          font-family:"Segoe UI","Arial",sans-serif !important;
          opacity:1 !important;
      }}

      div[data-baseweb="select"] > div {{
          background:#FFFFFF !important;
          color:#152238 !important;
          border-color:#B6C6D8 !important;
      }}
      div[data-baseweb="select"] > div > div {{
          background:transparent !important;
      }}
      div[data-baseweb="select"] span,
      div[data-baseweb="select"] input,
      div[data-baseweb="select"] p {{
          color:#152238 !important;
          -webkit-text-fill-color:#152238 !important;
          font-family:"Segoe UI","Arial",sans-serif !important;
          opacity:1 !important;
      }}
      div[data-baseweb="select"] svg {{
          color:#2F75B5 !important;
          fill:#2F75B5 !important;
          opacity:1 !important;
      }}

      [data-baseweb="tag"] {{
          background:#062D5D !important;
          border:1px solid #B6C6D8 !important;
          color:#152238 !important;
      }}
      [data-baseweb="tag"] * {{
          color:#152238 !important;
          -webkit-text-fill-color:#152238 !important;
      }}

      [data-baseweb="popover"] ul, [role="listbox"] {{
          background:#FFFFFF !important;
          color:#152238 !important;
          border:1px solid #B6C6D8 !important;
      }}
      [role="option"], [role="option"] * {{
          color:#152238 !important;
          -webkit-text-fill-color:#152238 !important;
      }}

      /* Preserve readable tables while using the same v3.28 palette */
      [data-testid="stDataFrame"] [role="columnheader"],
      [data-testid="stDataFrame"] [role="columnheader"] * {{
          color:#2F75B5 !important;
          -webkit-text-fill-color:#2F75B5 !important;
      }}
      [data-testid="stDataFrame"] [role="gridcell"],
      [data-testid="stDataFrame"] [role="gridcell"] * {{
          color:#152238 !important;
          -webkit-text-fill-color:#152238 !important;
      }}


      /* v3.37 — v3.18 FONT + COLOUR STYLE ONLY
         Functional/layout code is unchanged. */

      /* v3.18 palette:
         NAVY #062D5D | BLUE #2F75B5 | GREEN #079447
         ORANGE #F28C28 | RED #D94A3A | TEXT #152238
         LIGHT #F4F7FB | secondary text #506176/#718099 */

      html, body, .stApp, button, input, textarea, select,
      [data-testid="stMarkdownContainer"],
      [data-testid="stWidgetLabel"] {{
          font-family:"Source Sans Pro", sans-serif !important;
      }}

      .stApp {{
          background:#F4F7FB !important;
          color:#152238 !important;
      }}

      section[data-testid="stSidebar"] {{
          background:#EDF3FA !important;
      }}

      section[data-testid="stSidebar"],
      section[data-testid="stSidebar"] p,
      section[data-testid="stSidebar"] span,
      section[data-testid="stSidebar"] label {{
          color:#152238 !important;
      }}

      h1, h2, h3, h4, h5, h6 {{
          color:#062D5D !important;
          font-family:"Source Sans Pro", sans-serif !important;
      }}

      .main-title {{
          background:#062D5D !important;
          color:#FFFFFF !important;
          font-size:28px !important;
          font-weight:800 !important;
      }}

      .subtitle {{
          color:#506176 !important;
          font-size:14px !important;
      }}

      .panel-title {{
          background:#062D5D !important;
          color:#FFFFFF !important;
          font-size:14px !important;
          font-weight:800 !important;
      }}

      .kpi-label {{
          color:#20324A !important;
          font-size:10px !important;
          font-weight:800 !important;
      }}

      .kpi-value {{
          color:#071B3B !important;
          font-size:23px !important;
          font-weight:800 !important;
      }}

      .kpi-note {{
          color:#718099 !important;
          font-size:9px !important;
      }}

      .small-note {{
          color:#6C7C90 !important;
          font-size:12px !important;
      }}

      .good {{ color:#079447 !important; font-weight:700 !important; }}
      .warn {{ color:#F28C28 !important; font-weight:700 !important; }}
      .bad  {{ color:#D94A3A !important; font-weight:700 !important; }}

      div[data-testid="stMetric"] label,
      div[data-testid="stMetric"] label * {{
          color:#39495E !important;
      }}

      div[data-testid="stMetricValue"],
      div[data-testid="stMetricValue"] * {{
          color:#152238 !important;
      }}

      /* Filters: v3.18 light style + clearly visible arrow */
      [data-testid="stSelectbox"] label,
      [data-testid="stMultiSelect"] label,
      [data-testid="stSelectbox"] label *,
      [data-testid="stMultiSelect"] label * {{
          color:#152238 !important;
          -webkit-text-fill-color:#152238 !important;
          font-family:"Source Sans Pro", sans-serif !important;
          opacity:1 !important;
      }}

      div[data-baseweb="select"] > div {{
          background:#FFFFFF !important;
          color:#152238 !important;
          border-color:#DCE4EE !important;
      }}

      div[data-baseweb="select"] > div > div {{
          background:transparent !important;
      }}

      div[data-baseweb="select"] span,
      div[data-baseweb="select"] input,
      div[data-baseweb="select"] p {{
          color:#152238 !important;
          -webkit-text-fill-color:#152238 !important;
          opacity:1 !important;
      }}

      div[data-baseweb="select"] svg {{
          color:#062D5D !important;
          fill:#062D5D !important;
          opacity:1 !important;
      }}

      [data-baseweb="tag"] {{
          background:#EAF3FA !important;
          border:1px solid #B6C6D8 !important;
          color:#062D5D !important;
      }}

      [data-baseweb="tag"] *,
      [data-baseweb="tag"] svg {{
          color:#062D5D !important;
          fill:#062D5D !important;
          -webkit-text-fill-color:#062D5D !important;
      }}

      [data-baseweb="popover"] ul,
      [role="listbox"] {{
          background:#FFFFFF !important;
          color:#152238 !important;
          border:1px solid #DCE4EE !important;
      }}

      [role="option"],
      [role="option"] * {{
          color:#152238 !important;
          -webkit-text-fill-color:#152238 !important;
      }}

      [role="option"]:hover,
      [role="option"][aria-selected="true"] {{
          background:#EAF3FA !important;
          color:#062D5D !important;
      }}

      /* Navigation/buttons in the v3.18 navy/blue palette */
      div[role="radiogroup"] label {{
          color:#152238 !important;
      }}

      div[role="radiogroup"] label[data-checked="true"],
      .stButton > button,
      .stDownloadButton > button {{
          background:#062D5D !important;
          color:#FFFFFF !important;
          border-color:#2F75B5 !important;
      }}

      .stButton > button *,
      .stDownloadButton > button * {{
          color:#FFFFFF !important;
      }}

      /* Data tables: v3.18 navy headers / dark readable text */
      [data-testid="stDataFrame"] [role="columnheader"],
      [data-testid="stDataFrame"] [role="columnheader"] * {{
          background:#062D5D !important;
          color:#FFFFFF !important;
          -webkit-text-fill-color:#FFFFFF !important;
      }}

      [data-testid="stDataFrame"] [role="gridcell"],
      [data-testid="stDataFrame"] [role="gridcell"] * {{
          background:#FFFFFF !important;
          color:#152238 !important;
          -webkit-text-fill-color:#152238 !important;
      }}

      table thead th {{
          background:#062D5D !important;
          color:#FFFFFF !important;
      }}

      table tbody td {{
          background:#FFFFFF !important;
          color:#152238 !important;
      }}


      /* v3.38 — apply supplied v3.18 font + colour styling ACROSS ALL AREAS */

      html, body, .stApp, .stApp *,
      button, input, textarea, select {{
          font-family:"Source Sans Pro", sans-serif !important;
      }}

      .stApp {{
          background:#F4F7FB !important;
          color:#152238 !important;
      }}
      .block-container {{
          background:#F4F7FB !important;
      }}

      section[data-testid="stSidebar"] {{
          background:#EDF3FA !important;
      }}
      section[data-testid="stSidebar"] *,
      [data-testid="stMarkdownContainer"],
      [data-testid="stMarkdownContainer"] p,
      [data-testid="stMarkdownContainer"] span {{
          color:#152238 !important;
      }}

      h1,h2,h3,h4,h5,h6 {{
          color:#062D5D !important;
      }}

      .main-title, .panel-title, .lost-bar,
      .tp-section-bar, .section-title {{
          background:#062D5D !important;
          color:#FFFFFF !important;
      }}

      .subtitle, .small-note, .lost-sub,
      .kpi-note, .lost-card-note, .tp-kpi-note {{
          color:#718099 !important;
      }}

      .kpi, .lost-card, .tp-kpi, .tp-table-wrap,
      div[data-testid="stMetric"],
      [data-testid="stExpander"] details {{
          background:#FFFFFF !important;
          color:#152238 !important;
          border-color:#DCE4EE !important;
      }}

      .kpi-label, .lost-card-label, .tp-kpi-lbl {{
          color:#20324A !important;
      }}
      .kpi-value, .lost-card-value, .tp-kpi-val {{
          color:#071B3B !important;
      }}

      /* Navigation tabs/radio */
      div[role="radiogroup"] label {{
          background:#FFFFFF !important;
          color:#152238 !important;
          border:1px solid #DCE4EE !important;
      }}
      div[role="radiogroup"] label * {{
          color:#152238 !important;
      }}
      div[role="radiogroup"] label[data-checked="true"] {{
          background:#062D5D !important;
          border-color:#2F75B5 !important;
      }}
      div[role="radiogroup"] label[data-checked="true"] * {{
          color:#FFFFFF !important;
      }}

      /* All selectboxes and multiselects */
      [data-testid="stSelectbox"] label,
      [data-testid="stMultiSelect"] label,
      [data-testid="stDateInput"] label,
      [data-testid="stNumberInput"] label,
      [data-testid="stTextInput"] label {{
          color:#152238 !important;
      }}

      div[data-baseweb="select"] > div,
      div[data-baseweb="input"] > div,
      [data-baseweb="base-input"],
      input, textarea {{
          background:#FFFFFF !important;
          color:#152238 !important;
          border-color:#DCE4EE !important;
      }}
      div[data-baseweb="select"] > div > div {{
          background:transparent !important;
      }}
      div[data-baseweb="select"] span,
      div[data-baseweb="select"] p,
      div[data-baseweb="select"] input {{
          color:#152238 !important;
          -webkit-text-fill-color:#152238 !important;
      }}
      div[data-baseweb="select"] svg {{
          color:#062D5D !important;
          fill:#062D5D !important;
          opacity:1 !important;
      }}

      [data-baseweb="tag"] {{
          background:#EAF3FA !important;
          color:#062D5D !important;
          border:1px solid #B6C6D8 !important;
      }}
      [data-baseweb="tag"] *,
      [data-baseweb="tag"] svg {{
          color:#062D5D !important;
          fill:#062D5D !important;
          -webkit-text-fill-color:#062D5D !important;
      }}

      [data-baseweb="popover"],
      [data-baseweb="popover"] ul,
      [role="listbox"] {{
          background:#FFFFFF !important;
          color:#152238 !important;
      }}
      [role="option"], [role="option"] * {{
          color:#152238 !important;
          -webkit-text-fill-color:#152238 !important;
      }}
      [role="option"]:hover,
      [role="option"][aria-selected="true"] {{
          background:#EAF3FA !important;
          color:#062D5D !important;
      }}

      /* File uploader / buttons / download controls */
      [data-testid="stFileUploaderDropzone"],
      [data-testid="stFileUploader"] section {{
          background:#FFFFFF !important;
          color:#152238 !important;
          border-color:#B6C6D8 !important;
      }}
      [data-testid="stFileUploader"] *,
      [data-testid="stFileUploaderDropzone"] * {{
          color:#152238 !important;
      }}
      .stButton > button,
      .stDownloadButton > button,
      [data-testid="stFileUploader"] button {{
          background:#062D5D !important;
          color:#FFFFFF !important;
          border-color:#2F75B5 !important;
      }}
      .stButton > button *,
      .stDownloadButton > button *,
      [data-testid="stFileUploader"] button * {{
          color:#FFFFFF !important;
      }}

      /* Expanders/popovers */
      [data-testid="stExpander"] summary,
      [data-testid="stPopover"] button {{
          background:#FFFFFF !important;
          color:#152238 !important;
          border-color:#DCE4EE !important;
      }}
      [data-testid="stExpander"] summary *,
      [data-testid="stPopover"] button * {{
          color:#152238 !important;
      }}

      /* Dataframes + static tables */
      [data-testid="stDataFrame"] {{
          background:#FFFFFF !important;
          border-color:#DDE5EF !important;
      }}
      [data-testid="stDataFrame"] [role="columnheader"],
      [data-testid="stDataFrame"] [role="columnheader"] * {{
          background:#062D5D !important;
          color:#FFFFFF !important;
          -webkit-text-fill-color:#FFFFFF !important;
      }}
      [data-testid="stDataFrame"] [role="gridcell"],
      [data-testid="stDataFrame"] [role="gridcell"] * {{
          background:#FFFFFF !important;
          color:#152238 !important;
          -webkit-text-fill-color:#152238 !important;
      }}
      table {{
          background:#FFFFFF !important;
          color:#152238 !important;
      }}
      table thead th {{
          background:#062D5D !important;
          color:#FFFFFF !important;
      }}
      table tbody td {{
          background:#FFFFFF !important;
          color:#152238 !important;
      }}

      /* Alerts/captions/standard text */
      [data-testid="stCaptionContainer"],
      [data-testid="stCaptionContainer"] * {{
          color:#6C7C90 !important;
      }}

      /* Keep semantic retention colours from v3.18 */
      .good {{color:#079447 !important;}}
      .warn {{color:#F28C28 !important;}}
      .bad {{color:#D94A3A !important;}}



/* v3.39 VERIFIED v3.18 VISUAL CONSISTENCY */
.stApp, .stApp * {{
    font-family:"Source Sans Pro", sans-serif !important;
}}
.stApp {{
    background:#F4F7FB !important;
    color:#152238 !important;
}}
section[data-testid="stSidebar"] {{
    background:#EDF3FA !important;
}}
section[data-testid="stSidebar"] *,
.stApp p, .stApp li, .stApp label {{
    color:#152238 !important;
}}
h1,h2,h3,h4,h5,h6 {{
    color:#062D5D !important;
}}
.main-title,.panel-title,.tp-bar,.lost-bar {{
    background:#062D5D !important;
    color:#FFFFFF !important;
}}
.main-title *,.panel-title *,.tp-bar *,.lost-bar * {{
    color:#FFFFFF !important;
}}
.subtitle,.small-note,.tp-subtitle,.lost-sub,
.kpi-note,.tp-kpi-note,.lost-card-note {{
    color:#718099 !important;
}}
.kpi,.tp-kpi,.lost-card,.lost-action-box,
div[data-testid="stMetric"] {{
    background:#FFFFFF !important;
    border-color:#DCE4EE !important;
    color:#152238 !important;
}}
.kpi-label,.tp-kpi-lbl,.lost-card-label {{
    color:#20324A !important;
}}
.kpi-value,.tp-kpi-val,.lost-card-value {{
    color:#071B3B !important;
}}

/* Main section navigation */
div[role="radiogroup"] label {{
    background:#FFFFFF !important;
    color:#152238 !important;
    border:1px solid #DCE4EE !important;
}}
div[role="radiogroup"] label * {{
    color:#152238 !important;
}}
div[role="radiogroup"] label[data-checked="true"],
div[role="radiogroup"] label:has(input:checked) {{
    background:#062D5D !important;
    border-color:#2F75B5 !important;
}}
div[role="radiogroup"] label[data-checked="true"] *,
div[role="radiogroup"] label:has(input:checked) * {{
    color:#FFFFFF !important;
}}

/* Filters and inputs */
[data-testid="stSelectbox"] label *,
[data-testid="stMultiSelect"] label *,
[data-testid="stDateInput"] label *,
[data-testid="stNumberInput"] label *,
[data-testid="stTextInput"] label * {{
    color:#152238 !important;
    -webkit-text-fill-color:#152238 !important;
}}
div[data-baseweb="select"] > div,
div[data-baseweb="input"] > div,
[data-baseweb="base-input"],
input, textarea {{
    background:#FFFFFF !important;
    color:#152238 !important;
    border-color:#DCE4EE !important;
}}
div[data-baseweb="select"] > div > div {{
    background:transparent !important;
}}
div[data-baseweb="select"] span,
div[data-baseweb="select"] p,
div[data-baseweb="select"] input {{
    color:#152238 !important;
    -webkit-text-fill-color:#152238 !important;
    opacity:1 !important;
}}
div[data-baseweb="select"] svg {{
    color:#062D5D !important;
    fill:#062D5D !important;
    opacity:1 !important;
}}
[data-baseweb="tag"] {{
    background:#EAF3FA !important;
    border:1px solid #B6C6D8 !important;
}}
[data-baseweb="tag"] *,
[data-baseweb="tag"] svg {{
    color:#062D5D !important;
    fill:#062D5D !important;
    -webkit-text-fill-color:#062D5D !important;
}}
[data-baseweb="popover"],
[data-baseweb="popover"] ul,
[role="listbox"] {{
    background:#FFFFFF !important;
    color:#152238 !important;
    border-color:#DCE4EE !important;
}}
[role="option"],[role="option"] * {{
    color:#152238 !important;
    -webkit-text-fill-color:#152238 !important;
}}
[role="option"]:hover,
[role="option"][aria-selected="true"] {{
    background:#EAF3FA !important;
    color:#062D5D !important;
}}

/* Buttons, uploaders, popovers, expanders */
.stButton > button,
.stDownloadButton > button,
[data-testid="stFileUploader"] button {{
    background:#062D5D !important;
    color:#FFFFFF !important;
    border:1px solid #2F75B5 !important;
}}
.stButton > button *,
.stDownloadButton > button *,
[data-testid="stFileUploader"] button * {{
    color:#FFFFFF !important;
}}
[data-testid="stFileUploaderDropzone"],
[data-testid="stFileUploader"] section {{
    background:#FFFFFF !important;
    border-color:#B6C6D8 !important;
}}
[data-testid="stFileUploaderDropzone"] *,
[data-testid="stFileUploader"] section * {{
    color:#152238 !important;
}}
[data-testid="stExpander"] details,
[data-testid="stExpander"] summary,
[data-testid="stPopover"] button {{
    background:#FFFFFF !important;
    border-color:#DCE4EE !important;
    color:#152238 !important;
}}
[data-testid="stExpander"] summary *,
[data-testid="stPopover"] button * {{
    color:#152238 !important;
}}

/* Tables */
[data-testid="stDataFrame"] {{
    background:#FFFFFF !important;
    border-color:#DDE5EF !important;
}}
[data-testid="stDataFrame"] [role="columnheader"],
[data-testid="stDataFrame"] [role="columnheader"] * {{
    background:#062D5D !important;
    color:#FFFFFF !important;
    -webkit-text-fill-color:#FFFFFF !important;
}}
[data-testid="stDataFrame"] [role="gridcell"],
[data-testid="stDataFrame"] [role="gridcell"] * {{
    background:#FFFFFF !important;
    color:#152238 !important;
    -webkit-text-fill-color:#152238 !important;
}}
table {{ background:#FFFFFF !important; color:#152238 !important; }}
table thead th {{ background:#062D5D !important; color:#FFFFFF !important; }}
table tbody td {{ background:#FFFFFF !important; color:#152238 !important; }}

/* Throughput-specific HTML tables */
.tp-table,.tp-summary,.tp-table-wrap {{
    background:#FFFFFF !important;
    color:#152238 !important;
}}
.tp-table th,.tp-summary th {{
    background:#062D5D !important;
    color:#FFFFFF !important;
}}
.tp-table td,.tp-summary td {{
    background:#FFFFFF !important;
    color:#13243C !important;
}}
.tp-total td,.tp-summary .total td {{
    background:#EAF3FA !important;
    color:#062D5D !important;
}}

/* Semantic status colors retained from v3.18 */
.good {{ color:#079447 !important; }}
.warn {{ color:#F28C28 !important; }}
.bad  {{ color:#D94A3A !important; }}



/* v3.40 — EXACT FIX FOR USER-IDENTIFIED TABLE AREAS */

/* Month-on-month cohort table */
.monthly-v318-wrap {{
    width:100% !important;
    overflow-x:auto !important;
    background:#FFFFFF !important;
}}
table.monthly-v318 {{
    width:100% !important;
    border-collapse:collapse !important;
    background:#FFFFFF !important;
    color:#152238 !important;
    font-size:14px !important;
}}
table.monthly-v318 thead th {{
    background:#062D5D !important;
    color:#FFFFFF !important;
    -webkit-text-fill-color:#FFFFFF !important;
    border:1px solid #B6C6D8 !important;
    padding:7px 8px !important;
    font-weight:700 !important;
    text-align:left !important;
}}
table.monthly-v318 tbody tr td,
table.monthly-v318 tbody tr:nth-child(odd) td,
table.monthly-v318 tbody tr:nth-child(even) td {{
    background:#FFFFFF !important;
    color:#152238 !important;
    -webkit-text-fill-color:#152238 !important;
    border:1px solid #B6C6D8 !important;
    padding:6px 8px !important;
}}
table.monthly-v318 tbody td:not(:first-child) {{
    text-align:right !important;
}}

/* Throughput main + YOY summary tables */
table.tp-table tbody tr td,
table.tp-table tbody tr:nth-child(odd) td,
table.tp-table tbody tr:nth-child(even) td,
table.tp-summary tbody tr td,
table.tp-summary tbody tr:nth-child(odd) td,
table.tp-summary tbody tr:nth-child(even) td {{
    background:#FFFFFF !important;
    color:#13243C !important;
    -webkit-text-fill-color:#13243C !important;
}}
table.tp-table thead tr th,
table.tp-summary thead tr th {{
    background:#062D5D !important;
    color:#FFFFFF !important;
    -webkit-text-fill-color:#FFFFFF !important;
}}
table.tp-table td.tp-cat-cell,
table.tp-table tbody tr:nth-child(odd) td.tp-cat-cell,
table.tp-table tbody tr:nth-child(even) td.tp-cat-cell {{
    background:#FFFFFF !important;
    color:#13243C !important;
    -webkit-text-fill-color:#13243C !important;
}}
table.tp-table tbody tr.tp-total td,
table.tp-table tbody tr.tp-total:nth-child(odd) td,
table.tp-table tbody tr.tp-total:nth-child(even) td,
table.tp-summary tbody tr.total td,
table.tp-summary tbody tr.total:nth-child(odd) td,
table.tp-summary tbody tr.total:nth-child(even) td {{
    background:#EAF3FA !important;
    color:#062D5D !important;
    -webkit-text-fill-color:#062D5D !important;
    font-weight:850 !important;
}}



/* v3.43 — RESTORE STREAMLIT / MATERIAL ICON GLYPHS
   Prevent icon names such as keyboard_arrow_down from rendering as text. */
.material-symbols-rounded,
.material-symbols-outlined,
.material-icons,
.material-icons-outlined,
[data-testid="stIconMaterial"],
[data-testid="stIconMaterial"] *,
span[class*="material-symbols"],
span[class*="material-icons"] {{
    font-family:"Material Symbols Rounded","Material Symbols Outlined","Material Icons" !important;
    font-weight:normal !important;
    font-style:normal !important;
    font-size:inherit;
    line-height:1;
    letter-spacing:normal;
    text-transform:none;
    white-space:nowrap;
    word-wrap:normal;
    direction:ltr;
    -webkit-font-feature-settings:"liga";
    -webkit-font-smoothing:antialiased;
    font-feature-settings:"liga";
}}

/* BaseWeb/Streamlit dropdown arrows should stay as SVG icons. */
div[data-baseweb="select"] svg,
[data-testid="stSelectbox"] svg,
[data-testid="stMultiSelect"] svg,
[data-testid="stPopover"] svg,
[data-testid="stExpander"] svg {{
    font-family:initial !important;
    color:#062D5D !important;
    fill:#062D5D !important;
    opacity:1 !important;
}}

</style>



    """,
    unsafe_allow_html=True,
)


def clean_str(s: pd.Series) -> pd.Series:
    return s.fillna("").astype(str).str.strip()


def ensure_datetime(s: pd.Series) -> pd.Series:
    # Pandas/openpyxl usually returns Excel dates as datetimes; this handles serials too.
    out = pd.to_datetime(s, errors="coerce")
    numeric = pd.to_numeric(s, errors="coerce")
    mask = out.isna() & numeric.notna()
    if mask.any():
        out.loc[mask] = pd.Timestamp("1899-12-30") + pd.to_timedelta(numeric.loc[mask], unit="D")
    return out


def normalize_service_category(v: str) -> str:
    x = str(v).strip().lower()
    if not x:
        return "Other"
    if x in {"pms", "pm"} or "periodic" in x:
        return "PMS"
    if x == "gr" or "general" in x or "running" in x:
        return "GR"
    if "accident" in x or x == "bp" or "body" in x:
        return "Accident Repair"
    if "inspect" in x or "ins-ser" in x:
        return "Inspection"
    return str(v).strip()


def infer_brand(model: str) -> str:
    # Current workbook appears to contain Škoda retail models.
    m = str(model).upper()
    skoda_tokens = ["KUSHAQ", "SLAVIA", "KODIAQ", "SUPERB", "OCTAVIA", "KYLAQ", "RAPID"]
    if any(t in m for t in skoda_tokens):
        return "Škoda"
    return "Other / Unmapped"


@st.cache_data(show_spinner=False)
def load_from_bytes(file_bytes: bytes):
    bio = io.BytesIO(file_bytes)
    sold = pd.read_excel(bio, sheet_name="SOLD_DATA", header=1, engine="openpyxl")
    bio.seek(0)
    service = pd.read_excel(bio, sheet_name="SERVICE_DATA", header=1, engine="openpyxl")
    return sold, service


@st.cache_data(show_spinner=False)
def load_from_path(path: str, modified_time: float):
    sold = pd.read_excel(path, sheet_name="SOLD_DATA", header=1, engine="openpyxl")
    service = pd.read_excel(path, sheet_name="SERVICE_DATA", header=1, engine="openpyxl")
    return sold, service


@st.cache_data(show_spinner=False)
def load_prepared_smp(path: str, modified_time: float):
    """Read + clean the SMP workbook once per file version instead of on every rerun."""
    smp = pd.read_excel(path, engine="openpyxl")
    smp.columns = [str(c).strip() for c in smp.columns]

    for _c in ["VIN", "Customer Name", "Model", "Status", "Product Name",
               "Dealer Name", "Dealer City", "Sold From", "Regd Number"]:
        if _c in smp.columns:
            smp[_c] = smp[_c].fillna("").astype(str).str.strip()

    smp["VIN"] = smp["VIN"].str.upper()
    for _c in ["Sold Date", "Created Date", "Validity From Date", "Validity To Date"]:
        if _c in smp.columns:
            smp[_c] = pd.to_datetime(smp[_c], errors="coerce", dayfirst=False)

    if "Validity in Years" in smp.columns:
        smp["Validity in Years"] = pd.to_numeric(smp["Validity in Years"], errors="coerce")
    else:
        smp["Validity in Years"] = (
            (smp["Validity To Date"] - smp["Validity From Date"]).dt.days / 365.25
        ).round()

    smp = smp[
        smp["VIN"].ne("")
        & smp["Validity From Date"].notna()
        & smp["Validity To Date"].notna()
    ].copy()
    return smp



def club_service_rows_by_ro(service: pd.DataFrame) -> pd.DataFrame:
    """Return one service-event row per VIN + RO Number.

    The source can contain separate customer / warranty / internal invoices for the
    same repair order.  For dashboard purposes they are one workshop visit, so we
    club them while summing the monetary values.  Blank RO numbers are deliberately
    left as separate rows because there is no safe event key for them.
    """
    if service.empty or "VIN" not in service.columns or "RO Number" not in service.columns:
        return service

    df = service.copy().drop_duplicates().reset_index(drop=True)
    ro = df["RO Number"].fillna("").astype(str).str.strip()
    with_ro = df[ro.ne("")].copy()
    without_ro = df[ro.eq("")].copy()
    if with_ro.empty:
        return df

    money_cols = [c for c in [
        "Labor", "Parts", "Total", "Invoice Amount Inc Tax",
        "Amount", "Amount After Discount", "Amount (Inc Tax)",
    ] if c in with_ro.columns]
    for c in money_cols:
        with_ro[c] = pd.to_numeric(with_ro[c], errors="coerce")

    if "KMS" in with_ro.columns:
        with_ro["KMS"] = pd.to_numeric(with_ro["KMS"], errors="coerce")

    def first_nonblank(x):
        for v in x:
            if pd.notna(v) and str(v).strip() != "":
                return v
        return ""

    def join_unique(x):
        vals = []
        for v in x:
            if pd.notna(v):
                t = str(v).strip()
                if t and t not in vals:
                    vals.append(t)
        return " / ".join(vals)

    agg = {}
    for c in with_ro.columns:
        if c in ("VIN", "RO Number"):
            continue
        if c == "Service Date":
            # Earliest date best represents when the RO/visit was first reported.
            agg[c] = "min"
        elif c in money_cols:
            agg[c] = lambda x: x.sum(min_count=1)
        elif c == "KMS":
            agg[c] = "max"
        elif c == "Invoice No":
            agg[c] = join_unique
        else:
            agg[c] = first_nonblank

    clubbed = (
        with_ro.groupby(["VIN", "RO Number"], as_index=False, sort=False, dropna=False)
        .agg(agg)
    )

    # Restore the original column order and keep blank-RO rows unchanged.
    clubbed = clubbed.reindex(columns=df.columns)
    if not without_ro.empty:
        out = pd.concat([clubbed, without_ro.reindex(columns=df.columns)], ignore_index=True)
    else:
        out = clubbed
    if "Service Date" in out.columns:
        out = out.sort_values(["Service Date", "VIN", "RO Number"], kind="stable").reset_index(drop=True)
    return out

def _prepare_data_worker(sold: pd.DataFrame, service: pd.DataFrame):
    sold = sold.copy()
    service = service.copy()

    sold.columns = [str(c).strip() for c in sold.columns]
    service.columns = [str(c).strip() for c in service.columns]

    required_sold = ["Customer Name", "City", "Sold Date", "Model", "VIN", "Sold Branch"]
    required_service = ["VIN", "Service Date", "Service Location", "Service Type"]
    missing = [f"SOLD_DATA: {c}" for c in required_sold if c not in sold.columns]
    missing += [f"SERVICE_DATA: {c}" for c in required_service if c not in service.columns]
    if missing:
        raise ValueError("Missing required columns: " + ", ".join(missing))

    for c in ["Customer Name", "City", "Model", "VIN", "Consultant", "Sold Branch"]:
        if c in sold.columns:
            sold[c] = clean_str(sold[c])
    for c in ["VIN", "Vehicle No", "RO Number", "Service Location", "Service Type",
              "Repair Type", "Mapped Category", "Customer Name"]:
        if c in service.columns:
            service[c] = clean_str(service[c])

    sold["VIN"] = sold["VIN"].str.upper()
    service["VIN"] = service["VIN"].str.upper()

    sold["Sold Date"] = ensure_datetime(sold["Sold Date"])
    service["Service Date"] = ensure_datetime(service["Service Date"])

    sold = sold[sold["VIN"].ne("") & sold["Sold Date"].notna()].copy()
    service = service[service["VIN"].ne("") & service["Service Date"].notna()].copy()

    # One workshop visit = one VIN + RO.  Customer/warranty/internal invoice rows
    # under the same RO are clubbed and their amounts are summed before any KPI,
    # VIN Search, PMS history, retention or throughput calculation is performed.
    service = club_service_rows_by_ro(service)

    # Never trust workbook formula/cache columns for calculations.
    sold["Sold Year Calc"] = sold["Sold Date"].dt.year.astype("Int64")
    sold["Sold Month No Calc"] = sold["Sold Date"].dt.month.astype("Int64")
    sold["Sold Month Calc"] = sold["Sold Date"].dt.strftime("%b")
    sold["Cohort Calc"] = sold["Sold Date"].dt.strftime("%b-%y")

    if "Brand" in sold.columns:
        sold["Brand Calc"] = clean_str(sold["Brand"]).replace("", "Other / Unmapped")
    else:
        sold["Brand Calc"] = sold["Model"].map(infer_brand)

    if "Mapped Category" in service.columns:
        service["Category Calc"] = service["Mapped Category"].map(normalize_service_category)
    elif "Repair Type" in service.columns:
        service["Category Calc"] = service["Repair Type"].map(normalize_service_category)
    else:
        service["Category Calc"] = service["Service Type"].map(normalize_service_category)

    service["Service Year Calc"] = service["Service Date"].dt.year.astype("Int64")
    service["Service Month No Calc"] = service["Service Date"].dt.month.astype("Int64")
    service["Service Month Calc"] = service["Service Date"].dt.strftime("%b")

    # A VIN is a vehicle. Keep one sold row per VIN for population analysis.
    sold_unique = (
        sold.sort_values("Sold Date")
            .drop_duplicates("VIN", keep="last")
            .copy()
    )
    return sold, sold_unique, service


@st.cache_data(show_spinner=False)
def load_prepared_master(
    sales_path: str,
    sales_mtime: float,
    service_path: str,
    service_mtime: float,
):
    """Load + prepare master data once. Re-runs only after either master file changes."""
    sold_raw = pd.read_csv(sales_path, low_memory=False)
    service_raw = pd.read_csv(service_path, low_memory=False)
    sold_all, sold_unique, service = _prepare_data_worker(sold_raw, service_raw)
    return sold_raw, service_raw, sold_all, sold_unique, service

def prepare_data(sold: pd.DataFrame, service: pd.DataFrame):
    """Compatibility wrapper used by upload validation."""
    return _prepare_data_worker(sold, service)

def choices(series):
    return sorted([x for x in series.dropna().unique().tolist() if str(x).strip() != ""], key=lambda x: str(x))


# Keep Best/Lowest Cohort cards tall enough for a two-line value without clipping the title.
st.markdown("""
<style>
.kpi.cohort-extreme-kpi {
    min-height: 100px !important;
    height: 100px !important;
    overflow: visible !important;
    padding-top: 10px !important;
    padding-bottom: 10px !important;
}
.kpi.cohort-extreme-kpi .kpi-content {
    display: flex !important;
    flex-direction: column !important;
    justify-content: center !important;
    min-height: 78px !important;
}
.kpi.cohort-extreme-kpi .kpi-label {
    display: block !important;
    margin: 0 0 3px 0 !important;
    line-height: 1.15 !important;
}
.kpi.cohort-extreme-kpi .kpi-value {
    font-size: 22px !important;
    line-height: 1.02 !important;
    margin: 0 0 3px 0 !important;
}
.kpi.cohort-extreme-kpi .kpi-note {
    margin: 0 !important;
    line-height: 1.1 !important;
}
</style>
""", unsafe_allow_html=True)

def unique_count(df, col="VIN"):
    return int(df[col].nunique()) if col in df.columns and len(df) else 0


def kpi(label, value, note="", accent=BLUE, icon=""):
    icon_html = (
        f'<div class="kpi-icon" style="color:{accent};">{icon}</div>'
        if icon else ""
    )
    cohort_class = " cohort-extreme-kpi" if label in ("Best Cohort", "Lowest Cohort") else ""
    kpi_html = textwrap.dedent(f"""
        <div class="kpi{cohort_class}" style="border-top:4px solid {accent};">
            {icon_html}
            <div class="kpi-content">
                <div class="kpi-label">{label}</div>
                <div class="kpi-value">{value}</div>
                <div class="kpi-note">{note}</div>
            </div>
        </div>
    """).strip()
    st.markdown(kpi_html, unsafe_allow_html=True)

def style_fig(fig, chart_height=350, percent_y=False):
    fig.update_layout(
        height=chart_height,
        margin=dict(l=15, r=15, t=40, b=25),
        paper_bgcolor="#FFFFFF",
        plot_bgcolor="#FFFFFF",
        font=dict(family=APP_FONT, color="#152238", size=12),
        title_font=dict(family=APP_FONT, size=14, color="#2F75B5"),
        legend=dict(
            orientation="h", yanchor="bottom", y=1.01, xanchor="left", x=0,
            font=dict(family=APP_FONT, color="#39495E")
        ),
        hoverlabel=dict(
            bgcolor="#062D5D",
            bordercolor="#B6C6D8",
            font=dict(family=APP_FONT, color="#FFFFFF")
        ),
    )
    fig.update_xaxes(
        showgrid=False, linecolor="#B6C6D8",
        tickfont=dict(family=APP_FONT, color="#39495E", size=11),
        title_font=dict(family=APP_FONT, color="#39495E")
    )
    fig.update_yaxes(
        showgrid=True, gridcolor="#DDE5EF", zeroline=False,
        tickfont=dict(family=APP_FONT, color="#39495E", size=11),
        title_font=dict(family=APP_FONT, color="#39495E")
    )
    if percent_y:
        fig.update_yaxes(ticksuffix="%")
    return fig



# Sidebar navigation — styled to match the compact dashboard menu reference.
PAGE_LABELS = [
    "⬆️ Monthly Data Upload",
    "📊 Executive Dashboard",
    "🧩 Cohort Matrix",
    "📈 Throughput Tracking",
    "🚨 Action Dashboard",
    "🚫 Lost VIN Details",
    "🔎 VIN Search",
    "✅ Data Quality",
    "🚨 Lost VIN & Immediate Action",
    "🛡️ SMP Tracking",
]

# Friendlier display text for the left navigation while preserving the original
# page identifiers used by the dashboard logic below.
NAV_ITEMS = [
    ("📊  Executive Dashboard", PAGE_LABELS[1]),
    ("🧩  Cohort Matrix", PAGE_LABELS[2]),
    ("📈  Throughput Tracking", PAGE_LABELS[3]),
    ("🚨  Lost VIN & Action", PAGE_LABELS[8]),
    ("🔎  VIN Search", PAGE_LABELS[6]),
    ("✅  Data Quality", PAGE_LABELS[7]),
    ("⬆️  Data Upload", PAGE_LABELS[0]),
    ("🚨  Action Dashboard", PAGE_LABELS[4]),
    ("🚫  Lost VIN Details", PAGE_LABELS[5]),
    ("🛡️  SMP Tracking", PAGE_LABELS[9]),
]

st.markdown(
    """
    <style>
    /* Left dashboard menu — screenshot-style vertical tabs */
    section[data-testid="stSidebar"] div[role="radiogroup"] {
        display:flex !important;
        flex-direction:column !important;
        gap:4px !important;
        width:100% !important;
    }
    section[data-testid="stSidebar"] div[role="radiogroup"] label {
        width:100% !important;
        min-height:36px !important;
        display:flex !important;
        align-items:center !important;
        justify-content:flex-start !important;
        background:transparent !important;
        border:0 !important;
        border-radius:8px !important;
        padding:7px 11px !important;
        margin:0 !important;
        box-shadow:none !important;
        cursor:pointer !important;
    }
    section[data-testid="stSidebar"] div[role="radiogroup"] label:hover {
        background:#E7EEF7 !important;
    }
    section[data-testid="stSidebar"] div[role="radiogroup"] label[data-checked="true"],
    section[data-testid="stSidebar"] div[role="radiogroup"] label:has(input:checked) {
        background:#0B376B !important;
        border:0 !important;
        box-shadow:none !important;
    }
    section[data-testid="stSidebar"] div[role="radiogroup"] label *,
    section[data-testid="stSidebar"] div[role="radiogroup"] label p,
    section[data-testid="stSidebar"] div[role="radiogroup"] label span {
        color:#071B3B !important;
        -webkit-text-fill-color:#071B3B !important;
        font-weight:600 !important;
        font-size:14px !important;
        line-height:1.15 !important;
    }
    section[data-testid="stSidebar"] div[role="radiogroup"] label[data-checked="true"] *,
    section[data-testid="stSidebar"] div[role="radiogroup"] label:has(input:checked) * {
        color:#FFFFFF !important;
        -webkit-text-fill-color:#FFFFFF !important;
        font-weight:700 !important;
    }
    /* Hide Streamlit's radio dot/circle so each option reads like a tab. */
    section[data-testid="stSidebar"] div[role="radiogroup"] label > div:first-child {
        display:none !important;
    }
    .sidebar-dashboard-title {
        font-size:17px;
        font-weight:800;
        color:#071B3B;
        margin:2px 0 10px 2px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <style>
    /* FINAL ICON FIX: never expose Material icon ligature names as text.
       Streamlit renders the surrounding button, while this hides only the
       font-dependent ligature glyph that can become keyboard_arrow_* text. */
    [data-testid="stIconMaterial"],
    [data-testid="stIconMaterial"] * {
        font-size:0 !important;
        color:transparent !important;
        -webkit-text-fill-color:transparent !important;
        overflow:hidden !important;
        width:0 !important;
        max-width:0 !important;
    }
    /* Keep genuine SVG controls visible and independent of icon fonts. */
    button svg, [data-baseweb="select"] svg, [data-testid="stSelectbox"] svg,
    [data-testid="stMultiSelect"] svg, [data-testid="stPopover"] svg,
    [data-testid="stExpander"] svg {
        display:block !important;
        width:18px !important;
        height:18px !important;
        color:#062D5D !important;
        fill:currentColor !important;
        opacity:1 !important;
    }
    /* Growth indicators: thick, high-contrast and immune to broad text CSS. */
    .tp-kpi-note span, table.tp-table td span, table.tp-summary td span {
        opacity:1 !important;
        -webkit-text-fill-color:currentColor !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

with st.sidebar:
    st.markdown('<div class="sidebar-dashboard-title">Dashboard</div>', unsafe_allow_html=True)
    nav_display = st.radio(
        "Dashboard",
        [x[0] for x in NAV_ITEMS],
        index=0,
        label_visibility="collapsed",
        key="main_sidebar_nav_v1",
    )
    active_page = dict(NAV_ITEMS)[nav_display]

    st.markdown("---")
    st.markdown("### 📁 Data Source")
    st.success("Monthly mode enabled")
    st.caption("Upload only Sales Data + Service Data each month. The app keeps the historical master data.")
    if st.button("🔄 Refresh Dashboard", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

try:
    if not MASTER_SALES.exists() or not MASTER_SERVICE.exists():
        raise FileNotFoundError("Master data files have not been initialized.")
    sold_raw, service_raw, sold_all, sold, service = load_prepared_master(
        str(MASTER_SALES),
        MASTER_SALES.stat().st_mtime,
        str(MASTER_SERVICE),
        MASTER_SERVICE.stat().st_mtime,
    )
    source_name = "Monthly Master Data"
except Exception as e:
    st.error(f"Could not load master data: {e}")
    st.stop()

sold_years = sorted([int(x) for x in sold["Sold Year Calc"].dropna().unique()], reverse=True)
service_years = sorted([int(x) for x in service["Service Year Calc"].dropna().unique()], reverse=True)

default_sold_year = 2025 if 2025 in sold_years else sold_years[0]
default_service_year = default_sold_year if default_sold_year in service_years else service_years[0]

with st.sidebar:
    st.markdown("---")
    st.markdown("### 🎛️ Dashboard Filters")

    # Main filters in the requested order
    brand = st.selectbox("1. Brand", ["All"] + choices(sold["Brand Calc"]))
    branch = st.selectbox("2. Sold Branch", ["All"] + choices(sold["Sold Branch"]))

    st.markdown("**3. Sold Year - Month**")
    sold_year = st.selectbox(
        "Sold Year",
        sold_years,
        index=sold_years.index(default_sold_year),
        key="filter_sold_year",
    )
    sold_month = st.selectbox(
        "Sold Month",
        ["All"] + list(pd.date_range("2025-01-01", periods=12, freq="MS").strftime("%b")),
        key="filter_sold_month",
    )

    st.markdown("**4. Service Year - Month**")
    service_year = st.selectbox(
        "Service Year",
        service_years,
        index=service_years.index(default_service_year),
        key="filter_service_year",
    )
    service_month = st.selectbox(
        "Service Month",
        ["All"] + list(pd.date_range("2025-01-01", periods=12, freq="MS").strftime("%b")),
        key="filter_service_month",
    )

    service_type = st.selectbox("5. Service Type", ["All"] + choices(service["Service Type"]))
    category = st.selectbox("6. Service Category", ["All"] + choices(service["Category Calc"]))

    with st.expander("More Filters"):
        city = st.selectbox("Customer Location / City", ["All"] + choices(sold["City"]))
        model = st.selectbox("Model", ["All"] + choices(sold["Model"]))
        service_location = st.selectbox("Service Location", ["All"] + choices(service["Service Location"]))

# Sold population filters
sf = sold[sold["Sold Year Calc"].eq(sold_year)].copy()
if sold_month != "All":
    sf = sf[sf["Sold Month Calc"].eq(sold_month)]
if branch != "All":
    sf = sf[sf["Sold Branch"].eq(branch)]
if city != "All":
    sf = sf[sf["City"].eq(city)]
if brand != "All":
    sf = sf[sf["Brand Calc"].eq(brand)]
if model != "All":
    sf = sf[sf["Model"].eq(model)]

population_vins = set(sf["VIN"])

# Service filters
sv = service[service["VIN"].isin(population_vins) & service["Service Year Calc"].eq(service_year)].copy()
if service_month != "All":
    sv = sv[sv["Service Month Calc"].eq(service_month)]
if service_location != "All":
    sv = sv[sv["Service Location"].eq(service_location)]
if service_type != "All":
    sv = sv[sv["Service Type"].eq(service_type)]
if category != "All":
    sv = sv[sv["Category Calc"].eq(category)]

serviced_vins = set(sv["VIN"].unique())
total_sold = len(population_vins)
total_serviced = len(serviced_vins)
retention = total_serviced / total_sold if total_sold else 0
not_reported = max(total_sold - total_serviced, 0)

cat_counts = (
    sv.groupby("Category Calc")["VIN"].nunique().sort_values(ascending=False)
    if len(sv) else pd.Series(dtype=int)
)
pms = int(cat_counts.get("PMS", 0))
gr = int(cat_counts.get("GR", 0))
acc = int(cat_counts.get("Accident Repair", 0))
inspection = int(cat_counts.get("Inspection", 0))

st.markdown('<div class="main-title">🚘 MODY SKODA SOUTH — SOLD VEHICLE COHORT RETENTION DASHBOARD</div>', unsafe_allow_html=True)
st.markdown(
    f'<div class="subtitle">Live calculations from <b>{source_name}</b> • Unique VIN logic • Filters recalculate instantly</div>',
    unsafe_allow_html=True,
)

# active_page is selected from the screenshot-style sidebar navigation above.
# Only the selected page is rendered, keeping the original performance behavior.



def yoy_growth(current, previous):
    current = float(current or 0)
    previous = float(previous or 0)
    if previous == 0:
        return np.nan if current > 0 else 0.0
    return ((current - previous) / previous) * 100.0

def growth_text(growth):
    if pd.isna(growth):
        return "— New"
    if growth > 0.05:
        return f"▲ {growth:.1f}%"
    if growth < -0.05:
        return f"▼ {abs(growth):.1f}%"
    return f"▶ {growth:.1f}%"

if active_page == PAGE_LABELS[0]:
    st.subheader("Data Upload Center")
    st.caption(
        "Use Full Master Upload when loading/replacing the complete historical database. "
        "Use Monthly Incremental Upload for your regular monthly additions."
    )

    upload_mode = st.radio(
        "Choose upload type",
        ["📚 Full Master Upload", "➕ Monthly Incremental Upload"],
        horizontal=True,
        key="data_upload_mode_v325",
    )

    st.markdown(
        """
        **Mandatory columns**
        - Sales: `VIN`, `Sold Date`
        - Service: `VIN`, `Service Date`

        For a single combined workbook, use sheet names `SOLD_DATA` and `SERVICE_DATA`.

        Recommended columns are preserved automatically, including Customer Name, City, Model,
        Sold Branch, Consultant, Vehicle No, RO Number, Service Location, KMS, Service Type,
        Repair Type, Mapped Category and invoice values.
        """
    )

    if upload_mode == "📚 Full Master Upload":
        st.info(
            "Replace either the complete Sales master or the complete Service master independently. "
            "The other master remains unchanged and the current file is backed up automatically."
        )

        replace_choice = st.radio(
            "What do you want to replace?",
            [
                "🟦 Replace FULL SALES",
                "🟩 Replace FULL SERVICE",
            ],
            horizontal=True,
            key="full_replace_choice_v346",
        )

        if replace_choice == "🟦 Replace FULL SALES":
            full_file = st.file_uploader(
                "Upload FULL SALES history",
                type=["xlsx", "csv"],
                key="full_sales_v346",
                help="Upload Excel or CSV containing VIN and Sold Date.",
            )
            confirm_text = "I understand this will replace the current Sales master. Service data will remain unchanged."
            button_label = "🔄 Validate & Replace FULL SALES"
        else:
            full_file = st.file_uploader(
                "Upload FULL SERVICE history",
                type=["xlsx", "csv"],
                key="full_service_v346",
                help="Upload Excel or CSV containing VIN and Service Date.",
            )
            confirm_text = "I understand this will replace the current Service master. Sales data will remain unchanged."
            button_label = "🔄 Validate & Replace FULL SERVICE"

        confirm_replace = st.checkbox(confirm_text, key="confirm_full_replace_v346")

        if st.button(
            button_label,
            type="primary",
            use_container_width=True,
            disabled=not (full_file is not None and confirm_replace),
            key="replace_full_master_v346",
        ):
            try:
                backup_dir = APP_DIR / "backups"
                backup_dir.mkdir(exist_ok=True)
                stamp = datetime.now().strftime("%Y%m%d_%H%M%S")

                current_sales = pd.read_csv(MASTER_SALES, low_memory=False)
                current_service = pd.read_csv(MASTER_SERVICE, low_memory=False)

                if replace_choice == "🟦 Replace FULL SALES":
                    new_sales = read_monthly_file(full_file, "sales")
                    new_sales.columns = [str(c).strip() for c in new_sales.columns]

                    if "VIN" not in new_sales.columns or "Sold Date" not in new_sales.columns:
                        raise ValueError("Sales file must contain VIN and Sold Date.")

                    new_sales = new_sales[
                        new_sales["VIN"].fillna("").astype(str).str.strip().ne("")
                    ].drop_duplicates().reset_index(drop=True)

                    # Validate against the unchanged Service master before replacement.
                    _ = prepare_data(new_sales, current_service)

                    if MASTER_SALES.exists():
                        shutil.copy2(
                            MASTER_SALES,
                            backup_dir / f"MASTER_SALES_before_replace_{stamp}.csv",
                        )

                    save_master_csv(new_sales, MASTER_SALES)
                    save_default_full_workbook(new_sales, current_service)

                    import_type = "FULL SALES REPLACE"
                    sales_rows = len(new_sales)
                    service_rows = len(current_service)
                    sales_name = full_file.name
                    service_name = "UNCHANGED"

                    success_message = (
                        f"FULL SALES replaced successfully with {sales_rows:,} rows. "
                        f"Service master remains unchanged at {service_rows:,} rows."
                    )

                else:
                    new_service = read_monthly_file(full_file, "service")
                    new_service.columns = [str(c).strip() for c in new_service.columns]

                    if "VIN" not in new_service.columns or "Service Date" not in new_service.columns:
                        raise ValueError("Service file must contain VIN and Service Date.")

                    new_service = new_service[
                        new_service["VIN"].fillna("").astype(str).str.strip().ne("")
                    ].drop_duplicates().reset_index(drop=True)

                    # Validate against the unchanged Sales master before replacement.
                    _ = prepare_data(current_sales, new_service)

                    if MASTER_SERVICE.exists():
                        shutil.copy2(
                            MASTER_SERVICE,
                            backup_dir / f"MASTER_SERVICE_before_replace_{stamp}.csv",
                        )

                    save_master_csv(new_service, MASTER_SERVICE)
                    save_default_full_workbook(current_sales, new_service)

                    import_type = "FULL SERVICE REPLACE"
                    sales_rows = len(current_sales)
                    service_rows = len(new_service)
                    sales_name = "UNCHANGED"
                    service_name = full_file.name

                    success_message = (
                        f"FULL SERVICE replaced successfully with {service_rows:,} rows. "
                        f"Sales master remains unchanged at {sales_rows:,} rows."
                    )

                log_row = pd.DataFrame([{
                    "Imported At": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "Import Type": import_type,
                    "Sales File": sales_name,
                    "Service File": service_name,
                    "Master Sales Rows": sales_rows,
                    "Master Service Rows": service_rows,
                }])

                if IMPORT_LOG.exists():
                    old_log = pd.read_csv(IMPORT_LOG)
                    log_row = pd.concat([old_log, log_row], ignore_index=True)

                log_row.to_csv(IMPORT_LOG, index=False, encoding="utf-8-sig")
                st.cache_data.clear()
                st.success(success_message)
                st.success("Chandra.xlsx has been synchronized with the active master data.")
                st.rerun()

            except Exception as e:
                st.error(f"Full master upload failed: {e}")

    else:
        st.info(
            "Monthly upload works independently. Upload this month's Sales data OR this month's Service data. "
            "You do not need to upload both."
        )

        monthly_choice = st.radio(
            "What monthly data do you want to append?",
            [
                "🟦 Append MONTHLY SALES",
                "🟩 Append MONTHLY SERVICE",
            ],
            horizontal=True,
            key="monthly_append_choice_v347",
        )

        if monthly_choice == "🟦 Append MONTHLY SALES":
            monthly_file = st.file_uploader(
                "Upload this month's SALES data",
                type=["xlsx", "csv"],
                key="monthly_sales_only_v347",
                help="Excel or CSV containing VIN and Sold Date.",
            )
            confirm_monthly = st.checkbox(
                "I confirm this file contains monthly Sales data to append.",
                key="confirm_monthly_sales_v347",
            )
            monthly_button = "➕ Validate & Append MONTHLY SALES"

        else:
            monthly_file = st.file_uploader(
                "Upload this month's SERVICE data",
                type=["xlsx", "csv"],
                key="monthly_service_only_v347",
                help="Excel or CSV containing VIN and Service Date.",
            )
            confirm_monthly = st.checkbox(
                "I confirm this file contains monthly Service data to append.",
                key="confirm_monthly_service_v347",
            )
            monthly_button = "➕ Validate & Append MONTHLY SERVICE"

        if st.button(
            monthly_button,
            type="primary",
            use_container_width=True,
            disabled=not (monthly_file is not None and confirm_monthly),
            key="append_monthly_v347",
        ):
            try:
                backup_dir = APP_DIR / "backups"
                backup_dir.mkdir(exist_ok=True)
                stamp = datetime.now().strftime("%Y%m%d_%H%M%S")

                current_sales = read_master_csv(MASTER_SALES)
                current_service = read_master_csv(MASTER_SERVICE)

                if monthly_choice == "🟦 Append MONTHLY SALES":
                    new_sales = read_monthly_file(monthly_file, "sales")
                    new_sales.columns = [str(c).strip() for c in new_sales.columns]

                    if "VIN" not in new_sales.columns or "Sold Date" not in new_sales.columns:
                        raise ValueError("Monthly Sales file must contain VIN and Sold Date.")

                    merged_sales, sales_dupes = append_monthly_data(
                        current_sales, new_sales, "sales"
                    )

                    # Validate with the unchanged Service master before saving.
                    _ = prepare_data(merged_sales, current_service)

                    if MASTER_SALES.exists():
                        shutil.copy2(
                            MASTER_SALES,
                            backup_dir / f"MASTER_SALES_before_monthly_append_{stamp}.csv",
                        )

                    save_master_csv(merged_sales, MASTER_SALES)
                    save_default_full_workbook(merged_sales, current_service)

                    log_row = pd.DataFrame([{
                        "Imported At": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "Import Type": "MONTHLY SALES APPEND",
                        "Sales File": monthly_file.name,
                        "Service File": "UNCHANGED",
                        "Sales Rows Uploaded": len(new_sales),
                        "Service Rows Uploaded": 0,
                        "Sales Exact Duplicates Removed": sales_dupes,
                        "Service Exact Duplicates Removed": 0,
                        "Master Sales Rows": len(merged_sales),
                        "Master Service Rows": len(current_service),
                    }])

                    success_message = (
                        f"Monthly Sales appended successfully. "
                        f"Sales master now contains {len(merged_sales):,} rows. "
                        f"Service master remains unchanged at {len(current_service):,} rows."
                    )
                    duplicate_message = (
                        f"Exact duplicate Sales rows automatically removed: {sales_dupes:,}."
                        if sales_dupes else ""
                    )

                else:
                    new_service = read_monthly_file(monthly_file, "service")
                    new_service.columns = [str(c).strip() for c in new_service.columns]

                    if "VIN" not in new_service.columns or "Service Date" not in new_service.columns:
                        raise ValueError("Monthly Service file must contain VIN and Service Date.")

                    merged_service, service_dupes = append_monthly_data(
                        current_service, new_service, "service"
                    )

                    # Validate with the unchanged Sales master before saving.
                    _ = prepare_data(current_sales, merged_service)

                    if MASTER_SERVICE.exists():
                        shutil.copy2(
                            MASTER_SERVICE,
                            backup_dir / f"MASTER_SERVICE_before_monthly_append_{stamp}.csv",
                        )

                    save_master_csv(merged_service, MASTER_SERVICE)
                    save_default_full_workbook(current_sales, merged_service)

                    log_row = pd.DataFrame([{
                        "Imported At": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "Import Type": "MONTHLY SERVICE APPEND",
                        "Sales File": "UNCHANGED",
                        "Service File": monthly_file.name,
                        "Sales Rows Uploaded": 0,
                        "Service Rows Uploaded": len(new_service),
                        "Sales Exact Duplicates Removed": 0,
                        "Service Exact Duplicates Removed": service_dupes,
                        "Master Sales Rows": len(current_sales),
                        "Master Service Rows": len(merged_service),
                    }])

                    success_message = (
                        f"Monthly Service appended successfully. "
                        f"Service master now contains {len(merged_service):,} rows. "
                        f"Sales master remains unchanged at {len(current_sales):,} rows."
                    )
                    duplicate_message = (
                        f"Exact duplicate Service rows automatically removed: {service_dupes:,}."
                        if service_dupes else ""
                    )

                if IMPORT_LOG.exists():
                    old_log = pd.read_csv(IMPORT_LOG)
                    log_row = pd.concat([old_log, log_row], ignore_index=True)

                log_row.to_csv(IMPORT_LOG, index=False, encoding="utf-8-sig")

                st.cache_data.clear()
                st.success(success_message)
                if duplicate_message:
                    st.info(duplicate_message)
                st.success("Chandra.xlsx has been synchronized with the active master data.")
                st.rerun()

            except Exception as e:
                st.error(f"Monthly upload failed: {e}")

    st.markdown("---")
    a, b, c = st.columns(3)
    a.metric("Master Sales Rows", f"{len(sold_raw):,}")
    b.metric("Master Service Rows", f"{len(service_raw):,}")
    c.metric("Unique Sold VINs", f"{sold['VIN'].nunique():,}")

    if IMPORT_LOG.exists():
        st.markdown("#### Recent Data Imports")
        log = pd.read_csv(IMPORT_LOG)
        st.dataframe(log.tail(12).iloc[::-1], use_container_width=True, hide_index=True)


if active_page == PAGE_LABELS[1]:
    st.markdown("### 📊 Executive Summary")
    krow1 = st.columns(4, gap="small")
    with krow1[0]:
        kpi("TOTAL VINs SOLD", f"{total_sold:,}", f"Sold {sold_month} {sold_year}" if sold_month != "All" else f"Sold year {sold_year}", BLUE, "🚗")
    with krow1[1]:
        kpi("TOTAL VINs SERVICED", f"{total_serviced:,}", f"Service {service_month} {service_year}" if service_month != "All" else f"Service year {service_year}", GREEN, "👥")
    with krow1[2]:
        kpi("TOTAL RETENTION %", f"{retention:.1%}", retention_band(retention), retention_color(retention), "📈")
    with krow1[3]:
        kpi("PMS VINs", f"{pms:,}", "Unique VINs", TEAL, "🔧")

    krow2 = st.columns(4, gap="small")
    with krow2[0]:
        kpi("GR VINs", f"{gr:,}", "Unique VINs", ORANGE, "⚙️")
    with krow2[1]:
        kpi("Accident Repair VINs", f"{acc:,}", "Unique VINs", RED, "🚘")
    with krow2[2]:
        kpi("Inspection VINs", f"{inspection:,}", "Unique VINs", BLUE, "📋")
    with krow2[3]:
        st.markdown('<div style="height:100%;"></div>', unsafe_allow_html=True)

    # Monthly cohort table
    monthly_rows = []
    for m in range(1, 13):
        cohort = sf[sf["Sold Month No Calc"].eq(m)] if sold_month == "All" else sf
        if sold_month != "All" and m != pd.to_datetime(sold_month, format="%b").month:
            continue
        cvins = set(cohort["VIN"])
        csv = sv[sv["VIN"].isin(cvins)]
        cserv = csv["VIN"].nunique()
        csold = len(cvins)
        row = {
            "Sold Month Cohort": pd.Timestamp(2000, m, 1).strftime("%b") + f"-{str(sold_year)[-2:]}",
            "Original Sold VINs": csold,
            "Unique VINs Serviced": cserv,
            "Retention %": cserv / csold if csold else 0,
            "PMS VINs": csv.loc[csv["Category Calc"].eq("PMS"), "VIN"].nunique(),
            "GR VINs": csv.loc[csv["Category Calc"].eq("GR"), "VIN"].nunique(),
            "Accident Repair VINs": csv.loc[csv["Category Calc"].eq("Accident Repair"), "VIN"].nunique(),
            "Inspection VINs": csv.loc[csv["Category Calc"].eq("Inspection"), "VIN"].nunique(),
            "Not Reported VINs": max(csold - cserv, 0),
        }
        monthly_rows.append(row)
        if sold_month != "All":
            break

    monthly = pd.DataFrame(monthly_rows)

    st.markdown(f'<div class="panel-title">MONTH ON MONTH RETENTION BY SOLD MONTH COHORT — {sold_year}</div>', unsafe_allow_html=True)
    display_monthly = monthly.copy()
    display_monthly["Retention %"] = display_monthly["Retention %"].map(lambda x: f"{x:.1%}")

    # Full-width static table with dedicated v3.18 styling.
    monthly_html = display_monthly.to_html(
        index=False,
        border=0,
        classes="monthly-v318",
        justify="right",
    )
    st.markdown(
        f'<div class="monthly-v318-wrap">{monthly_html}</div>',
        unsafe_allow_html=True,
    )

    fig = px.bar(
        monthly,
        x="Sold Month Cohort",
        y=monthly["Retention %"] * 100,
        text=monthly["Retention %"].map(lambda x: f"{x:.1%}"),
        title=f"RETENTION % BY SOLD MONTH COHORT — {sold_year}",
    )
    fig.update_traces(
        marker_color=[retention_color(v) for v in monthly["Retention %"]],
        textposition="outside",
        cliponaxis=False
    )
    fig.update_yaxes(range=[0, 100], title="Retention %")
    st.plotly_chart(style_fig(fig, 320, True), use_container_width=True, config={"displayModeBar": False})

    c1, c2, c3 = st.columns([1, 1.15, 1])
    with c1:
        if len(cat_counts):
            cat_df = cat_counts.rename("Unique VINs").reset_index().rename(columns={"Category Calc": "Service Category"})
            cat_df["% of Serviced VINs"] = cat_df["Unique VINs"] / max(total_serviced, 1) * 100
            fig = px.bar(
                cat_df.sort_values("Unique VINs"),
                x="Unique VINs", y="Service Category", orientation="h",
                text=cat_df.sort_values("Unique VINs")["% of Serviced VINs"].map(lambda x: f"{x:.1f}%"),
                title=f"SERVICE CATEGORY PENETRATION — {service_year}",
            )
            fig.update_traces(marker_color=[TEAL if x=="PMS" else ORANGE if x=="GR" else RED if x=="Accident Repair" else BLUE for x in cat_df.sort_values("Unique VINs")["Service Category"]])
            st.plotly_chart(style_fig(fig, 330), use_container_width=True, config={"displayModeBar": False})
            st.caption("Categories can overlap: one VIN may have PMS and GR in the selected period.")
        else:
            st.info("No service category data for these filters.")

    with c2:
        fig = go.Figure()
        fig.add_trace(go.Scatter(
            x=monthly["Sold Month Cohort"],
            y=monthly["Retention %"] * 100,
            mode="lines+markers+text",
            text=monthly["Retention %"].map(lambda x: f"{x:.1%}"),
            textposition="top center",
            line=dict(color=BLUE, width=3),
            marker=dict(size=8),
            name="Retention %",
        ))
        fig.update_layout(title="RETENTION % TREND BY SOLD COHORT")
        st.plotly_chart(style_fig(fig, 330, True), use_container_width=True, config={"displayModeBar": False})

    with c3:
        st.markdown('<div class="panel-title">SUMMARY</div>', unsafe_allow_html=True)
        s1, s2 = st.columns(2)
        with s1:
            st.metric("VINs Sold", f"{total_sold:,}")
            st.metric("Overall Retention", f"{retention:.1%}")
        with s2:
            st.metric("VINs Serviced", f"{total_serviced:,}")
            st.metric("Not Reported", f"{not_reported:,}")
        st.markdown(
            f"""
            <div class="small-note">
            <b>Sold period:</b> {sold_month if sold_month != "All" else "All months"} {sold_year}<br>
            <b>Service period:</b> {service_month if service_month != "All" else "All months"} {service_year}<br>
            <b>Dealer:</b> {branch}<br>
            <b>Location:</b> {city}<br>
            <b>Service location:</b> {service_location}
            </div>
            """,
            unsafe_allow_html=True,
        )

    with st.expander("⬇️ View / export exact VINs behind this dashboard"):
        vin_detail = sf.copy()
        vin_detail["Serviced in Selected Period"] = vin_detail["VIN"].isin(serviced_vins)
        keep = [c for c in ["VIN","Customer Name","City","Sold Date","Model","Sold Branch","Brand Calc"] if c in vin_detail.columns]
        vin_detail = vin_detail[keep + ["Serviced in Selected Period"]].sort_values(["Serviced in Selected Period","VIN"])
        st.dataframe(vin_detail, use_container_width=True, hide_index=True)
        st.download_button(
            "Download dashboard VIN list (CSV)",
            vin_detail.to_csv(index=False).encode("utf-8-sig"),
            "dashboard_vin_list.csv",
            "text/csv",
        )

if active_page == PAGE_LABELS[2]:
    st.subheader("🧩 Cohort Matrix — When Sold VINs Came Back for Service")
    st.caption("Each cell = unique VINs serviced in that reported/service month ÷ unique VINs sold in that sold-month cohort. Color scale is fixed from 0% to 100%: very low retention is very light blue and higher retention becomes progressively darker blue.")

    # Matrix uses selected sold filters but all service months in selected service year.
    base_service = service[service["VIN"].isin(population_vins) & service["Service Year Calc"].eq(service_year)].copy()
    if service_location != "All":
        base_service = base_service[base_service["Service Location"].eq(service_location)]
    if service_type != "All":
        base_service = base_service[base_service["Service Type"].eq(service_type)]
    if category != "All":
        base_service = base_service[base_service["Category Calc"].eq(category)]

    matrix = []
    row_labels = []
    for sm in range(1, 13):
        cohort = sf[sf["Sold Month No Calc"].eq(sm)]
        cv = set(cohort["VIN"])
        denom = len(cv)
        vals = []
        for svc_m in range(1, 13):
            n = base_service.loc[
                base_service["VIN"].isin(cv) & base_service["Service Month No Calc"].eq(svc_m),
                "VIN"
            ].nunique()
            vals.append((n / denom * 100) if denom else 0)
        matrix.append(vals)
        row_labels.append(pd.Timestamp(2000, sm, 1).strftime("%b") + f"-{str(sold_year)[-2:]}")


    st.markdown("### 📌 Cohort Matrix Summary")
    st.info(
        "This summary shows how many VINs were sold in the selected cohorts, "
        "how many came back for service in the selected service period, and the TOTAL RETENTION %."
    )

    # Cohort Matrix summary
    cohort_summary_rows = []
    for sm in range(1, 13):
        cohort = sf[sf["Sold Month No Calc"].eq(sm)]
        cv = set(cohort["VIN"])
        sold_n = len(cv)
        serviced_n = base_service.loc[base_service["VIN"].isin(cv), "VIN"].nunique()
        cohort_summary_rows.append({
            "Sold Month Cohort": pd.Timestamp(2000, sm, 1).strftime("%b") + f"-{str(sold_year)[-2:]}",
            "VINs Sold": sold_n,
            "Reported / Serviced VINs": serviced_n,
            "Retention": serviced_n / sold_n if sold_n else 0,
        })

    cohort_summary = pd.DataFrame(cohort_summary_rows)
    cohort_total_sold = int(cohort_summary["VINs Sold"].sum())
    cohort_total_serviced = base_service["VIN"].nunique()
    cohort_retention = cohort_total_serviced / cohort_total_sold if cohort_total_sold else 0
    valid = cohort_summary[cohort_summary["VINs Sold"] > 0]
    best = valid.loc[valid["Retention"].idxmax()] if len(valid) else None
    worst = valid.loc[valid["Retention"].idxmin()] if len(valid) else None

    cs1, cs2, cs3, cs4, cs5 = st.columns(5)
    with cs1: kpi("Cohort VINs Sold", f"{cohort_total_sold:,}", f"Sold year {sold_year}", BLUE, "🚗")
    with cs2: kpi("Reported / Serviced VINs", f"{cohort_total_serviced:,}", f"Service year {service_year}", GREEN, "👥")
    with cs3: kpi("TOTAL RETENTION %", f"{cohort_retention:.1%}", retention_band(cohort_retention), retention_color(cohort_retention), "📈")
    with cs4: kpi("Best Cohort", f"{best['Sold Month Cohort']}<br>({best['Retention']:.1%})" if best is not None else "-", "Highest retention", GREEN, "🏆")
    with cs5: kpi("Lowest Cohort", f"{worst['Sold Month Cohort']}<br>({worst['Retention']:.1%})" if worst is not None else "-", "Needs attention", RED, "📉")

    fig = go.Figure(data=go.Heatmap(
        z=matrix,
        x=list(pd.date_range("2025-01-01", periods=12, freq="MS").strftime("%b")),
        y=row_labels,
        colorscale=[
            [0.00, "#F7FBFF"],
            [0.10, "#E7F1FA"],
            [0.20, "#D2E5F5"],
            [0.30, "#B7D5EC"],
            [0.40, "#8FBCDD"],
            [0.50, "#6AA5CF"],
            [0.60, "#4A8CC2"],
            [0.70, "#3174B3"],
            [0.80, "#205D9B"],
            [0.90, "#15477D"],
            [1.00, "#0A2F5A"],
        ],
        zmin=0,
        zmax=100,
        text=[[f"{v:.1f}%" if v else "" for v in row] for row in matrix],
        texttemplate="%{text}",
        textfont=dict(size=11),
        colorbar=dict(title="Retention %", tickmode="array", tickvals=[0,10,20,30,40,50,60,70,80,90,100], ticktext=["0","10","20","30","40","50","60","70","80","90","100"]),
        hovertemplate="Sold cohort: %{y}<br>Service month: %{x}<br>Retention: %{z:.1f}%<extra></extra>",
    ))
    fig.update_layout(
        title=f"REPORTED / SERVICE YEAR {service_year}",
        height=590,
        xaxis=dict(side="top", title=""),
        yaxis=dict(title="SOLD MONTH COHORT", autorange="reversed"),
    )
    st.plotly_chart(style_fig(fig, 590), use_container_width=True, config={"displayModeBar": True})

    st.markdown("#### Month-wise Cohort Summary")
    cohort_summary_display = cohort_summary.copy()
    cohort_summary_display["Retention %"] = cohort_summary_display["Retention"].map(lambda x: f"{x:.1%}")
    cohort_summary_display["Status"] = cohort_summary_display["Retention"].map(retention_band)
    cohort_summary_display = cohort_summary_display.drop(columns=["Retention"])

    total_row = pd.DataFrame([{
        "Sold Month Cohort": "TOTAL",
        "VINs Sold": cohort_total_sold,
        "Reported / Serviced VINs": cohort_total_serviced,
        "Retention %": f"{cohort_retention:.1%}",
        "Status": retention_band(cohort_retention),
    }])
    cohort_summary_display = pd.concat([cohort_summary_display, total_row], ignore_index=True)

    def _retention_pct_style(value):
        try:
            pct = float(str(value).replace("%", "").strip())
        except Exception:
            return ""
        if pct >= 80:
            return "background-color: #C6EFCE; color: #006100; font-weight: 800;"
        if pct >= 50:
            return "background-color: #FCE4B2; color: #9C5700; font-weight: 800;"
        return "background-color: #FFC7CE; color: #9C0006; font-weight: 800;"

    cohort_summary_styled = cohort_summary_display.style.map(
        _retention_pct_style,
        subset=["Retention %"],
    )

    st.dataframe(
        cohort_summary_styled,
        use_container_width=True,
        hide_index=True,
        height=500,
    )





def throughput_growth_note(compare_year, compare_count, growth):
    if pd.isna(growth):
        return (
            f'vs {compare_year}: {compare_count:,} • '
            '<span style="color:#7A8799;font-weight:800;">— New</span>'
        )
    if growth > 0.05:
        return (
            f'vs {compare_year}: {compare_count:,} • '
            f'<span style="color:{GREEN};font-weight:900;white-space:nowrap;"><span style="font-size:15px;line-height:1;vertical-align:-1px;">▲</span> {growth:.1f}%</span>'
        )
    if growth < -0.05:
        return (
            f'vs {compare_year}: {compare_count:,} • '
            f'<span style="color:{RED};font-weight:900;white-space:nowrap;"><span style="font-size:15px;line-height:1;vertical-align:-1px;">▼</span> {abs(growth):.1f}%</span>'
        )
    return (
        f'vs {compare_year}: {compare_count:,} • '
        f'<span style="color:{ORANGE};font-weight:900;white-space:nowrap;"><span style="font-size:14px;line-height:1;vertical-align:-1px;">▶</span> {growth:.1f}%</span>'
    )


def service_svg_icon(category, size=34):
    # Clean single-color SVG icons to match the supplied dashboard screenshot.
    colors = {
        "PMS": "#18A7B8",
        "GR": "#FF7A00",
        "Accident Repair": "#D94A3A",
        "Inspection": "#0D63D6",
    }
    color = colors.get(category, "#2F75B5")

    if category == "PMS":
        # Wrench
        return f"""
        <svg width="{size}" height="{size}" viewBox="0 0 64 64" aria-hidden="true">
          <path fill="{color}" d="M52.8 7.2c-5.7-5.7-14.6-6.9-21.6-2.9l9.3 9.3-8.9 8.9-9.3-9.3c-4 7-2.8 15.9 2.9 21.6 1.4 1.4 3 2.6 4.7 3.4L9.2 58.9c-2.6 2.6-6.8 2.6-9.4 0s-2.6-6.8 0-9.4l20.7-20.7c-.8-1.7-2-3.3-3.4-4.7-5.7-5.7-14.6-6.9-21.6-2.9L14.8 40.5l8.9-8.9L4.4 12.3c-4 7-2.8 15.9 2.9 21.6 5.8 5.8 14.9 6.9 22 2.7L50.6 58c2.6 2.6 6.8 2.6 9.4 0s2.6-6.8 0-9.4L38.7 27.3c4.2-7.1 3.1-16.2-2.7-22z"/>
        </svg>
        """
    if category == "GR":
        # Gear
        return f"""
        <svg width="{size}" height="{size}" viewBox="0 0 64 64" aria-hidden="true">
          <path fill="{color}" d="M57.7 36.7c.2-1.5.3-3.1.3-4.7s-.1-3.2-.4-4.7l6-4.6-6-10.4-7.2 2.9c-2.4-2-5.1-3.6-8-4.7L41.3 3H29.2l-1.1 7.5c-3 1.1-5.7 2.7-8.1 4.7l-7.1-2.9-6 10.4 6 4.6c-.2 1.5-.4 3.1-.4 4.7s.1 3.2.4 4.7l-6 4.6 6 10.4 7.1-2.9c2.4 2 5.1 3.6 8.1 4.7l1.1 7.5h12.1l1.1-7.5c2.9-1.1 5.6-2.7 8-4.7l7.2 2.9 6-10.4-5.9-4.6zM35.2 43.5A11.5 11.5 0 1 1 35.2 20a11.5 11.5 0 0 1 0 23.5z"/>
        </svg>
        """
    if category == "Accident Repair":
        # Car with impact marks
        return f"""
        <svg width="{size}" height="{size}" viewBox="0 0 72 64" aria-hidden="true">
          <path fill="{color}" d="M14 39l4-13c1-4 4-6 8-6h20c4 0 7 2 8 6l4 13c4 1 7 5 7 9v7h-6v5h-8v-5H21v5h-8v-5H7v-7c0-4 3-8 7-9zm9-12-3 11h36l-3-11c-.5-2-2-3-4-3H27c-2 0-3.5 1-4 3zm-6 18a5 5 0 1 0 0 10 5 5 0 0 0 0-10zm38 0a5 5 0 1 0 0 10 5 5 0 0 0 0-10z"/>
          <path fill="{color}" d="M9 4l3 7 7 3-7 3-3 7-3-7-7-3 7-3 3-7zm48 0l2 5 5 2-5 2-2 5-2-5-5-2 5-2 2-5z"/>
        </svg>
        """
    # Inspection clipboard + magnifier
    return f"""
    <svg width="{size}" height="{size}" viewBox="0 0 64 64" aria-hidden="true">
      <path fill="{color}" d="M18 6h9a7 7 0 0 1 14 0h9a4 4 0 0 1 4 4v27h-6V12H16v40h20v6H12a4 4 0 0 1-4-4V10a4 4 0 0 1 4-4h6zm5 6h18V8h-4a5 5 0 0 1-10 0h-4v4z"/>
      <path fill="{color}" d="M24 23h18v5H24zm0 10h14v5H24z"/>
      <path fill="{color}" d="M46 37a11 11 0 1 1 0 22 11 11 0 0 1 0-22zm0 5a6 6 0 1 0 0 12 6 6 0 0 0 0-12z"/>
      <path fill="{color}" d="M53 53l11 11-4 4-11-11z"/>
    </svg>
    """


@st.cache_data(show_spinner=False)
def asset_data_uri(filename):
    import base64
    p = APP_DIR / "assets" / filename
    if not p.exists():
        return ""
    return "data:image/png;base64," + base64.b64encode(p.read_bytes()).decode("ascii")

def tp_growth_html(g):
    """Growth indicator using CSS shapes, not font arrow glyphs.

    This keeps the direction marker thick and the color reliable even when
    broad Streamlit/theme CSS overrides text fonts or inherited colors.
    """
    if pd.isna(g):
        return '<span class="tp-growth tp-growth-new">— New</span>'
    if g > 0.05:
        return (f'<span class="tp-growth tp-growth-up">'
                f'<span class="tp-growth-shape" aria-hidden="true"></span>'
                f'<span class="tp-growth-value">{g:.1f}%</span></span>')
    if g < -0.05:
        return (f'<span class="tp-growth tp-growth-down">'
                f'<span class="tp-growth-shape" aria-hidden="true"></span>'
                f'<span class="tp-growth-value">{abs(g):.1f}%</span></span>')
    return (f'<span class="tp-growth tp-growth-flat">'
            f'<span class="tp-growth-shape" aria-hidden="true"></span>'
            f'<span class="tp-growth-value">{g:.1f}%</span></span>')

def tp_growth_text(g):
    if pd.isna(g):
        return "— New"
    if g > 0.05:
        return f"▲ {g:.1f}%"
    if g < -0.05:
        return f"▼ {abs(g):.1f}%"
    return f"▶ {g:.1f}%"


if active_page == PAGE_LABELS[3]:
    # ------------------------------------------------------------------
    # THROUGHPUT TRACKING — styled to match the supplied reference image
    # ------------------------------------------------------------------
    st.markdown("""
    <style>
    .tp-title-wrap{
        display:flex;align-items:center;gap:12px;margin:0 0 4px 0;
    }
    .tp-title{
        font-size:27px;font-weight:850;color:#2F75B5;line-height:1.0;
    }
    .tp-subtitle{
        font-size:12px;color:#718099;margin:2px 0 10px 48px;
    }
    .tp-bar{
        background:#062D5D;color:white;font-weight:800;text-align:center;
        padding:6px 8px;border-radius:4px 4px 0 0;font-size:12px;
        letter-spacing:.1px;margin-top:8px;
    }
    .tp-kpi{
        background:#FFFFFF;border:1px solid #B6C6D8;border-radius:8px;
        min-height:92px;padding:10px 12px;display:flex;align-items:center;
        gap:10px;box-shadow:0 1px 5px rgba(6,45,93,.06);
    }
    .tp-kpi img{width:38px;height:42px;object-fit:contain;flex:0 0 38px;}
    .tp-kpi-lbl{font-size:10px;font-weight:800;color:#39495E;line-height:1.15;}
    .tp-kpi-val{font-size:24px;font-weight:900;color:#2F75B5;line-height:1.0;margin:4px 0 6px;}
    .tp-kpi-note{font-size:10.5px;color:#718099;line-height:1.25;}
    .tp-table-wrap{overflow-x:auto;background:#FFFFFF;}
    .tp-table{width:100%;border-collapse:collapse;font-size:10.5px;background:#FFFFFF;}
    .tp-table th{
        background:#062D5D;color:#fff;border:1px solid #B6C6D8;
        padding:5px 4px;text-align:center;font-weight:800;
    }
    .tp-table td{
        border:1px solid #DDE5EF;padding:4px 5px;text-align:center;color:#152238;
    }
    .tp-cat-cell{
        min-width:116px;vertical-align:middle!important;background:#FFFFFF!important;
        font-weight:800;color:#152238!important;
    }
    .tp-cat-cell img{width:34px;height:38px;object-fit:contain;display:block;margin:0 auto 7px;}
    .tp-total td{background:#062D5D!important;font-weight:850;color:#062D5D!important;}
    .tp-summary{
        width:100%;border-collapse:collapse;background:#FFFFFF;font-size:11px;
    }
    .tp-summary th{
        background:#062D5D;color:#fff;padding:7px 5px;border:1px solid #B6C6D8;text-align:center;
    }
    .tp-summary td{
        border:1px solid #DDE5EF;padding:7px 6px;text-align:center;color:#152238;
    }
    .tp-summary .total td{background:#062D5D;font-weight:850;color:#062D5D;}
    .tp-insights{
        border:1px solid #DDE5EF;border-top:0;background:#FFFFFF;padding:10px 12px 6px;
    }
    .tp-insight{
        display:flex;gap:9px;align-items:flex-start;margin:8px 0;font-size:11px;color:#39495E;
    }
    .tp-insight .ico{font-size:20px;line-height:1;}


    

/* v3.39 VERIFIED v3.18 VISUAL CONSISTENCY */
.stApp, .stApp * {
    font-family:"Source Sans Pro", sans-serif !important;
}
.stApp {
    background:#F4F7FB !important;
    color:#152238 !important;
}
section[data-testid="stSidebar"] {
    background:#EDF3FA !important;
}
section[data-testid="stSidebar"] *,
.stApp p, .stApp li, .stApp label {
    color:#152238 !important;
}
h1,h2,h3,h4,h5,h6 {
    color:#062D5D !important;
}
.main-title,.panel-title,.tp-bar,.lost-bar {
    background:#062D5D !important;
    color:#FFFFFF !important;
}
.main-title *,.panel-title *,.tp-bar *,.lost-bar * {
    color:#FFFFFF !important;
}
.subtitle,.small-note,.tp-subtitle,.lost-sub,
.kpi-note,.tp-kpi-note,.lost-card-note {
    color:#718099 !important;
}
.kpi,.tp-kpi,.lost-card,.lost-action-box,
div[data-testid="stMetric"] {
    background:#FFFFFF !important;
    border-color:#DCE4EE !important;
    color:#152238 !important;
}
.kpi-label,.tp-kpi-lbl,.lost-card-label {
    color:#20324A !important;
}
.kpi-value,.tp-kpi-val,.lost-card-value {
    color:#071B3B !important;
}

/* Main section navigation */
div[role="radiogroup"] label {
    background:#FFFFFF !important;
    color:#152238 !important;
    border:1px solid #DCE4EE !important;
}
div[role="radiogroup"] label * {
    color:#152238 !important;
}
div[role="radiogroup"] label[data-checked="true"],
div[role="radiogroup"] label:has(input:checked) {
    background:#062D5D !important;
    border-color:#2F75B5 !important;
}
div[role="radiogroup"] label[data-checked="true"] *,
div[role="radiogroup"] label:has(input:checked) * {
    color:#FFFFFF !important;
}

/* Filters and inputs */
[data-testid="stSelectbox"] label *,
[data-testid="stMultiSelect"] label *,
[data-testid="stDateInput"] label *,
[data-testid="stNumberInput"] label *,
[data-testid="stTextInput"] label * {
    color:#152238 !important;
    -webkit-text-fill-color:#152238 !important;
}
div[data-baseweb="select"] > div,
div[data-baseweb="input"] > div,
[data-baseweb="base-input"],
input, textarea {
    background:#FFFFFF !important;
    color:#152238 !important;
    border-color:#DCE4EE !important;
}
div[data-baseweb="select"] > div > div {
    background:transparent !important;
}
div[data-baseweb="select"] span,
div[data-baseweb="select"] p,
div[data-baseweb="select"] input {
    color:#152238 !important;
    -webkit-text-fill-color:#152238 !important;
    opacity:1 !important;
}
div[data-baseweb="select"] svg {
    color:#062D5D !important;
    fill:#062D5D !important;
    opacity:1 !important;
}
[data-baseweb="tag"] {
    background:#EAF3FA !important;
    border:1px solid #B6C6D8 !important;
}
[data-baseweb="tag"] *,
[data-baseweb="tag"] svg {
    color:#062D5D !important;
    fill:#062D5D !important;
    -webkit-text-fill-color:#062D5D !important;
}
[data-baseweb="popover"],
[data-baseweb="popover"] ul,
[role="listbox"] {
    background:#FFFFFF !important;
    color:#152238 !important;
    border-color:#DCE4EE !important;
}
[role="option"],[role="option"] * {
    color:#152238 !important;
    -webkit-text-fill-color:#152238 !important;
}
[role="option"]:hover,
[role="option"][aria-selected="true"] {
    background:#EAF3FA !important;
    color:#062D5D !important;
}

/* Buttons, uploaders, popovers, expanders */
.stButton > button,
.stDownloadButton > button,
[data-testid="stFileUploader"] button {
    background:#062D5D !important;
    color:#FFFFFF !important;
    border:1px solid #2F75B5 !important;
}
.stButton > button *,
.stDownloadButton > button *,
[data-testid="stFileUploader"] button * {
    color:#FFFFFF !important;
}
[data-testid="stFileUploaderDropzone"],
[data-testid="stFileUploader"] section {
    background:#FFFFFF !important;
    border-color:#B6C6D8 !important;
}
[data-testid="stFileUploaderDropzone"] *,
[data-testid="stFileUploader"] section * {
    color:#152238 !important;
}
[data-testid="stExpander"] details,
[data-testid="stExpander"] summary,
[data-testid="stPopover"] button {
    background:#FFFFFF !important;
    border-color:#DCE4EE !important;
    color:#152238 !important;
}
[data-testid="stExpander"] summary *,
[data-testid="stPopover"] button * {
    color:#152238 !important;
}

/* Tables */
[data-testid="stDataFrame"] {
    background:#FFFFFF !important;
    border-color:#DDE5EF !important;
}
[data-testid="stDataFrame"] [role="columnheader"],
[data-testid="stDataFrame"] [role="columnheader"] * {
    background:#062D5D !important;
    color:#FFFFFF !important;
    -webkit-text-fill-color:#FFFFFF !important;
}
[data-testid="stDataFrame"] [role="gridcell"],
[data-testid="stDataFrame"] [role="gridcell"] * {
    background:#FFFFFF !important;
    color:#152238 !important;
    -webkit-text-fill-color:#152238 !important;
}
table { background:#FFFFFF !important; color:#152238 !important; }
table thead th { background:#062D5D !important; color:#FFFFFF !important; }
table tbody td { background:#FFFFFF !important; color:#152238 !important; }

/* Throughput-specific HTML tables */
.tp-table,.tp-summary,.tp-table-wrap {
    background:#FFFFFF !important;
    color:#152238 !important;
}
.tp-table th,.tp-summary th {
    background:#062D5D !important;
    color:#FFFFFF !important;
}
.tp-table td,.tp-summary td {
    background:#FFFFFF !important;
    color:#13243C !important;
}
.tp-total td,.tp-summary .total td {
    background:#EAF3FA !important;
    color:#062D5D !important;
}

/* Semantic status colors retained from v3.18 */
.good { color:#079447 !important; }
.warn { color:#F28C28 !important; }
.bad  { color:#D94A3A !important; }



/* v3.40 — EXACT FIX FOR USER-IDENTIFIED TABLE AREAS */

/* Month-on-month cohort table */
.monthly-v318-wrap {{
    width:100% !important;
    overflow-x:auto !important;
    background:#FFFFFF !important;
}}
table.monthly-v318 {{
    width:100% !important;
    border-collapse:collapse !important;
    background:#FFFFFF !important;
    color:#152238 !important;
    font-size:14px !important;
}}
table.monthly-v318 thead th {{
    background:#062D5D !important;
    color:#FFFFFF !important;
    -webkit-text-fill-color:#FFFFFF !important;
    border:1px solid #B6C6D8 !important;
    padding:7px 8px !important;
    font-weight:700 !important;
    text-align:left !important;
}}
table.monthly-v318 tbody tr td,
table.monthly-v318 tbody tr:nth-child(odd) td,
table.monthly-v318 tbody tr:nth-child(even) td {{
    background:#FFFFFF !important;
    color:#152238 !important;
    -webkit-text-fill-color:#152238 !important;
    border:1px solid #B6C6D8 !important;
    padding:6px 8px !important;
}}
table.monthly-v318 tbody td:not(:first-child) {{
    text-align:right !important;
}}

/* Throughput main + YOY summary tables */
table.tp-table tbody tr td,
table.tp-table tbody tr:nth-child(odd) td,
table.tp-table tbody tr:nth-child(even) td,
table.tp-summary tbody tr td,
table.tp-summary tbody tr:nth-child(odd) td,
table.tp-summary tbody tr:nth-child(even) td {{
    background:#FFFFFF !important;
    color:#13243C !important;
    -webkit-text-fill-color:#13243C !important;
}}
table.tp-table thead tr th,
table.tp-summary thead tr th {{
    background:#062D5D !important;
    color:#FFFFFF !important;
    -webkit-text-fill-color:#FFFFFF !important;
}}
table.tp-table td.tp-cat-cell,
table.tp-table tbody tr:nth-child(odd) td.tp-cat-cell,
table.tp-table tbody tr:nth-child(even) td.tp-cat-cell {{
    background:#FFFFFF !important;
    color:#13243C !important;
    -webkit-text-fill-color:#13243C !important;
}}
table.tp-table tbody tr.tp-total td,
table.tp-table tbody tr.tp-total:nth-child(odd) td,
table.tp-table tbody tr.tp-total:nth-child(even) td,
table.tp-summary tbody tr.total td,
table.tp-summary tbody tr.total:nth-child(odd) td,
table.tp-summary tbody tr.total:nth-child(even) td {{
    background:#EAF3FA !important;
    color:#062D5D !important;
    -webkit-text-fill-color:#062D5D !important;
    font-weight:850 !important;
}}

</style>
    """, unsafe_allow_html=True)

    serviced_img = asset_data_uri("serviced.png")
    pms_img = asset_data_uri("pms.png")
    gr_img = asset_data_uri("gr.png")
    accident_img = asset_data_uri("accident.png")
    inspection_img = asset_data_uri("inspection.png")

    icon_uri = {
        "PMS": pms_img,
        "GR": gr_img,
        "Accident Repair": accident_img,
        "Inspection": inspection_img,
    }

    st.markdown(
        '<div class="tp-title-wrap"><div style="font-size:29px;color:#062D5D;">▣</div>'
        '<div class="tp-title">Throughput Tracking</div></div>'
        '<div class="tp-subtitle">Service Throughput & Selected-Year Comparison</div>',
        unsafe_allow_html=True
    )

    years_tp = sorted(int(y) for y in service["Service Year Calc"].dropna().unique())
    default_current = service_year if service_year in years_tp else max(years_tp)

    f1,f2,f3,f4,f5,f6 = st.columns([1,1,1.45,1.15,1.1,1.15], gap="small")
    with f1:
        tp_year = st.selectbox(
            "Current Service Year",
            years_tp,
            index=years_tp.index(default_current),
            key="tp_current_year_v311",
        )

    compare_options = [y for y in years_tp if y != tp_year]
    lower = [y for y in compare_options if y < tp_year]
    default_compare = max(lower) if lower else (compare_options[0] if compare_options else tp_year)
    with f2:
        # Multi-year comparison selector, using the same compact popover pattern as Service Months.
        compare_state_prefix = "tp_compare_year_pick_"
        compare_init_key = "tp_compare_year_selector_initialized_current_year"

        if st.session_state.get(compare_init_key) != tp_year:
            for _y in years_tp:
                st.session_state[compare_state_prefix + str(_y)] = (_y == default_compare)
            st.session_state[compare_init_key] = tp_year

        _compare_selected_count = sum(
            bool(st.session_state.get(compare_state_prefix + str(_y), False))
            for _y in compare_options
        )
        st.markdown(
            '<div style="font-size:14px;margin-bottom:4px;color:#152238;">Compare With Year</div>',
            unsafe_allow_html=True,
        )
        with st.popover(
            f"Select Years ({_compare_selected_count})",
            use_container_width=True,
            help="Select one or more comparison years. The latest selected year is used as the primary detailed comparison; all selected years appear in the multi-year summary.",
        ):
            st.caption("Select comparison year(s).")
            _yb1, _yb2 = st.columns(2)
            with _yb1:
                if st.button("Select All", key="tp_compare_year_select_all", use_container_width=True):
                    for _y in compare_options:
                        st.session_state[compare_state_prefix + str(_y)] = True
                    st.rerun()
            with _yb2:
                if st.button("Clear All", key="tp_compare_year_clear_all", use_container_width=True):
                    for _y in compare_options:
                        st.session_state[compare_state_prefix + str(_y)] = False
                    st.rerun()
            for _y in sorted(compare_options, reverse=True):
                st.checkbox(str(_y), key=compare_state_prefix + str(_y))

        compare_years = [
            _y for _y in compare_options
            if bool(st.session_state.get(compare_state_prefix + str(_y), False))
        ]
        if not compare_years and compare_options:
            compare_years = [default_compare]
        compare_years = sorted(compare_years, reverse=True)
        compare_year = compare_years[0] if compare_years else tp_year

    month_names = list(pd.date_range("2025-01-01", periods=12, freq="MS").strftime("%b"))
    month_num_map = {m:i for i,m in enumerate(month_names, start=1)}
    # For the latest/current year, default to months actually available in data.
    months_available_current = sorted(
        int(m) for m in service.loc[
            service["Service Year Calc"].eq(tp_year), "Service Month No Calc"
        ].dropna().unique()
    )
    default_month_names = [month_names[m-1] for m in months_available_current] or month_names

    with f3:
        # Compact month selector: month choices are visible ONLY inside this dropdown/popover.
        # When closed, no Jan/Feb/Mar chips or month list are shown on the dashboard.
        month_state_prefix = "tp_month_pick_"
        init_year_key = "tp_month_selector_initialized_year"

        if st.session_state.get(init_year_key) != tp_year:
            for _m in month_names:
                st.session_state[month_state_prefix + _m] = (_m in default_month_names)
            st.session_state[init_year_key] = tp_year

        selected_count = sum(
            bool(st.session_state.get(month_state_prefix + _m, False))
            for _m in month_names
        )

        st.markdown(
            '<div style="font-size:14px;margin-bottom:4px;color:#152238;">Service Months</div>',
            unsafe_allow_html=True,
        )

        with st.popover(
            f"Select Months ({selected_count})",
            use_container_width=True,
            help="Month choices stay hidden inside this dropdown. The same selected months are applied to both comparison years.",
        ):
            st.caption("Select the months to compare in both years.")

            _b1, _b2 = st.columns(2)
            with _b1:
                if st.button("Select All", key="tp_month_select_all", use_container_width=True):
                    for _m in month_names:
                        st.session_state[month_state_prefix + _m] = True
                    st.rerun()
            with _b2:
                if st.button("Clear All", key="tp_month_clear_all", use_container_width=True):
                    for _m in month_names:
                        st.session_state[month_state_prefix + _m] = False
                    st.rerun()

            _mc1, _mc2, _mc3 = st.columns(3)
            _month_cols = [_mc1, _mc2, _mc3]
            for _idx, _m in enumerate(month_names):
                with _month_cols[_idx % 3]:
                    st.checkbox(
                        _m,
                        key=month_state_prefix + _m,
                    )

        tp_months = [
            _m for _m in month_names
            if bool(st.session_state.get(month_state_prefix + _m, False))
        ]
    with f4:
        tp_location = st.selectbox(
            "Service Location",
            ["All"] + choices(service["Service Location"]),
            key="tp_location_v311",
        )
    with f5:
        tp_type = st.selectbox(
            "Service Type",
            ["All"] + choices(service["Service Type"]),
            key="tp_type_v311",
        )
    with f6:
        tp_category = st.selectbox(
            "Service Category",
            ["All"] + choices(service["Category Calc"]),
            key="tp_category_v311",
        )

    selected_month_nums = [month_num_map[m] for m in tp_months]

    tp = service.copy()
    if tp_location != "All":
        tp = tp[tp["Service Location"].eq(tp_location)]
    if tp_type != "All":
        tp = tp[tp["Service Type"].eq(tp_type)]
    if tp_category != "All":
        tp = tp[tp["Category Calc"].eq(tp_category)]
    if selected_month_nums:
        tp = tp[tp["Service Month No Calc"].isin(selected_month_nums)]
    else:
        tp = tp.iloc[0:0].copy()

    cur = tp[tp["Service Year Calc"].eq(tp_year)].copy()
    cmp = tp[tp["Service Year Calc"].eq(compare_year)].copy()

    categories = ["PMS","GR","Accident Repair","Inspection"]
    cur_total = cur["VIN"].nunique()
    cmp_total = cmp["VIN"].nunique()
    total_growth = yoy_growth(cur_total, cmp_total)

    # KPI row
    cards = []
    cards.append(("Unique VINs Serviced", serviced_img, cur_total, cmp_total, total_growth))
    for cat in categories:
        c = cur.loc[cur["Category Calc"].eq(cat),"VIN"].nunique()
        p = cmp.loc[cmp["Category Calc"].eq(cat),"VIN"].nunique()
        cards.append((cat, icon_uri[cat], c, p, yoy_growth(c,p)))

    kcols = st.columns(5, gap="small")
    for i,(label,img_uri,curr,prev,growth) in enumerate(cards):
        with kcols[i]:
            note = (
                f'vs {compare_year}: {prev:,} • {tp_growth_html(growth)}'
            )
            st.markdown(
                f'<div class="tp-kpi">'
                f'<img src="{img_uri}" alt="">'
                f'<div><div class="tp-kpi-lbl">{label} ({tp_year})</div>'
                f'<div class="tp-kpi-val">{curr:,}</div>'
                f'<div class="tp-kpi-note">{note}</div></div>'
                f'</div>',
                unsafe_allow_html=True
            )

    selected_period = ", ".join(tp_months) if tp_months else "No months selected"
    _compare_year_text = ", ".join(str(y) for y in compare_years) if compare_years else "None"
    st.caption(
        f"Comparison period applied equally to all selected years: {selected_period} — Current {tp_year}; Compare {_compare_year_text}. "
        f"Primary detailed comparison: {compare_year}."
    )

    # Multi-year comparison summary: every selected comparison year is shown here.
    if compare_years:
        _my_rows = []
        for _cat in ["ALL"] + categories:
            _cur_src = cur if _cat == "ALL" else cur[cur["Category Calc"].eq(_cat)]
            _cur_n = _cur_src["VIN"].nunique()
            _row = {"Service Category": "Unique VINs Serviced" if _cat == "ALL" else _cat, str(tp_year): int(_cur_n)}
            for _y in compare_years:
                _ysrc = tp[tp["Service Year Calc"].eq(_y)]
                if _cat != "ALL":
                    _ysrc = _ysrc[_ysrc["Category Calc"].eq(_cat)]
                _yn = _ysrc["VIN"].nunique()
                _row[str(_y)] = int(_yn)
                _g = yoy_growth(_cur_n, _yn)
                _row[f"Growth vs {_y}"] = "N/A" if pd.isna(_g) else f"{_g:+.1f}%"
            _my_rows.append(_row)
        st.markdown('<div class="tp-bar">MULTI-YEAR COMPARISON SUMMARY</div>', unsafe_allow_html=True)
        st.dataframe(pd.DataFrame(_my_rows), use_container_width=True, hide_index=True)

    # Detailed table data
    rows_by_cat = {}
    totals_by_cat = {}
    for cat in categories:
        ccat = cur[cur["Category Calc"].eq(cat)]
        pcat = cmp[cmp["Category Calc"].eq(cat)]
        ctot = ccat["VIN"].nunique()
        ptot = pcat["VIN"].nunique()
        cat_rows = []
        for m in selected_month_nums:
            mname = month_names[m-1]
            c = ccat.loc[ccat["Service Month No Calc"].eq(m),"VIN"].nunique()
            p = pcat.loc[pcat["Service Month No Calc"].eq(m),"VIN"].nunique()
            cat_rows.append({
                "month":mname,
                "c":int(c),
                "cp":(c/ctot*100.0) if ctot else 0.0,
                "p":int(p),
                "pp":(p/ptot*100.0) if ptot else 0.0,
                "g":yoy_growth(c,p),
            })
        rows_by_cat[cat] = cat_rows
        totals_by_cat[cat] = {
            "c":int(ctot),"p":int(ptot),"g":yoy_growth(ctot,ptot)
        }

    left,right = st.columns([1.55,1], gap="small")

    with left:
        st.markdown(
            f'<div class="tp-bar">SERVICE THROUGHPUT TRACKING (MONTH WISE) — {tp_year} vs {compare_year}</div>',
            unsafe_allow_html=True
        )
        html = [
            '<div class="tp-table-wrap"><table class="tp-table">',
            '<thead><tr>',
            '<th>Service Category</th><th>Service Month</th>',
            f'<th>{tp_year}<br>Count</th><th>{tp_year}<br>% of Total</th>',
            f'<th>{compare_year}<br>Count</th><th>{compare_year}<br>% of Total</th>',
            '<th>Growth %</th>',
            '</tr></thead><tbody>'
        ]

        for cat in categories:
            cat_rows = rows_by_cat[cat]
            rowspan = max(len(cat_rows),1)
            for ri,row in enumerate(cat_rows):
                html.append('<tr>')
                if ri == 0:
                    html.append(
                        f'<td class="tp-cat-cell" rowspan="{rowspan}">'
                        f'<img src="{icon_uri[cat]}" alt="">'
                        f'<div>{cat}</div></td>'
                    )
                html.append(
                    f'<td>{row["month"]}</td>'
                    f'<td>{row["c"]:,}</td><td>{row["cp"]:.1f}%</td>'
                    f'<td>{row["p"]:,}</td><td>{row["pp"]:.1f}%</td>'
                    f'<td>{tp_growth_html(row["g"])}</td></tr>'
                )
            tot = totals_by_cat[cat]
            html.append(
                f'<tr class="tp-total"><td>Total {cat}</td><td></td>'
                f'<td>{tot["c"]:,}</td><td>100%</td>'
                f'<td>{tot["p"]:,}</td><td>100%</td>'
                f'<td>{tp_growth_html(tot["g"])}</td></tr>'
            )
        html.append('</tbody></table></div>')
        st.markdown("".join(html), unsafe_allow_html=True)

    with right:
        st.markdown(
            f'<div class="tp-bar">YEAR ON YEAR SUMMARY ({tp_year} vs {compare_year})</div>',
            unsafe_allow_html=True
        )
        html = [
            '<table class="tp-summary"><thead><tr>',
            '<th>Service Category</th>',
            f'<th>{tp_year} Count</th><th>{compare_year} Count</th>',
            '<th>Growth %</th></tr></thead><tbody>'
        ]
        for cat in categories:
            t = totals_by_cat[cat]
            html.append(
                f'<tr><td style="text-align:left;padding-left:12px;">'
                f'<img src="{icon_uri[cat]}" style="width:22px;height:25px;object-fit:contain;vertical-align:middle;margin-right:8px;">'
                f'{cat}</td><td>{t["c"]:,}</td><td>{t["p"]:,}</td>'
                f'<td>{tp_growth_html(t["g"])}</td></tr>'
            )
        html.append(
            f'<tr class="total"><td>TOTAL</td><td>{cur_total:,}</td>'
            f'<td>{cmp_total:,}</td><td>{tp_growth_html(total_growth)}</td></tr>'
        )
        html.append('</tbody></table>')
        st.markdown("".join(html), unsafe_allow_html=True)

        st.markdown(
            '<div class="tp-bar">MONTH WISE TOTAL THROUGHPUT (ALL CATEGORIES)</div>',
            unsafe_allow_html=True
        )
        chart_rows=[]
        for m in selected_month_nums:
            mname=month_names[m-1]
            c=cur.loc[cur["Service Month No Calc"].eq(m),"VIN"].nunique()
            p=cmp.loc[cmp["Service Month No Calc"].eq(m),"VIN"].nunique()
            chart_rows.append({
                "Month":mname,
                str(tp_year):c,
                str(compare_year):p,
                "Growth":0 if pd.isna(yoy_growth(c,p)) else yoy_growth(c,p)
            })
        chart_df=pd.DataFrame(chart_rows)

        fig=go.Figure()
        if len(chart_df):
            fig.add_trace(go.Bar(
                x=chart_df["Month"],y=chart_df[str(tp_year)],
                name=str(tp_year),marker_color="#082D63",
                text=chart_df[str(tp_year)],textposition="outside"
            ))
            fig.add_trace(go.Bar(
                x=chart_df["Month"],y=chart_df[str(compare_year)],
                name=str(compare_year),marker_color="#8BBCE5",
                text=chart_df[str(compare_year)],textposition="outside"
            ))
            fig.add_trace(go.Scatter(
                x=chart_df["Month"],y=chart_df["Growth"],
                name="Growth %",mode="lines+markers",
                line=dict(color="#079447",width=2),
                marker=dict(size=6),yaxis="y2"
            ))
        fig.update_layout(
            barmode="group",height=300,margin=dict(l=20,r=20,t=35,b=25),
            paper_bgcolor="white",plot_bgcolor="white",
            legend=dict(orientation="h",y=1.13,x=.28),
            yaxis=dict(title="Count",gridcolor="#E6ECF2"),
            yaxis2=dict(title="Growth %",overlaying="y",side="right",showgrid=False),
            font=dict(family=APP_FONT,size=10,color="#152238")
        )
        st.plotly_chart(fig,use_container_width=True,config={"displayModeBar":False})

        st.markdown('<div class="tp-bar">KEY INSIGHTS</div>',unsafe_allow_html=True)

        growths=[]
        for cat in categories:
            g=totals_by_cat[cat]["g"]
            if pd.notna(g):
                growths.append((cat,g))

        if pd.isna(total_growth):
            overall_text=f"No comparable baseline exists for {compare_year}."
            overall_icon="ℹ️"
        elif total_growth >= 0:
            overall_text=f"Overall throughput increased by {total_growth:.1f}% in {tp_year} compared with {compare_year}."
            overall_icon="📈"
        else:
            overall_text=f"Overall throughput decreased by {abs(total_growth):.1f}% in {tp_year} compared with {compare_year}."
            overall_icon="📉"

        insights=[(overall_icon,overall_text)]
        if growths:
            best_cat,best_g=max(growths,key=lambda x:x[1])
            worst_cat,worst_g=min(growths,key=lambda x:x[1])
            insights.append(("✅",f"{best_cat} has the strongest comparison at {tp_growth_text(best_g)}."))
            insights.append(("🎯",f"{worst_cat} is the lowest comparison at {tp_growth_text(worst_g)}."))
        insights.append(("📅",f"Both years use the same selected months: {selected_period}."))

        html=['<div class="tp-insights">']
        for ico,text in insights:
            html.append(
                f'<div class="tp-insight"><div class="ico">{ico}</div><div>{text}</div></div>'
            )
        html.append('</div>')
        st.markdown("".join(html),unsafe_allow_html=True)


    st.markdown('<div class="tp-bar">MISSED PMS TRACKING — VIN LEVEL</div>', unsafe_allow_html=True)
    st.caption(
        f"Exact VIN comparison: a PMS VIN from {compare_year} is treated as RETURNED when that VIN has a PMS visit in ANY of the selected month(s) of {tp_year}. "
        "This prevents a genuine PMS return from being marked missed only because the service shifted to another selected month."
    )

    _pms = service[service["Category Calc"].eq("PMS")].copy()
    _miss_rows, _miss_details = [], []

    # Current-year PMS pool across the WHOLE selected period (not month-for-month only).
    _current_period_pms = _pms[
        (_pms["Service Year Calc"] == tp_year) &
        (_pms["Service Month No Calc"].isin(selected_month_nums))
    ].copy()
    if tp_location != "All":
        _current_period_pms = _current_period_pms[_current_period_pms["Service Location"].eq(tp_location)]
    if tp_type != "All":
        _current_period_pms = _current_period_pms[_current_period_pms["Service Type"].eq(tp_type)]
    _current_period_pms_vins = set(_current_period_pms["VIN"].dropna().astype(str))
    _current_period_pms_latest = (
        _current_period_pms.sort_values("Service Date").groupby("VIN", as_index=False).tail(1)
        if len(_current_period_pms) else pd.DataFrame()
    )
    _current_pms_date_map = (
        _current_period_pms_latest.set_index(_current_period_pms_latest["VIN"].astype(str))["Service Date"].to_dict()
        if len(_current_period_pms_latest) else {}
    )

    for _m in selected_month_nums:
        _mn = month_names[_m-1]
        _b = _pms[(_pms["Service Year Calc"] == compare_year) & (_pms["Service Month No Calc"] == _m)].copy()
        if tp_location != "All":
            _b = _b[_b["Service Location"].eq(tp_location)]
        if tp_type != "All":
            _b = _b[_b["Service Type"].eq(tp_type)]

        _bv = set(_b["VIN"].dropna().astype(str))
        _rv, _mv = _bv & _current_period_pms_vins, _bv - _current_period_pms_vins
        _miss_rows.append({
            "Month": _mn, f"{compare_year} PMS VINs": len(_bv),
            f"Returned PMS in {tp_year}": len(_rv), "Missed PMS VINs": len(_mv),
            "PMS Return %": (len(_rv)/len(_bv)*100 if _bv else 0),
            "Missed PMS %": (len(_mv)/len(_bv)*100 if _bv else 0),
        })

        if _mv:
            _d = _b[_b["VIN"].astype(str).isin(_mv)].sort_values("Service Date").groupby("VIN", as_index=False).tail(1).copy()
            _d["Comparison Month"] = _mn
            _d["Comparison PMS Date"] = _d["Service Date"]
            _d["Current-Year PMS Date"] = _d["VIN"].astype(str).map(_current_pms_date_map)
            _any = service[(service["Service Year Calc"] == tp_year) & (service["Service Month No Calc"].isin(selected_month_nums)) & service["VIN"].astype(str).isin(_mv)].copy()
            if tp_location != "All":
                _any = _any[_any["Service Location"].eq(tp_location)]
            _anyv = set(_any["VIN"].dropna().astype(str))
            _d["Visited Workshop Selected Period"] = _d["VIN"].astype(str).isin(_anyv).map({True:"Yes",False:"No"})
            _latest = _any.sort_values("Service Date").groupby("VIN", as_index=False).tail(1) if len(_any) else pd.DataFrame()
            if len(_latest):
                _dm = _latest.set_index(_latest["VIN"].astype(str))["Service Date"].to_dict()
                _cm = _latest.set_index(_latest["VIN"].astype(str))["Category Calc"].to_dict()
                _d["Current-Year Other Visit Date"] = _d["VIN"].astype(str).map(_dm)
                _d["Current-Year Other Category"] = _d["VIN"].astype(str).map(_cm)
            else:
                _d["Current-Year Other Visit Date"] = pd.NaT
                _d["Current-Year Other Category"] = ""
            _d["Missed PMS Status"] = np.where(_d["Visited Workshop Selected Period"].eq("Yes"), "Missed PMS — Other Service Visit", "Missed PMS — No Workshop Visit")
            _miss_details.append(_d)

    _ms = pd.DataFrame(_miss_rows)
    if len(_ms):
        _show = _ms.copy()
        _show["PMS Return %"] = _show["PMS Return %"].map(lambda x:f"{x:.1f}%")
        _show["Missed PMS %"] = _show["Missed PMS %"].map(lambda x:f"{x:.1f}%")
        st.dataframe(_show, use_container_width=True, hide_index=True)

    if _miss_details:
        _md = pd.concat(_miss_details, ignore_index=True)
        _cols = ["Comparison Month","VIN","Vehicle No","Customer Name","Comparison PMS Date","Current-Year PMS Date","Service Location","KMS","Visited Workshop Selected Period","Current-Year Other Visit Date","Current-Year Other Category","Missed PMS Status"]
        _cols = [x for x in _cols if x in _md.columns]
        with st.expander(f"View Missed PMS VIN Details ({len(_md):,} records)"):
            st.dataframe(_md[_cols], use_container_width=True, hide_index=True, height=350)
        st.download_button("⬇️ Download Missed PMS VINs", _md[_cols].to_csv(index=False).encode("utf-8-sig"), file_name=f"missed_pms_{compare_year}_vs_{tp_year}.csv", mime="text/csv", key="missed_pms_dl_v318")

    st.markdown('<div class="tp-bar">EXTERNAL / NOT IN OUR SOLD DATA — ACQUISITION & RETENTION</div>', unsafe_allow_html=True)
    st.caption("External VIN means a serviced VIN that is not present in the uploaded SOLD_DATA. It may have been sold by another dealer, but the workbook cannot prove the original selling dealer.")

    # Independent filter row for External VIN analysis.
    _ext_years = sorted([int(y) for y in service["Service Year Calc"].dropna().unique()])
    _ext_default_cur = tp_year if tp_year in _ext_years else (_ext_years[-1] if _ext_years else tp_year)
    _ext_default_cmp = compare_year if compare_year in _ext_years else (_ext_years[-2] if len(_ext_years) > 1 else _ext_default_cur)

    _ef1, _ef2, _ef3, _ef4, _ef5, _ef6 = st.columns([1.05,1.05,1.55,1.35,1.35,1.35], gap="small")

    with _ef1:
        ext_tp_year = st.selectbox(
            "Current Service Year",
            _ext_years,
            index=_ext_years.index(_ext_default_cur),
            key="ext_current_service_year_v319",
        )

    with _ef2:
        _cmp_choices = [y for y in _ext_years if y != ext_tp_year] or _ext_years
        _cmp_idx = _cmp_choices.index(_ext_default_cmp) if _ext_default_cmp in _cmp_choices else max(0, len(_cmp_choices)-1)
        ext_compare_year = st.selectbox(
            "Compare With Year",
            _cmp_choices,
            index=_cmp_idx,
            key="ext_compare_service_year_v319",
        )

    with _ef3:
        with st.popover(f"Select Months ({len(selected_month_nums)})", use_container_width=True):
            st.caption("Choose the same months to compare in both years.")
            _ext_all = st.checkbox("Select All", value=len(selected_month_nums) == 12, key="ext_months_all_v319")
            _ext_month_flags = []
            for _i, _mn in enumerate(month_names, start=1):
                _ext_month_flags.append(
                    st.checkbox(_mn, value=(_i in selected_month_nums), key=f"ext_month_{_i}_v319")
                )
        if _ext_all:
            ext_selected_month_nums = list(range(1, 13))
        else:
            ext_selected_month_nums = [i for i, flag in enumerate(_ext_month_flags, start=1) if flag]
            if not ext_selected_month_nums:
                ext_selected_month_nums = selected_month_nums[:] if selected_month_nums else list(range(1, 13))

    with _ef4:
        _ext_locations = ["All"] + sorted([x for x in service["Service Location"].dropna().astype(str).unique() if x])
        ext_tp_location = st.selectbox(
            "Service Location",
            _ext_locations,
            index=_ext_locations.index(tp_location) if tp_location in _ext_locations else 0,
            key="ext_service_location_v319",
        )

    with _ef5:
        _ext_types = ["All"] + sorted([x for x in service["Service Type"].dropna().astype(str).unique() if x])
        ext_tp_type = st.selectbox(
            "Service Type",
            _ext_types,
            index=_ext_types.index(tp_type) if tp_type in _ext_types else 0,
            key="ext_service_type_v319",
        )

    with _ef6:
        _ext_categories = ["All"] + sorted([x for x in service["Category Calc"].dropna().astype(str).unique() if x])
        ext_tp_category = st.selectbox(
            "Service Category",
            _ext_categories,
            index=_ext_categories.index(tp_category) if tp_category in _ext_categories else 0,
            key="ext_service_category_v319",
        )

    st.caption(
        "External VIN comparison period: "
        + ", ".join([month_names[m-1] for m in ext_selected_month_nums])
        + f" — {ext_tp_year} vs {ext_compare_year}"
    )

    _our = set(sold["VIN"].dropna().astype(str))
    _ext = service[~service["VIN"].astype(str).isin(_our)].copy()
    if ext_tp_location != "All":
        _ext = _ext[_ext["Service Location"].eq(ext_tp_location)]
    if ext_tp_type != "All":
        _ext = _ext[_ext["Service Type"].eq(ext_tp_type)]
    if ext_tp_category != "All":
        _ext = _ext[_ext["Category Calc"].eq(ext_tp_category)]

    _eb = _ext[(_ext["Service Year Calc"] == ext_compare_year) & _ext["Service Month No Calc"].isin(ext_selected_month_nums)].copy()
    _ec = _ext[(_ext["Service Year Calc"] == ext_tp_year) & _ext["Service Month No Calc"].isin(ext_selected_month_nums)].copy()
    _ebv, _ecv = set(_eb["VIN"].dropna().astype(str)), set(_ec["VIN"].dropna().astype(str))
    _erv = _ebv & _ecv
    _erp = len(_erv)/len(_ebv)*100 if _ebv else 0

    _first = _ext.sort_values("Service Date").groupby("VIN", as_index=False).head(1)
    _new = _first[(_first["Service Year Calc"] == tp_year) & _first["Service Month No Calc"].isin(selected_month_nums)].copy()

    _e1,_e2,_e3,_e4 = st.columns(4)
    _e1.metric(f"External VINs {ext_tp_year}", f"{len(_ecv):,}")
    _e2.metric(f"New External VINs {ext_tp_year}", f"{_new['VIN'].nunique():,}")
    _e3.metric(f"{ext_compare_year} External VIN Base", f"{len(_ebv):,}")
    _e4.metric("External VIN Retention", f"{_erp:.1f}%", f"{len(_erv):,} returned")

    _catrows = []
    for _cat in ["PMS","GR","Accident Repair","Inspection"]:
        _x = set(_eb.loc[_eb["Category Calc"].eq(_cat),"VIN"].dropna().astype(str))
        _y = set(_ec.loc[_ec["Category Calc"].eq(_cat),"VIN"].dropna().astype(str))
        _r = _x & _y
        _catrows.append({"Service Category":_cat,f"{ext_compare_year} External VINs":len(_x),f"{ext_tp_year} External VINs":len(_y),"Same VINs Returned":len(_r),"Retention %":f"{(len(_r)/len(_x)*100 if _x else 0):.1f}%"})
    st.dataframe(pd.DataFrame(_catrows), use_container_width=True, hide_index=True)

    if len(_eb):
        _ed = _eb.sort_values("Service Date").groupby("VIN", as_index=False).tail(1).copy()
        _ed["Returned in Current Period"] = _ed["VIN"].astype(str).isin(_erv).map({True:"Yes",False:"No"})
        _ec_latest = _ec.sort_values("Service Date").groupby("VIN", as_index=False).tail(1)
        if len(_ec_latest):
            _edm = _ec_latest.set_index(_ec_latest["VIN"].astype(str))["Service Date"].to_dict()
            _ecm = _ec_latest.set_index(_ec_latest["VIN"].astype(str))["Category Calc"].to_dict()
            _ed["Current Period Service Date"] = _ed["VIN"].astype(str).map(_edm)
            _ed["Current Period Category"] = _ed["VIN"].astype(str).map(_ecm)
        _ed["External Retention Status"] = np.where(_ed["Returned in Current Period"].eq("Yes"),"Retained External VIN","External VIN Not Returned")
        _ecols = ["VIN","Vehicle No","Service Date","Service Location","Category Calc","Service Type","KMS","Returned in Current Period","Current Period Service Date","Current Period Category","External Retention Status"]
        _ecols = [x for x in _ecols if x in _ed.columns]
        with st.expander(f"View External VIN Retention Details ({len(_ed):,} VINs)"):
            st.dataframe(_ed[_ecols], use_container_width=True, hide_index=True, height=350)
        st.download_button("⬇️ Download External VIN Retention Data", _ed[_ecols].to_csv(index=False).encode("utf-8-sig"), file_name=f"external_vin_retention_{ext_compare_year}_vs_{ext_tp_year}.csv", mime="text/csv", key="external_retention_dl_v318")

    if len(_new):
        _ncols = [x for x in ["VIN","Vehicle No","Service Date","Service Location","Category Calc","Service Type","KMS"] if x in _new.columns]
        st.download_button(f"⬇️ Download New External VINs {ext_tp_year}", _new[_ncols].to_csv(index=False).encode("utf-8-sig"), file_name=f"new_external_vins_{ext_tp_year}.csv", mime="text/csv", key="new_external_dl_v318")

if active_page == PAGE_LABELS[4]:
    st.subheader("Lost VIN & Immediate Action Dashboard")
    st.caption("Freshly recalculated from real service dates. This fixes the workbook's no-service/date-0 issue.")

    # Latest service record per VIN from ALL service history.
    svc_sorted = service.sort_values(["VIN", "Service Date"])
    last_service = svc_sorted.drop_duplicates("VIN", keep="last")[
        [c for c in ["VIN","Service Date","Service Location","KMS","Vehicle No","RO Number"] if c in svc_sorted.columns]
    ].copy()
    rename = {"Service Date":"Last Service Date Calc","Service Location":"Last Service Location Calc","KMS":"Latest KMS Calc"}
    last_service = last_service.rename(columns=rename)

    action = sold.copy().merge(last_service, on="VIN", how="left")
    today = pd.Timestamp(date.today())
    action["Next Due Date Calc"] = np.where(
        action["Last Service Date Calc"].notna(),
        action["Last Service Date Calc"] + pd.Timedelta(days=365),
        action["Sold Date"] + pd.Timedelta(days=365),
    )
    action["Next Due Date Calc"] = pd.to_datetime(action["Next Due Date Calc"])
    action["Overdue Days Calc"] = (today - action["Next Due Date Calc"]).dt.days

    def status_from_days(d):
        if pd.isna(d): return "Unknown"
        if d > 180: return "Critical Lost"
        if d > 90: return "High Risk"
        if d > 0: return "Overdue"
        if d >= -30: return "Due Soon"
        return "Not Yet Due"

    action["Status Calc"] = action["Overdue Days Calc"].map(status_from_days)
    action["Relationship Calc"] = np.where(
        action["Last Service Date Calc"].isna(),
        "Never Serviced",
        np.where((today - action["Last Service Date Calc"]).dt.days <= 365, "Active / Retained", "Inactive")
    )
    action["Recommended Action Calc"] = action["Status Calc"].map({
        "Critical Lost":"Call customer now / recovery campaign",
        "High Risk":"Priority recovery call",
        "Overdue":"Advisor follow-up",
        "Due Soon":"Service reminder / appointment",
        "Not Yet Due":"No immediate action",
        "Unknown":"Review data",
    })

    af = action[action["Sold Year Calc"].eq(sold_year)]
    if branch != "All": af = af[af["Sold Branch"].eq(branch)]
    if city != "All": af = af[af["City"].eq(city)]
    if brand != "All": af = af[af["Brand Calc"].eq(brand)]
    if model != "All": af = af[af["Model"].eq(model)]

    status_order = ["Critical Lost","High Risk","Overdue","Due Soon","Not Yet Due"]
    counts = af["Status Calc"].value_counts()

    # Use native Streamlit metrics here instead of HTML KPI cards.
    # This avoids Markdown rendering raw <div> tags on some Streamlit versions.
    cc = st.columns(5, gap="small")
    metric_labels = {
        "Critical Lost":"CRITICAL LOST",
        "High Risk":"HIGH RISK",
        "Overdue":"OVERDUE",
        "Due Soon":"DUE SOON",
        "Not Yet Due":"NOT YET DUE",
    }
    for col, stat in zip(cc, status_order):
        with col:
            st.metric(
                metric_labels[stat],
                f"{int(counts.get(stat,0)):,}",
                "VINs",
                delta_color="off",
            )

    focus = st.multiselect("Show statuses", status_order, default=["Critical Lost","High Risk","Overdue","Due Soon"])
    show = af[af["Status Calc"].isin(focus)].copy() if focus else af.copy()
    show = show.sort_values(["Overdue Days Calc"], ascending=False)
    action_cols = [c for c in [
        "VIN","Customer Name","City","Model","Sold Branch","Sold Date",
        "Last Service Date Calc","Last Service Location Calc","Latest KMS Calc",
        "Next Due Date Calc","Overdue Days Calc","Status Calc","Relationship Calc","Recommended Action Calc"
    ] if c in show.columns]
    st.dataframe(show[action_cols], use_container_width=True, hide_index=True, height=520)
    st.download_button(
        "⬇️ Download action list",
        show[action_cols].to_csv(index=False).encode("utf-8-sig"),
        f"action_list_{sold_year}.csv",
        "text/csv",
    )


if active_page == PAGE_LABELS[5]:
    st.subheader("🚨 Lost VIN Details — Recommended Action")
    st.caption("Shows exact lost/risk VINs, last service information, and the action your team should take.")

    svc_last = service.sort_values(["VIN", "Service Date"]).drop_duplicates("VIN", keep="last").copy()
    last_cols = [c for c in ["VIN","Service Date","Service Location","KMS","Vehicle No","RO Number","Service Type","Category Calc"] if c in svc_last.columns]
    svc_last = svc_last[last_cols].rename(columns={
        "Service Date":"Last Service Date",
        "Service Location":"Last Service Location",
        "KMS":"Latest KMS",
        "Service Type":"Last Service Type",
        "Category Calc":"Last Service Category",
    })

    # SOLD_DATA may already contain calculated service-history columns.
    # Drop them before merging the fresh service history so pandas does not create _x/_y names.
    sold_for_lost = sold.copy()
    existing_service_cols = [
        "Last Service Date", "Last Service Location", "Latest KMS",
        "Vehicle No", "RO Number", "Last Service Type", "Last Service Category"
    ]
    sold_for_lost = sold_for_lost.drop(
        columns=[c for c in existing_service_cols if c in sold_for_lost.columns],
        errors="ignore",
    )
    lost_df = sold_for_lost.merge(svc_last, on="VIN", how="left")

    # Guarantee the columns exist even when a monthly service file omits optional fields.
    if "Last Service Date" not in lost_df.columns:
        lost_df["Last Service Date"] = pd.NaT
    if "Last Service Location" not in lost_df.columns:
        lost_df["Last Service Location"] = ""
    if "Latest KMS" not in lost_df.columns:
        lost_df["Latest KMS"] = np.nan

    today_lost = pd.Timestamp(date.today())
    lost_df["Days Since Sale"] = (today_lost - lost_df["Sold Date"]).dt.days
    lost_df["Days Since Last Service"] = np.where(
        lost_df["Last Service Date"].notna(),
        (today_lost - lost_df["Last Service Date"]).dt.days,
        np.nan
    )

    _never_serviced = lost_df["Last Service Date"].isna()
    _dsl_lost = lost_df["Days Since Last Service"]
    _dss_lost = lost_df["Days Since Sale"]
    lost_df["Lost Status"] = np.select(
        [
            _never_serviced & (_dss_lost > 365),
            _never_serviced,
            _dsl_lost > 730,
            _dsl_lost > 545,
            _dsl_lost > 365,
        ],
        [
            "Never Serviced — Lost",
            "Never Serviced — Follow Up",
            "Critical Lost",
            "Lost",
            "High Risk",
        ],
        default="Retained / Active",
    )

    action_map = {
        "Critical Lost":"Senior advisor recovery call; identify defection reason; offer targeted win-back campaign.",
        "Lost":"Priority recovery call; check competitor servicing, relocation, vehicle sale, or unresolved complaint.",
        "Never Serviced — Lost":"Immediate recovery call; verify customer/vehicle status; offer first-service inspection/campaign.",
        "High Risk":"Proactive service reminder; confirm mileage and book an appointment.",
        "Never Serviced — Follow Up":"Welcome/service reminder call; confirm first service due date and preferred workshop.",
        "Retained / Active":"Continue normal retention follow-up."
    }

    lost_df["Recommended Action"] = lost_df["Lost Status"].map(action_map)

    lf = lost_df[lost_df["Sold Year Calc"].eq(sold_year)].copy()
    if branch != "All": lf = lf[lf["Sold Branch"].eq(branch)]
    if city != "All": lf = lf[lf["City"].eq(city)]
    if brand != "All": lf = lf[lf["Brand Calc"].eq(brand)]
    if model != "All": lf = lf[lf["Model"].eq(model)]

    lost_statuses = ["Critical Lost","Lost","Never Serviced — Lost","High Risk","Never Serviced — Follow Up"]
    lost_only = lf[lf["Lost Status"].isin(lost_statuses)].copy()
    cnt = lost_only["Lost Status"].value_counts()

    l1,l2,l3,l4,l5 = st.columns(5)
    with l1: kpi("Critical Lost", f"{int(cnt.get('Critical Lost',0)):,}", "Highest priority", RED, "🚨")
    with l2: kpi("Lost", f"{int(cnt.get('Lost',0)):,}", "Win-back required", RED, "❌")
    with l3: kpi("Never Serviced — Lost", f"{int(cnt.get('Never Serviced — Lost',0)):,}", "No service history", ORANGE, "🕳️")
    with l4: kpi("High Risk", f"{int(cnt.get('High Risk',0)):,}", "Prevent loss now", ORANGE, "⚠️")
    with l5: kpi("Total Lost / Risk VINs", f"{len(lost_only):,}", f"Sold year {sold_year}", NAVY, "🔎")

    chosen = st.multiselect("Show lost statuses", lost_statuses, default=["Critical Lost","Lost","Never Serviced — Lost","High Risk"])
    show_lost = lost_only[lost_only["Lost Status"].isin(chosen)].copy() if chosen else lost_only.copy()
    priority = {"Critical Lost":1,"Lost":2,"Never Serviced — Lost":3,"High Risk":4,"Never Serviced — Follow Up":5}
    show_lost["Priority"] = show_lost["Lost Status"].map(priority).fillna(99)
    show_lost = show_lost.sort_values(["Priority","Days Since Last Service","Days Since Sale"], ascending=[True,False,False])

    cols = [c for c in [
        "VIN","Customer Name","City","Model","Sold Branch","Sold Date","Vehicle No",
        "Last Service Date","Last Service Location","Last Service Type","Last Service Category",
        "Latest KMS","Days Since Last Service","Days Since Sale","Lost Status","Recommended Action"
    ] if c in show_lost.columns]
    st.dataframe(show_lost[cols], use_container_width=True, hide_index=True, height=560)

    st.download_button(
        "⬇️ Download Lost VIN Action List",
        show_lost[cols].to_csv(index=False).encode("utf-8-sig"),
        f"lost_vin_action_list_{sold_year}.csv",
        "text/csv",
        use_container_width=True
    )

    st.markdown("#### Action Guide")
    st.dataframe(pd.DataFrame([
        ["Critical Lost","No service for more than ~24 months","Senior advisor recovery + win-back campaign"],
        ["Lost","No service for ~18–24 months","Priority recovery call + appointment offer"],
        ["Never Serviced — Lost","Sold >12 months ago; no service found","Verify customer/vehicle + first-service recovery"],
        ["High Risk","No service for ~12–18 months","Immediate service reminder + appointment"],
        ["Never Serviced — Follow Up","Recent sale; no service yet","Welcome call + confirm first service due"]
    ], columns=["Status","Meaning","Recommended Action"]), use_container_width=True, hide_index=True)


if active_page == PAGE_LABELS[6]:
    st.subheader("Exact VIN / Vehicle / Customer Search")
    q = st.text_input("Search VIN, vehicle number, customer, model, RO number, city or consultant", placeholder="Type at least 2 characters...")
    if len(q.strip()) >= 2:
        query = q.strip().lower()
        sold_search_cols = [c for c in ["VIN","Customer Name","City","Model","Consultant","Sold Branch"] if c in sold.columns]
        smask = sold[sold_search_cols].astype(str).apply(lambda s: s.str.lower().str.contains(query, regex=False, na=False)).any(axis=1)
        sold_hits = sold[smask].copy()

        service_search_cols = [c for c in ["VIN","Vehicle No","RO Number","Customer Name","Service Location","Service Type"] if c in service.columns]
        vmask = service[service_search_cols].astype(str).apply(lambda s: s.str.lower().str.contains(query, regex=False, na=False)).any(axis=1)
        service_hits = service[vmask].copy()

        hit_vins = set(sold_hits["VIN"]) | set(service_hits["VIN"])
        st.markdown(f"**{len(hit_vins):,} matching VIN(s)**")

        for vin in list(sorted(hit_vins))[:25]:
            sr = sold[sold["VIN"].eq(vin)].head(1)
            vh = service[service["VIN"].eq(vin)].sort_values("Service Date", ascending=False)
            label = vin
            if len(sr):
                label += f" — {sr.iloc[0].get('Model','')} — {sr.iloc[0].get('Customer Name','')}"
            with st.expander(label):
                if len(sr):
                    r = sr.iloc[0]
                    a,b,c,d = st.columns(4)
                    a.metric("Sold Date", r["Sold Date"].strftime("%d-%b-%Y") if pd.notna(r["Sold Date"]) else "-")
                    b.metric("Branch", r.get("Sold Branch","-"))
                    c.metric("City", r.get("City","-"))
                    d.metric("Service Visits", f"{vh['RO Number'].nunique() if 'RO Number' in vh.columns else len(vh):,}")
                cols = [c for c in ["Service Date","RO Number","Vehicle No","Service Location","Service Type","Category Calc","KMS","Invoice Amount Inc Tax"] if c in vh.columns]
                st.dataframe(vh[cols], use_container_width=True, hide_index=True)
    else:
        st.info("Enter VIN, vehicle number, customer name, model, RO number, city or consultant.")

if active_page == PAGE_LABELS[7]:
    st.subheader("Data Quality & Accuracy Checks")
    sold_dupes = int(sold_all["VIN"].duplicated(keep=False).sum())
    missing_sold_vin = int(clean_str(sold_raw.get("VIN", pd.Series(dtype=str))).eq("").sum()) if "VIN" in sold_raw.columns else 0
    missing_service_vin = int(clean_str(service_raw.get("VIN", pd.Series(dtype=str))).eq("").sum()) if "VIN" in service_raw.columns else 0
    unmatched_service = service.loc[~service["VIN"].isin(set(sold["VIN"])), "VIN"].nunique()
    duplicate_ro = 0
    if "RO Number" in service.columns:
        duplicate_ro = int(service.duplicated(subset=["VIN","RO Number","Service Date"], keep=False).sum())

    qcols = st.columns(5)
    qcols[0].metric("Sold rows", f"{len(sold_all):,}")
    qcols[1].metric("Service rows", f"{len(service):,}")
    qcols[2].metric("Duplicate sold VIN rows", f"{sold_dupes:,}")
    qcols[3].metric("Unmatched service VINs", f"{unmatched_service:,}")
    qcols[4].metric("Repeated VIN/RO/date rows", f"{duplicate_ro:,}")

    st.markdown("#### Why this app is more reliable than the workbook dashboard")
    st.markdown(
        """
        - **Sold Year, Sold Month, cohort and retention are recalculated from the real Sold Date.**
        - **Last Service Date is recalculated from SERVICE_DATA**, not from cached `MAXIFS` results.
        - A vehicle with **no service is handled as missing**, not as Excel date 0.
        - All KPI percentages use **unique VINs**, not repair-order line counts.
        - Service category counts are unique VIN counts and may overlap; the chart does not pretend they are mutually exclusive.
        - Every dashboard number can be traced to an exportable VIN list.
        """
    )

    issues = []
    if sold_dupes: issues.append(f"{sold_dupes:,} sold rows belong to VINs that appear more than once.")
    if missing_sold_vin: issues.append(f"{missing_sold_vin:,} sold rows have a blank VIN.")
    if missing_service_vin: issues.append(f"{missing_service_vin:,} service rows have a blank VIN.")
    if unmatched_service: issues.append(f"{unmatched_service:,} unique service VINs do not match SOLD_DATA.")
    if duplicate_ro: issues.append(f"{duplicate_ro:,} service rows repeat the same VIN + RO + service date (often multiple invoice/customer lines).")
    if issues:
        st.warning("\n\n".join("• " + x for x in issues))
    else:
        st.success("No major structural data-quality issues detected.")

st.markdown("---")
st.caption("Dashboard engine: Python + Streamlit. All calculations are performed from raw workbook data at runtime. Retention colors: 80%+ Green • 50–79% Orange • below 50% Red.")


if active_page == PAGE_LABELS[8]:
    st.markdown("""
    <style>
    .lost-title{font-size:25px;font-weight:900;color:#062D5D;margin-bottom:0;}
    .lost-sub{font-size:12px;color:#718099;margin-bottom:8px;}
    .lost-bar{background:#062D5D;color:#fff;font-size:12px;font-weight:850;text-align:center;padding:6px;border-radius:4px 4px 0 0;margin-top:8px;}
    .lost-card{background:#FFFFFF;border:1px solid #B6C6D8;border-radius:8px;padding:11px 12px;min-height:100px;box-shadow:0 1px 5px rgba(6,45,93,.05);}
    .lost-card-label{font-size:10px;font-weight:850;color:#39495E;}
    .lost-card-value{font-size:25px;font-weight:900;line-height:1.1;margin:5px 0;}
    .lost-card-note{font-size:9.5px;color:#718099;}
    .lost-action-box{background:#FFFFFF;border:1px solid #B6C6D8;padding:9px 12px;font-size:11px;}
    

/* v3.39 VERIFIED v3.18 VISUAL CONSISTENCY */
.stApp, .stApp * {
    font-family:"Source Sans Pro", sans-serif !important;
}
.stApp {
    background:#F4F7FB !important;
    color:#152238 !important;
}
section[data-testid="stSidebar"] {
    background:#EDF3FA !important;
}
section[data-testid="stSidebar"] *,
.stApp p, .stApp li, .stApp label {
    color:#152238 !important;
}
h1,h2,h3,h4,h5,h6 {
    color:#062D5D !important;
}
.main-title,.panel-title,.tp-bar,.lost-bar {
    background:#062D5D !important;
    color:#FFFFFF !important;
}
.main-title *,.panel-title *,.tp-bar *,.lost-bar * {
    color:#FFFFFF !important;
}
.subtitle,.small-note,.tp-subtitle,.lost-sub,
.kpi-note,.tp-kpi-note,.lost-card-note {
    color:#718099 !important;
}
.kpi,.tp-kpi,.lost-card,.lost-action-box,
div[data-testid="stMetric"] {
    background:#FFFFFF !important;
    border-color:#DCE4EE !important;
    color:#152238 !important;
}
.kpi-label,.tp-kpi-lbl,.lost-card-label {
    color:#20324A !important;
}
.kpi-value,.tp-kpi-val,.lost-card-value {
    color:#071B3B !important;
}

/* Main section navigation */
div[role="radiogroup"] label {
    background:#FFFFFF !important;
    color:#152238 !important;
    border:1px solid #DCE4EE !important;
}
div[role="radiogroup"] label * {
    color:#152238 !important;
}
div[role="radiogroup"] label[data-checked="true"],
div[role="radiogroup"] label:has(input:checked) {
    background:#062D5D !important;
    border-color:#2F75B5 !important;
}
div[role="radiogroup"] label[data-checked="true"] *,
div[role="radiogroup"] label:has(input:checked) * {
    color:#FFFFFF !important;
}

/* Filters and inputs */
[data-testid="stSelectbox"] label *,
[data-testid="stMultiSelect"] label *,
[data-testid="stDateInput"] label *,
[data-testid="stNumberInput"] label *,
[data-testid="stTextInput"] label * {
    color:#152238 !important;
    -webkit-text-fill-color:#152238 !important;
}
div[data-baseweb="select"] > div,
div[data-baseweb="input"] > div,
[data-baseweb="base-input"],
input, textarea {
    background:#FFFFFF !important;
    color:#152238 !important;
    border-color:#DCE4EE !important;
}
div[data-baseweb="select"] > div > div {
    background:transparent !important;
}
div[data-baseweb="select"] span,
div[data-baseweb="select"] p,
div[data-baseweb="select"] input {
    color:#152238 !important;
    -webkit-text-fill-color:#152238 !important;
    opacity:1 !important;
}
div[data-baseweb="select"] svg {
    color:#062D5D !important;
    fill:#062D5D !important;
    opacity:1 !important;
}
[data-baseweb="tag"] {
    background:#EAF3FA !important;
    border:1px solid #B6C6D8 !important;
}
[data-baseweb="tag"] *,
[data-baseweb="tag"] svg {
    color:#062D5D !important;
    fill:#062D5D !important;
    -webkit-text-fill-color:#062D5D !important;
}
[data-baseweb="popover"],
[data-baseweb="popover"] ul,
[role="listbox"] {
    background:#FFFFFF !important;
    color:#152238 !important;
    border-color:#DCE4EE !important;
}
[role="option"],[role="option"] * {
    color:#152238 !important;
    -webkit-text-fill-color:#152238 !important;
}
[role="option"]:hover,
[role="option"][aria-selected="true"] {
    background:#EAF3FA !important;
    color:#062D5D !important;
}

/* Buttons, uploaders, popovers, expanders */
.stButton > button,
.stDownloadButton > button,
[data-testid="stFileUploader"] button {
    background:#062D5D !important;
    color:#FFFFFF !important;
    border:1px solid #2F75B5 !important;
}
.stButton > button *,
.stDownloadButton > button *,
[data-testid="stFileUploader"] button * {
    color:#FFFFFF !important;
}
[data-testid="stFileUploaderDropzone"],
[data-testid="stFileUploader"] section {
    background:#FFFFFF !important;
    border-color:#B6C6D8 !important;
}
[data-testid="stFileUploaderDropzone"] *,
[data-testid="stFileUploader"] section * {
    color:#152238 !important;
}
[data-testid="stExpander"] details,
[data-testid="stExpander"] summary,
[data-testid="stPopover"] button {
    background:#FFFFFF !important;
    border-color:#DCE4EE !important;
    color:#152238 !important;
}
[data-testid="stExpander"] summary *,
[data-testid="stPopover"] button * {
    color:#152238 !important;
}

/* Tables */
[data-testid="stDataFrame"] {
    background:#FFFFFF !important;
    border-color:#DDE5EF !important;
}
[data-testid="stDataFrame"] [role="columnheader"],
[data-testid="stDataFrame"] [role="columnheader"] * {
    background:#062D5D !important;
    color:#FFFFFF !important;
    -webkit-text-fill-color:#FFFFFF !important;
}
[data-testid="stDataFrame"] [role="gridcell"],
[data-testid="stDataFrame"] [role="gridcell"] * {
    background:#FFFFFF !important;
    color:#152238 !important;
    -webkit-text-fill-color:#152238 !important;
}
table { background:#FFFFFF !important; color:#152238 !important; }
table thead th { background:#062D5D !important; color:#FFFFFF !important; }
table tbody td { background:#FFFFFF !important; color:#152238 !important; }

/* Throughput-specific HTML tables */
.tp-table,.tp-summary,.tp-table-wrap {
    background:#FFFFFF !important;
    color:#152238 !important;
}
.tp-table th,.tp-summary th {
    background:#062D5D !important;
    color:#FFFFFF !important;
}
.tp-table td,.tp-summary td {
    background:#FFFFFF !important;
    color:#13243C !important;
}
.tp-total td,.tp-summary .total td {
    background:#EAF3FA !important;
    color:#062D5D !important;
}

/* Semantic status colors retained from v3.18 */
.good { color:#079447 !important; }
.warn { color:#F28C28 !important; }
.bad  { color:#D94A3A !important; }



/* v3.40 — EXACT FIX FOR USER-IDENTIFIED TABLE AREAS */

/* Month-on-month cohort table */
.monthly-v318-wrap {{
    width:100% !important;
    overflow-x:auto !important;
    background:#FFFFFF !important;
}}
table.monthly-v318 {{
    width:100% !important;
    border-collapse:collapse !important;
    background:#FFFFFF !important;
    color:#152238 !important;
    font-size:14px !important;
}}
table.monthly-v318 thead th {{
    background:#062D5D !important;
    color:#FFFFFF !important;
    -webkit-text-fill-color:#FFFFFF !important;
    border:1px solid #B6C6D8 !important;
    padding:7px 8px !important;
    font-weight:700 !important;
    text-align:left !important;
}}
table.monthly-v318 tbody tr td,
table.monthly-v318 tbody tr:nth-child(odd) td,
table.monthly-v318 tbody tr:nth-child(even) td {{
    background:#FFFFFF !important;
    color:#152238 !important;
    -webkit-text-fill-color:#152238 !important;
    border:1px solid #B6C6D8 !important;
    padding:6px 8px !important;
}}
table.monthly-v318 tbody td:not(:first-child) {{
    text-align:right !important;
}}

/* Throughput main + YOY summary tables */
table.tp-table tbody tr td,
table.tp-table tbody tr:nth-child(odd) td,
table.tp-table tbody tr:nth-child(even) td,
table.tp-summary tbody tr td,
table.tp-summary tbody tr:nth-child(odd) td,
table.tp-summary tbody tr:nth-child(even) td {{
    background:#FFFFFF !important;
    color:#13243C !important;
    -webkit-text-fill-color:#13243C !important;
}}
table.tp-table thead tr th,
table.tp-summary thead tr th {{
    background:#062D5D !important;
    color:#FFFFFF !important;
    -webkit-text-fill-color:#FFFFFF !important;
}}
table.tp-table td.tp-cat-cell,
table.tp-table tbody tr:nth-child(odd) td.tp-cat-cell,
table.tp-table tbody tr:nth-child(even) td.tp-cat-cell {{
    background:#FFFFFF !important;
    color:#13243C !important;
    -webkit-text-fill-color:#13243C !important;
}}
table.tp-table tbody tr.tp-total td,
table.tp-table tbody tr.tp-total:nth-child(odd) td,
table.tp-table tbody tr.tp-total:nth-child(even) td,
table.tp-summary tbody tr.total td,
table.tp-summary tbody tr.total:nth-child(odd) td,
table.tp-summary tbody tr.total:nth-child(even) td {{
    background:#EAF3FA !important;
    color:#062D5D !important;
    -webkit-text-fill-color:#062D5D !important;
    font-weight:850 !important;
}}

</style>
    """, unsafe_allow_html=True)

    st.markdown('<div class="lost-title">🚨 LOST VIN & IMMEDIATE ACTION DASHBOARD</div>', unsafe_allow_html=True)
    st.markdown('<div class="lost-sub">Identify At-Risk VINs, Focus Areas & Take Action — SOLD + SERVICE DATA</div>', unsafe_allow_html=True)

    # Filters
    lf1, lf2, lf3, lf4, lf5 = st.columns(5, gap="small")
    sold_years_lost = sorted(int(y) for y in sold["Sold Year Calc"].dropna().unique())
    with lf1:
        lost_sold_year = st.selectbox("Sold Year", ["All"] + sold_years_lost, key="lost_sold_year")
    with lf2:
        lost_branch = st.selectbox("Sold Branch", ["All"] + choices(sold["Sold Branch"]), key="lost_branch")
    with lf3:
        lost_city = st.selectbox("Customer Location", ["All"] + choices(sold["City"]), key="lost_city")
    with lf4:
        lost_model = st.selectbox("Model", ["All"] + choices(sold["Model"]), key="lost_model")
    with lf5:
        lost_as_on = st.date_input("As On Date", value=pd.Timestamp.today().date(), key="lost_as_on")

    # Base sold VIN population.
    lp = sold.copy()
    if lost_sold_year != "All":
        lp = lp[lp["Sold Year Calc"].eq(lost_sold_year)]
    if lost_branch != "All":
        lp = lp[lp["Sold Branch"].eq(lost_branch)]
    if lost_city != "All":
        lp = lp[lp["City"].eq(lost_city)]
    if lost_model != "All":
        lp = lp[lp["Model"].eq(lost_model)]

    lp = lp.sort_values("Sold Date").drop_duplicates("VIN", keep="last").copy()

    # Rebuild service history from raw service rows as of the selected date.
    as_on_ts = pd.Timestamp(lost_as_on)
    svc_hist = service[service["Service Date"].notna() & (service["Service Date"] <= as_on_ts)].copy()

    if len(svc_hist):
        svc_hist = svc_hist.sort_values(["VIN", "Service Date"])
        last_rows = svc_hist.groupby("VIN", as_index=False).tail(1)
        last_map = last_rows.set_index("VIN")["Service Date"].to_dict()
        loc_map = last_rows.set_index("VIN")["Service Location"].to_dict()
        cat_map = last_rows.set_index("VIN")["Category Calc"].to_dict()
        type_map = last_rows.set_index("VIN")["Service Type"].to_dict()
        kms_map = last_rows.set_index("VIN")["KMS"].to_dict() if "KMS" in last_rows.columns else {}
        serviced_vins = set(svc_hist["VIN"].dropna())
    else:
        last_map, loc_map, cat_map, type_map, kms_map = {}, {}, {}, {}, {}
        serviced_vins = set()

    lp["Last Service Date Action"] = lp["VIN"].map(last_map)
    lp["Last Service Location Action"] = lp["VIN"].map(loc_map)
    lp["Last Service Category Action"] = lp["VIN"].map(cat_map)
    lp["Last Service Type Action"] = lp["VIN"].map(type_map)
    lp["Latest KMS Action"] = lp["VIN"].map(kms_map)
    lp["Has Serviced"] = lp["VIN"].isin(serviced_vins)

    # Due logic: last service + 365 days; never-serviced VIN = sold date + 365 days.
    lp["Due Date Action"] = lp["Last Service Date Action"] + pd.to_timedelta(365, unit="D")
    never_mask = lp["Last Service Date Action"].isna()
    lp.loc[never_mask, "Due Date Action"] = lp.loc[never_mask, "Sold Date"] + pd.to_timedelta(365, unit="D")
    lp["Overdue Days Action"] = (as_on_ts - lp["Due Date Action"]).dt.days

    _d_action = lp["Overdue Days Action"]
    lp["Action Status"] = np.select(
        [
            _d_action.isna(),
            _d_action > 180,
            _d_action > 0,
            _d_action >= -30,
            lp["Has Serviced"],
        ],
        [
            "Not Yet Due",
            "Critical Lost",
            "Overdue",
            "Due Next 30 Days",
            "Retained",
        ],
        default="Not Yet Due",
    )

    def action_recommendation(status):
        return {
            "Critical Lost": "Call Now — personal advisor follow-up",
            "Overdue": "Campaign Now — call + service offer",
            "Due Next 30 Days": "Reminder — call / WhatsApp / SMS",
            "Retained": "Maintain relationship / loyalty follow-up",
            "Not Yet Due": "Monitor until due date",
        }.get(status, "Review")

    lp["Recommended Action Dashboard"] = lp["Action Status"].map(action_recommendation)

    critical = int((lp["Action Status"] == "Critical Lost").sum())
    overdue = int((lp["Action Status"] == "Overdue").sum())
    due30 = int((lp["Action Status"] == "Due Next 30 Days").sum())
    retained = int((lp["Action Status"] == "Retained").sum())
    notdue = int((lp["Action Status"] == "Not Yet Due").sum())
    eligible = len(lp)
    at_risk = critical + overdue
    at_risk_pct = (at_risk / eligible * 100) if eligible else 0

    cards = [
        ("🚨", "CRITICAL LOST VINs", critical, "> 180 Days Overdue", "#D94A3A"),
        ("⏰", "OVERDUE VINs", overdue, "1–180 Days Overdue", "#F28C28"),
        ("📅", "DUE NEXT 30 DAYS", due30, "0–30 Days to Due", "#F28C28"),
        ("✅", "RETAINED VINs", retained, "Serviced & currently within due cycle", "#079447"),
        ("⌛", "NOT YET DUE", notdue, "Too early for service", "#2F75B5"),
    ]
    kc = st.columns(5, gap="small")
    for i,(ico,lbl,val,note,col) in enumerate(cards):
        with kc[i]:
            st.markdown(
                f'<div class="lost-card"><div style="display:flex;gap:9px;align-items:center;">'
                f'<div style="font-size:28px;">{ico}</div><div>'
                f'<div class="lost-card-label">{lbl}</div>'
                f'<div class="lost-card-value" style="color:{col};">{val:,}</div>'
                f'<div class="lost-card-note">{note}</div></div></div></div>',
                unsafe_allow_html=True
            )

    # Top analytical row
    lcol, rcol = st.columns([1,1], gap="small")

    with lcol:
        st.markdown('<div class="lost-bar">PRIORITY ACTION — WHERE TO FOCUS FIRST</div>', unsafe_allow_html=True)
        priority = (
            lp.groupby(["Sold Year Calc", "Sold Month Calc"], dropna=False)
              .agg(
                  Eligible_VINs=("VIN","nunique"),
                  Retained_VINs=("Has Serviced","sum"),
                  At_Risk_VINs=("Action Status", lambda x: x.isin(["Critical Lost","Overdue"]).sum())
              ).reset_index()
        )
        if len(priority):
            priority["Retention %"] = (priority["Retained_VINs"] / priority["Eligible_VINs"] * 100).round(1)
            priority["Priority"] = np.select(
                [priority["At_Risk_VINs"] >= priority["At_Risk_VINs"].quantile(.75),
                 priority["At_Risk_VINs"] >= priority["At_Risk_VINs"].median()],
                ["Critical","High"], default="Watch"
            )
            priority["Action"] = priority["Priority"].map({"Critical":"Immediate Calls","High":"Campaign Now","Watch":"Monitor & Nudge"})
            # Sold Month Calc may contain either month numbers or names such as Jan/Feb/Dec.
            _month_num = pd.to_numeric(priority["Sold Month Calc"], errors="coerce")
            _month_name_num = pd.to_datetime(
                priority["Sold Month Calc"].astype(str).str.strip().str[:3],
                format="%b",
                errors="coerce"
            ).dt.month
            _resolved_month = _month_num.fillna(_month_name_num)

            _year_num = pd.to_numeric(priority["Sold Year Calc"], errors="coerce")
            _valid_cohort = _year_num.notna() & _resolved_month.notna()

            priority["Sold Month Cohort"] = "-"
            if _valid_cohort.any():
                priority.loc[_valid_cohort, "Sold Month Cohort"] = pd.to_datetime(
                    dict(
                        year=_year_num.loc[_valid_cohort].astype(int),
                        month=_resolved_month.loc[_valid_cohort].astype(int),
                        day=1
                    ),
                    errors="coerce"
                ).dt.strftime("%b-%y").fillna("-")
            st.dataframe(
                priority[["Priority","Sold Month Cohort","Eligible_VINs","Retained_VINs","At_Risk_VINs","Retention %","Action"]]
                .rename(columns={"Eligible_VINs":"Eligible VINs","Retained_VINs":"Retained VINs","At_Risk_VINs":"Lost / Overdue VINs"}),
                use_container_width=True, hide_index=True, height=260
            )
        else:
            st.info("No VINs for the selected filters.")

    with rcol:
        st.markdown('<div class="lost-bar">VIN STATUS DISTRIBUTION</div>', unsafe_allow_html=True)
        status_order = ["Critical Lost","Overdue","Due Next 30 Days","Retained","Not Yet Due"]
        status_counts = lp["Action Status"].value_counts().reindex(status_order, fill_value=0)
        fig_status = go.Figure(go.Pie(
            labels=status_counts.index,
            values=status_counts.values,
            hole=.48,
            marker=dict(colors=["#D94A3A","#F28C28","#F28C28","#079447","#2F75B5"]),
            textinfo="percent",
        ))
        fig_status.update_layout(
            height=260, margin=dict(l=10,r=10,t=20,b=10),
            legend=dict(orientation="v", x=1.0, y=.9),
            annotations=[dict(text=f"<b>{eligible:,}</b><br>VINs", x=.5,y=.5,showarrow=False)]
        )
        st.plotly_chart(fig_status, use_container_width=True, config={"displayModeBar":False})
        st.caption(f"Total at risk: {at_risk:,} VINs • {at_risk_pct:.1f}% of selected sold VIN population")

    # Location + trend row
    c1,c2,c3 = st.columns(3, gap="small")
    with c1:
        st.markdown('<div class="lost-bar">LOST / OVERDUE VINs BY SOLD LOCATION</div>', unsafe_allow_html=True)
        loc = (
            lp.assign(AtRisk=lp["Action Status"].isin(["Critical Lost","Overdue"]).astype(int))
              .groupby("Sold Branch", dropna=False)
              .agg(At_Risk_VINs=("AtRisk","sum"), Total_VINs=("VIN","nunique"))
              .reset_index()
        )
        loc["At Risk %"] = np.where(loc["Total_VINs"]>0, loc["At_Risk_VINs"]/loc["Total_VINs"]*100, 0)
        loc = loc.sort_values("At_Risk_VINs", ascending=False)
        fig_loc = go.Figure(go.Bar(x=loc["Sold Branch"], y=loc["At_Risk_VINs"], marker_color="#D94A3A", text=loc["At_Risk_VINs"], textposition="outside"))
        fig_loc.update_layout(height=260, margin=dict(l=20,r=10,t=25,b=55), xaxis_title="", yaxis_title="VINs")
        st.plotly_chart(fig_loc, use_container_width=True, config={"displayModeBar":False})

    with c2:
        st.markdown('<div class="lost-bar">RETENTION % BY SOLD LOCATION</div>', unsafe_allow_html=True)
        lr = (
            lp.groupby("Sold Branch", dropna=False)
              .agg(Total=("VIN","nunique"), Serviced=("Has Serviced","sum"))
              .reset_index()
        )
        lr["Retention"] = np.where(lr["Total"]>0, lr["Serviced"]/lr["Total"]*100, 0)
        lr = lr.sort_values("Retention")
        fig_lr = go.Figure(go.Bar(x=lr["Retention"], y=lr["Sold Branch"], orientation="h", text=lr["Retention"].map(lambda x:f"{x:.1f}%"), textposition="outside"))
        fig_lr.update_layout(height=260, margin=dict(l=20,r=45,t=25,b=30), xaxis=dict(range=[0,100], ticksuffix="%"), yaxis_title="")
        st.plotly_chart(fig_lr, use_container_width=True, config={"displayModeBar":False})

    with c3:
        st.markdown('<div class="lost-bar">RETENTION TREND BY SOLD MONTH COHORT</div>', unsafe_allow_html=True)
        trend = (
            lp.groupby(["Sold Year Calc","Sold Month Calc"])
              .agg(Total=("VIN","nunique"), Serviced=("Has Serviced","sum"))
              .reset_index()
        )
        if len(trend):
            trend["Retention"] = np.where(trend["Total"]>0, trend["Serviced"]/trend["Total"]*100, 0)
            _trend_month_num = pd.to_numeric(trend["Sold Month Calc"], errors="coerce")
            _trend_month_name_num = pd.to_datetime(
                trend["Sold Month Calc"].astype(str).str.strip().str[:3],
                format="%b",
                errors="coerce"
            ).dt.month
            _trend_resolved_month = _trend_month_num.fillna(_trend_month_name_num)
            _trend_year_num = pd.to_numeric(trend["Sold Year Calc"], errors="coerce")

            trend["Cohort"] = pd.NaT
            _trend_valid = _trend_year_num.notna() & _trend_resolved_month.notna()
            if _trend_valid.any():
                trend.loc[_trend_valid, "Cohort"] = pd.to_datetime(
                    dict(
                        year=_trend_year_num.loc[_trend_valid].astype(int),
                        month=_trend_resolved_month.loc[_trend_valid].astype(int),
                        day=1
                    ),
                    errors="coerce"
                )
            trend = trend.dropna(subset=["Cohort"]).sort_values("Cohort")
            fig_tr = go.Figure(go.Scatter(x=trend["Cohort"], y=trend["Retention"], mode="lines+markers+text", text=trend["Retention"].map(lambda x:f"{x:.1f}%"), textposition="top center", line=dict(color="#062D5D")))
            fig_tr.update_layout(height=260, margin=dict(l=20,r=15,t=30,b=35), yaxis=dict(range=[0,100],ticksuffix="%"), xaxis_title="")
            st.plotly_chart(fig_tr, use_container_width=True, config={"displayModeBar":False})

    # Detailed action list + action guidance
    al, guide = st.columns([2.25, .9], gap="small")
    with al:
        st.markdown('<div class="lost-bar">LOST / OVERDUE / DUE VINs — ACTION LIST</div>', unsafe_allow_html=True)
        action_df = lp[lp["Action Status"].isin(["Critical Lost","Overdue","Due Next 30 Days"])].copy()
        priority_map = {"Critical Lost":1,"Overdue":2,"Due Next 30 Days":3}
        action_df["_p"] = action_df["Action Status"].map(priority_map)
        action_df = action_df.sort_values(["_p","Overdue Days Action"], ascending=[True,False])

        customer_col = "Customer Name" if "Customer Name" in action_df.columns else None
        vehicle_col = "Vehicle No" if "Vehicle No" in action_df.columns else None
        cols = ["Action Status","VIN"]
        if vehicle_col: cols.append(vehicle_col)
        if customer_col: cols.append(customer_col)
        cols += ["Sold Date","Model","Sold Branch","Last Service Date Action","Last Service Category Action","Last Service Type Action","Latest KMS Action","Due Date Action","Overdue Days Action","Recommended Action Dashboard"]
        cols = [c for c in cols if c in action_df.columns]

        st.dataframe(
            action_df[cols].rename(columns={
                "Action Status":"Priority",
                "Last Service Date Action":"Last Service Date",
                "Last Service Category Action":"Last Service Category",
                "Last Service Type Action":"Last Service Type",
                "Latest KMS Action":"Latest KMS",
                "Due Date Action":"Next Due Date",
                "Overdue Days Action":"Overdue Days",
                "Recommended Action Dashboard":"Action",
            }),
            use_container_width=True, hide_index=True, height=330
        )
        st.download_button(
            "⬇️ Download Action List CSV",
            action_df[cols].to_csv(index=False).encode("utf-8"),
            file_name=f"lost_vin_action_list_{lost_as_on}.csv",
            mime="text/csv",
            key="lost_action_download",
        )

    with guide:
        st.markdown('<div class="lost-bar">WHAT ACTION SHOULD WE TAKE?</div>', unsafe_allow_html=True)
        st.markdown("""
        <div class="lost-action-box">🚨 <b style="color:#D92F2F;">Critical Lost (&gt;180 Days)</b><br>Immediate call and personal advisor follow-up.</div>
        <div class="lost-action-box">⏰ <b style="color:#F28C28;">Overdue</b><br>Call + targeted service campaign / offer.</div>
        <div class="lost-action-box">🔔 <b style="color:#D6A600;">Due Next 30 Days</b><br>Reminder by call, WhatsApp or SMS.</div>
        <div class="lost-action-box">✅ <b style="color:#079447;">Retained</b><br>Maintain relationship and loyalty engagement.</div>
        <div class="lost-action-box">⌛ <b style="color:#2F75B5;">Not Yet Due</b><br>No immediate action. Monitor until due.</div>
        """, unsafe_allow_html=True)

    st.caption(
        "Status is calculated from the selected sold VIN population and actual service history as of the selected date. "
        "Current rule: next due = 365 days after last service, or 365 days after sold date for never-serviced VINs."
    )

if active_page == PAGE_LABELS[9]:
    st.subheader("🛡️ SMP Tracking — Service Maintenance Pack")
    st.caption(
        "Maps SMP contracts to actual PMS service history by VIN. "
        "Services availed are counted as unique PMS repair orders within the SMP validity period."
    )

    # ---------------------------------------------------------------
    # SMP data update
    # ---------------------------------------------------------------
    with st.expander("📁 Update / Replace SMP Data", expanded=False):
        st.caption(
            "Upload the latest SMP Excel/CSV whenever you receive a refreshed SMP contract list. "
            "The app will replace the local SMP master used by this tab."
        )
        _smp_upload = st.file_uploader(
            "Upload SMP data",
            type=["xlsx", "xls", "csv"],
            key="smp_upload_v324",
        )
        if _smp_upload is not None:
            if st.button("✅ Replace SMP Data", type="primary", key="replace_smp_v324"):
                try:
                    _name = _smp_upload.name.lower()
                    if _name.endswith(".csv"):
                        _new_smp = pd.read_csv(io.BytesIO(_smp_upload.getvalue()), low_memory=False)
                    else:
                        _new_smp = pd.read_excel(io.BytesIO(_smp_upload.getvalue()), engine="openpyxl")
                    _new_smp.columns = [str(c).strip() for c in _new_smp.columns]
                    _required = {"VIN", "Validity From Date", "Validity To Date"}
                    _missing = _required - set(_new_smp.columns)
                    if _missing:
                        raise ValueError("SMP file is missing: " + ", ".join(sorted(_missing)))
                    _new_smp.to_excel(SMP_FILE, index=False, engine="openpyxl")
                    st.success(f"SMP data updated successfully: {len(_new_smp):,} rows.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Could not update SMP data: {e}")

    if not SMP_FILE.exists():
        st.warning("SMP_DATA.xlsx was not found. Upload SMP data above.")
    else:
        try:
            smp = load_prepared_smp(str(SMP_FILE), SMP_FILE.stat().st_mtime)

            # One current/latest SMP row per contract VIN if duplicate contract rows exist.
            _sort_cols = [c for c in ["VIN", "Created Date", "Validity From Date"] if c in smp.columns]
            if _sort_cols:
                smp = smp.sort_values(_sort_cols)
            if "Contract Number" in smp.columns:
                smp = smp.drop_duplicates("Contract Number", keep="last")
            else:
                smp = smp.drop_duplicates(
                    subset=["VIN", "Validity From Date", "Validity To Date"],
                    keep="last"
                )

            # -------------------------------------------------------
            # Filters / selected due month
            # -------------------------------------------------------
            _today_smp = pd.Timestamp(date.today()).normalize()
            _month_options = pd.period_range(
                start=min(smp["Validity From Date"].min(), _today_smp).to_period("M"),
                end=max(smp["Validity To Date"].max(), _today_smp).to_period("M"),
                freq="M",
            )
            _month_labels = [p.strftime("%b %Y") for p in _month_options]
            _default_period = _today_smp.to_period("M")
            _default_idx = (
                list(_month_options).index(_default_period)
                if _default_period in list(_month_options)
                else len(_month_options) - 1
            )

            f1, f2, f3, f4, f5 = st.columns([1.2, 1.05, 1.25, 1.6, 1.3], gap="small")
            with f1:
                _due_label = st.selectbox(
                    "Due Month",
                    _month_labels,
                    index=_default_idx,
                    key="smp_due_month_v324",
                )
            with f2:
                _years_opts = ["All"] + sorted(
                    [int(x) for x in smp["Validity in Years"].dropna().unique()]
                )
                _term = st.selectbox("SMP Validity", _years_opts, key="smp_term_v324")
            with f3:
                _status_opts = ["All"] + sorted(
                    [x for x in smp.get("Status", pd.Series(dtype=str)).dropna().astype(str).unique() if x]
                )
                _pack_status = st.selectbox("SMP Status", _status_opts, key="smp_status_v324")
            with f4:
                _model_opts = ["All"] + sorted(
                    [x for x in smp.get("Model", pd.Series(dtype=str)).dropna().astype(str).unique() if x]
                )
                _smp_model = st.selectbox("Model", _model_opts, key="smp_model_v324")
            with f5:
                _svc_loc_opts = ["All"] + sorted(
                    [x for x in service["Service Location"].dropna().astype(str).unique() if x]
                )
                _smp_service_loc = st.selectbox(
                    "Service Location",
                    _svc_loc_opts,
                    key="smp_service_location_v324",
                )

            _selected_period = pd.Period(_due_label, freq="M")
            _period_start = _selected_period.start_time.normalize()
            _period_end = _selected_period.end_time.normalize()

            sf_smp = smp.copy()
            if _term != "All":
                sf_smp = sf_smp[sf_smp["Validity in Years"].eq(float(_term))]
            if _pack_status != "All" and "Status" in sf_smp.columns:
                sf_smp = sf_smp[sf_smp["Status"].eq(_pack_status)]
            if _smp_model != "All" and "Model" in sf_smp.columns:
                sf_smp = sf_smp[sf_smp["Model"].eq(_smp_model)]

            # -------------------------------------------------------
            # Match PMS service history to each SMP VIN.
            # -------------------------------------------------------
            _pms_hist = service[service["Category Calc"].eq("PMS")].copy()
            _pms_hist["VIN"] = _pms_hist["VIN"].fillna("").astype(str).str.strip().str.upper()

            if _smp_service_loc != "All":
                _pms_hist = _pms_hist[_pms_hist["Service Location"].eq(_smp_service_loc)]

            # Deduplicate invoice lines: one PMS event per VIN + RO where RO is available.
            if "RO Number" in _pms_hist.columns:
                _pms_hist["_event_key"] = (
                    _pms_hist["VIN"].astype(str) + "|" +
                    _pms_hist["RO Number"].fillna("").astype(str).str.strip()
                )
                _blank_ro = _pms_hist["RO Number"].fillna("").astype(str).str.strip().eq("")
                _pms_hist.loc[_blank_ro, "_event_key"] = (
                    _pms_hist.loc[_blank_ro, "VIN"].astype(str) + "|" +
                    _pms_hist.loc[_blank_ro, "Service Date"].astype(str)
                )
            else:
                _pms_hist["_event_key"] = (
                    _pms_hist["VIN"].astype(str) + "|" +
                    _pms_hist["Service Date"].astype(str)
                )
            _pms_events = (
                _pms_hist.sort_values("Service Date")
                .drop_duplicates("_event_key", keep="last")
                .copy()
            )

            # Group once per VIN so each loop iteration only scans that VIN's own
            # (usually tiny) PMS history instead of re-filtering the full table.
            _pms_by_vin = {
                _vin_key: _grp.sort_values("Service Date")
                for _vin_key, _grp in _pms_events.groupby("VIN", sort=False)
            }
            _empty_pms = _pms_events.iloc[0:0]

            _rows = []
            for _, r in sf_smp.iterrows():
                _vin = r["VIN"]
                _vf = pd.Timestamp(r["Validity From Date"]).normalize()
                _vt = pd.Timestamp(r["Validity To Date"]).normalize()

                _vin_pms = _pms_by_vin.get(_vin, _empty_pms)
                _vh = _vin_pms[
                    _vin_pms["Service Date"].between(_vf, _vt, inclusive="both")
                ]

                _availed = len(_vh)
                _last_pms = _vh["Service Date"].max() if _availed else pd.NaT

                # Annual service cadence:
                # no PMS availed -> first due one year after SMP validity starts.
                # after a PMS -> next due one year after last PMS.
                _next_due = (
                    pd.Timestamp(_last_pms).normalize() + pd.Timedelta(days=365)
                    if pd.notna(_last_pms)
                    else _vf + pd.Timedelta(days=365)
                )

                # Estimate service entitlement from validity years at annual cadence.
                _term_years = r.get("Validity in Years", np.nan)
                _allowed = int(_term_years) if pd.notna(_term_years) and float(_term_years) > 0 else np.nan
                _balance = max(_allowed - _availed, 0) if pd.notna(_allowed) else np.nan

                if _period_end < _vf:
                    _due_status = "Pack Not Started"
                elif _period_start > _vt:
                    _due_status = "SMP Expired"
                elif _next_due > _vt:
                    _due_status = "No Further PMS Due in SMP"
                elif _period_start <= _next_due <= _period_end:
                    _due_status = "Due This Month"
                elif _next_due < _period_start:
                    _due_status = "Overdue"
                else:
                    _due_status = "Active / Not Due"

                _last_loc = ""
                _latest_kms = np.nan
                if _availed:
                    _lastrow = _vh.iloc[-1]
                    _last_loc = _lastrow.get("Service Location", "")
                    _latest_kms = _lastrow.get("KMS", np.nan)

                _days = (_period_end - _next_due).days if pd.notna(_next_due) else np.nan

                _rows.append({
                    "VIN": _vin,
                    "Regd Number": r.get("Regd Number", ""),
                    "Customer Name": r.get("Customer Name", ""),
                    "Model": r.get("Model", ""),
                    "Sold Date": r.get("Sold Date", pd.NaT),
                    "SMP Status": r.get("Status", ""),
                    "SMP Validity Years": _term_years,
                    "SMP Start Date": _vf,
                    "SMP End Date": _vt,
                    "Product Name": r.get("Product Name", ""),
                    "Services Allowed*": _allowed,
                    "PMS Services Availed": _availed,
                    "Balance Services*": _balance,
                    "Last PMS Date": _last_pms,
                    "Next PMS Due Date": _next_due,
                    "Due Status": _due_status,
                    "Due / Overdue Days": _days,
                    "Last PMS Location": _last_loc,
                    "Latest PMS KMS": _latest_kms,
                })

            smp_track = pd.DataFrame(_rows)

            # -------------------------------------------------------
            # Summary KPIs
            # -------------------------------------------------------
            _total_smp = smp_track["VIN"].nunique()
            _active_smp = smp_track[
                (smp_track["SMP Start Date"] <= _period_end)
                & (smp_track["SMP End Date"] >= _period_start)
            ]["VIN"].nunique()
            _due_n = smp_track.loc[smp_track["Due Status"].eq("Due This Month"), "VIN"].nunique()
            _overdue_n = smp_track.loc[smp_track["Due Status"].eq("Overdue"), "VIN"].nunique()
            _availed_total = int(smp_track["PMS Services Availed"].sum())
            _unused = smp_track.loc[smp_track["PMS Services Availed"].eq(0), "VIN"].nunique()

            k1,k2,k3,k4,k5,k6 = st.columns(6, gap="small")
            k1.metric("SMP VINs", f"{_total_smp:,}")
            k2.metric("Active in Selected Month", f"{_active_smp:,}")
            k3.metric("PMS Services Availed", f"{_availed_total:,}")
            k4.metric("Due This Month", f"{_due_n:,}")
            k5.metric("Overdue", f"{_overdue_n:,}")
            k6.metric("No PMS Availed", f"{_unused:,}")

            st.caption(
                "*Services Allowed / Balance is estimated from SMP validity years using one PMS per year. "
                "If your official SMP entitlement differs by product, we can replace this with the exact entitlement mapping."
            )

            # -------------------------------------------------------
            # Summary table
            # -------------------------------------------------------
            st.markdown("#### SMP Summary")
            _summary = (
                smp_track.groupby("SMP Validity Years", dropna=False)
                .agg(
                    SMP_VINs=("VIN","nunique"),
                    PMS_Services_Availed=("PMS Services Availed","sum"),
                    Due_This_Month=("Due Status", lambda s: int((s == "Due This Month").sum())),
                    Overdue=("Due Status", lambda s: int((s == "Overdue").sum())),
                    No_PMS_Availed=("PMS Services Availed", lambda s: int((s == 0).sum())),
                )
                .reset_index()
                .rename(columns={
                    "SMP Validity Years":"SMP Validity",
                    "SMP_VINs":"SMP VINs",
                    "PMS_Services_Availed":"PMS Services Availed",
                    "Due_This_Month":"Due This Month",
                    "No_PMS_Availed":"No PMS Availed",
                })
            )
            _summary["SMP Validity"] = _summary["SMP Validity"].map(
                lambda x: f"{int(x)} Year" if pd.notna(x) and int(x) == 1
                else (f"{int(x)} Years" if pd.notna(x) else "Unknown")
            )
            st.dataframe(_summary, use_container_width=True, hide_index=True)

            # -------------------------------------------------------
            # Due / overdue detail
            # -------------------------------------------------------
            st.markdown(f"#### Due Data — {_due_label}")
            _due_filter = st.multiselect(
                "Show Due Status",
                ["Due This Month", "Overdue", "Active / Not Due", "No Further PMS Due in SMP", "SMP Expired", "Pack Not Started"],
                default=["Due This Month", "Overdue"],
                key="smp_due_status_filter_v324",
            )

            _detail = (
                smp_track[smp_track["Due Status"].isin(_due_filter)].copy()
                if _due_filter else smp_track.copy()
            )
            _priority = {
                "Overdue":1,
                "Due This Month":2,
                "Active / Not Due":3,
                "No Further PMS Due in SMP":4,
                "SMP Expired":5,
                "Pack Not Started":6,
            }
            _detail["_priority"] = _detail["Due Status"].map(_priority).fillna(99)
            _detail = _detail.sort_values(
                ["_priority", "Next PMS Due Date", "VIN"],
                ascending=[True, True, True]
            ).drop(columns="_priority")

            st.dataframe(
                _detail,
                use_container_width=True,
                hide_index=True,
                height=520,
            )

            st.download_button(
                "⬇️ Download SMP Due Data",
                _detail.to_csv(index=False).encode("utf-8-sig"),
                file_name=f"smp_due_data_{_selected_period.strftime('%Y_%m')}.csv",
                mime="text/csv",
                key="smp_due_download_v324",
            )

            st.caption(
                "Due logic in this first version uses a 365-day PMS interval. "
                "It checks actual PMS events only inside each VIN's SMP validity window."
            )

        except Exception as e:
            st.error(f"SMP Tracking could not be calculated: {e}")



# -----------------------------------------------------------------------------
# FINAL VISUAL OVERRIDES — intentionally last in the script.
# Keeps growth colors/shapes and the sidebar top collapse/expand arrow reliable.
# -----------------------------------------------------------------------------
st.markdown(
    """
    <style>
    /* Growth: use CSS geometry instead of Unicode arrows. This avoids thin
       font glyphs and makes the direction color independent of font loading. */
    .tp-growth {
        display:inline-flex !important;
        align-items:center !important;
        gap:4px !important;
        white-space:nowrap !important;
        font-weight:900 !important;
        line-height:1 !important;
    }
    .tp-growth .tp-growth-value {
        font-weight:900 !important;
        font-size:11px !important;
    }
    .tp-growth-up,
    .tp-growth-up .tp-growth-value {
        color:#079447 !important;
        -webkit-text-fill-color:#079447 !important;
    }
    .tp-growth-down,
    .tp-growth-down .tp-growth-value {
        color:#D83932 !important;
        -webkit-text-fill-color:#D83932 !important;
    }
    .tp-growth-flat,
    .tp-growth-flat .tp-growth-value {
        color:#E18A00 !important;
        -webkit-text-fill-color:#E18A00 !important;
    }
    .tp-growth-new,
    .tp-growth-new .tp-growth-value {
        color:#718099 !important;
        -webkit-text-fill-color:#718099 !important;
    }
    .tp-growth-shape {
        display:inline-block !important;
        width:0 !important;
        height:0 !important;
        flex:0 0 auto !important;
        -webkit-text-fill-color:initial !important;
    }
    .tp-growth-up .tp-growth-shape {
        border-left:6px solid transparent !important;
        border-right:6px solid transparent !important;
        border-bottom:10px solid #079447 !important;
    }
    .tp-growth-down .tp-growth-shape {
        border-left:6px solid transparent !important;
        border-right:6px solid transparent !important;
        border-top:10px solid #D83932 !important;
    }
    .tp-growth-flat .tp-growth-shape {
        border-top:6px solid transparent !important;
        border-bottom:6px solid transparent !important;
        border-left:10px solid #E18A00 !important;
    }
    /* Explicitly beat earlier KPI/table !important color rules. */
    .tp-kpi-note .tp-growth *,
    table.tp-table td .tp-growth *,
    table.tp-summary td .tp-growth * {
        opacity:1 !important;
    }

    /* Sidebar top arrow: do not depend on Material Icons. Earlier CSS hid the
       ligature to prevent 'keyboard_arrow_*' text; draw a real CSS chevron. */
    button[data-testid="stSidebarCollapseButton"],
    [data-testid="stSidebarCollapseButton"] > button,
    button[data-testid="collapsedControl"],
    [data-testid="collapsedControl"] > button {
        position:relative !important;
        min-width:28px !important;
        min-height:28px !important;
        overflow:visible !important;
    }
    button[data-testid="stSidebarCollapseButton"] [data-testid="stIconMaterial"],
    [data-testid="stSidebarCollapseButton"] > button [data-testid="stIconMaterial"],
    button[data-testid="collapsedControl"] [data-testid="stIconMaterial"],
    [data-testid="collapsedControl"] > button [data-testid="stIconMaterial"] {
        display:none !important;
    }
    button[data-testid="stSidebarCollapseButton"]::after,
    [data-testid="stSidebarCollapseButton"] > button::after {
        content:"" !important;
        display:block !important;
        width:9px !important;
        height:9px !important;
        border-left:3px solid #062D5D !important;
        border-bottom:3px solid #062D5D !important;
        transform:rotate(45deg) !important;
        margin:auto !important;
        box-sizing:border-box !important;
    }
    button[data-testid="collapsedControl"]::after,
    [data-testid="collapsedControl"] > button::after {
        content:"" !important;
        display:block !important;
        width:9px !important;
        height:9px !important;
        border-right:3px solid #062D5D !important;
        border-top:3px solid #062D5D !important;
        transform:rotate(45deg) !important;
        margin:auto !important;
        box-sizing:border-box !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)
