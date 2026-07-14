"""
╔══════════════════════════════════════════════════════════════╗
║   🐦  Bird Species Observation Analysis — Streamlit App     ║
║   Domain : Environmental Studies · Biodiversity · Ecology   ║
║   Skills : Data Cleaning · EDA · SQL · Plotly · Streamlit   ║
╚══════════════════════════════════════════════════════════════╝
"""

import os
import sqlite3
import warnings
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from plotly.subplots import make_subplots

warnings.filterwarnings("ignore")

# ─────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="🐦 Bird Species Analysis",
    page_icon="🐦",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────────────────────
# DATASET PATHS  (auto-load from data/ folder)
# ─────────────────────────────────────────────────────────────
BASE_DIR       = Path(__file__).parent
FOREST_PATH    = BASE_DIR / "data" / "Bird_Monitoring_Data_FOREST.XLSX"
GRASSLAND_PATH = BASE_DIR / "data" / "Bird_Monitoring_Data_GRASSLAND.XLSX"

# ─────────────────────────────────────────────────────────────
# CSS — Premium dark-nature theme
# ─────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=Playfair+Display:wght@700&display=swap');

html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
.main  { background-color: #0d1117; }
.block-container { padding: 1.5rem 2.5rem 3rem 2.5rem; }
#MainMenu { visibility: hidden; }
footer    { visibility: hidden; }

/* Hide header decorations but always keep sidebar toggle visible */
div[data-testid="stDecoration"] { visibility: hidden; }
div[data-testid="stStatusWidget"] { visibility: hidden; }
div[data-testid="stToolbar"] { right: 2rem; }
div[data-testid="stToolbarActions"] { visibility: hidden; }

header[data-testid="stHeader"] {
    background: transparent !important;
    box-shadow: none !important;
}

[data-testid="collapsedControl"],
[data-testid="stSidebarCollapsedControl"],
button[kind="header"] {
    visibility: visible !important;
    display: flex !important;
    opacity: 1 !important;
}

/* HERO */
.hero-banner {
    background: linear-gradient(135deg,#0f2027 0%,#203a43 50%,#2c5364 100%);
    border-radius: 18px; padding: 2.8rem 3rem; margin-bottom: 2rem;
    border: 1px solid rgba(255,255,255,0.06);
    box-shadow: 0 8px 40px rgba(0,0,0,0.5);
    position: relative; overflow: hidden;
}
.hero-banner::before {
    content:"🌿"; position:absolute; right:2.5rem; top:50%;
    transform:translateY(-50%); font-size:6rem; opacity:0.12;
}
.hero-title {
    font-family:'Playfair Display',serif; font-size:2.8rem;
    font-weight:700; color:#e6f0fa; margin:0 0 0.5rem 0; letter-spacing:-0.5px;
}
.hero-sub  { font-size:1rem; color:#8ab4c9; margin:0; font-weight:400; line-height:1.7; }
.hero-badge {
    display:inline-block; background:rgba(100,200,100,0.15);
    border:1px solid rgba(100,200,100,0.4); border-radius:50px;
    padding:5px 16px; font-size:0.76rem; color:#7dd87d;
    margin-top:1rem; font-weight:600; letter-spacing:0.4px;
}

/* KPI CARDS */
.kpi-card {
    background:linear-gradient(145deg,#161b22,#1c2333);
    border:1px solid rgba(255,255,255,0.08); border-radius:16px;
    padding:1.5rem 1.4rem; text-align:center;
    box-shadow:0 4px 20px rgba(0,0,0,0.35);
    transition:transform 0.2s,box-shadow 0.2s;
}
.kpi-card:hover { transform:translateY(-4px); box-shadow:0 10px 28px rgba(0,0,0,0.45); }
.kpi-icon  { font-size:2.1rem; margin-bottom:0.5rem; }
.kpi-value { font-size:2rem; font-weight:700; color:#58a6ff; line-height:1; margin-bottom:0.25rem; }
.kpi-label { font-size:0.75rem; color:#8b949e; text-transform:uppercase; letter-spacing:0.9px; font-weight:600; }
.kpi-sub   { font-size:0.74rem; margin-top:0.4rem; color:#3fb950; font-weight:500; }

/* SECTION HEADERS */
.sec-hdr {
    font-family:'Playfair Display',serif; font-size:1.5rem; color:#c9d1d9;
    border-left:4px solid #388bfd; padding-left:1rem;
    margin:2rem 0 1rem 0; font-weight:700;
}

/* TABS */
.stTabs [data-baseweb="tab-list"] {
    gap:6px; background:#161b22; padding:8px;
    border-radius:14px; border:1px solid rgba(255,255,255,0.06);
}
.stTabs [data-baseweb="tab"] {
    border-radius:8px; padding:8px 18px;
    font-size:0.86rem; font-weight:500; color:#8b949e; background:transparent;
}
.stTabs [aria-selected="true"] { background:#1f6feb !important; color:white !important; }

/* SIDEBAR */
[data-testid="stSidebar"] {
    background:linear-gradient(180deg,#0d1117 0%,#161b22 100%) !important;
    border-right:1px solid rgba(255,255,255,0.06) !important;
}

[data-testid="stSidebarNav"] {
    display: block !important;
}

/* Sidebar toggle button - force visible */
button[kind="header"] {
    display: block !important;
    visibility: visible !important;
}
.sb-logo { text-align:center; padding:1rem 0 1.5rem 0;
    border-bottom:1px solid rgba(255,255,255,0.08); margin-bottom:1.5rem; }
.sb-logo h2 { color:#58a6ff; font-family:'Playfair Display',serif; font-size:1.3rem; margin:0.4rem 0 0 0; }
.sb-logo p  { color:#8b949e; font-size:0.76rem; margin:0; }

/* INSIGHT CARDS */
.ins-card {
    background:linear-gradient(145deg,#161b22,#1c2333);
    border:1px solid rgba(255,255,255,0.07); border-radius:12px;
    padding:1.1rem 1.4rem; margin:0.45rem 0;
}
.ins-card h4 { color:#c9d1d9; margin:0 0 0.45rem 0; font-size:0.94rem; }
.ins-card p  { color:#8b949e; font-size:0.84rem; margin:0; line-height:1.65; }

/* SQL CODE BOX */
.sql-box {
    background:#0d1117; border:1px solid #30363d; border-radius:10px;
    padding:1rem 1.4rem; font-family:'Courier New',monospace;
    font-size:0.81rem; color:#7ee787; white-space:pre-wrap;
    overflow-x:auto; margin:0.8rem 0;
}

/* MISC */
.stDataFrame { border-radius:10px; overflow:hidden; }
hr { border-color:rgba(255,255,255,0.07) !important; margin:1.5rem 0; }
.footer-bar {
    background:#161b22; border-top:1px solid rgba(255,255,255,0.06);
    border-radius:10px; padding:1rem 2rem; text-align:center;
    color:#8b949e; font-size:0.77rem; margin-top:3rem;
}
</style>
""", unsafe_allow_html=True)

# Colour maps
HC = {"Forest": "#2E86AB", "Grassland": "#A23B72"}


# ─────────────────────────────────────────────────────────────
# DATA LOADING
# ─────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_data(forest_path: str, grassland_path: str) -> pd.DataFrame:
    def _load(fpath, label):
        xl = pd.ExcelFile(fpath)
        frames = []
        for sh in xl.sheet_names:
            try:
                d = pd.read_excel(fpath, sheet_name=sh)
                d["Source_Sheet"]    = sh
                d["Source_Location"] = label
                frames.append(d)
            except Exception:
                pass
        return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()

    df = pd.concat([_load(forest_path, "Forest"),
                    _load(grassland_path, "Grassland")], ignore_index=True)

    # — Standardize column names
    df.columns = df.columns.str.strip().str.replace(" ", "_").str.replace("-", "_")

    # — Duplicates
    df.drop_duplicates(inplace=True)

    # — Location Type
    if "Location_Type" in df.columns:
        df["Location_Type"] = df["Location_Type"].str.strip().str.title()
        df["Location_Type"] = df.apply(
            lambda r: r["Source_Location"]
            if str(r["Location_Type"]).lower() in ("unknown", "", "nan", "none")
            else r["Location_Type"], axis=1)

    # — Date & temporal features
    if "Date" in df.columns:
        df["Date"]        = pd.to_datetime(df["Date"], errors="coerce")
        df["Month"]       = df["Date"].dt.month
        df["Month_Name"]  = df["Date"].dt.strftime("%B")
        df["Day_of_Week"] = df["Date"].dt.day_name()
        df["Season"]      = df["Month"].map({
            12:"Winter",1:"Winter",2:"Winter",
            3:"Spring", 4:"Spring",5:"Spring",
            6:"Summer", 7:"Summer",8:"Summer",
            9:"Fall",  10:"Fall", 11:"Fall"
        })

    # — Year
    if "Year" in df.columns:
        df["Year"] = pd.to_numeric(df["Year"], errors="coerce").astype("Int64")

    # — Fill categoricals
    for c in ["Sex","ID_Method","Distance","Sky","Wind","Disturbance",
              "PIF_Watchlist_Status","Regional_Stewardship_Status"]:
        if c in df.columns:
            df[c] = df[c].fillna("Unknown")

    # — Fill numerics
    for c in ["Temperature","Humidity","Initial_Three_Min_Cnt"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")
            df[c] = df[c].fillna(df[c].median())

    # — Drop rows with no species
    if "Common_Name" in df.columns:
        df = df[df["Common_Name"].notna() & (df["Common_Name"].astype(str).str.strip() != "")]

    # — Encode bool-like columns
    if "Flyover_Observed" in df.columns:
        raw = df["Flyover_Observed"].astype(str).str.strip().str.upper()
        df["Flyover_Observed"] = raw.map(
            {"TRUE":"Yes","FALSE":"No","1":"Yes","0":"No",
             "TRUE":"Yes","FALSE":"No"}).fillna("Unknown")

    for c in ["PIF_Watchlist_Status","Regional_Stewardship_Status"]:
        if c in df.columns:
            df[c] = df[c].astype(str).str.strip().str.upper()
            df[c] = df[c].replace({"TRUE":"True","FALSE":"False","NAN":"Unknown"})

    return df


@st.cache_resource
def build_sql(_df: pd.DataFrame):
    conn = sqlite3.connect(":memory:", check_same_thread=False)
    _df.to_sql("bird_observations", conn, if_exists="replace", index=False)
    return conn


def qry(conn, sql: str) -> pd.DataFrame:
    try:
        return pd.read_sql_query(sql, conn)
    except Exception as e:
        return pd.DataFrame({"Error": [str(e)]})


# ─────────────────────────────────────────────────────────────
# CHART THEME HELPER
# ─────────────────────────────────────────────────────────────
def T(fig, h=420):
    fig.update_layout(
        height=h,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(22,27,34,0.85)",
        font=dict(color="#c9d1d9", size=12),
        margin=dict(l=40, r=20, t=50, b=40),
        legend=dict(bgcolor="rgba(22,27,34,0.8)",
                    bordercolor="rgba(255,255,255,0.1)",
                    borderwidth=1),
    )
    fig.update_xaxes(gridcolor="rgba(255,255,255,0.05)", zeroline=False)
    fig.update_yaxes(gridcolor="rgba(255,255,255,0.05)", zeroline=False)
    return fig


# ═══════════════════════════════════════════════════════════════
# SIDEBAR
# ═══════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("""
    <div class="sb-logo">
        <div style="font-size:3rem">🐦</div>
        <h2>Bird Analysis</h2>
        <p>Observation Dashboard v2.0</p>
    </div>""", unsafe_allow_html=True)

    # ── Check dataset files exist
    datasets_ok = FOREST_PATH.exists() and GRASSLAND_PATH.exists()

    if datasets_ok:
        st.success("✅ Datasets found — loading automatically")
    else:
        st.error("❌ Dataset files not found in `data/` folder")
        st.markdown("""
        <div style='font-size:0.78rem;color:#8b949e;line-height:1.7;'>
        Place these files inside the <b>data/</b> folder:<br>
        • Bird_Monitoring_Data_FOREST.XLSX<br>
        • Bird_Monitoring_Data_GRASSLAND.XLSX
        </div>""", unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("""
    <div style='font-size:0.76rem;color:#8b949e;line-height:1.9;'>
    <b style='color:#c9d1d9;font-size:0.85rem;'>📌 Project Info</b><br>
    Domain: Environmental Studies<br>
    Skills: EDA · SQL · Plotly · Streamlit<br>
    Habitats: Forest &amp; Grassland<br>
    Admin Units: ANTI CATO CHOH GWMP<br>
    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    HAFE MANA MONO NACE<br>
    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    PRWI ROCR WOTR
    </div>""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────
# HERO BANNER
# ─────────────────────────────────────────────────────────────
st.markdown("""
<div class="hero-banner">
    <p class="hero-title">🐦 Bird Species Observation Analysis</p>
    <p class="hero-sub">
        Comprehensive EDA &nbsp;·&nbsp; SQL Database Queries &nbsp;·&nbsp; Interactive Plotly Visualizations<br>
        Temporal &nbsp;·&nbsp; Spatial &nbsp;·&nbsp; Species &nbsp;·&nbsp; Environmental &nbsp;·&nbsp; Conservation Insights
    </p>
    <span class="hero-badge">🌿 Environmental Studies &amp; Biodiversity Conservation &amp; Ecology</span>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────
# GATE — datasets must exist
# ─────────────────────────────────────────────────────────────
if not (FOREST_PATH.exists() and GRASSLAND_PATH.exists()):
    st.error("⚠️ Dataset files not found. Please place the Excel files in the `data/` folder.")
    st.stop()

# ─────────────────────────────────────────────────────────────
# LOAD DATA
# ─────────────────────────────────────────────────────────────
with st.spinner("🔄 Loading & cleaning datasets…"):
    df   = load_data(str(FOREST_PATH), str(GRASSLAND_PATH))
    conn = build_sql(df)

# ─────────────────────────────────────────────────────────────
# SIDEBAR FILTERS
# ─────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 🔎 Filters")

    habitats  = sorted(df["Location_Type"].dropna().unique().tolist())
    sel_hab   = st.multiselect("🏕️ Habitat", habitats, default=habitats)

    units     = sorted(df["Admin_Unit_Code"].dropna().unique().tolist()) if "Admin_Unit_Code" in df.columns else []
    sel_units = st.multiselect("🗺️ Admin Unit", units, default=units)

    years = sorted(df["Year"].dropna().unique().tolist()) if "Year" in df.columns else []
    if len(years) >= 2:
        yr_range = st.slider("📅 Year", int(min(years)), int(max(years)),
                             (int(min(years)), int(max(years))))
    elif len(years) == 1:
        yr_range = (int(years[0]), int(years[0]))
    else:
        yr_range = (2000, 2030)

# Apply filters
dff = df.copy()
if sel_hab:
    dff = dff[dff["Location_Type"].isin(sel_hab)]
if sel_units and "Admin_Unit_Code" in dff.columns:
    dff = dff[dff["Admin_Unit_Code"].isin(sel_units)]
if "Year" in dff.columns and years:
    dff = dff[dff["Year"].between(yr_range[0], yr_range[1], inclusive="both")]

# ─────────────────────────────────────────────────────────────
# KPI CARDS
# ─────────────────────────────────────────────────────────────
total_obs    = len(dff)
unique_sp    = dff["Common_Name"].nunique()      if "Common_Name"      in dff.columns else 0
admin_cnt    = dff["Admin_Unit_Code"].nunique()  if "Admin_Unit_Code"  in dff.columns else 0
observer_cnt = dff["Observer"].nunique()         if "Observer"         in dff.columns else 0
hab_cnt      = dff["Location_Type"].nunique()    if "Location_Type"    in dff.columns else 0

watchlist = 0
if "PIF_Watchlist_Status" in dff.columns:
    watchlist = dff[dff["PIF_Watchlist_Status"].str.upper() == "TRUE"]["Common_Name"].nunique()

kpis = [
    ("📊", f"{total_obs:,}",    "Total Observations",  ""),
    ("🦜", f"{unique_sp:,}",    "Unique Species",       ""),
    ("🗺️", f"{admin_cnt}",      "Admin Units",          ""),
    ("🚨", f"{watchlist}",      "At-Risk Species",      "PIF Watchlist"),
    ("🏕️", f"{hab_cnt}",        "Habitats",             ""),
    ("👁️", f"{observer_cnt}",   "Observers",            ""),
]

cols = st.columns(6)
for col, (icon, val, label, sub) in zip(cols, kpis):
    col.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-icon">{icon}</div>
        <div class="kpi-value">{val}</div>
        <div class="kpi-label">{label}</div>
        {"<div class='kpi-sub'>"+sub+"</div>" if sub else ""}
    </div>""", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────
# TABS
# ─────────────────────────────────────────────────────────────
tabs = st.tabs([
    "🏠 Overview",
    "📅 Temporal",
    "🗺️ Spatial",
    "🦜 Species",
    "🌡️ Environment",
    "📏 Behaviour",
    "🚨 Conservation",
    "🗃️ SQL Queries",
    "💡 Insights",
])

# ══════════════════════════════════════════════════════════════
# TAB 0 ─ OVERVIEW
# ══════════════════════════════════════════════════════════════
with tabs[0]:
    st.markdown('<div class="sec-hdr">Dataset Overview</div>', unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    with c1:
        hab_cnt_df = dff["Location_Type"].value_counts().reset_index()
        hab_cnt_df.columns = ["Habitat","Count"]
        fig = px.pie(hab_cnt_df, values="Count", names="Habitat",
                     title="🌿 Observations by Habitat", hole=0.55,
                     color="Habitat", color_discrete_map=HC)
        fig.update_traces(textposition="outside", textinfo="percent+label",
                          marker=dict(line=dict(color="#0d1117",width=2)))
        T(fig, 380); st.plotly_chart(fig, use_container_width=True)

    with c2:
        sp_hab = dff.groupby("Location_Type")["Common_Name"].nunique().reset_index()
        sp_hab.columns = ["Habitat","Unique_Species"]
        fig2 = px.bar(sp_hab, x="Habitat", y="Unique_Species",
                      title="🦜 Unique Species per Habitat",
                      color="Habitat", color_discrete_map=HC, text_auto=True)
        T(fig2, 380); st.plotly_chart(fig2, use_container_width=True)

    st.markdown("---")
    st.markdown('<div class="sec-hdr">Dataset Preview</div>', unsafe_allow_html=True)
    prev_cols = [c for c in ["Admin_Unit_Code","Location_Type","Date","Season",
                              "Common_Name","Scientific_Name","Observer",
                              "Temperature","Humidity","PIF_Watchlist_Status"]
                 if c in dff.columns]
    st.dataframe(dff[prev_cols].head(200), use_container_width=True, height=340)

    st.markdown(f"""
    <div class="ins-card" style="margin-top:1rem">
        <h4>📋 Cleaning Summary</h4>
        <p>
        <b>{dff.shape[0]:,}</b> observations &nbsp;·&nbsp;
        <b>{dff.shape[1]}</b> columns &nbsp;·&nbsp;
        Forest: <b>{len(dff[dff['Location_Type']=='Forest']):,}</b> rows &nbsp;·&nbsp;
        Grassland: <b>{len(dff[dff['Location_Type']=='Grassland']):,}</b> rows &nbsp;·&nbsp;
        Remaining nulls: <b>{dff.isnull().sum().sum():,}</b>
        </p>
    </div>""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
# TAB 1 ─ TEMPORAL
# ══════════════════════════════════════════════════════════════
with tabs[1]:
    st.markdown('<div class="sec-hdr">Temporal Analysis</div>', unsafe_allow_html=True)

    # Yearly
    if "Year" in dff.columns:
        yd = dff.groupby(["Year","Location_Type"]).size().reset_index(name="Observations")
        fig = px.line(yd, x="Year", y="Observations", color="Location_Type", markers=True,
                      title="📅 Yearly Bird Observations by Habitat",
                      color_discrete_map=HC)
        T(fig, 380); st.plotly_chart(fig, use_container_width=True)

    # Monthly
    if "Month_Name" in dff.columns:
        MO = ["January","February","March","April","May","June",
              "July","August","September","October","November","December"]
        md = dff.groupby(["Month_Name","Location_Type"]).size().reset_index(name="Observations")
        md["Month_Name"] = pd.Categorical(md["Month_Name"], categories=MO, ordered=True)
        md.sort_values("Month_Name", inplace=True)
        fig2 = px.bar(md, x="Month_Name", y="Observations", color="Location_Type",
                      barmode="group", title="🗓️ Monthly Bird Observations by Habitat",
                      color_discrete_map=HC)
        T(fig2, 380); st.plotly_chart(fig2, use_container_width=True)

    # Seasonal
    if "Season" in dff.columns:
        SO = ["Spring","Summer","Fall","Winter"]
        c1, c2 = st.columns(2)
        with c1:
            sd = dff.groupby(["Season","Location_Type"]).size().reset_index(name="Observations")
            sd["Season"] = pd.Categorical(sd["Season"], categories=SO, ordered=True)
            sd.sort_values("Season", inplace=True)
            fig3 = px.bar(sd, x="Season", y="Observations", color="Location_Type",
                          barmode="group", title="🍂 Seasonal Activity by Habitat",
                          color_discrete_map=HC, text_auto=True)
            T(fig3, 360); st.plotly_chart(fig3, use_container_width=True)
        with c2:
            ss = dff.groupby(["Season","Location_Type"])["Common_Name"].nunique().reset_index(name="Unique_Species")
            ss["Season"] = pd.Categorical(ss["Season"], categories=SO, ordered=True)
            ss.sort_values("Season", inplace=True)
            fig4 = px.line(ss, x="Season", y="Unique_Species", color="Location_Type",
                           markers=True, title="🦜 Unique Species by Season",
                           color_discrete_map=HC)
            T(fig4, 360); st.plotly_chart(fig4, use_container_width=True)

    # Monthly × Admin Heatmap
    if "Month_Name" in dff.columns and "Admin_Unit_Code" in dff.columns:
        st.markdown('<div class="sec-hdr">Monthly × Admin Unit Heatmap</div>', unsafe_allow_html=True)
        MO = ["January","February","March","April","May","June",
              "July","August","September","October","November","December"]
        heat = (dff.groupby(["Admin_Unit_Code","Month_Name"]).size()
                   .reset_index(name="Obs")
                   .pivot(index="Admin_Unit_Code", columns="Month_Name", values="Obs")
                   .reindex(columns=MO, fill_value=0).fillna(0))
        fig5 = px.imshow(heat, text_auto=True,
                         title="🗓️ Monthly Observation Heatmap: Admin Unit × Month",
                         color_continuous_scale="YlOrRd", aspect="auto")
        T(fig5, 440); st.plotly_chart(fig5, use_container_width=True)

# ══════════════════════════════════════════════════════════════
# TAB 2 ─ SPATIAL
# ══════════════════════════════════════════════════════════════
with tabs[2]:
    st.markdown('<div class="sec-hdr">Spatial Analysis</div>', unsafe_allow_html=True)

    if "Admin_Unit_Code" in dff.columns:
        adm = (dff.groupby(["Admin_Unit_Code","Location_Type"])
                  .agg(Observations=("Common_Name","count"),
                       Unique_Species=("Common_Name","nunique"))
                  .reset_index())
        c1, c2 = st.columns(2)
        with c1:
            fig = px.bar(adm, x="Admin_Unit_Code", y="Observations", color="Location_Type",
                         barmode="group", title="🗺️ Observations per Admin Unit",
                         color_discrete_map=HC)
            T(fig, 390); st.plotly_chart(fig, use_container_width=True)
        with c2:
            fig2 = px.bar(adm, x="Admin_Unit_Code", y="Unique_Species", color="Location_Type",
                          barmode="group", title="🦜 Unique Species per Admin Unit",
                          color_discrete_map={"Forest":"#F18F01","Grassland":"#C73E1D"})
            T(fig2, 390); st.plotly_chart(fig2, use_container_width=True)

    # Treemap
    st.markdown('<div class="sec-hdr">Species Richness Treemap</div>', unsafe_allow_html=True)
    tm = dff.groupby(["Location_Type","Admin_Unit_Code","Common_Name"]).size().reset_index(name="Count")
    fig3 = px.treemap(tm, path=["Location_Type","Admin_Unit_Code","Common_Name"],
                      values="Count",
                      title="🌲 Species Richness: Habitat → Admin Unit → Species",
                      color="Count", color_continuous_scale="Viridis")
    T(fig3, 570); st.plotly_chart(fig3, use_container_width=True)

    # Top 20 hotspots
    if "Plot_Name" in dff.columns:
        st.markdown('<div class="sec-hdr">Top 20 Observation Hotspots</div>', unsafe_allow_html=True)
        ph = (dff.groupby(["Plot_Name","Location_Type"])
                 .agg(Observations=("Common_Name","count"),
                      Unique_Species=("Common_Name","nunique"))
                 .reset_index()
                 .sort_values("Observations", ascending=False).head(20))
        fig4 = px.bar(ph, x="Observations", y="Plot_Name", orientation="h",
                      color="Location_Type", title="📍 Top 20 Hotspots (Plot Level)",
                      color_discrete_map=HC)
        fig4.update_layout(yaxis={"categoryorder":"total ascending"})
        T(fig4, 540); st.plotly_chart(fig4, use_container_width=True)

# ══════════════════════════════════════════════════════════════
# TAB 3 ─ SPECIES
# ══════════════════════════════════════════════════════════════
with tabs[3]:
    st.markdown('<div class="sec-hdr">Species Analysis</div>', unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    with c1:
        ts = dff["Common_Name"].value_counts().head(20).reset_index()
        ts.columns = ["Common_Name","Count"]
        fig = px.bar(ts, x="Count", y="Common_Name", orientation="h",
                     title="🏆 Top 20 Most Observed Species",
                     color="Count", color_continuous_scale="Viridis")
        fig.update_layout(yaxis={"categoryorder":"total ascending"})
        T(fig, 540); st.plotly_chart(fig, use_container_width=True)

    with c2:
        f_sp = set(dff[dff["Location_Type"]=="Forest"]["Common_Name"].unique())
        g_sp = set(dff[dff["Location_Type"]=="Grassland"]["Common_Name"].unique())
        sh_sp= f_sp & g_sp
        venn = pd.DataFrame({
            "Category":["Forest Only","Grassland Only","Both Habitats"],
            "Count":[len(f_sp-g_sp), len(g_sp-f_sp), len(sh_sp)]
        })
        fig2 = px.pie(venn, values="Count", names="Category",
                      title="🌍 Species Distribution by Habitat Exclusivity",
                      color_discrete_sequence=["#2E86AB","#A23B72","#F18F01"], hole=0.42)
        fig2.update_traces(textposition="outside", textinfo="percent+label")
        T(fig2, 540); st.plotly_chart(fig2, use_container_width=True)

    # Sunburst — FIXED: use sort+head instead of apply (which drops cols in pandas 2.x)
    if "Season" in dff.columns:
        st.markdown('<div class="sec-hdr">Habitat → Season → Top Species (Sunburst)</div>', unsafe_allow_html=True)
        sb = dff.groupby(["Location_Type","Season","Common_Name"]).size().reset_index(name="Count")
        sb_top = (sb.sort_values("Count", ascending=False)
                    .groupby(["Location_Type","Season"])
                    .head(5)
                    .reset_index(drop=True))
        fig3 = px.sunburst(sb_top, path=["Location_Type","Season","Common_Name"],
                           values="Count",
                           title="☀️ Sunburst: Habitat → Season → Top 5 Species",
                           color="Count", color_continuous_scale="RdYlGn")
        T(fig3, 580); st.plotly_chart(fig3, use_container_width=True)

    # Sex ratio
    if "Sex" in dff.columns:
        st.markdown('<div class="sec-hdr">Sex Ratio Analysis</div>', unsafe_allow_html=True)
        c1, c2 = st.columns(2)
        with c1:
            sx = dff["Sex"].value_counts().reset_index()
            sx.columns = ["Sex","Count"]
            fig4 = px.bar(sx, x="Sex", y="Count",
                          title="⚥ Sex Distribution of Observed Birds",
                          color="Sex", text_auto=True,
                          color_discrete_sequence=px.colors.qualitative.Set2)
            T(fig4, 370); st.plotly_chart(fig4, use_container_width=True)
        with c2:
            fig5 = px.pie(sx, values="Count", names="Sex",
                          title="⚥ Sex Ratio Pie",
                          hole=0.4,
                          color_discrete_sequence=px.colors.qualitative.Set2)
            fig5.update_traces(textposition="outside", textinfo="percent+label")
            T(fig5, 370); st.plotly_chart(fig5, use_container_width=True)

    # ID Method
    if "ID_Method" in dff.columns:
        st.markdown('<div class="sec-hdr">Identification Method</div>', unsafe_allow_html=True)
        idm = dff["ID_Method"].value_counts().head(15).reset_index()
        idm.columns = ["ID_Method","Count"]
        fig6 = px.bar(idm, x="ID_Method", y="Count",
                      title="🔬 Bird Identification Methods Used", text_auto=True,
                      color="Count", color_continuous_scale="Blues")
        T(fig6, 370); st.plotly_chart(fig6, use_container_width=True)

# ══════════════════════════════════════════════════════════════
# TAB 4 ─ ENVIRONMENT
# ══════════════════════════════════════════════════════════════
with tabs[4]:
    st.markdown('<div class="sec-hdr">Environmental Conditions</div>', unsafe_allow_html=True)

    SO = ["Spring","Summer","Fall","Winter"]

    if "Temperature" in dff.columns and "Season" in dff.columns:
        subset = dff[dff["Season"].isin(SO)]
        c1, c2 = st.columns(2)
        with c1:
            fig = px.box(subset, x="Season", y="Temperature", color="Location_Type",
                         title="🌡️ Temperature by Season & Habitat",
                         category_orders={"Season":SO}, color_discrete_map=HC)
            T(fig, 410); st.plotly_chart(fig, use_container_width=True)
        with c2:
            if "Humidity" in dff.columns:
                fig2 = px.box(subset, x="Season", y="Humidity", color="Location_Type",
                              title="💧 Humidity by Season & Habitat",
                              category_orders={"Season":SO}, color_discrete_map=HC)
                T(fig2, 410); st.plotly_chart(fig2, use_container_width=True)

    # Temp vs Humidity histogram overlay
    st.markdown('<div class="sec-hdr">Temperature & Humidity Distributions</div>', unsafe_allow_html=True)
    fig3 = make_subplots(rows=1, cols=2,
                         subplot_titles=["🌡️ Temperature Distribution","💧 Humidity Distribution"])
    for loc, color in [("Forest","#2E86AB"),("Grassland","#A23B72")]:
        sub = dff[dff["Location_Type"]==loc]
        if "Temperature" in dff.columns:
            fig3.add_trace(go.Histogram(x=sub["Temperature"], name=f"{loc} Temp",
                                        opacity=0.7, marker_color=color, showlegend=True), row=1,col=1)
        if "Humidity" in dff.columns:
            fig3.add_trace(go.Histogram(x=sub["Humidity"], name=f"{loc} Humidity",
                                        opacity=0.7, marker_color=color, showlegend=False), row=1,col=2)
    fig3.update_layout(barmode="overlay", title_text="🌤️ Environmental Distributions by Habitat")
    T(fig3, 420); st.plotly_chart(fig3, use_container_width=True)

    # Temperature vs Observations scatter
    if "Temperature" in dff.columns:
        st.markdown('<div class="sec-hdr">Temperature vs Observations</div>', unsafe_allow_html=True)
        ta = dff.groupby(["Temperature","Location_Type"]).agg(Observations=("Common_Name","count")).reset_index()
        fig4 = px.scatter(ta, x="Temperature", y="Observations", color="Location_Type",
                          trendline="ols", title="🌡️ Temperature vs Bird Observations",
                          color_discrete_map=HC)
        T(fig4, 420); st.plotly_chart(fig4, use_container_width=True)

    # Correlation heatmap
    st.markdown('<div class="sec-hdr">Correlation Matrix</div>', unsafe_allow_html=True)
    nc = dff.select_dtypes(include="number").columns.tolist()
    if len(nc) >= 2:
        corr = dff[nc].corr()
        fig5 = px.imshow(corr, text_auto=".2f",
                         title="🔗 Correlation of Numerical Features",
                         color_continuous_scale="RdBu", aspect="auto")
        T(fig5, 520); st.plotly_chart(fig5, use_container_width=True)

    # Sky & Wind
    c1, c2 = st.columns(2)
    if "Sky" in dff.columns:
        with c1:
            sky = dff["Sky"].value_counts().head(10).reset_index()
            sky.columns = ["Sky","Count"]
            fig6 = px.bar(sky, x="Count", y="Sky", orientation="h",
                          title="🌤️ Sky Conditions", text_auto=True,
                          color="Count", color_continuous_scale="Blues")
            T(fig6, 360); st.plotly_chart(fig6, use_container_width=True)
    if "Wind" in dff.columns:
        with c2:
            wd = dff["Wind"].value_counts().head(10).reset_index()
            wd.columns = ["Wind","Count"]
            fig7 = px.bar(wd, x="Count", y="Wind", orientation="h",
                          title="💨 Wind Conditions", text_auto=True,
                          color="Count", color_continuous_scale="Teal")
            T(fig7, 360); st.plotly_chart(fig7, use_container_width=True)

    # Disturbance
    if "Disturbance" in dff.columns:
        st.markdown('<div class="sec-hdr">Disturbance Impact</div>', unsafe_allow_html=True)
        dstb = (dff.groupby("Disturbance")
                   .agg(Observations=("Common_Name","count"),
                        Unique_Species=("Common_Name","nunique"))
                   .reset_index().sort_values("Observations", ascending=False))
        fig8 = px.bar(dstb, x="Disturbance", y="Observations",
                      title="⚠️ Disturbance Impact on Observation Count",
                      text_auto=True, color="Unique_Species",
                      color_continuous_scale="Oranges")
        T(fig8, 390); st.plotly_chart(fig8, use_container_width=True)

# ══════════════════════════════════════════════════════════════
# TAB 5 ─ BEHAVIOUR
# ══════════════════════════════════════════════════════════════
with tabs[5]:
    st.markdown('<div class="sec-hdr">Distance & Behaviour Analysis</div>', unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    if "Distance" in dff.columns:
        with c1:
            di = dff["Distance"].value_counts().reset_index()
            di.columns = ["Distance","Count"]
            fig = px.bar(di, x="Distance", y="Count", title="📏 Distance from Observer",
                         text_auto=True, color="Count", color_continuous_scale="Teal")
            T(fig, 390); st.plotly_chart(fig, use_container_width=True)

    if "Flyover_Observed" in dff.columns:
        with c2:
            fly = dff["Flyover_Observed"].value_counts().reset_index()
            fly.columns = ["Flyover","Count"]
            fig2 = px.pie(fly, values="Count", names="Flyover",
                          title="🛫 Flyover Behaviour", hole=0.42,
                          color_discrete_sequence=["#2E86AB","#F18F01","#A23B72"])
            fig2.update_traces(textposition="outside", textinfo="percent+label")
            T(fig2, 390); st.plotly_chart(fig2, use_container_width=True)

    if "ID_Method" in dff.columns:
        st.markdown('<div class="sec-hdr">ID Method by Habitat</div>', unsafe_allow_html=True)
        idmh = dff.groupby(["ID_Method","Location_Type"]).size().reset_index(name="Count")
        fig3 = px.bar(idmh, x="ID_Method", y="Count", color="Location_Type",
                      barmode="group", title="🔍 Identification Method by Habitat",
                      color_discrete_map=HC, text_auto=True)
        T(fig3, 410); st.plotly_chart(fig3, use_container_width=True)

    if "Interval_Length" in dff.columns:
        st.markdown('<div class="sec-hdr">Observation Interval Analysis</div>', unsafe_allow_html=True)
        ivl = dff.groupby(["Interval_Length","Location_Type"]).size().reset_index(name="Count")
        fig4 = px.bar(ivl, x="Interval_Length", y="Count", color="Location_Type",
                      barmode="group", title="⏱️ Observation Interval Length",
                      color_discrete_map=HC)
        T(fig4, 390); st.plotly_chart(fig4, use_container_width=True)

    # Visit impact
    if "Visit" in dff.columns:
        st.markdown('<div class="sec-hdr">Repeated Visit Impact</div>', unsafe_allow_html=True)
        vis = (dff.groupby(["Visit","Location_Type"])
                  .agg(Observations=("Common_Name","count"),
                       Unique_Species=("Common_Name","nunique"))
                  .reset_index())
        fig5 = px.line(vis, x="Visit", y="Unique_Species", color="Location_Type",
                       markers=True, title="🔁 Repeated Visit Impact on Species Count",
                       color_discrete_map=HC)
        T(fig5, 390); st.plotly_chart(fig5, use_container_width=True)

    # Observer
    if "Observer" in dff.columns:
        st.markdown('<div class="sec-hdr">Top Observers</div>', unsafe_allow_html=True)
        obs_agg = {
            "Observations": ("Common_Name","count"),
            "Unique_Species": ("Common_Name","nunique"),
        }
        if "Plot_Name" in dff.columns:
            obs_agg["Sites_Visited"] = ("Plot_Name","nunique")
        obsd = (dff.groupby("Observer").agg(**obs_agg)
                   .reset_index()
                   .sort_values("Observations", ascending=False).head(20))
        fig6 = px.bar(obsd, x="Observer", y="Observations",
                      title="👁️ Top 20 Observers by Total Observations", text_auto=True,
                      color="Unique_Species", color_continuous_scale="Plasma")
        T(fig6, 410); st.plotly_chart(fig6, use_container_width=True)

# ══════════════════════════════════════════════════════════════
# TAB 6 ─ CONSERVATION
# ══════════════════════════════════════════════════════════════
with tabs[6]:
    st.markdown('<div class="sec-hdr">Conservation Insights</div>', unsafe_allow_html=True)

    if "PIF_Watchlist_Status" in dff.columns:
        c1, c2 = st.columns(2)
        with c1:
            wl = (dff.groupby(["PIF_Watchlist_Status","Location_Type"])
                     .agg(Species=("Common_Name","nunique")).reset_index())
            fig = px.bar(wl, x="PIF_Watchlist_Status", y="Species",
                         color="Location_Type", barmode="group",
                         title="🚨 PIF Watchlist Species by Habitat",
                         color_discrete_map=HC, text_auto=True)
            T(fig, 390); st.plotly_chart(fig, use_container_width=True)

        if "Regional_Stewardship_Status" in dff.columns:
            with c2:
                rs = (dff.groupby(["Regional_Stewardship_Status","Location_Type"])
                         .agg(Species=("Common_Name","nunique")).reset_index())
                fig2 = px.bar(rs, x="Regional_Stewardship_Status", y="Species",
                              color="Location_Type", barmode="group",
                              title="🛡️ Regional Stewardship by Habitat",
                              color_discrete_map={"Forest":"#F18F01","Grassland":"#C73E1D"},
                              text_auto=True)
                T(fig2, 390); st.plotly_chart(fig2, use_container_width=True)

    # At-risk heatmap
    if "PIF_Watchlist_Status" in dff.columns and "Admin_Unit_Code" in dff.columns:
        st.markdown('<div class="sec-hdr">At-Risk Species Heatmap</div>', unsafe_allow_html=True)
        ar = dff[dff["PIF_Watchlist_Status"].str.upper() == "TRUE"]
        if len(ar) > 0:
            pv = (ar.groupby(["Admin_Unit_Code","Location_Type"])
                    .agg(At_Risk=("Common_Name","nunique"))
                    .reset_index()
                    .pivot(index="Admin_Unit_Code", columns="Location_Type", values="At_Risk")
                    .fillna(0))
            fig3 = px.imshow(pv, text_auto=True,
                             title="🗺️ At-Risk Species: Admin Unit × Habitat",
                             color_continuous_scale="Reds")
            T(fig3, 440); st.plotly_chart(fig3, use_container_width=True)

    # Watchlist table
    st.markdown('<div class="sec-hdr">Watchlist Species Detail Table</div>', unsafe_allow_html=True)
    if "PIF_Watchlist_Status" in dff.columns:
        wlt = (dff[dff["PIF_Watchlist_Status"].str.upper() == "TRUE"]
               .groupby(["Common_Name","Scientific_Name","Location_Type"])
               .agg(Observations=("Common_Name","count"))
               .reset_index()
               .sort_values("Observations", ascending=False).head(30))
        st.dataframe(wlt, use_container_width=True, height=360)

# ══════════════════════════════════════════════════════════════
# TAB 7 ─ SQL QUERIES
# ══════════════════════════════════════════════════════════════
with tabs[7]:
    st.markdown('<div class="sec-hdr">SQL Database Queries</div>', unsafe_allow_html=True)

    SQL_QUERIES = {
        "📊 Observations by Habitat": """
SELECT Location_Type,
       COUNT(*)                        AS Total_Observations,
       COUNT(DISTINCT Common_Name)     AS Unique_Species,
       COUNT(DISTINCT Admin_Unit_Code) AS Admin_Units
FROM bird_observations
GROUP BY Location_Type
ORDER BY Total_Observations DESC""",

        "🏆 Top 15 Most Observed Species": """
SELECT Common_Name, Scientific_Name,
       COUNT(*) AS Observations,
       COUNT(DISTINCT Location_Type) AS Habitat_Types
FROM bird_observations
GROUP BY Common_Name, Scientific_Name
ORDER BY Observations DESC
LIMIT 15""",

        "🍂 Seasonal Trends by Habitat": """
SELECT Season, Location_Type,
       COUNT(*) AS Observations,
       COUNT(DISTINCT Common_Name) AS Unique_Species
FROM bird_observations
WHERE Season IS NOT NULL
GROUP BY Season, Location_Type
ORDER BY Season, Location_Type""",

        "🚨 At-Risk / Watchlist Species": """
SELECT Common_Name, Scientific_Name, AOU_Code,
       PIF_Watchlist_Status, Regional_Stewardship_Status,
       COUNT(*) AS Observations
FROM bird_observations
WHERE UPPER(PIF_Watchlist_Status)        = 'TRUE'
   OR UPPER(Regional_Stewardship_Status) = 'TRUE'
GROUP BY Common_Name, Scientific_Name, AOU_Code,
         PIF_Watchlist_Status, Regional_Stewardship_Status
ORDER BY Observations DESC
LIMIT 20""",

        "🌡️ Avg Environmental Conditions": """
SELECT Location_Type,
       ROUND(AVG(Temperature), 2) AS Avg_Temperature,
       ROUND(AVG(Humidity), 2)    AS Avg_Humidity,
       COUNT(*)                   AS Records
FROM bird_observations
GROUP BY Location_Type""",

        "📐 Biodiversity Index per Admin Unit": """
SELECT Admin_Unit_Code, Location_Type,
       COUNT(*)                    AS Total_Obs,
       COUNT(DISTINCT Common_Name) AS Unique_Species,
       ROUND(COUNT(DISTINCT Common_Name)*100.0/COUNT(*), 3) AS Diversity_Index
FROM bird_observations
GROUP BY Admin_Unit_Code, Location_Type
ORDER BY Diversity_Index DESC""",

        "🌿 Species in Both Habitats": """
SELECT Common_Name, Scientific_Name, AOU_Code,
       COUNT(DISTINCT Location_Type) AS Habitat_Count,
       COUNT(*) AS Total_Observations
FROM bird_observations
GROUP BY Common_Name, Scientific_Name, AOU_Code
HAVING Habitat_Count = 2
ORDER BY Total_Observations DESC
LIMIT 20""",

        "📅 Peak Observation Month per Admin Unit": """
SELECT Admin_Unit_Code, Location_Type, Month_Name, Observations
FROM (
    SELECT Admin_Unit_Code, Location_Type, Month_Name,
           COUNT(*) AS Observations,
           RANK() OVER (
               PARTITION BY Admin_Unit_Code, Location_Type
               ORDER BY COUNT(*) DESC
           ) AS rnk
    FROM bird_observations
    WHERE Month_Name IS NOT NULL
    GROUP BY Admin_Unit_Code, Location_Type, Month_Name
)
WHERE rnk = 1
ORDER BY Admin_Unit_Code""",
    }

    sel = st.selectbox("Select a pre-built query:", list(SQL_QUERIES.keys()))
    sql_code = SQL_QUERIES[sel]
    st.markdown(f'<div class="sql-box">{sql_code}</div>', unsafe_allow_html=True)
    res = qry(conn, sql_code)
    st.dataframe(res, use_container_width=True, height=320)

    st.markdown("---")
    st.markdown("#### ✏️ Custom SQL Query")
    custom_sql = st.text_area(
        "Write your own SQL (table: `bird_observations`)",
        "SELECT Common_Name, COUNT(*) AS Obs FROM bird_observations GROUP BY Common_Name ORDER BY Obs DESC LIMIT 10",
        height=110)
    if st.button("▶️ Run Query"):
        st.dataframe(qry(conn, custom_sql), use_container_width=True, height=320)

# ══════════════════════════════════════════════════════════════
# TAB 8 ─ INSIGHTS
# ══════════════════════════════════════════════════════════════
with tabs[8]:
    st.markdown('<div class="sec-hdr">Key Findings & Business Insights</div>', unsafe_allow_html=True)

    top_season  = dff["Season"].value_counts().idxmax()       if "Season"         in dff.columns else "N/A"
    top_admin   = dff["Admin_Unit_Code"].value_counts().idxmax() if "Admin_Unit_Code" in dff.columns else "N/A"
    top_species = dff["Common_Name"].value_counts().idxmax()  if "Common_Name"     in dff.columns else "N/A"
    top_sp_cnt  = int(dff["Common_Name"].value_counts().iloc[0]) if "Common_Name"  in dff.columns else 0

    f_sp2 = set(dff[dff["Location_Type"]=="Forest"]["Common_Name"].unique())
    g_sp2 = set(dff[dff["Location_Type"]=="Grassland"]["Common_Name"].unique())
    sh2   = f_sp2 & g_sp2

    insights = [
        ("📊 Dataset Scale",
         f"<b>{total_obs:,}</b> observations across <b>{admin_cnt}</b> admin units in <b>Forest</b> and <b>Grassland</b> habitats, covering <b>{unique_sp:,}</b> unique bird species."),
        ("🌿 Habitat Distribution",
         f"Forest: <b>{len(dff[dff['Location_Type']=='Forest']):,}</b> obs ({len(dff[dff['Location_Type']=='Forest'])/total_obs*100:.1f}%) &nbsp;·&nbsp; Grassland: <b>{len(dff[dff['Location_Type']=='Grassland']):,}</b> obs ({len(dff[dff['Location_Type']=='Grassland'])/total_obs*100:.1f}%). Forest shows higher overall observation density."),
        ("🍂 Peak Season",
         f"<b>{top_season}</b> records the highest bird activity. This aligns with breeding and migration patterns of temperate avian communities."),
        ("📍 Most Active Admin Unit",
         f"<b>{top_admin}</b> shows the highest observation count — a priority site for eco-tourism and conservation focus."),
        ("🏆 Most Common Species",
         f"<b>{top_species}</b> ({top_sp_cnt:,} observations) is the most frequently recorded species — a habitat generalist observed across both ecosystems."),
        ("🌍 Cross-Habitat Species",
         f"<b>{len(sh2):,}</b> species were recorded in BOTH Forest and Grassland habitats. These generalists require corridor-based conservation strategies."),
        ("🚨 At-Risk Species",
         f"<b>{watchlist}</b> species are on the PIF Watchlist. Targeted habitat protection and population monitoring programmes are essential."),
        ("👁️ Observer Network",
         f"Data contributed by <b>{observer_cnt}</b> field observers — reflecting a robust citizen science and field survey programme across 11 national park units."),
    ]

    for title, body in insights:
        st.markdown(f'<div class="ins-card"><h4>{title}</h4><p>{body}</p></div>',
                    unsafe_allow_html=True)

    st.markdown("---")
    st.markdown('<div class="sec-hdr">✅ Actionable Recommendations</div>', unsafe_allow_html=True)

    recs = [
        ("🛡️ 1. Conservation Priority",
         "Focus protection on at-risk watchlist species hotspots. Admin units with high at-risk counts and fragmented habitats need immediate intervention."),
        ("🌿 2. Eco-Tourism Development",
         f"Develop bird-watching tourism in high-diversity admin units during {top_season} (peak season) to boost local economies and fund conservation."),
        ("📏 3. Buffer Zones",
         "Establish buffer zones in high-disturbance sites to reduce anthropogenic impact on bird sightings and nesting behaviour."),
        ("🔗 4. Habitat Corridors",
         f"The {len(sh2):,} cross-habitat species need connected forest-grassland landscapes. Corridor development between fragmented habitats is a priority."),
        ("📊 5. Longitudinal Monitoring",
         "Maintain standardised monitoring across all 11 admin units with consistent protocols to detect population trends and early warning indicators."),
        ("🌡️ 6. Climate Adaptation",
         "Temperature and humidity data show measurable habitat differences (Forest: 21.87°C vs Grassland: 23.27°C). Model climate-change impacts on species distributions."),
    ]

    c1, c2 = st.columns(2)
    for i, (title, body) in enumerate(recs):
        with (c1 if i % 2 == 0 else c2):
            st.markdown(f'<div class="ins-card"><h4>{title}</h4><p>{body}</p></div>',
                        unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────
# FOOTER
# ─────────────────────────────────────────────────────────────
st.markdown("""
<div class="footer-bar">
    🐦 <b>Bird Species Observation Analysis</b> &nbsp;·&nbsp;
    Environmental Studies &amp; Biodiversity Conservation &amp; Ecology &nbsp;·&nbsp;
    Skills: Data Cleaning · EDA · SQL · Plotly · Streamlit
</div>
""", unsafe_allow_html=True)
