<div align="center">

# 🐦 Bird Species Observation Analysis

### An interactive biodiversity intelligence dashboard for Forest & Grassland habitats

[![Live Demo](https://img.shields.io/badge/🚀%20Live%20Demo-bird--analysis--app--027.streamlit.app-FF4B4B?style=for-the-badge&logo=streamlit&logoColor=white)](https://bird-analysis-app-027.streamlit.app/)
[![Python](https://img.shields.io/badge/Python-3.9+-3776AB?style=for-the-badge&logo=python&logoColor=white)](https://python.org)
[![Streamlit](https://img.shields.io/badge/Streamlit-1.x-FF4B4B?style=for-the-badge&logo=streamlit&logoColor=white)](https://streamlit.io)
[![Plotly](https://img.shields.io/badge/Plotly-Interactive-3F4F75?style=for-the-badge&logo=plotly&logoColor=white)](https://plotly.com)
[![SQLite](https://img.shields.io/badge/SQLite-In--Memory%20DB-003B57?style=for-the-badge&logo=sqlite&logoColor=white)](https://sqlite.org)
[![License: MIT](https://img.shields.io/badge/License-MIT-22c55e?style=for-the-badge)](LICENSE)

<br/>

> **Comprehensive EDA · SQL Queries · Interactive Plotly Visualizations**  
> Temporal · Spatial · Species · Environmental · Conservation Insights

<br/>

![Bird Analysis Banner](https://img.shields.io/badge/Domain-Environmental%20Studies%20%7C%20Biodiversity%20%7C%20Ecology-2E86AB?style=flat-square)
![Admin Units](https://img.shields.io/badge/Admin%20Units-11%20National%20Parks-A23B72?style=flat-square)
![Habitats](https://img.shields.io/badge/Habitats-Forest%20%26%20Grassland-3fb950?style=flat-square)

</div>

---

## 📌 Table of Contents

- [🌿 Project Overview](#-project-overview)
- [🚀 Live Demo](#-live-demo)
- [🗂️ Dashboard Tabs](#️-dashboard-tabs)
- [📊 Analysis Covered](#-analysis-covered)
- [🔧 Tech Stack](#-tech-stack)
- [⚡ Quick Start](#-quick-start)
- [📂 Project Structure](#-project-structure)
- [🗃️ SQL Queries](#️-sql-queries)
- [💡 Key Insights](#-key-insights)

---

## 🌿 Project Overview

This **interactive Streamlit dashboard** analyzes the distribution and biodiversity of bird species across **Forest** and **Grassland** habitats spanning **11 National Park administrative units** in the US:

<div align="center">

`ANTI` · `CATO` · `CHOH` · `GWMP` · `HAFE` · `MANA` · `MONO` · `NACE` · `PRWI` · `ROCR` · `WOTR`

</div>

The app combines data cleaning, exploratory data analysis, in-memory SQL querying, and interactive Plotly visualizations — all in one polished dark-themed dashboard built for ecologists, data analysts, and conservationists.

---

## 🚀 Live Demo

<div align="center">

### 👉 [https://bird-analysis-app-027.streamlit.app/](https://bird-analysis-app-027.streamlit.app/)

*No setup needed — open and explore instantly in your browser!*

</div>

---

## 🗂️ Dashboard Tabs

| # | Tab | What's Inside |
|---|-----|---------------|
| 🏠 | **Overview** | KPI cards, habitat distribution pie, unique species bar, data preview table, cleaning summary |
| 📅 | **Temporal** | Yearly trends, monthly grouped bars, seasonal activity, Admin Unit × Month heatmap |
| 🗺️ | **Spatial** | Observations per admin unit, species richness treemap, top 20 observation hotspots |
| 🦜 | **Species** | Top 20 species, habitat exclusivity pie, Habitat→Season→Species sunburst, sex ratio |
| 🌡️ | **Environment** | Temp/humidity boxplots, distribution histograms, OLS scatter, correlation matrix, sky/wind/disturbance |
| 📏 | **Behaviour** | Distance from observer, flyover analysis, ID method by habitat, interval length, observer leaderboard |
| 🚨 | **Conservation** | PIF Watchlist by habitat, Regional Stewardship, at-risk heatmap, watchlist detail table |
| 🗃️ | **SQL Queries** | 8 pre-built SQL queries + live custom SQL editor on in-memory SQLite |
| 💡 | **Insights** | Auto-generated key findings, 6 actionable conservation recommendations |

---

## 📊 Analysis Covered

```
✅ Data Cleaning & Preprocessing          ✅ Temporal Analysis (Year / Month / Season)
✅ In-Memory SQL Database (SQLite)        ✅ Spatial Analysis (Admin Unit / Plot Level)
✅ Advanced SQL (Window Functions, RANK)  ✅ Species Diversity & Exclusivity Analysis
✅ Environmental Conditions Analysis      ✅ Distance & Behaviour Analysis
✅ Observer Network Trends                ✅ PIF Watchlist & Conservation Insights
✅ Interactive Filters (Habitat / Unit / Year Range)
```

---

## 🔧 Tech Stack

| Tool | Purpose |
|------|---------|
| 🐍 **Python 3.9+** | Core language |
| 🎈 **Streamlit** | Interactive web dashboard |
| 📊 **Plotly Express / Graph Objects** | Bar, line, pie, box, scatter, treemap, sunburst, heatmap |
| 🐼 **Pandas** | Data loading, cleaning, transformation |
| 🗄️ **SQLite3** | In-memory SQL database for querying |
| 🔢 **NumPy** | Numerical operations |
| 🎨 **Custom CSS** | Premium dark-nature GitHub-inspired theme |

---

## ⚡ Quick Start

### Prerequisites
- Python **3.9+** installed
- The two `.XLSX` dataset files (Forest & Grassland)

### 1️⃣ Clone the repository

```bash
git clone https://github.com/D0027/bird-analysis-app.git
cd bird-analysis-app
```

### 2️⃣ Install dependencies

```bash
pip install -r requirements.txt
```

### 3️⃣ Add the datasets

Place both Excel files inside the `data/` folder:

```
bird-analysis-app/
└── data/
    ├── Bird_Monitoring_Data_FOREST.XLSX
    └── Bird_Monitoring_Data_GRASSLAND.XLSX
```

### 4️⃣ Run the app

```bash
streamlit run app.py
```

Your browser will open at **http://localhost:8501** 🎉

---

## 📂 Project Structure

```
bird-analysis-app/
│
├── 📄 app.py                             # Main Streamlit application
├── 📋 requirements.txt                   # Python dependencies
├── 📖 README.md                          # You're reading this!
│
├── 📁 data/
│   ├── 🐦 Bird_Monitoring_Data_FOREST.XLSX
│   └── 🌿 Bird_Monitoring_Data_GRASSLAND.XLSX
│
└── 📁 .streamlit/
    └── ⚙️ config.toml                    # Streamlit theme config
```

---

## 🗃️ SQL Queries

The app includes 8 pre-built SQL queries running on an in-memory SQLite database, plus a live custom SQL editor:

```sql
-- Example: Biodiversity Index per Admin Unit
SELECT Admin_Unit_Code, Location_Type,
       COUNT(*)                    AS Total_Obs,
       COUNT(DISTINCT Common_Name) AS Unique_Species,
       ROUND(COUNT(DISTINCT Common_Name)*100.0/COUNT(*), 3) AS Diversity_Index
FROM bird_observations
GROUP BY Admin_Unit_Code, Location_Type
ORDER BY Diversity_Index DESC;
```

Other included queries: Observations by Habitat · Top 15 Species · Seasonal Trends · At-Risk Watchlist · Avg Environmental Conditions · Species in Both Habitats · Peak Observation Month per Admin Unit

---

## 💡 Key Insights

- 🌿 **Forest habitat** shows higher overall observation density compared to Grassland
- 🍂 **Spring** records peak bird activity — aligning with breeding and migration seasons  
- 🚨 Multiple species flagged on the **PIF Watchlist** across both habitats
- 🌍 A significant number of species appear in **both Forest and Grassland** — requiring corridor-based conservation
- 🌡️ Forest avg temp (~21.87°C) is measurably cooler than Grassland (~23.27°C), affecting species distribution

---

## 👤 Author

<div align="center">

**Deepak Yadav**  
MCA · AI/ML & Full-Stack Developer  
P P Savani University, Bharuch, Gujarat

[![GitHub](https://img.shields.io/badge/GitHub-D0027-181717?style=for-the-badge&logo=github)](https://github.com/D0027)
[![LinkedIn](https://img.shields.io/badge/LinkedIn-deepakyadav027-0A66C2?style=for-the-badge&logo=linkedin)](https://linkedin.com/in/deepakyadav027)

</div>

---

<div align="center">

*🐦 Bird Species Observation Analysis · Environmental Studies & Biodiversity Conservation & Ecology*  
*Skills: Data Cleaning · EDA · SQL · Plotly · Streamlit*

**⭐ Star this repo if you found it useful!**

</div>
