# Marketing Campaign Performance Analysis

**Academic Project — MCA 2nd Year (Data Analytics)**  
**Institution:** ABES Engineering College, Ghaziabad  
**Student:** Madhav Dutta  
**Roll Number:** 2503201450085  
**Project Mentor:** Prof. Kalpana Dwivedi  
**Date:** November 21, 2026  
**Repository Slug:** `2503201450085-marketing-campaign-performance-analysis`

---

## 1. Executive Summary & Working Scenario

This project delivers an enterprise-grade **Marketing Campaign Performance Analysis** web application and comprehensive analytics suite. The solution empowers marketing executives, campaign managers, and finance controllers to ingest, clean, explore, benchmark, and extract actionable insights from multi-channel digital campaigns.

### Working Organisation Assumption
- **Organisation Name:** **AuraNova Digital Solutions Pvt. Ltd.**
- **Headquarters:** Sector 62, Noida, Uttar Pradesh, India
- **Scale:** 500-employee IT consulting, cloud migration, and enterprise digital-services firm.
- **Scenario Notice:** All business scenarios, client portfolios, metrics, and data records are **100% synthetic** and generated strictly for academic and analytical demonstration. No real personal identifiable information (PII) or proprietary commercial data is used.

---

## 2. Architecture & Tech Stack

- **Application Framework:** Streamlit (v1.58.0) Multi-Page Architecture
- **Language & Runtime:** Python 3.12 / 3.13
- **Data Manipulation & Analysis:** pandas, NumPy, SciPy, statsmodels
- **Interactive Visualisations:** Plotly Express & Plotly Graph Objects
- **Database:** SQLite (`data/app.db`) with Foreign Key integrity constraints (`PRAGMA foreign_keys = ON`)
- **Testing:** pytest & Streamlit AppTest framework
- **Styling & Cards:** Responsive Tailwind CSS components embedded via supported HTML component APIs (`st.components.v1.html`)
- **Deployment Target:** Streamlit Community Cloud

---

## 3. Project Directory Structure

```text
├── app.py                         # Application Entry Point & Executive KPI Dashboard
├── pages/
│   ├── 1_Trends.py                # Longitudinal Time-Series Analysis
│   ├── 2_Comparison.py            # Channel & Campaign Efficiency Benchmarking
│   ├── 3_Statistics.py            # Inferential Statistics & OLS Regression
│   ├── 4_Audience_Creative.py     # Demographics, Segments & Creative Asset Matrix
│   ├── 5_Insights.py              # Automated Actionable Recommendations Engine
│   ├── 6_Data_Explorer.py         # Tabular Filtering & Excel/CSV Exports
│   ├── 7_Data_Quality.py          # Data Health, Null Profiles & Cleaning Log
│   └── 8_About.py                 # Methodology, KPI Formulas & Academic Details
├── src/
│   ├── config.py                  # Project Configuration, Metadata & Master ID Registry
│   ├── db.py                      # SQLite Connection & Query Manager
│   ├── data_loader.py             # Data Fetching & Caching Layer
│   ├── cleaning.py                # Data Cleaning & Outlier Treatment Engine
│   ├── validation.py              # Schema & Business Constraint Validation
│   ├── kpis.py                    # Standardised KPI Calculation Engine
│   ├── analysis.py                # Aggregations, Groupings & Channel Comparisons
│   ├── statistics.py              # Parametric/Non-Parametric Tests & Regression
│   ├── insights.py                # Rule-Based Actionable Recommendations
│   └── ui.py                      # Tailwind CSS Header & KPI Card Components
├── data/
│   ├── app.db                     # Normalized SQLite Database
│   ├── marketing_campaign_data_raw.xlsx  # Raw Synthetic Dataset (Deliberate Issues)
│   ├── marketing_campaign_data.xlsx      # Cleaned Multi-Sheet Canonical Workbook
│   ├── validation_report.md       # Data Integrity & Cleaning Audit Report
│   └── *.csv                      # Table-specific canonical CSV exports
├── tests/
│   ├── test_data.py               # Database Schema, Integrity & Cleaning Tests
│   ├── test_analysis.py           # Mathematical Verification of KPIs & Stats
│   └── test_app.py                # Streamlit Page Smoke & Component Tests
├── docs/                          # Academic Documentation Deliverables (DOCX, XLSX, PPTX)
│   ├── 01_Project_Understanding.docx
│   ├── 02_HMW_Statements.docx
│   ├── 03_Needs_Analysis.docx
│   ├── 04_Traceability_Matrix.xlsx
│   ├── 05_BRD.docx
│   ├── 06_PRD.docx
│   ├── 07_TRD.docx
│   ├── 08_Database_Design.xlsx
│   ├── 09_ER_Diagram.docx / .mmd / .png
│   ├── 11_User_Stories.docx
│   ├── 12_UI_Framework.docx
│   ├── 13_Test_Cases.docx
│   ├── 14_BUILD_PLAN.docx
│   ├── 2503201450085_Project_Document.docx
│   ├── 2503201450085_Presentation.pptx
│   └── screenshots/               # Verified Application UI Screenshots
├── requirements.txt               # Pinned Python Dependencies
├── BUILD_PLAN.md                  # Machine-readable Step-by-Step Build Plan
├── .gitignore                     # Git Exclusion Rules
└── README.md                      # Primary Project Documentation
```

---

## 4. Master ID Conventions

| Entity | ID Pattern | Description |
| :--- | :--- | :--- |
| **Stakeholders** | `SH-01` to `SH-05` | Key business & technical roles |
| **Operational Questions** | `OQ-01` to `OQ-10` | Core analytical queries |
| **How Might We** | `HMW-01` to `HMW-10` | Solution-neutral problem statements |
| **Business Requirements** | `BR-01` to `BR-12` | Capabilities mapped to value |
| **Product Features** | `FEAT-01` to `FEAT-12` | Functional application features |
| **User Stories** | `US-01` to `US-12` | Agile personas with acceptance criteria |
| **UI Screens** | `UI-01` to `UI-09` | Multi-page Streamlit dashboards |
| **Functional Tests** | `TC-F-01` to `TC-F-12` | Black-box user workflow tests |
| **Technical Tests** | `TC-T-01` to `TC-T-10` | Deterministic unit & integration tests |
| **Build Tasks** | `BUILD-01` to `BUILD-09` | Phased implementation milestones |

---

## 5. Standard KPI Formulations

All ratio metrics represent **ratios of aggregate sums** across selected slices rather than naive averages of row-level ratios:

- **Click-Through Rate (CTR):** $\text{CTR} = \frac{\sum \text{Clicks}}{\sum \text{Impressions}} \times 100$
- **Cost Per Click (CPC):** $\text{CPC} = \frac{\sum \text{Spend}}{\sum \text{Clicks}}$ (₹)
- **Cost Per Mille (CPM):** $\text{CPM} = \frac{\sum \text{Spend}}{\sum \text{Impressions}} \times 1000$ (₹)
- **Conversion Rate (CVR):** $\text{CVR} = \frac{\sum \text{Conversions}}{\sum \text{Clicks}} \times 100$
- **Cost Per Acquisition (CPA):** $\text{CPA} = \frac{\sum \text{Spend}}{\sum \text{Conversions}}$ (₹)
- **Return on Ad Spend (ROAS):** $\text{ROAS} = \frac{\sum \text{Revenue}}{\sum \text{Spend}}$
- **Return on Investment (ROI):** $\text{ROI} = \frac{\sum \text{Revenue} - \sum \text{Spend}}{\sum \text{Spend}} \times 100$ (%)

*Zero Division Guard:* When denominators evaluate to zero, all functions return `None` (rendered in UI tables as `"—"`), avoiding `ZeroDivisionError` or infinite values.

---

## 6. Installation & Execution (Windows PowerShell)

### Prerequisites
- Python 3.12 or 3.13 installed with PATH configured.

### Step 1: Open PowerShell and Navigate to Project
```powershell
cd "c:\Users\HP\Desktop\hcl details"
```

### Step 2: Install Dependencies
```powershell
pip install -r requirements.txt
```

### Step 3: Generate Canonical Datasets & SQLite Database
```powershell
python src/generate_data.py
```

### Step 4: Run the Complete Test Suite
```powershell
pytest -q
```

### Step 5: Launch the Streamlit Web Application
```powershell
streamlit run app.py
```
The application will open automatically in your browser at `http://localhost:8501`.

---

## 7. Deployment Readiness

- Repository is structured with `app.py` and `requirements.txt` at the root.
- All file operations use deterministic `Path(__file__).resolve()` paths.
- Database (`data/app.db`) is committed and bundled for zero-configuration Community Cloud deployment.
- Live URL: `PENDING` (Ready for push to `https://github.com/madhavdutta/2503201450085-marketing-campaign-performance-analysis`).
