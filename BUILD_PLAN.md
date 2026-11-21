# Project Build Plan & Engineering Roadmap

**Project:** Marketing Campaign Performance Analysis  
**Student:** Madhav Dutta (Roll: 2503201450085)  
**Programme:** MCA, 2nd Year, Data Analytics | ABES Engineering College, Ghaziabad  
**Mentor:** Prof. Kalpana Dwivedi  
**Date:** 2026-11-21  

---

## 1. Phased Build Tasks (BUILD-01 to BUILD-09)

| Task ID | Task Milestone | Primary Files | Status |
| :--- | :--- | :--- | :--- |
| **BUILD-01** | Environment, Pinned Dependencies & Configuration | `requirements.txt`, `src/config.py`, `.gitignore` | COMPLETED |
| **BUILD-02** | Database Schema, DDL & Relational Loader | `src/db.py`, `data/app.db`, `docs/08_Database_Design.xlsx` | COMPLETED |
| **BUILD-03** | Synthetic Data Generation, Noise Injection & Cleaning | `src/generate_data.py`, `data/marketing_campaign_data.xlsx` | COMPLETED |
| **BUILD-04** | Centralised KPI & Ratio-of-Totals Analytics Engine | `src/kpis.py`, `src/analysis.py` | COMPLETED |
| **BUILD-05** | Inferential Statistics & OLS Regression Module | `src/statistics.py` | COMPLETED |
| **BUILD-06** | Actionable Recommendations & Diagnostic Engine | `src/insights.py` | COMPLETED |
| **BUILD-07** | Tailwind UI Components & Multi-Page Dashboards | `src/ui.py`, `app.py`, `pages/*.py` | COMPLETED |
| **BUILD-08** | Automated Test Suites & Fixes (`pytest -q`) | `tests/test_data.py`, `test_analysis.py`, `test_app.py` | COMPLETED (11/11 PASSED) |
| **BUILD-09** | Deployment Packaging, README & Academic Reports | `README.md`, `BUILD_PLAN.md`, `docs/*.docx`, `docs/*.pptx` | COMPLETED |

---

## 2. Windows PowerShell Execution Commands

```powershell
# 1. Install dependencies
pip install -r requirements.txt

# 2. Generate canonical datasets and SQLite DB
python src/generate_data.py

# 3. Run automated pytest test suite
python -m pytest -q

# 4. Launch interactive Streamlit web dashboard
streamlit run app.py
```

---

## 3. Definition of Done (DoD)
- 100% of Must Business Requirements trace to code, database fields, UI screens, and tests.
- PRAGMA foreign_key_check passes with 0 violations.
- App starts cleanly on Windows PowerShell with zero runtime exceptions.
- Zero fabricated live URLs, test results, or business facts.
