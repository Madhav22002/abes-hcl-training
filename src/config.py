"""
Project Configuration, Master ID Register, and Assumptions
Project: Marketing Campaign Performance Analysis
Student: Madhav Dutta (Roll No: 2503201450085)
Institution: ABES Engineering College, Ghaziabad
Mentor: Kalpana Dwivedi
Organisation: AuraNova Digital Solutions Pvt. Ltd. (Noida, UP, India)
"""

from pathlib import Path

# Paths
BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = BASE_DIR / "data"
DOCS_DIR = BASE_DIR / "docs"
SCREENSHOTS_DIR = DOCS_DIR / "screenshots"
DB_PATH = DATA_DIR / "app.db"
RAW_DATA_PATH = DATA_DIR / "marketing_campaign_data_raw.xlsx"
CLEAN_DATA_PATH = DATA_DIR / "marketing_campaign_data.xlsx"
VALIDATION_REPORT_PATH = DATA_DIR / "validation_report.md"

# Project Metadata
PROJECT_METADATA = {
    "title": "Marketing Campaign Performance Analysis",
    "domain": "Data Analytics / Marketing Analytics",
    "student_name": "Madhav Dutta",
    "roll_number": "2503201450085",
    "programme": "MCA, 2nd Year, Data Analytics",
    "institution": "ABES Engineering College, Ghaziabad",
    "mentor_name": "Kalpana Dwivedi",
    "date": "2026-11-21",
    "repo_slug": "Madhav22002/abes-hcl-training",
    "github_url": "https://github.com/Madhav22002/abes-hcl-training",
    "streamlit_url": "https://25003201450087-p078.streamlit.app/",
    "organisation": "AuraNova Digital Solutions Pvt. Ltd.",
    "org_description": (
        "A fictional 500-employee IT consulting, cloud solutions, and digital-services "
        "provider headquartered in Sector 62, Noida, Uttar Pradesh, India. "
        "All data and scenarios are strictly synthetic for academic demonstration."
    ),
    "date_range": ("2026-01-01", "2026-09-30"),
    "random_seed": 42,
    "target_fact_rows": 260
}

# Master ID Register
ID_REGISTER = {
    "stakeholders": [
        {"id": "SH-01", "name": "Chief Marketing Officer (CMO)", "role": "Executive Sponsor", "influence": "High"},
        {"id": "SH-02", "name": "Head of Digital Marketing / Performance Lead", "role": "Operational Owner", "influence": "High"},
        {"id": "SH-03", "name": "Campaign Managers / Media Buyers", "role": "Execution Specialist", "influence": "Medium"},
        {"id": "SH-04", "name": "Finance & Budget Controller", "role": "Financial Auditor", "influence": "High"},
        {"id": "SH-05", "name": "Chief Technology Officer / Solution Architect", "role": "Technical Advisor", "influence": "Medium"}
    ],
    "operational_questions": [
        {"id": "OQ-01", "question": "What is the overall performance across all digital campaigns in spend, conversions, and revenue?"},
        {"id": "OQ-02", "question": "How do spend, impressions, clicks, conversions, and revenue trend over time across monthly/quarterly intervals?"},
        {"id": "OQ-03", "question": "Which channels generate the highest ROAS and lowest CPA, and how do their efficiency profiles compare?"},
        {"id": "OQ-04", "question": "Which specific campaigns are top-performing vs. severely underperforming relative to targets and ROAS benchmarks?"},
        {"id": "OQ-05", "question": "How effectively is the allocated campaign budget being utilised, and are any campaigns over-exhausted or underspent?"},
        {"id": "OQ-06", "question": "Which regional audience segments and age brackets drive the best conversion rates and revenue?"},
        {"id": "OQ-07", "question": "Which creative types, formats, and messaging themes produce the highest click-through rates and lowest CPA?"},
        {"id": "OQ-08", "question": "Is there a statistically significant relationship between marketing spend and attributed revenue across campaigns?"},
        {"id": "OQ-09", "question": "Which channel and campaign combinations represent underperformance that should be paused or reallocated?"},
        {"id": "OQ-10", "question": "What is the health, completeness, and integrity profile of the marketing analytics dataset?"}
    ],
    "how_might_we": [
        {"id": "HMW-01", "statement": "How might we give executive leaders an instant consolidated view of multi-channel marketing ROI and ROAS?", "sh": "SH-01", "oq": "OQ-01", "priority": "P1"},
        {"id": "HMW-02", "statement": "How might we enable performance leads to track longitudinal daily and monthly trends across key KPIs?", "sh": "SH-02", "oq": "OQ-02", "priority": "P2"},
        {"id": "HMW-03", "statement": "How might we empower media buyers to objectively benchmark channel cost-efficiency (CPC, CPA, CPM) for budget reallocation?", "sh": "SH-03", "oq": "OQ-03", "priority": "P1"},
        {"id": "HMW-04", "statement": "How might we systematically identify underperforming campaigns that burn budget without reaching minimum ROAS thresholds?", "sh": "SH-02", "oq": "OQ-04", "priority": "P1"},
        {"id": "HMW-05", "statement": "How might we allow finance teams to audit planned vs. actual budget burn across campaigns in real time?", "sh": "SH-04", "oq": "OQ-05", "priority": "P2"},
        {"id": "HMW-06", "statement": "How might we dissect audience demographics (regions, age bands, devices) to uncover top-converting customer cohorts?", "sh": "SH-02", "oq": "OQ-06", "priority": "P3"},
        {"id": "HMW-07", "statement": "How might we evaluate creative formats and offer types to optimise ad fatigue and engagement?", "sh": "SH-03", "oq": "OQ-07", "priority": "P3"},
        {"id": "HMW-08", "statement": "How might we statistically validate the spend-to-revenue elasticity and correlation rigorously?", "sh": "SH-01", "oq": "OQ-08", "priority": "P2"},
        {"id": "HMW-09", "statement": "How might we automate actionable recommendations for campaign pausing, scaling, and budget reallocation?", "sh": "SH-02", "oq": "OQ-09", "priority": "P1"},
        {"id": "HMW-10", "statement": "How might we provide data-quality auditing and cleaning traceability to assure data integrity?", "sh": "SH-05", "oq": "OQ-10", "priority": "P1"}
    ],
    "business_requirements": [
        {"id": "BR-01", "req": "The business needs an executive KPI scorecard of Spend, Revenue, Conversions, ROAS, and ROI so that executive leadership can evaluate overall portfolio profitability at a glance.", "priority": "Must", "sh": "SH-01", "hmw": "HMW-01"},
        {"id": "BR-02", "req": "The business needs interactive multi-metric time-series trend visualisations so that marketing leads can spot seasonality, momentum, and demand peaks.", "priority": "Must", "sh": "SH-02", "hmw": "HMW-02"},
        {"id": "BR-03", "req": "The business needs cross-channel efficiency benchmarking (CPC, CPA, CTR, ROAS) so that media buyers can redirect spend towards highest-yield channels.", "priority": "Must", "sh": "SH-03", "hmw": "HMW-03"},
        {"id": "BR-04", "req": "The business needs campaign-level head-to-head performance comparisons with benchmark flags so that underperforming campaigns can be flagged.", "priority": "Must", "sh": "SH-02", "hmw": "HMW-04"},
        {"id": "BR-05", "req": "The business needs planned vs. actual budget utilisation tracking so that finance controllers prevent cost overruns.", "priority": "Must", "sh": "SH-04", "hmw": "HMW-05"},
        {"id": "BR-06", "req": "The business needs audience demographic and regional segmentation breakdown so that campaigns can target high-converting business clusters.", "priority": "Should", "sh": "SH-02", "hmw": "HMW-06"},
        {"id": "BR-07", "req": "The business needs creative asset performance analysis by format and messaging theme so that creative designers iterate on winning concepts.", "priority": "Should", "sh": "SH-03", "hmw": "HMW-07"},
        {"id": "BR-08", "req": "The business needs statistical correlation and hypothesis testing between marketing spend and revenue so that investment decisions are scientifically backed.", "priority": "Must", "sh": "SH-01", "hmw": "HMW-08"},
        {"id": "BR-09", "req": "The business needs automated rule-based actionable insights for budget reallocation and campaign optimisation so that operational teams save analysis time.", "priority": "Must", "sh": "SH-02", "hmw": "HMW-09"},
        {"id": "BR-10", "req": "The business needs a transparent data quality scorecard reporting nulls, duplicates, outliers, and cleaning rules applied so that data governance standards are maintained.", "priority": "Must", "sh": "SH-05", "hmw": "HMW-10"},
        {"id": "BR-11", "req": "The business needs a granular data explorer with filtering and CSV/Excel export capabilities so that analysts can perform ad-hoc offline reporting.", "priority": "Should", "sh": "SH-03", "hmw": "HMW-03"},
        {"id": "BR-12", "req": "The business needs an in-app methodology and documentation section so that users understand KPI calculations, caveats, and data provenance.", "priority": "Could", "sh": "SH-01", "hmw": "HMW-01"}
    ],
    "features": [
        {"id": "FEAT-01", "name": "Executive KPI Scorecard & Summary Dashboard", "br": "BR-01", "priority": "Must"},
        {"id": "FEAT-02", "name": "Time-Series Trend Explorer", "br": "BR-02", "priority": "Must"},
        {"id": "FEAT-03", "name": "Channel Performance Benchmarking", "br": "BR-03", "priority": "Must"},
        {"id": "FEAT-04", "name": "Campaign Comparison & Target Benchmarking", "br": "BR-04", "priority": "Must"},
        {"id": "FEAT-05", "name": "Budget Utilisation & Variance Tracker", "br": "BR-05", "priority": "Must"},
        {"id": "FEAT-06", "name": "Audience & Regional Segmentation Module", "br": "BR-06", "priority": "Should"},
        {"id": "FEAT-07", "name": "Creative Asset Performance Matrix", "br": "BR-07", "priority": "Should"},
        {"id": "FEAT-08", "name": "Inferential Statistics & Correlation Engine", "br": "BR-08", "priority": "Must"},
        {"id": "FEAT-09", "name": "Automated Actionable Insights Generator", "br": "BR-09", "priority": "Must"},
        {"id": "FEAT-10", "name": "Data Quality & Data Provenance Inspector", "br": "BR-10", "priority": "Must"},
        {"id": "FEAT-11", "name": "Data Explorer & Export Center", "br": "BR-11", "priority": "Should"},
        {"id": "FEAT-12", "name": "Methodology & Assumptions Guide", "br": "BR-12", "priority": "Could"}
    ],
    "ui_screens": [
        {"id": "UI-01", "name": "Overview & KPI Dashboard", "page": "app.py"},
        {"id": "UI-02", "name": "Campaign & Channel Trends", "page": "pages/1_Trends.py"},
        {"id": "UI-03", "name": "Channel & Campaign Comparison", "page": "pages/2_Comparison.py"},
        {"id": "UI-04", "name": "EDA & Inferential Statistics", "page": "pages/3_Statistics.py"},
        {"id": "UI-05", "name": "Audience & Creative Analysis", "page": "pages/4_Audience_Creative.py"},
        {"id": "UI-06", "name": "Insights & Recommendations", "page": "pages/5_Insights.py"},
        {"id": "UI-07", "name": "Data Explorer & Downloads", "page": "pages/6_Data_Explorer.py"},
        {"id": "UI-08", "name": "Data Quality & Health Audit", "page": "pages/7_Data_Quality.py"},
        {"id": "UI-09", "name": "About & Methodology", "page": "pages/8_About.py"}
    ]
}

# Standard KPI Math Formulations
# All ratio metrics use total aggregate sums to prevent Simpson's Paradox / distorted averages
def calc_kpis(impressions: float, clicks: float, conversions: float, spend: float, revenue: float):
    ctr = (clicks / impressions * 100.0) if impressions > 0 else None
    cpc = (spend / clicks) if clicks > 0 else None
    cpm = (spend / impressions * 1000.0) if impressions > 0 else None
    cvr = (conversions / clicks * 100.0) if clicks > 0 else None
    cpa = (spend / conversions) if conversions > 0 else None
    roas = (revenue / spend) if spend > 0 else None
    roi = ((revenue - spend) / spend * 100.0) if spend > 0 else None

    return {
        "Impressions": impressions,
        "Clicks": clicks,
        "Conversions": conversions,
        "Spend": spend,
        "Revenue": revenue,
        "CTR": ctr,
        "CPC": cpc,
        "CPM": cpm,
        "CVR": cvr,
        "CPA": cpa,
        "ROAS": roas,
        "ROI": roi
    }
