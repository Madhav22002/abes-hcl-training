"""
Page 8: About, Academic Details and Methodology (pages/8_About.py)
Supports BR-12, US-12, FEAT-12, UI-09
"""

import streamlit as st
from src.config import PROJECT_METADATA

st.set_page_config(page_title="About & Methodology | AuraNova", page_icon="ℹ️", layout="wide")

st.title("ℹ️ About & Analytics Methodology")
st.caption("Academic project documentation, mathematical formulations, architectural governance, and synthetic data provenance.")

col1, col2 = st.columns(2)

with col1:
    st.subheader("🎓 Academic Metadata")
    st.markdown(
        f"""
        - **Project Title:** {PROJECT_METADATA['title']}
        - **Domain:** {PROJECT_METADATA['domain']}
        - **Student Name:** **{PROJECT_METADATA['student_name']}**
        - **University Roll Number:** `{PROJECT_METADATA['roll_number']}`
        - **Programme:** {PROJECT_METADATA['programme']}
        - **Academic Institution:** **{PROJECT_METADATA['institution']}**
        - **Project Mentor:** **Prof. {PROJECT_METADATA['mentor_name']}**
        - **Execution Date:** {PROJECT_METADATA['date']}
        - **Repository Identifier:** `{PROJECT_METADATA['repo_slug']}`
        """
    )

with col2:
    st.subheader("🏢 Fictional Case Organisation")
    st.markdown(
        f"""
        - **Organisation Name:** **{PROJECT_METADATA['organisation']}**
        - **Headquarters:** Sector 62, Noida, Uttar Pradesh, India
        - **Workforce Size:** ~500 Employees
        - **Industry:** Enterprise IT Consulting, Cloud Solutions & Digital Transformation Services
        - **Scenario Notice:** {PROJECT_METADATA['org_description']}
        - **Data Provenance:** Purely synthetic generation (Seed = `{PROJECT_METADATA['random_seed']}`).
        """
    )

st.divider()

st.subheader("📐 Standard KPI Mathematical Definitions")
st.markdown(
    """
    All aggregate ratio metrics are computed as **ratios of totals** across selected subsets rather than arithmetic averages of row ratios, preventing distortion and Simpson's Paradox:

    | KPI Name | Mathematical Formula | Business Interpretation | Zero Denominator Handling |
    | :--- | :--- | :--- | :--- |
    | **Click-Through Rate (CTR)** | $\\frac{\\sum \\text{Clicks}}{\\sum \\text{Impressions}} \\times 100$ | Creative resonance and audience interaction rate | Returns `None` (`"—"`) if Impressions = 0 |
    | **Cost Per Click (CPC)** | $\\frac{\\sum \\text{Spend}}{\\sum \\text{Clicks}}$ | Unit cost incurred per visitor traffic interaction | Returns `None` (`"—"`) if Clicks = 0 |
    | **Cost Per Mille (CPM)** | $\\frac{\\sum \\text{Spend}}{\\sum \\text{Impressions}} \\times 1000$ | Cost per 1,000 ad impression displays served | Returns `None` (`"—"`) if Impressions = 0 |
    | **Conversion Rate (CVR)** | $\\frac{\\sum \\text{Conversions}}{\\sum \\text{Clicks}} \\times 100$ | Traffic-to-lead qualification efficiency | Returns `None` (`"—"`) if Clicks = 0 |
    | **Cost Per Acquisition (CPA)**| $\\frac{\\sum \\text{Spend}}{\\sum \\text{Conversions}}$ | Capital required to secure one qualified customer action | Returns `None` (`"—"`) if Conversions = 0 |
    | **Return on Ad Spend (ROAS)** | $\\frac{\\sum \\text{Revenue}}{\\sum \\text{Spend}}$ | Gross revenue multiple generated per unit marketing spend | Returns `None` (`"—"`) if Spend = 0 |
    | **Return on Investment (ROI)** | $\\frac{\\sum \\text{Revenue} - \\sum \\text{Spend}}{\\sum \\text{Spend}} \\times 100$ | Net profit or loss percentage yielded by campaign | Returns `None` (`"—"`) if Spend = 0 |
    """
)

st.divider()

st.subheader("🛠️ Software Architecture & Tech Stack")
st.markdown(
    """
    - **Language & Runtime:** Python 3.12 / 3.13
    - **Web Application:** Streamlit Multipage Architecture (v1.58.0)
    - **Visualisation Engine:** Plotly Graph Objects & Plotly Express (v7.1.0)
    - **Database Tier:** SQLite 3 (`data/app.db`) with Foreign Key constraints enabled
    - **Statistical Tooling:** SciPy (v1.17.1), statsmodels (v0.15.0), NumPy, pandas
    - **Styling Components:** Tailwind CSS CDN embedded through `st.components.v1.html`
    - **Automated Verification:** pytest test suite & Streamlit AppTest framework
    """
)
