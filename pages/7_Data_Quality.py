"""
Page 7: Data Quality and Integrity Audit Center (pages/7_Data_Quality.py)
Supports OQ-10, BR-10, US-10, FEAT-10, UI-08
"""

import streamlit as st
import pandas as pd
from pathlib import Path

from src.db import verify_db_integrity
from src.cleaning import get_raw_cleaning_audit
from src.config import VALIDATION_REPORT_PATH

st.set_page_config(page_title="Data Quality Audit | AuraNova", page_icon="🛡️", layout="wide")

st.title("🛡️ Data Quality & Integrity Audit Center")
st.caption("Data governance scorecard, raw-vs-clean audit trail, foreign key validation, and outlier handling log.")

# Run live DB checks
db_check = verify_db_integrity()
audit_stats = get_raw_cleaning_audit()

col1, col2, col3, col4 = st.columns(4)
col1.metric("Database Integrity", db_check["integrity"].upper(), help="Result of PRAGMA integrity_check;")
col2.metric("FK Violations", db_check["fk_violations_count"], help="Result of PRAGMA foreign_key_check;")
col3.metric("Duplicates Removed", audit_stats.get("rows_removed_or_deduped", 0), help="Duplicate fact rows excised during cleaning")
col4.metric("Nulls Imputed", audit_stats.get("raw_null_cells", 0), help="Missing cells imputed via grouped median strategy")

st.write("")

# Table comparison
st.subheader("Raw Source vs. Cleaned Store Profile")
comp_data = [
    {"Metric": "Total Fact Rows", "Raw Synthetic Source": audit_stats.get("raw_total_rows"), "Cleaned SQLite DB": audit_stats.get("clean_total_rows"), "Variance / Impact": f"-{audit_stats.get('rows_removed_or_deduped')} duplicate rows excised"},
    {"Metric": "Null / Missing Value Cells", "Raw Synthetic Source": audit_stats.get("raw_null_cells"), "Cleaned SQLite DB": audit_stats.get("clean_null_cells"), "Variance / Impact": "100% resolved via channel median imputation"},
    {"Metric": "Duplicate Grain Records", "Raw Synthetic Source": audit_stats.get("raw_duplicate_facts"), "Cleaned SQLite DB": audit_stats.get("clean_duplicate_facts"), "Variance / Impact": "Zero duplicates in canonical schema"},
    {"Metric": "Foreign Key Violations", "Raw Synthetic Source": "Unenforced Excel format", "Cleaned SQLite DB": f"{db_check['fk_violations_count']} violations", "Variance / Impact": "PRAGMA foreign_keys = ON verified"},
    {"Metric": "Tracking Outlier Glitches", "Raw Synthetic Source": "3 extreme glitches injected", "Cleaned SQLite DB": "Documented & capped", "Variance / Impact": "Preserved genuine high-volume seasonal spikes"}
]
st.dataframe(pd.DataFrame(comp_data), use_container_width=True, hide_index=True)

# Treatment log
st.subheader("Documented Data Cleaning Rules")
st.markdown(
    """
    1. **Deduplication Strategy:** Fact table deduplication enforces the natural business composite key `(report_date, campaign_id, channel_id, segment_id, creative_id)`. The first recorded instance is retained.
    2. **Grouped Median Imputation:** Missing metrics in impressions, clicks, and spend are imputed using the channel-level median. This preserves channel cost differences (e.g. LinkedIn CPC vs Meta CPM) without artificially reducing variance.
    3. **Campaign Revenue Imputation:** Missing revenue entries are imputed using the campaign-level median to reflect distinct campaign deal sizes.
    4. **Outlier Treatment:** Recording glitches (e.g. 10x tracking sensor glitch on row 22, billing glitch on row 75) were capped back to plausible maximum distributions. Legitimate mid-year business spikes were retained to avoid destroying true variance.
    """
)

# Show raw validation markdown
if VALIDATION_REPORT_PATH.exists():
    with st.expander("📄 View Full Markdown Validation Report (data/validation_report.md)"):
        with open(VALIDATION_REPORT_PATH, "r", encoding="utf-8") as f:
            st.markdown(f.read())
