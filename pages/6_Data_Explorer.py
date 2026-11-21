"""
Page 6: Data Explorer and Multi-Format Downloads (pages/6_Data_Explorer.py)
Supports BR-11, US-11, FEAT-11, UI-07
"""

import io
import streamlit as st
import pandas as pd

from src.data_loader import (
    load_full_performance_data,
    get_filter_choices,
    apply_filters
)

st.set_page_config(page_title="Data Explorer | AuraNova", page_icon="📁", layout="wide")

st.title("📁 Data Explorer & Analytics Export Center")
st.caption("Granular record-level exploration, custom column projections, and validated CSV/Excel data exports.")

df_raw = load_full_performance_data()
choices = get_filter_choices(df_raw)

# Sidebar
st.sidebar.markdown("### 🎯 Active Filters")
selected_camps = st.sidebar.multiselect("Campaigns", options=choices["campaigns"], default=st.session_state.get("filter_campaigns", []))
selected_chans = st.sidebar.multiselect("Channels", options=choices["channels"], default=st.session_state.get("filter_channels", []))
selected_regs = st.sidebar.multiselect("Regions", options=choices["regions"], default=st.session_state.get("filter_regions", []))

df = apply_filters(df_raw, selected_campaigns=selected_camps, selected_channels=selected_chans, selected_regions=selected_regs)

if len(df) == 0:
    st.warning("No data found matching active filters.")
    st.stop()

# Controls
st.markdown(f"**Total Records Matching Filters:** `{len(df)}` fact records")

default_cols = [
    "report_date", "campaign_name", "channel_name", "region", "creative_name",
    "impressions", "clicks", "conversions", "spend_inr", "revenue_inr", "ctr_pct", "cpc_inr", "roas"
]
available_cols = list(df.columns)
selected_cols = st.multiselect("Select Visible Columns", options=available_cols, default=[c for c in default_cols if c in available_cols])

if not selected_cols:
    selected_cols = default_cols

view_df = df[selected_cols].copy()

# Sort controls
sort_col = st.selectbox("Sort Table By", options=selected_cols, index=0)
sort_asc = st.checkbox("Ascending Order", value=False)
view_df = view_df.sort_values(sort_col, ascending=sort_asc)

# Display Table
st.dataframe(view_df, use_container_width=True, hide_index=True)

# Export Buttons
col_exp1, col_exp2 = st.columns(2)

with col_exp1:
    csv_bytes = view_df.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="📥 Download Filtered Data as CSV",
        data=csv_bytes,
        file_name="marketing_campaign_performance_filtered.csv",
        mime="text/csv",
        use_container_width=True
    )

with col_exp2:
    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
        view_df.to_excel(writer, index=False, sheet_name="filtered_performance")
    st.download_button(
        label="📊 Download Filtered Data as Excel (.xlsx)",
        data=excel_buffer.getvalue(),
        file_name="marketing_campaign_performance_filtered.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

st.info("🔒 **Data Governance Guarantee:** Exports contain only cleaned, deduplicated canonical records. Raw uncleaned source records are strictly quarantined.")
