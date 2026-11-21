"""
Marketing Campaign Performance Analysis — Executive KPI Dashboard (app.py)
Project: Marketing Campaign Performance Analysis
Student: Madhav Dutta (Roll No: 2503201450085) | Institution: ABES Engineering College
Mentor: Kalpana Dwivedi | Organisation: AuraNova Digital Solutions Pvt. Ltd.
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from pathlib import Path

# Page config
st.set_page_config(
    page_title="Marketing Campaign Analytics | AuraNova",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

from src.data_loader import (
    load_full_performance_data,
    load_campaigns_metadata,
    get_filter_choices,
    apply_filters
)
from src.kpis import calculate_portfolio_kpis, fmt_curr, fmt_pct, fmt_ratio, fmt_num
from src.analysis import aggregate_by_channel, aggregate_by_campaign, aggregate_by_date
from src.ui import render_app_header, render_kpi_cards

# Initialize session state for filters if not set
if "filter_campaigns" not in st.session_state:
    st.session_state.filter_campaigns = []
if "filter_channels" not in st.session_state:
    st.session_state.filter_channels = []
if "filter_regions" not in st.session_state:
    st.session_state.filter_regions = []

def reset_filters():
    st.session_state.filter_campaigns = []
    st.session_state.filter_channels = []
    st.session_state.filter_regions = []

# Load data safely
try:
    df_raw = load_full_performance_data()
    campaigns_meta = load_campaigns_metadata()
except Exception as e:
    st.error(f"Error loading database: {e}")
    st.info("Please verify that data/app.db exists by running: python src/generate_data.py")
    st.stop()

# Extract filter choices
choices = get_filter_choices(df_raw)

# Sidebar
st.sidebar.markdown("### 🎯 Global Filters")
st.sidebar.caption("Filters apply consistently across all pages via shared session state.")

selected_camps = st.sidebar.multiselect(
    "Select Campaigns",
    options=choices["campaigns"],
    default=st.session_state.filter_campaigns,
    key="sb_camps"
)
st.session_state.filter_campaigns = selected_camps

selected_chans = st.sidebar.multiselect(
    "Select Channels",
    options=choices["channels"],
    default=st.session_state.filter_channels,
    key="sb_chans"
)
st.session_state.filter_channels = selected_chans

selected_regs = st.sidebar.multiselect(
    "Select Regions",
    options=choices["regions"],
    default=st.session_state.filter_regions,
    key="sb_regs"
)
st.session_state.filter_regions = selected_regs

# Date range filter
date_range = st.sidebar.date_input(
    "Date Range",
    value=(choices["min_date"], choices["max_date"]),
    min_value=choices["min_date"],
    max_value=choices["max_date"]
)

if st.sidebar.button("🔄 Reset All Filters", use_container_width=True):
    reset_filters()
    st.rerun()

st.sidebar.divider()
st.sidebar.markdown(
    """
    **Project Metadata:**
    - **Student:** Madhav Dutta
    - **Roll:** `2503201450085`
    - **Course:** MCA 2nd Year (Data Analytics)
    - **College:** ABES Engineering College
    - **Mentor:** Prof. Kalpana Dwivedi
    """
)

# Apply active filters
df = apply_filters(
    df_raw,
    selected_campaigns=selected_camps,
    selected_channels=selected_chans,
    selected_regions=selected_regs,
    date_range=date_range if isinstance(date_range, tuple) and len(date_range) == 2 else None
)

# Render Tailwind Header
render_app_header(
    title="Marketing Campaign Performance Dashboard",
    subtitle="Enterprise multi-channel marketing intelligence, budget pacing, and ROI analytics for AuraNova Digital Solutions.",
    badge_text="ACADEMIC DEMO • 100% SYNTHETIC"
)

# Empty state handling
if len(df) == 0:
    st.warning("⚠️ No campaign performance records found for the selected filter combination.")
    st.info("Try selecting different filters or click 'Reset All Filters' in the sidebar.")
    st.stop()

# Compute KPIs
kpis = calculate_portfolio_kpis(df)

# Render KPI Cards
kpi_card_data = [
    {"label": "Total Spend", "value": fmt_curr(kpis["spend"]), "subtext": "Sanctioned ad spend", "color": "blue"},
    {"label": "Attributed Revenue", "value": fmt_curr(kpis["revenue"]), "subtext": f"Net Profit: {fmt_curr(kpis['net_profit'])}", "color": "green"},
    {"label": "Return on Ad Spend (ROAS)", "value": fmt_ratio(kpis["roas"]), "subtext": f"Portfolio ROI: {fmt_pct(kpis['roi'])}", "color": "emerald"},
    {"label": "Total Conversions", "value": fmt_num(kpis["conversions"]), "subtext": f"CVR: {fmt_pct(kpis['cvr'])}", "color": "purple"},
    {"label": "Impressions", "value": fmt_num(kpis["impressions"]), "subtext": f"Reach: {fmt_num(kpis['reach'])}", "color": "indigo"},
    {"label": "Clicks", "value": fmt_num(kpis["clicks"]), "subtext": f"CTR: {fmt_pct(kpis['ctr'])}", "color": "blue"},
    {"label": "Cost Per Click (CPC)", "value": fmt_curr(kpis["cpc"]), "subtext": f"CPM: {fmt_curr(kpis['cpm'])}", "color": "amber"},
    {"label": "Cost Per Acquisition (CPA)", "value": fmt_curr(kpis["cpa"]), "subtext": "Per qualified conversion", "color": "rose"}
]
render_kpi_cards(kpi_card_data)

st.write("") # Spacing

# Row 1: Charts
col_left, col_right = st.columns([3, 2])

with col_left:
    st.subheader("📈 Monthly Performance Trajectory")
    monthly_df = aggregate_by_date(df, freq="M")
    
    fig_trend = go.Figure()
    fig_trend.add_trace(go.Bar(
        x=monthly_df["period"],
        y=monthly_df["spend_inr"],
        name="Spend (₹)",
        marker_color="#3B82F6",
        opacity=0.85
    ))
    fig_trend.add_trace(go.Scatter(
        x=monthly_df["period"],
        y=monthly_df["revenue_inr"],
        name="Attributed Revenue (₹)",
        mode="lines+markers",
        line=dict(color="#10B981", width=3),
        marker=dict(size=7)
    ))
    fig_trend.update_layout(
        title="Monthly Spend vs. Attributed Revenue",
        xaxis_title="Month",
        yaxis_title="Amount in INR (₹)",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        template="plotly_white",
        height=380,
        margin=dict(l=40, r=20, t=50, b=40)
    )
    st.plotly_chart(fig_trend, use_container_width=True)

with col_right:
    st.subheader("🍩 Channel Spend & ROAS Distribution")
    ch_df = aggregate_by_channel(df)
    fig_ch = px.pie(
        ch_df,
        values="spend_inr",
        names="channel_name",
        color="channel_name",
        color_discrete_sequence=px.colors.qualitative.Safe,
        hole=0.45,
        title="Ad Spend Share by Channel"
    )
    fig_ch.update_traces(textposition='inside', textinfo='percent+label')
    fig_ch.update_layout(
        template="plotly_white",
        height=380,
        margin=dict(l=20, r=20, t=50, b=40),
        showlegend=False
    )
    st.plotly_chart(fig_ch, use_container_width=True)

# Row 2: Campaign Breakdown Table
st.subheader("📋 Campaign Head-to-Head Scorecard")
cmp_df = aggregate_by_campaign(df)

display_cmp = cmp_df.copy()
display_cmp["spend_inr"] = display_cmp["spend_inr"].apply(lambda v: fmt_curr(v))
display_cmp["revenue_inr"] = display_cmp["revenue_inr"].apply(lambda v: fmt_curr(v))
display_cmp["campaign_budget_inr"] = display_cmp["campaign_budget_inr"].apply(lambda v: fmt_curr(v, decimals=0))
display_cmp["ctr_pct"] = display_cmp["ctr_pct"].apply(lambda v: fmt_pct(v))
display_cmp["cpc_inr"] = display_cmp["cpc_inr"].apply(lambda v: fmt_curr(v))
display_cmp["cpa_inr"] = display_cmp["cpa_inr"].apply(lambda v: fmt_curr(v))
display_cmp["roas"] = display_cmp["roas"].apply(lambda v: fmt_ratio(v))
display_cmp["roi_pct"] = display_cmp["roi_pct"].apply(lambda v: fmt_pct(v))

display_cmp.rename(columns={
    "campaign_name": "Campaign Name",
    "campaign_objective": "Objective",
    "campaign_budget_inr": "Budget",
    "impressions": "Impressions",
    "clicks": "Clicks",
    "conversions": "Conversions",
    "spend_inr": "Actual Spend",
    "revenue_inr": "Revenue",
    "ctr_pct": "CTR",
    "cpc_inr": "CPC",
    "cpa_inr": "CPA",
    "roas": "ROAS",
    "roi_pct": "ROI"
}, inplace=True)

drop_cols = ["campaign_id", "campaign_status", "reach", "cpm_inr", "cvr_pct"]
display_cmp = display_cmp.drop(columns=[c for c in drop_cols if c in display_cmp.columns])

st.dataframe(display_cmp, use_container_width=True, hide_index=True)

# Footer notice
st.caption(
    "⚠️ **Academic Disclaimer:** All campaign metrics, channel data, and company entities represent synthetic records generated "
    "for academic evaluation (Seed = 42). Division-by-zero occurrences evaluate to '—'."
)
