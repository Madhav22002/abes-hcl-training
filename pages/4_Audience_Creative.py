"""
Page 4: Audience Demographics and Creative Asset Matrix (pages/4_Audience_Creative.py)
Supports OQ-06, OQ-07, BR-06, BR-07, US-06, US-07, FEAT-06, FEAT-07, UI-05
"""

import streamlit as st
import pandas as pd
import plotly.express as px

from src.data_loader import (
    load_full_performance_data,
    get_filter_choices,
    apply_filters
)
from src.analysis import aggregate_by_segment, aggregate_by_creative
from src.kpis import fmt_curr, fmt_pct, fmt_ratio, fmt_num

st.set_page_config(page_title="Audience & Creative | AuraNova", page_icon="👥", layout="wide")

st.title("👥 Audience Demographics & Creative Asset Matrix")
st.caption("Deep-dive segmentation of geographic cohorts, target devices, creative formats, and messaging themes.")

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

tab_aud, tab_crt = st.tabs(["🗺️ Audience & Regional Cohorts", "🎨 Creative Asset Performance"])

with tab_aud:
    st.subheader("Regional Performance & Device Dynamics")
    seg_df = aggregate_by_segment(df)

    col1, col2 = st.columns(2)
    with col1:
        # Conversion rate by Region
        fig_reg = px.bar(
            seg_df,
            x="region",
            y="cvr_pct",
            color="cvr_pct",
            color_continuous_scale="Blues",
            text="cvr_pct",
            title="Lead Conversion Rate (CVR %) by Geographic Cluster",
            labels={"cvr_pct": "Conversion Rate (%)", "region": "Region"}
        )
        fig_reg.update_traces(texttemplate='%{text:.2f}%', textposition='outside')
        fig_reg.update_layout(template="plotly_white", height=380)
        st.plotly_chart(fig_reg, use_container_width=True)

    with col2:
        # Device Breakdown
        dev_df = df.groupby("primary_device", as_index=False).agg({
            "impressions": "sum", "clicks": "sum", "conversions": "sum", "revenue_inr": "sum"
        })
        dev_df["cvr_pct"] = (dev_df["conversions"] / dev_df["clicks"] * 100.0).round(2)
        fig_dev = px.pie(
            dev_df,
            values="conversions",
            names="primary_device",
            hole=0.4,
            title="Conversions by Access Device",
            color_discrete_sequence=px.colors.qualitative.Prism
        )
        fig_dev.update_layout(template="plotly_white", height=380)
        st.plotly_chart(fig_dev, use_container_width=True)

    # Detailed Audience Table
    st.markdown("##### Regional Cohort Scorecard")
    disp_seg = seg_df[[
        "segment_name", "region", "target_age_band", "primary_device",
        "impressions", "clicks", "conversions", "spend_inr", "revenue_inr", "cvr_pct", "roas"
    ]].copy()
    disp_seg["spend_inr"] = disp_seg["spend_inr"].apply(fmt_curr)
    disp_seg["revenue_inr"] = disp_seg["revenue_inr"].apply(fmt_curr)
    disp_seg["cvr_pct"] = disp_seg["cvr_pct"].apply(fmt_pct)
    disp_seg["roas"] = disp_seg["roas"].apply(fmt_ratio)
    st.dataframe(disp_seg, use_container_width=True, hide_index=True)

with tab_crt:
    st.subheader("Creative Format & Messaging Theme Effectiveness")
    crt_df = aggregate_by_creative(df)

    col3, col4 = st.columns(2)
    with col3:
        # CTR by format
        fig_fmt = px.bar(
            crt_df,
            x="creative_format",
            y="ctr_pct",
            color="creative_format",
            text="ctr_pct",
            title="Click-Through Rate (CTR %) by Asset Format",
            labels={"ctr_pct": "CTR (%)", "creative_format": "Format"}
        )
        fig_fmt.update_traces(texttemplate='%{text:.2f}%', textposition='outside')
        fig_fmt.update_layout(template="plotly_white", height=380, showlegend=False)
        st.plotly_chart(fig_fmt, use_container_width=True)

    with col4:
        # CTR vs CPA Scatter
        fig_crt_eff = px.scatter(
            crt_df,
            x="ctr_pct",
            y="cpa_inr",
            size="conversions",
            color="messaging_theme",
            text="creative_name",
            title="Creative Engagement vs. Acquisition Cost (Bubble = Conversions)",
            labels={"ctr_pct": "CTR (%)", "cpa_inr": "CPA (₹)"}
        )
        fig_crt_eff.update_traces(textposition='top center')
        fig_crt_eff.update_layout(template="plotly_white", height=380)
        st.plotly_chart(fig_crt_eff, use_container_width=True)

    # Detailed Creative Table
    st.markdown("##### Creative Asset Diagnostics")
    disp_crt = crt_df[[
        "creative_name", "creative_format", "messaging_theme", "offer_type",
        "impressions", "clicks", "conversions", "ctr_pct", "cpc_inr", "cpa_inr", "roas"
    ]].copy()
    disp_crt["ctr_pct"] = disp_crt["ctr_pct"].apply(fmt_pct)
    disp_crt["cpc_inr"] = disp_crt["cpc_inr"].apply(fmt_curr)
    disp_crt["cpa_inr"] = disp_crt["cpa_inr"].apply(fmt_curr)
    disp_crt["roas"] = disp_crt["roas"].apply(fmt_ratio)
    st.dataframe(disp_crt, use_container_width=True, hide_index=True)
