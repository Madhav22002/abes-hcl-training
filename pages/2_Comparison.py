"""
Page 2: Channel and Campaign Efficiency Benchmarking (pages/2_Comparison.py)
Supports OQ-03, OQ-04, OQ-05, BR-03, BR-04, BR-05, US-03, US-04, US-05, UI-03
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

from src.data_loader import (
    load_full_performance_data,
    load_campaigns_metadata,
    get_filter_choices,
    apply_filters
)
from src.analysis import aggregate_by_channel, aggregate_by_campaign, calculate_budget_pacing
from src.kpis import fmt_curr, fmt_pct, fmt_ratio, fmt_num

st.set_page_config(page_title="Comparison & Benchmarks | AuraNova", page_icon="⚖️", layout="wide")

st.title("⚖️ Channel & Campaign Efficiency Benchmarking")
st.caption("Cross-channel cost efficiency, campaign head-to-head benchmarking, and budget pacing analysis.")

df_raw = load_full_performance_data()
campaigns_meta = load_campaigns_metadata()
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

tab_chan, tab_camp, tab_budget = st.tabs(["📢 Channel Benchmarks", "🏆 Campaign Benchmarks", "💰 Budget Utilisation"])

with tab_chan:
    st.subheader("Cross-Channel Cost & Efficiency Profile")
    ch_df = aggregate_by_channel(df)

    col1, col2 = st.columns(2)
    with col1:
        # ROAS by Channel
        fig_roas = px.bar(
            ch_df,
            x="channel_name",
            y="roas",
            color="roas",
            color_continuous_scale="Greens",
            text="roas",
            title="Channel ROAS Multiple (Target: > 2.0x)",
            labels={"roas": "ROAS (x)", "channel_name": "Channel"}
        )
        fig_roas.add_hline(y=2.0, line_dash="dash", line_color="red", annotation_text="Benchmark (2.0x)")
        fig_roas.update_traces(texttemplate='%{text:.2f}x', textposition='outside')
        fig_roas.update_layout(template="plotly_white", height=380)
        st.plotly_chart(fig_roas, use_container_width=True)

    with col2:
        # CPA vs CPC
        fig_cpa = px.scatter(
            ch_df,
            x="cpc_inr",
            y="cpa_inr",
            size="conversions",
            color="channel_name",
            text="channel_name",
            title="Channel Cost Efficiency: CPC vs. CPA (Bubble Size = Conversions)",
            labels={"cpc_inr": "Cost Per Click (₹)", "cpa_inr": "Cost Per Acquisition (₹)"}
        )
        fig_cpa.update_traces(textposition='top center')
        fig_cpa.update_layout(template="plotly_white", height=380)
        st.plotly_chart(fig_cpa, use_container_width=True)

    # Detailed Table
    st.markdown("##### Detailed Channel Performance Matrix")
    display_ch = ch_df[[
        "channel_name", "channel_type", "cost_model", "impressions",
        "clicks", "conversions", "spend_inr", "revenue_inr", "ctr_pct", "cpc_inr", "cpa_inr", "roas"
    ]].copy()
    display_ch["spend_inr"] = display_ch["spend_inr"].apply(fmt_curr)
    display_ch["revenue_inr"] = display_ch["revenue_inr"].apply(fmt_curr)
    display_ch["ctr_pct"] = display_ch["ctr_pct"].apply(fmt_pct)
    display_ch["cpc_inr"] = display_ch["cpc_inr"].apply(fmt_curr)
    display_ch["cpa_inr"] = display_ch["cpa_inr"].apply(fmt_curr)
    display_ch["roas"] = display_ch["roas"].apply(fmt_ratio)
    st.dataframe(display_ch, use_container_width=True, hide_index=True)

with tab_camp:
    st.subheader("Campaign Comparison Matrix")
    cmp_df = aggregate_by_campaign(df)

    fig_cmp = px.bar(
        cmp_df,
        x="campaign_name",
        y=["spend_inr", "revenue_inr"],
        barmode="group",
        title="Campaign Spend vs. Attributed Revenue",
        labels={"value": "Amount in INR (₹)", "campaign_name": "Campaign", "variable": "Metric"},
        color_discrete_map={"spend_inr": "#3B82F6", "revenue_inr": "#10B981"}
    )
    fig_cmp.update_layout(template="plotly_white", height=400)
    st.plotly_chart(fig_cmp, use_container_width=True)

    # Efficiency scatter
    fig_eff = px.scatter(
        cmp_df,
        x="spend_inr",
        y="roas",
        size="conversions",
        color="campaign_name",
        text="campaign_name",
        title="Spend vs. ROAS Positioning",
        labels={"spend_inr": "Actual Spend (₹)", "roas": "ROAS Multiple"}
    )
    fig_eff.add_hline(y=1.0, line_dash="dash", line_color="orange", annotation_text="Break-even ROAS (1.0x)")
    fig_eff.update_traces(textposition='top center')
    fig_eff.update_layout(template="plotly_white", height=380)
    st.plotly_chart(fig_eff, use_container_width=True)

with tab_budget:
    st.subheader("Campaign Budget Utilisation & Variance Analysis")
    pacing_df = calculate_budget_pacing(df, campaigns_meta)

    # Gauge / Bar chart of utilisation
    fig_pace = px.bar(
        pacing_df,
        x="campaign_name",
        y="utilisation_pct",
        color="pacing_status",
        text="utilisation_pct",
        title="Campaign Budget Utilisation Rate (%)",
        labels={"utilisation_pct": "Budget Utilised (%)", "campaign_name": "Campaign"},
        color_discrete_map={"Exceeded": "#EF4444", "On Track": "#10B981", "Underspent": "#F59E0B"}
    )
    fig_pace.add_hline(y=100.0, line_dash="dash", line_color="red", annotation_text="100% Budget Cap")
    fig_pace.update_traces(texttemplate='%{text:.1f}%', textposition='outside')
    fig_pace.update_layout(template="plotly_white", height=380)
    st.plotly_chart(fig_pace, use_container_width=True)

    # Pacing Table
    disp_pace = pacing_df[[
        "campaign_name", "objective", "budget_inr", "actual_spend_inr",
        "budget_variance_inr", "utilisation_pct", "pacing_status"
    ]].copy()
    disp_pace["budget_inr"] = disp_pace["budget_inr"].apply(lambda v: fmt_curr(v, decimals=0))
    disp_pace["actual_spend_inr"] = disp_pace["actual_spend_inr"].apply(fmt_curr)
    disp_pace["budget_variance_inr"] = disp_pace["budget_variance_inr"].apply(fmt_curr)
    disp_pace["utilisation_pct"] = disp_pace["utilisation_pct"].apply(fmt_pct)
    st.dataframe(disp_pace, use_container_width=True, hide_index=True)
