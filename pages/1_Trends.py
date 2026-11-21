"""
Page 1: Campaign and Channel Longitudinal Trends (pages/1_Trends.py)
Supports OQ-02, BR-02, US-02, FEAT-02, UI-02
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

from src.data_loader import (
    load_full_performance_data,
    get_filter_choices,
    apply_filters
)
from src.analysis import aggregate_by_date
from src.kpis import fmt_curr, fmt_pct, fmt_ratio, fmt_num

st.set_page_config(page_title="Trends Explorer | AuraNova", page_icon="📈", layout="wide")

st.title("📈 Campaign & Channel Performance Trends")
st.caption("Longitudinal time-series analysis of marketing spend, audience engagement, and revenue velocity.")

df_raw = load_full_performance_data()
choices = get_filter_choices(df_raw)

# Sidebar filters inherited
st.sidebar.markdown("### 🎯 Active Filters")
selected_camps = st.sidebar.multiselect("Campaigns", options=choices["campaigns"], default=st.session_state.get("filter_campaigns", []))
selected_chans = st.sidebar.multiselect("Channels", options=choices["channels"], default=st.session_state.get("filter_channels", []))
selected_regs = st.sidebar.multiselect("Regions", options=choices["regions"], default=st.session_state.get("filter_regions", []))

df = apply_filters(df_raw, selected_campaigns=selected_camps, selected_channels=selected_chans, selected_regions=selected_regs)

if len(df) == 0:
    st.warning("No data found matching active filters. Please adjust your selections.")
    st.stop()

# Controls
ctrl_col1, ctrl_col2, ctrl_col3 = st.columns([2, 2, 2])
with ctrl_col1:
    freq_choice = st.radio("Temporal Aggregation", options=["Daily", "Monthly"], index=1, horizontal=True)
with ctrl_col2:
    metric_choice = st.selectbox(
        "Primary Metric to Track",
        options=["Revenue (₹)", "Spend (₹)", "Conversions", "ROAS", "Clicks", "Impressions", "CTR (%)"]
    )
with ctrl_col3:
    group_choice = st.selectbox(
        "Segment Trajectory By",
        options=["Overall Portfolio", "By Channel", "By Campaign"]
    )

metric_mapping = {
    "Revenue (₹)": "revenue_inr",
    "Spend (₹)": "spend_inr",
    "Conversions": "conversions",
    "ROAS": "roas",
    "Clicks": "clicks",
    "Impressions": "impressions",
    "CTR (%)": "ctr_pct"
}
target_metric = metric_mapping[metric_choice]

# Trend Chart Generation
if group_choice == "Overall Portfolio":
    trend_df = aggregate_by_date(df, freq="M" if freq_choice == "Monthly" else "D")
    fig = px.line(
        trend_df,
        x="period",
        y=target_metric,
        markers=True,
        title=f"Overall {metric_choice} Trend ({freq_choice})",
        labels={"period": "Timeline", target_metric: metric_choice}
    )
    fig.update_traces(line=dict(color="#2563EB", width=3), marker=dict(size=6))
else:
    group_col = "channel_name" if group_choice == "By Channel" else "campaign_name"
    temp_df = df.copy()
    if freq_choice == "Monthly":
        temp_df["period"] = temp_df["report_date"].dt.strftime("%Y-%m")
    else:
        temp_df["period"] = temp_df["report_date"].dt.strftime("%Y-%m-%d")

    grouped_trend = temp_df.groupby(["period", group_col], as_index=False).agg({
        "impressions": "sum",
        "clicks": "sum",
        "conversions": "sum",
        "spend_inr": "sum",
        "revenue_inr": "sum"
    })
    grouped_trend["ctr_pct"] = (grouped_trend["clicks"] / grouped_trend["impressions"] * 100.0).round(2)
    grouped_trend["roas"] = (grouped_trend["revenue_inr"] / grouped_trend["spend_inr"]).round(2)

    fig = px.line(
        grouped_trend,
        x="period",
        y=target_metric,
        color=group_col,
        markers=True,
        title=f"{metric_choice} Trend by {group_choice.replace('By ', '')} ({freq_choice})",
        labels={"period": "Timeline", target_metric: metric_choice}
    )

fig.update_layout(template="plotly_white", height=450, legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
st.plotly_chart(fig, use_container_width=True)

# Cumulative Trajectory
st.subheader("📊 Cumulative Spend vs. Cumulative Revenue")
cum_df = df.sort_values("report_date").copy()
cum_df["cum_spend"] = cum_df["spend_inr"].cumsum()
cum_df["cum_revenue"] = cum_df["revenue_inr"].cumsum()

fig_cum = go.Figure()
fig_cum.add_trace(go.Scatter(
    x=cum_df["report_date"],
    y=cum_df["cum_spend"],
    name="Cumulative Spend (₹)",
    line=dict(color="#EF4444", width=2.5, dash="dash")
))
fig_cum.add_trace(go.Scatter(
    x=cum_df["report_date"],
    y=cum_df["cum_revenue"],
    name="Cumulative Revenue (₹)",
    line=dict(color="#10B981", width=3)
))
fig_cum.update_layout(
    title="Cumulative Capital Inflow vs. Outflow Over Campaign Lifecycle",
    xaxis_title="Date",
    yaxis_title="Cumulative INR (₹)",
    template="plotly_white",
    height=400,
    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
)
st.plotly_chart(fig_cum, use_container_width=True)
