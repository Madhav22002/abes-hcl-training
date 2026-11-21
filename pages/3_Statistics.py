"""
Page 3: Exploratory Data Analysis & Inferential Statistics (pages/3_Statistics.py)
Supports OQ-08, BR-08, US-08, FEAT-08, UI-04
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go

from src.data_loader import (
    load_full_performance_data,
    get_filter_choices,
    apply_filters
)
from src.statistics import (
    compute_correlation_matrix,
    compute_spend_revenue_regression,
    perform_channel_roas_anova
)
from src.kpis import fmt_curr, fmt_pct, fmt_ratio

st.set_page_config(page_title="EDA & Statistics | AuraNova", page_icon="🔬", layout="wide")

st.title("🔬 Exploratory Data Analysis & Inferential Statistics")
st.caption("Rigorous statistical hypothesis testing, correlation analysis, and regression modeling with explicit diagnostic boundaries.")

df_raw = load_full_performance_data()
choices = get_filter_choices(df_raw)

# Sidebar
st.sidebar.markdown("### 🎯 Active Filters")
selected_camps = st.sidebar.multiselect("Campaigns", options=choices["campaigns"], default=st.session_state.get("filter_campaigns", []))
selected_chans = st.sidebar.multiselect("Channels", options=choices["channels"], default=st.session_state.get("filter_channels", []))

df = apply_filters(df_raw, selected_campaigns=selected_camps, selected_channels=selected_chans)

if len(df) < 5:
    st.warning("Insufficient data points for meaningful statistical analysis. Please broaden filter selections.")
    st.stop()

tab_corr, tab_reg, tab_anova = st.tabs(["📊 Correlation Matrix", "📈 Spend vs. Revenue OLS Regression", "🧪 Channel ANOVA & Kruskal-Wallis"])

with tab_corr:
    st.subheader("Pearson Correlation Heatmap")
    st.caption("Evaluates linear bivariate association across key marketing volume and financial metrics.")
    
    corr_df = compute_correlation_matrix(df)
    if len(corr_df) > 0:
        fig_corr = px.imshow(
            corr_df,
            text_auto=True,
            aspect="auto",
            color_continuous_scale="RdBu_r",
            zmin=-1,
            zmax=1,
            title="Pearson Correlation Coefficient (r) Matrix"
        )
        fig_corr.update_layout(height=450, template="plotly_white")
        st.plotly_chart(fig_corr, use_container_width=True)

        st.info(
            "📌 **Analytical Note:** A strong positive correlation (e.g. Spend vs. Revenue, r > 0.70) "
            "indicates significant co-movement, but does **not** prove that spend alone causes revenue growth without considering audience intent and media quality."
        )

with tab_reg:
    st.subheader("Ordinary Least Squares (OLS) Regression Analysis")
    st.markdown("$$\\text{Revenue} = \\beta_0 + \\beta_1 \\times \\text{Spend} + \\epsilon$$")

    reg_res = compute_spend_revenue_regression(df)
    if reg_res.get("valid"):
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Regression Slope (β₁)", f"₹{reg_res['slope']:.2f}", help="Expected marginal revenue per ₹1 spend")
        col2.metric("Coefficient of Determination (R²)", f"{reg_res['r_squared']:.1%}", help="Proportion of revenue variance explained by spend")
        col3.metric("p-value", f"{reg_res['p_value']:.4e}", help="Probability of observing relationship under null hypothesis")
        col4.metric("Standard Error", f"±₹{reg_res['std_err']:.3f}", help="Standard error of slope estimate")

        # Scatter with OLS line
        x_vals = df["spend_inr"].values
        y_vals = df["revenue_inr"].values
        x_line = np.linspace(x_vals.min(), x_vals.max(), 100)
        y_line = reg_res["intercept"] + reg_res["slope"] * x_line

        fig_reg = go.Figure()
        fig_reg.add_trace(go.Scatter(
            x=x_vals,
            y=y_vals,
            mode="markers",
            name="Daily Observations",
            marker=dict(color="#3B82F6", size=7, opacity=0.7)
        ))
        fig_reg.add_trace(go.Scatter(
            x=x_line,
            y=y_line,
            mode="lines",
            name=f"OLS Fit: y = {reg_res['intercept']:.1f} + {reg_res['slope']:.2f}x",
            line=dict(color="#EF4444", width=3)
        ))
        fig_reg.update_layout(
            title="Ad Spend vs. Attributed Revenue Linear Regression",
            xaxis_title="Spend in INR (₹)",
            yaxis_title="Revenue in INR (₹)",
            template="plotly_white",
            height=450,
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
        )
        st.plotly_chart(fig_reg, use_container_width=True)

        st.markdown(f"**Interpretation:** {reg_res['interpretation']}")
        with st.expander("⚠️ Model Assumptions, Limitations & Diagnostics"):
            for lim in reg_res["limitations"]:
                st.write(f"- {lim}")
    else:
        st.error(reg_res.get("error", "Unable to fit regression model."))

with tab_anova:
    st.subheader("Channel Efficiency Variation (ANOVA & Kruskal-Wallis)")
    st.caption("Evaluates whether Return on Ad Spend (ROAS) differs systematically across delivery channels.")

    anova_res = perform_channel_roas_anova(df)
    if anova_res.get("valid"):
        col_a1, col_a2, col_a3, col_a4 = st.columns(4)
        col_a1.metric("Channels Compared", anova_res["n_channels"])
        col_a2.metric("One-Way ANOVA F-Stat", f"{anova_res['f_statistic']:.2f}")
        col_a3.metric("ANOVA p-value", f"{anova_res['p_value']:.4e}")
        col_a4.metric("Kruskal-Wallis H", f"{anova_res['kruskal_statistic']:.2f}")

        # Box plot of channel ROAS
        fig_box = px.box(
            df,
            x="channel_name",
            y="roas",
            color="channel_name",
            points="all",
            title="ROAS Distribution Dispersion Across Marketing Channels",
            labels={"roas": "ROAS Multiple", "channel_name": "Channel"}
        )
        fig_box.update_layout(template="plotly_white", height=420, showlegend=False)
        st.plotly_chart(fig_box, use_container_width=True)

        st.success(anova_res["summary"])
        with st.expander("ℹ️ Test Assumptions & Limitations"):
            for lim in anova_res["limitations"]:
                st.write(f"- {lim}")
    else:
        st.warning(anova_res.get("error", "ANOVA test could not be performed with active filter slice."))
