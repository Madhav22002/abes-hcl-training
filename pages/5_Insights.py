"""
Page 5: Automated Actionable Recommendations Engine (pages/5_Insights.py)
Supports OQ-09, BR-09, US-09, FEAT-09, UI-06
"""

import streamlit as st
import pandas as pd

from src.data_loader import (
    load_full_performance_data,
    load_campaigns_metadata,
    get_filter_choices,
    apply_filters
)
from src.insights import generate_dynamic_insights

st.set_page_config(page_title="Actionable Insights | AuraNova", page_icon="💡", layout="wide")

st.title("💡 Automated Actionable Recommendations")
st.caption("Evidence-backed decision support engine synthesizing active campaign performance, cost efficiency, and pacing data.")

df_raw = load_full_performance_data()
campaigns_meta = load_campaigns_metadata()
choices = get_filter_choices(df_raw)

# Sidebar
st.sidebar.markdown("### 🎯 Active Filters")
selected_camps = st.sidebar.multiselect("Campaigns", options=choices["campaigns"], default=st.session_state.get("filter_campaigns", []))
selected_chans = st.sidebar.multiselect("Channels", options=choices["channels"], default=st.session_state.get("filter_channels", []))

df = apply_filters(df_raw, selected_campaigns=selected_camps, selected_channels=selected_chans)

if len(df) == 0:
    st.warning("No data found matching active filters.")
    st.stop()

# Generate dynamic insights
insights = generate_dynamic_insights(df, campaigns_meta)

# Category filter
types = ["All Recommendations"] + sorted(list(set(i.get("type", "General") for i in insights)))
selected_type = st.selectbox("Filter Recommendations by Category", options=types)

filtered_insights = insights if selected_type == "All Recommendations" else [i for i in insights if i.get("type") == selected_type]

st.markdown(f"**Showing {len(filtered_insights)} strategic insights for the current portfolio slice:**")
st.write("")

for idx, item in enumerate(filtered_insights, start=1):
    badge = item.get("badge", "Insight")
    badge_color = item.get("badge_color", "#2563EB")
    
    with st.container():
        st.markdown(
            f"""
            <div style="background-color: #FFFFFF; border-left: 5px solid {badge_color}; border-radius: 8px; padding: 18px; margin-bottom: 16px; box-shadow: 0 1px 3px rgba(0,0,0,0.08); border-top: 1px solid #E2E8F0; border-right: 1px solid #E2E8F0; border-bottom: 1px solid #E2E8F0;">
                <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 6px;">
                    <span style="font-size: 1.1rem; font-weight: 700; color: #1E293B;">{idx}. {item.get('title', 'Recommendation')}</span>
                    <span style="background-color: {badge_color}1A; color: {badge_color}; font-weight: 600; font-size: 0.75rem; padding: 3px 10px; border-radius: 9999px; border: 1px solid {badge_color}4D;">{badge}</span>
                </div>
                <div style="font-size: 0.85rem; font-weight: 600; color: #475569; margin-bottom: 8px;">
                    📊 Key Evidence: <span style="color: #0F172A;">{item.get('metric_highlight', '—')}</span>
                </div>
                <p style="font-size: 0.9rem; color: #334155; margin-bottom: 10px; line-height: 1.5;">{item.get('evidence', '')}</p>
                <div style="background-color: #F8FAFC; border-radius: 6px; padding: 10px 12px; border-left: 3px solid #3B82F6;">
                    <span style="font-weight: 700; color: #1D4ED8; font-size: 0.85rem;">RECOMMENDED ACTION:</span>
                    <span style="color: #1E293B; font-size: 0.88rem;"> {item.get('action', '')}</span>
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

st.caption("ℹ️ Insights are generated dynamically using mathematical heuristics and threshold comparisons against current filtered medians.")
