"""
Data Ingestion, Caching, and Multi-Criteria Filtering Engine
Supports both Streamlit runtime caching and offline CLI/test access.
"""

import pandas as pd
from typing import List, Optional, Tuple
from src.db import get_analysis_view, get_table_data

# Optional Streamlit cache decorator
try:
    import streamlit as st
    cache_decorator = st.cache_data
except ImportError:
    def cache_decorator(func):
        return func

@cache_decorator
def load_full_performance_data() -> pd.DataFrame:
    """Fetch complete joined analytical view from SQLite."""
    df = get_analysis_view()
    if "report_date" in df.columns:
        df["report_date"] = pd.to_datetime(df["report_date"])
    return df

@cache_decorator
def load_campaigns_metadata() -> pd.DataFrame:
    """Fetch campaign dimension table."""
    return get_table_data("campaigns")

@cache_decorator
def load_channels_metadata() -> pd.DataFrame:
    """Fetch channels dimension table."""
    return get_table_data("channels")

@cache_decorator
def load_segments_metadata() -> pd.DataFrame:
    """Fetch audience segments dimension table."""
    return get_table_data("audience_segments")

@cache_decorator
def load_creatives_metadata() -> pd.DataFrame:
    """Fetch creatives dimension table."""
    return get_table_data("creatives")

def get_filter_choices(df: pd.DataFrame):
    """Extract distinct filter choices from dataframe."""
    campaigns = sorted(df["campaign_name"].dropna().unique().tolist())
    channels = sorted(df["channel_name"].dropna().unique().tolist())
    regions = sorted(df["region"].dropna().unique().tolist())
    formats = sorted(df["creative_format"].dropna().unique().tolist())
    min_date = df["report_date"].min().date() if len(df) > 0 else None
    max_date = df["report_date"].max().date() if len(df) > 0 else None

    return {
        "campaigns": campaigns,
        "channels": channels,
        "regions": regions,
        "formats": formats,
        "min_date": min_date,
        "max_date": max_date
    }

def apply_filters(
    df: pd.DataFrame,
    selected_campaigns: Optional[List[str]] = None,
    selected_channels: Optional[List[str]] = None,
    selected_regions: Optional[List[str]] = None,
    selected_formats: Optional[List[str]] = None,
    date_range: Optional[Tuple] = None
) -> pd.DataFrame:
    """Apply multi-dimensional filtering to performance DataFrame."""
    if df is None or len(df) == 0:
        return df

    filtered = df.copy()

    if selected_campaigns:
        filtered = filtered[filtered["campaign_name"].isin(selected_campaigns)]

    if selected_channels:
        filtered = filtered[filtered["channel_name"].isin(selected_channels)]

    if selected_regions:
        filtered = filtered[filtered["region"].isin(selected_regions)]

    if selected_formats:
        filtered = filtered[filtered["creative_format"].isin(selected_formats)]

    if date_range and len(date_range) == 2 and date_range[0] is not None and date_range[1] is not None:
        start_dt = pd.to_datetime(date_range[0])
        end_dt = pd.to_datetime(date_range[1])
        filtered = filtered[(filtered["report_date"] >= start_dt) & (filtered["report_date"] <= end_dt)]

    return filtered
