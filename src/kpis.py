"""
Centralised KPI Calculation Engine and Standard Formatters
Strictly enforces ratio-of-totals to avoid distortion.
Provides robust division-by-zero protection.
"""

import pandas as pd
from typing import Dict, Any, Optional

def calculate_portfolio_kpis(df: pd.DataFrame) -> Dict[str, Any]:
    """Calculate aggregate portfolio KPIs from filtered DataFrame."""
    if df is None or len(df) == 0:
        return {
            "impressions": 0,
            "clicks": 0,
            "conversions": 0,
            "spend": 0.0,
            "revenue": 0.0,
            "reach": 0,
            "ctr": None,
            "cpc": None,
            "cpm": None,
            "cvr": None,
            "cpa": None,
            "roas": None,
            "roi": None,
            "net_profit": 0.0
        }

    impressions = int(df["impressions"].sum())
    clicks = int(df["clicks"].sum())
    conversions = int(df["conversions"].sum())
    spend = float(df["spend_inr"].sum())
    revenue = float(df["revenue_inr"].sum())
    reach = int(df["reach"].sum()) if "reach" in df.columns else 0

    # Ratios of totals with zero-division guards
    ctr = (clicks / impressions * 100.0) if impressions > 0 else None
    cpc = (spend / clicks) if clicks > 0 else None
    cpm = (spend / impressions * 1000.0) if impressions > 0 else None
    cvr = (conversions / clicks * 100.0) if clicks > 0 else None
    cpa = (spend / conversions) if conversions > 0 else None
    roas = (revenue / spend) if spend > 0 else None
    roi = ((revenue - spend) / spend * 100.0) if spend > 0 else None
    net_profit = revenue - spend

    return {
        "impressions": impressions,
        "clicks": clicks,
        "conversions": conversions,
        "spend": spend,
        "revenue": revenue,
        "reach": reach,
        "ctr": ctr,
        "cpc": cpc,
        "cpm": cpm,
        "cvr": cvr,
        "cpa": cpa,
        "roas": roas,
        "roi": roi,
        "net_profit": net_profit
    }

def fmt_curr(val: Optional[float], decimals: int = 2) -> str:
    """Format currency in Indian Rupee format or return '—'."""
    if val is None or pd.isna(val):
        return "—"
    return f"₹{val:,.{decimals}f}"

def fmt_pct(val: Optional[float], decimals: int = 2) -> str:
    """Format percentage or return '—'."""
    if val is None or pd.isna(val):
        return "—"
    return f"{val:.{decimals}f}%"

def fmt_ratio(val: Optional[float], decimals: int = 2) -> str:
    """Format ratio multiple or return '—'."""
    if val is None or pd.isna(val):
        return "—"
    return f"{val:.{decimals}f}x"

def fmt_num(val: Optional[int]) -> str:
    """Format integer with thousands separator or return '—'."""
    if val is None or pd.isna(val):
        return "—"
    return f"{int(val):,}"
