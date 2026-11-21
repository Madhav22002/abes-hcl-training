"""
Aggregation, Dimensional Grouping, and Comparative Analysis Engine
Calculates clean ratio-of-totals metrics across channel, campaign, cohort, and creative slices.
"""

import pandas as pd
from typing import Dict, Any

def _compute_summary_ratios(grouped_df: pd.DataFrame) -> pd.DataFrame:
    """Helper to compute standard ratio metrics on grouped sums."""
    res = grouped_df.copy()
    res["ctr_pct"] = res.apply(lambda r: round((r["clicks"] / r["impressions"] * 100.0), 2) if r["impressions"] > 0 else None, axis=1)
    res["cpc_inr"] = res.apply(lambda r: round((r["spend_inr"] / r["clicks"]), 2) if r["clicks"] > 0 else None, axis=1)
    res["cpm_inr"] = res.apply(lambda r: round((r["spend_inr"] / r["impressions"] * 1000.0), 2) if r["impressions"] > 0 else None, axis=1)
    res["cvr_pct"] = res.apply(lambda r: round((r["conversions"] / r["clicks"] * 100.0), 2) if r["clicks"] > 0 else None, axis=1)
    res["cpa_inr"] = res.apply(lambda r: round((r["spend_inr"] / r["conversions"]), 2) if r["conversions"] > 0 else None, axis=1)
    res["roas"] = res.apply(lambda r: round((r["revenue_inr"] / r["spend_inr"]), 2) if r["spend_inr"] > 0 else None, axis=1)
    res["roi_pct"] = res.apply(lambda r: round(((r["revenue_inr"] - r["spend_inr"]) / r["spend_inr"] * 100.0), 2) if r["spend_inr"] > 0 else None, axis=1)
    return res

def aggregate_by_channel(df: pd.DataFrame) -> pd.DataFrame:
    """Aggregate metrics across channels."""
    if len(df) == 0:
        return pd.DataFrame()
    grouped = df.groupby(["channel_id", "channel_name", "channel_type", "cost_model"], as_index=False).agg({
        "impressions": "sum",
        "clicks": "sum",
        "conversions": "sum",
        "spend_inr": "sum",
        "revenue_inr": "sum",
        "reach": "sum"
    })
    return _compute_summary_ratios(grouped).sort_values("spend_inr", ascending=False).reset_index(drop=True)

def aggregate_by_campaign(df: pd.DataFrame) -> pd.DataFrame:
    """Aggregate metrics across campaigns."""
    if len(df) == 0:
        return pd.DataFrame()
    grouped = df.groupby(["campaign_id", "campaign_name", "campaign_objective", "campaign_budget_inr", "campaign_status"], as_index=False).agg({
        "impressions": "sum",
        "clicks": "sum",
        "conversions": "sum",
        "spend_inr": "sum",
        "revenue_inr": "sum",
        "reach": "sum"
    })
    return _compute_summary_ratios(grouped).sort_values("spend_inr", ascending=False).reset_index(drop=True)

def aggregate_by_date(df: pd.DataFrame, freq: str = "D") -> pd.DataFrame:
    """Aggregate metrics over time (Daily 'D' or Monthly 'M')."""
    if len(df) == 0:
        return pd.DataFrame()
    temp = df.copy()
    if freq == "M":
        temp["period"] = temp["report_date"].dt.strftime("%Y-%m")
    else:
        temp["period"] = temp["report_date"].dt.strftime("%Y-%m-%d")

    grouped = temp.groupby("period", as_index=False).agg({
        "impressions": "sum",
        "clicks": "sum",
        "conversions": "sum",
        "spend_inr": "sum",
        "revenue_inr": "sum",
        "reach": "sum"
    })
    res = _compute_summary_ratios(grouped)
    return res.sort_values("period", ascending=True).reset_index(drop=True)

def aggregate_by_segment(df: pd.DataFrame) -> pd.DataFrame:
    """Aggregate metrics across audience cohorts and regions."""
    if len(df) == 0:
        return pd.DataFrame()
    grouped = df.groupby(["segment_id", "segment_name", "region", "target_age_band", "primary_device"], as_index=False).agg({
        "impressions": "sum",
        "clicks": "sum",
        "conversions": "sum",
        "spend_inr": "sum",
        "revenue_inr": "sum"
    })
    return _compute_summary_ratios(grouped).sort_values("revenue_inr", ascending=False).reset_index(drop=True)

def aggregate_by_creative(df: pd.DataFrame) -> pd.DataFrame:
    """Aggregate metrics across creative formats and themes."""
    if len(df) == 0:
        return pd.DataFrame()
    grouped = df.groupby(["creative_id", "creative_name", "creative_format", "messaging_theme", "offer_type"], as_index=False).agg({
        "impressions": "sum",
        "clicks": "sum",
        "conversions": "sum",
        "spend_inr": "sum",
        "revenue_inr": "sum"
    })
    return _compute_summary_ratios(grouped).sort_values("ctr_pct", ascending=False).reset_index(drop=True)

def calculate_budget_pacing(perf_df: pd.DataFrame, campaigns_df: pd.DataFrame) -> pd.DataFrame:
    """Calculate actual spend against approved budgets and variance."""
    if len(campaigns_df) == 0:
        return pd.DataFrame()

    actual_spend = perf_df.groupby("campaign_id")["spend_inr"].sum().reset_index()
    actual_spend.rename(columns={"spend_inr": "actual_spend_inr"}, inplace=True)

    merged = pd.merge(campaigns_df, actual_spend, on="campaign_id", how="left")
    merged["actual_spend_inr"] = merged["actual_spend_inr"].fillna(0.0)
    merged["budget_variance_inr"] = merged["budget_inr"] - merged["actual_spend_inr"]
    merged["utilisation_pct"] = (merged["actual_spend_inr"] / merged["budget_inr"] * 100.0).round(2)
    merged["pacing_status"] = merged.apply(
        lambda r: "Exceeded" if r["utilisation_pct"] > 100.0 else ("On Track" if r["utilisation_pct"] >= 80.0 else "Underspent"),
        axis=1
    )
    return merged.sort_values("budget_inr", ascending=False).reset_index(drop=True)
