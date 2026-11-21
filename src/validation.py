"""
Validation and Quality Checking Engine
Validates schema compliance, numerical boundaries, and relational integrity.
"""

import pandas as pd
from typing import Dict, Any, List

def validate_fact_dataframe(df: pd.DataFrame) -> Dict[str, Any]:
    """Validate core fact dataframe against required schema and constraints."""
    required_cols = [
        "fact_id", "report_date", "campaign_id", "channel_id",
        "segment_id", "creative_id", "impressions", "clicks",
        "conversions", "spend_inr", "revenue_inr"
    ]
    missing_cols = [c for c in required_cols if c not in df.columns]
    if missing_cols:
        return {"valid": False, "error": f"Missing required columns: {missing_cols}"}

    # Null check
    null_counts = df[required_cols].isna().sum().to_dict()
    total_nulls = sum(null_counts.values())

    # Range and logic checks
    negative_impressions = int((df["impressions"] < 0).sum())
    negative_clicks = int((df["clicks"] < 0).sum())
    negative_conversions = int((df["conversions"] < 0).sum())
    negative_spend = int((df["spend_inr"] < 0).sum())
    negative_revenue = int((df["revenue_inr"] < 0).sum())

    clicks_exceed_impressions = int((df["clicks"] > df["impressions"]).sum())
    conversions_exceed_clicks = int((df["conversions"] > df["clicks"]).sum())

    is_valid = (
        total_nulls == 0 and
        negative_impressions == 0 and
        negative_clicks == 0 and
        negative_conversions == 0 and
        negative_spend == 0 and
        negative_revenue == 0 and
        clicks_exceed_impressions == 0 and
        conversions_exceed_clicks == 0
    )

    return {
        "valid": is_valid,
        "row_count": len(df),
        "total_nulls": total_nulls,
        "null_counts": null_counts,
        "anomalies": {
            "clicks_exceed_impressions": clicks_exceed_impressions,
            "conversions_exceed_clicks": conversions_exceed_clicks,
            "negative_spend": negative_spend,
            "negative_revenue": negative_revenue
        }
    }
