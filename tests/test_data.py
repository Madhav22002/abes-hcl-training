"""
Technical Test Suite: Database Schema, Integrity, and Data Cleanliness (TC-T-01 to TC-T-03)
"""

import sqlite3
import pytest
import pandas as pd
from pathlib import Path

from src.config import DB_PATH, RAW_DATA_PATH, CLEAN_DATA_PATH
from src.db import get_connection, verify_db_integrity, get_table_data, get_analysis_view
from src.cleaning import get_raw_cleaning_audit
from src.validation import validate_fact_dataframe

def test_tc_t_01_db_integrity_and_fk_enforcement():
    """TC-T-01: Verify SQLite integrity and foreign key constraints."""
    assert DB_PATH.exists(), f"Database file does not exist at {DB_PATH}"
    
    conn = get_connection()
    cur = conn.cursor()
    
    # Check FK is ON
    cur.execute("PRAGMA foreign_keys;")
    fk_on = cur.fetchone()[0]
    assert fk_on == 1, "PRAGMA foreign_keys must be active."
    
    # Run integrity and FK checks
    integrity_res = verify_db_integrity()
    assert integrity_res["integrity"] == "ok", f"Integrity check failed: {integrity_res['integrity']}"
    assert integrity_res["fk_violations_count"] == 0, f"Found {integrity_res['fk_violations_count']} FK violations."
    assert integrity_res["status"] == "PASS"
    conn.close()

def test_tc_t_02_schema_and_master_tables():
    """TC-T-02: Verify master dimension tables and expected row counts."""
    campaigns = get_table_data("campaigns")
    assert len(campaigns) == 5, f"Expected 5 campaigns, got {len(campaigns)}"
    assert set(["campaign_id", "campaign_name", "budget_inr", "status"]).issubset(campaigns.columns)

    channels = get_table_data("channels")
    assert len(channels) == 5, f"Expected 5 channels, got {len(channels)}"
    assert set(["channel_id", "channel_name", "cost_model"]).issubset(channels.columns)

    segments = get_table_data("audience_segments")
    assert len(segments) == 5, f"Expected 5 audience segments, got {len(segments)}"

    creatives = get_table_data("creatives")
    assert len(creatives) == 5, f"Expected 5 creatives, got {len(creatives)}"

def test_tc_t_03_fact_table_grain_and_cleanliness():
    """TC-T-03: Verify fact table row count (200-300 target), uniqueness of grain, and no nulls."""
    daily_perf = get_table_data("daily_performance")
    assert 200 <= len(daily_perf) <= 300, f"Expected fact rows between 200 and 300, got {len(daily_perf)}"
    assert len(daily_perf) == 260, f"Expected 260 deterministic fact rows, got {len(daily_perf)}"

    # Grain uniqueness check
    grain_cols = ["report_date", "campaign_id", "channel_id", "segment_id", "creative_id"]
    duplicates = daily_perf.duplicated(subset=grain_cols).sum()
    assert duplicates == 0, f"Found {duplicates} duplicate records violating composite grain!"

    # Validation check
    val_res = validate_fact_dataframe(daily_perf)
    assert val_res["valid"] is True, f"Validation failed: {val_res}"
    assert val_res["total_nulls"] == 0

def test_tc_t_06_analysis_view():
    """TC-T-06: Verify joined analytical view v_campaign_performance_details."""
    view_df = get_analysis_view()
    assert len(view_df) == 260, f"Expected 260 rows in joined view, got {len(view_df)}"
    expected_cols = [
        "fact_id", "report_date", "campaign_name", "channel_name",
        "region", "creative_name", "impressions", "clicks", "conversions",
        "spend_inr", "revenue_inr", "ctr_pct", "cpc_inr", "roas", "roi_pct"
    ]
    assert set(expected_cols).issubset(view_df.columns)
