"""
Data Cleaning and Transformation Audit Engine
Provides audit trail between raw synthetic dataset and cleaned canonical store.
"""

import pandas as pd
from pathlib import Path
from src.config import RAW_DATA_PATH, CLEAN_DATA_PATH

def get_raw_cleaning_audit():
    """Retrieve raw vs cleaned comparison statistics."""
    if not RAW_DATA_PATH.exists() or not CLEAN_DATA_PATH.exists():
        return {"status": "Datasets not yet generated"}

    raw_excel = pd.ExcelFile(RAW_DATA_PATH)
    raw_df = pd.read_excel(raw_excel, sheet_name="raw_daily_performance")
    
    clean_excel = pd.ExcelFile(CLEAN_DATA_PATH)
    clean_df = pd.read_excel(clean_excel, sheet_name="daily_performance")

    raw_nulls = int(raw_df.isna().sum().sum())
    clean_nulls = int(clean_df.isna().sum().sum())

    grain_cols = ["report_date", "campaign_id", "channel_id", "segment_id", "creative_id"]
    raw_duplicates = int(raw_df.duplicated(subset=grain_cols).sum())
    clean_duplicates = int(clean_df.duplicated(subset=grain_cols).sum())

    return {
        "raw_total_rows": len(raw_df),
        "clean_total_rows": len(clean_df),
        "rows_removed_or_deduped": len(raw_df) - len(clean_df),
        "raw_null_cells": raw_nulls,
        "clean_null_cells": clean_nulls,
        "raw_duplicate_facts": raw_duplicates,
        "clean_duplicate_facts": clean_duplicates,
        "raw_columns": list(raw_df.columns),
        "clean_columns": list(clean_df.columns)
    }
