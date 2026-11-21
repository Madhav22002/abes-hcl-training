"""
Technical Test Suite: KPI Formulas, Zero-Division, and Statistical Logic (TC-T-04 to TC-T-05, TC-T-07 to TC-T-09)
"""

import pytest
import pandas as pd
import numpy as np

from src.kpis import calculate_portfolio_kpis, fmt_curr, fmt_pct, fmt_ratio
from src.analysis import (
    aggregate_by_channel,
    aggregate_by_campaign,
    aggregate_by_date,
    calculate_budget_pacing
)
from src.statistics import (
    compute_correlation_matrix,
    compute_spend_revenue_regression,
    perform_channel_roas_anova
)

def test_tc_t_04_kpi_mathematical_accuracy():
    """TC-T-04: Verify KPI calculations against hand-computed known test dataset."""
    known_data = pd.DataFrame([
        {"impressions": 1000, "clicks": 50, "conversions": 5, "spend_inr": 200.0, "revenue_inr": 600.0, "reach": 800},
        {"impressions": 2000, "clicks": 50, "conversions": 5, "spend_inr": 300.0, "revenue_inr": 900.0, "reach": 1600}
    ])
    # Totals: Impressions=3000, Clicks=100, Conversions=10, Spend=500.0, Revenue=1500.0
    # Expected:
    # CTR = 100 / 3000 * 100 = 3.3333%
    # CPC = 500 / 100 = 5.0
    # CPM = 500 / 3000 * 1000 = 166.6667
    # CVR = 10 / 100 * 100 = 10.0%
    # CPA = 500 / 10 = 50.0
    # ROAS = 1500 / 500 = 3.0
    # ROI = (1500 - 500) / 500 * 100 = 200.0%
    # Net Profit = 1000.0
    kpis = calculate_portfolio_kpis(known_data)

    assert kpis["impressions"] == 3000
    assert kpis["clicks"] == 100
    assert kpis["conversions"] == 10
    assert kpis["spend"] == 500.0
    assert kpis["revenue"] == 1500.0
    assert pytest.approx(kpis["ctr"], 0.001) == 3.3333
    assert pytest.approx(kpis["cpc"], 0.001) == 5.0
    assert pytest.approx(kpis["cpm"], 0.001) == 166.6667
    assert pytest.approx(kpis["cvr"], 0.001) == 10.0
    assert pytest.approx(kpis["cpa"], 0.001) == 50.0
    assert pytest.approx(kpis["roas"], 0.001) == 3.0
    assert pytest.approx(kpis["roi"], 0.001) == 200.0
    assert pytest.approx(kpis["net_profit"], 0.001) == 1000.0

def test_tc_t_05_kpi_zero_division_guard():
    """TC-T-05: Verify zero division safety returns None / '—' without exceptions."""
    zero_data = pd.DataFrame([
        {"impressions": 0, "clicks": 0, "conversions": 0, "spend_inr": 0.0, "revenue_inr": 0.0, "reach": 0}
    ])
    kpis = calculate_portfolio_kpis(zero_data)
    assert kpis["ctr"] is None
    assert kpis["cpc"] is None
    assert kpis["cpm"] is None
    assert kpis["cvr"] is None
    assert kpis["cpa"] is None
    assert kpis["roas"] is None
    assert kpis["roi"] is None

    # Formatter safety
    assert fmt_curr(kpis["spend"]) == "₹0.00"
    assert fmt_ratio(kpis["roas"]) == "—"
    assert fmt_pct(kpis["ctr"]) == "—"

def test_tc_t_07_budget_pacing():
    """TC-T-07: Verify budget pacing and variance calculations."""
    perf_data = pd.DataFrame([
        {"campaign_id": "CMP-001", "spend_inr": 400000.0},
        {"campaign_id": "CMP-002", "spend_inr": 700000.0}
    ])
    meta_data = pd.DataFrame([
        {"campaign_id": "CMP-001", "campaign_name": "Camp 1", "budget_inr": 500000.0, "objective": "Lead Gen"},
        {"campaign_id": "CMP-002", "campaign_name": "Camp 2", "budget_inr": 600000.0, "objective": "Trials"}
    ])
    pacing = calculate_budget_pacing(perf_data, meta_data)
    row1 = pacing[pacing["campaign_id"] == "CMP-001"].iloc[0]
    row2 = pacing[pacing["campaign_id"] == "CMP-002"].iloc[0]

    assert row1["actual_spend_inr"] == 400000.0
    assert row1["budget_variance_inr"] == 100000.0
    assert row1["utilisation_pct"] == 80.0
    assert row1["pacing_status"] == "On Track"

    assert row2["actual_spend_inr"] == 700000.0
    assert row2["budget_variance_inr"] == -100000.0
    assert pytest.approx(row2["utilisation_pct"], 0.1) == 116.7
    assert row2["pacing_status"] == "Exceeded"

def test_tc_t_08_statistics_ols_regression():
    """TC-T-08: Verify linear regression function accuracy and boundary checks."""
    test_df = pd.DataFrame({
        "spend_inr": [100.0, 200.0, 300.0, 400.0, 500.0],
        "revenue_inr": [250.0, 450.0, 650.0, 850.0, 1050.0]
    })
    # Perfect linear: y = 50 + 2*x
    res = compute_spend_revenue_regression(test_df)
    assert res["valid"] is True
    assert pytest.approx(res["slope"], 0.01) == 2.0
    assert pytest.approx(res["intercept"], 0.01) == 50.0
    assert pytest.approx(res["r_squared"], 0.001) == 1.0
    assert res["p_value"] < 0.001

def test_tc_t_09_anova_test():
    """TC-T-09: Verify ANOVA test computation on channel ROAS."""
    test_df = pd.DataFrame({
        "channel_name": ["Google", "Google", "Google", "Meta", "Meta", "Meta"],
        "roas": [3.2, 3.4, 3.1, 1.2, 1.3, 1.1]
    })
    res = perform_channel_roas_anova(test_df)
    assert res["valid"] is True
    assert res["n_channels"] == 2
    assert res["f_statistic"] > 50.0
    assert res["p_value"] < 0.01
    assert res["significant"] is True
