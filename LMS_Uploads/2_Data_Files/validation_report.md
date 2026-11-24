# Data Integrity, Cleaning & Validation Audit Report

**Project:** Marketing Campaign Performance Analysis  
**Student:** Madhav Dutta (Roll No: 2503201450085)  
**Institution:** ABES Engineering College, Ghaziabad  
**Mentor:** Prof. Kalpana Dwivedi  
**Organisation:** AuraNova Digital Solutions Pvt. Ltd., Noida, UP (Synthetic Scenario)  
**Date:** 2026-11-21 10:52:57  
**Random Seed:** 42

---

## 1. Dataset Dimensions & Summary

| Stage | Total Rows | Columns | Integrity Status |
| :--- | :--- | :--- | :--- |
| **Raw Synthetic Dataset** | 263 | 12 | Contained deliberate anomalies & nulls |
| **Cleaned Canonical Dataset** | 260 | 12 | Passed schema & integrity constraints |
| **SQLite Database** (`data/app.db`) | 260 | 5 Normalised Tables + 1 Analysis View | PRAGMA foreign_keys = ON Verified |

---

## 2. Injected Anomalies & Cleaning Treatment Log

### A. Duplicate Rows
- **Injected in Raw:** 3 exact duplicate rows across composite grain `(report_date, campaign_id, channel_id, segment_id, creative_id)`.
- **Treatment:** Deduplicated using `.drop_duplicates(subset=grain_cols, keep='first')`.
- **Rows Removed:** 3.
- **Post-Cleaning Duplicates:** 0.

### B. Missing Values (Null Profile)
- **Injected in Raw:**
  - Impressions: 6 nulls (2.28%)
  - Clicks: 5 nulls (1.90%)
  - Spend (INR): 4 nulls (1.52%)
  - Attributed Revenue (INR): 4 nulls (1.52%)
- **Treatment:** Grouped median imputation by `channel_id` (for channel-driven metrics like impressions, clicks, spend) and by `campaign_id` (for campaign-driven revenue). Grouped imputation preserves natural channel cost structures without shrinking standard error artificially.
- **Post-Cleaning Nulls:** 0 across all fields.

### C. Outliers & Tracking Glitches
- **Row 22:** Sensor glitch caused a 10x artificial impression spike (>200,000 impressions). Corrected by scaling back to baseline distribution.
- **Row 75:** Billing integration glitch generated an 8.5x errant spend spike (>₹100,000). Corrected via baseline spend capping.
- **Row 140:** Click-bot spike created an impossible CTR (>35%). Corrected by adjusting clicks relative to channel CTR bounds.
- **Valid Outliers Retained:** High-performing days during the mid-year campaign surge (July/August) were retained to reflect authentic business seasonality.

---

## 3. Database Integrity & Constraints Verification

- **PRAGMA foreign_keys = ON:** Active on connection.
- **PRAGMA foreign_key_check Result:** PASSED (0 violations)
- **Unique Constraint on Grain:** Verified on `daily_performance(report_date, campaign_id, channel_id, segment_id, creative_id)`.
- **Check Constraints:** Verified `budget_inr > 0`, `impressions >= 0`, `clicks >= 0`, `conversions >= 0`, `spend_inr >= 0`.

---

## 4. Key Discoverable Synthetic Patterns (Verification Guidance)

The synthetic dataset contains 5 realistic, discoverable business patterns:
1. **Channel Efficiency Gradient:** Google Search (`CH-01`) and LinkedIn Ads (`CH-02`) yield high CVR (6.5% - 8.2%) and strong ROAS (~2.8x - 3.2x), whereas YouTube Video Ads (`CH-05`) drives heavy brand awareness (high impressions) but low conversion efficiency (CVR ~1.8%, CPA ~₹2,200).
2. **Top-Performing Campaign:** Campaign `CMP-004` (Custom DevOps Retargeting) produces the highest ROAS (~4.5x) and lowest acquisition costs due to mid-funnel warm audience engagement.
3. **Underperforming Budget Drain:** Campaign `CMP-005` (Digital Transformation Summit Noida) burns substantial spend (~₹9.5 Lakhs budget) with a sub-par ROAS (~0.85x - 1.1x), flagging a need for budget reallocation.
4. **Regional IT Hub Superiority:** South India (`SEG-02` - Bengaluru/Hyderabad) and North India (`SEG-01` - Delhi-NCR) deliver 25% higher conversion rates than East India (`SEG-04`).
5. **Seasonal Momentum:** Spend and conversions ramp steadily from Q1 to Q3 2026, peaking in mid-year July/August.
