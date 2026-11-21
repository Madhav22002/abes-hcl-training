"""
Generate Excel Workbooks for:
- Step 4: docs/04_Traceability_Matrix.xlsx
- Step 8: docs/08_Database_Design.xlsx
"""

from pathlib import Path
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

DOCS_DIR = Path(__file__).resolve().parent.parent / "docs"
DOCS_DIR.mkdir(parents=True, exist_ok=True)

# Styling helpers
HEADER_FONT = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
HEADER_FILL = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid") # Dark Blue
ALT_FILL = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid") # Light Slate
DATA_FONT = Font(name="Segoe UI", size=10)
TITLE_FONT = Font(name="Segoe UI", size=14, bold=True, color="1E3A8A")
THIN_BORDER = Border(
    left=Side(style='thin', color='CBD5E1'),
    right=Side(style='thin', color='CBD5E1'),
    top=Side(style='thin', color='CBD5E1'),
    bottom=Side(style='thin', color='CBD5E1')
)

def create_rtm_workbook():
    wb = openpyxl.Workbook()
    ws_rtm = wb.active
    ws_rtm.title = "RTM"
    ws_rtm.views.sheetView[0].showGridLines = True

    # RTM Headers
    headers = [
        "OQ", "HMW", "Req ID (BR)", "Requirement Description", "Stakeholder (SH)",
        "Priority", "Status", "Feature ID", "DB table.column", "User Story",
        "UI Screen", "Test Case (TC-F / TC-T)"
    ]
    
    ws_rtm.append(headers)
    for col_idx in range(1, len(headers) + 1):
        cell = ws_rtm.cell(row=1, column=col_idx)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # 12 RTM Rows mapping OQ -> HMW -> BR -> FEAT -> DB -> US -> UI -> TC
    rtm_data = [
        ("OQ-01", "HMW-01", "BR-01", "Executive KPI scorecard of Spend, Revenue, Conversions, ROAS, and ROI", "SH-01", "Must", "Implemented", "FEAT-01", "daily_performance.spend_inr, revenue_inr", "US-01", "UI-01", "TC-F-01, TC-T-04"),
        ("OQ-02", "HMW-02", "BR-02", "Interactive multi-metric time-series trend visualisations (daily/monthly)", "SH-02", "Must", "Implemented", "FEAT-02", "daily_performance.report_date, spend_inr", "US-02", "UI-02", "TC-F-02, TC-T-05"),
        ("OQ-03", "HMW-03", "BR-03", "Cross-channel cost-efficiency benchmarking (CPC, CPA, CTR, ROAS)", "SH-03", "Must", "Implemented", "FEAT-03", "channels.channel_name, daily_performance.*", "US-03", "UI-03", "TC-F-03, TC-T-04"),
        ("OQ-04", "HMW-04", "BR-04", "Campaign head-to-head performance comparisons with target benchmark flags", "SH-02", "Must", "Implemented", "FEAT-04", "campaigns.campaign_name, daily_performance.*", "US-04", "UI-03", "TC-F-04, TC-T-04"),
        ("OQ-05", "HMW-05", "BR-05", "Planned vs. actual campaign budget utilisation and burn tracking", "SH-04", "Must", "Implemented", "FEAT-05", "campaigns.budget_inr, daily_performance.spend_inr", "US-05", "UI-01, UI-03", "TC-F-05, TC-T-06"),
        ("OQ-06", "HMW-06", "BR-06", "Audience demographic and regional segmentation breakdown", "SH-02", "Should", "Implemented", "FEAT-06", "audience_segments.region, target_age_band", "US-06", "UI-05", "TC-F-06, TC-T-07"),
        ("OQ-07", "HMW-07", "BR-07", "Creative asset performance analysis by format and messaging theme", "SH-03", "Should", "Implemented", "FEAT-07", "creatives.creative_format, messaging_theme", "US-07", "UI-05", "TC-F-07, TC-T-07"),
        ("OQ-08", "HMW-08", "BR-08", "Statistical correlation and regression between marketing spend and revenue", "SH-01", "Must", "Implemented", "FEAT-08", "daily_performance.spend_inr, revenue_inr", "US-08", "UI-04", "TC-F-08, TC-T-08"),
        ("OQ-09", "HMW-09", "BR-09", "Automated rule-based actionable insights for budget reallocation", "SH-02", "Must", "Implemented", "FEAT-09", "v_campaign_performance_details", "US-09", "UI-06", "TC-F-09, TC-T-09"),
        ("OQ-10", "HMW-10", "BR-10", "Transparent data quality scorecard reporting nulls, duplicates, outliers", "SH-05", "Must", "Implemented", "FEAT-10", "app.db metadata, data_quality_log", "US-10", "UI-08", "TC-F-10, TC-T-01"),
        ("OQ-03", "HMW-03", "BR-11", "Granular data explorer with multi-attribute filtering and CSV/Excel downloads", "SH-03", "Should", "Implemented", "FEAT-11", "v_campaign_performance_details", "US-11", "UI-07", "TC-F-11, TC-T-10"),
        ("OQ-01", "HMW-01", "BR-12", "In-app methodology and documentation section explaining math & provenance", "SH-01", "Could", "Implemented", "FEAT-12", "src.config.py metadata", "US-12", "UI-09", "TC-F-12, TC-T-10")
    ]

    for row_idx, r in enumerate(rtm_data, start=2):
        for col_idx, val in enumerate(r, start=1):
            cell = ws_rtm.cell(row=row_idx, column=col_idx, value=val)
            cell.font = DATA_FONT
            cell.border = THIN_BORDER
            if row_idx % 2 == 1:
                cell.fill = ALT_FILL
            if col_idx in [1, 2, 3, 5, 6, 7, 8, 10, 11, 12]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")

    ws_rtm.freeze_panes = "A2"
    ws_rtm.auto_filter.ref = f"A1:L{len(rtm_data) + 1}"

    # Auto-fit column widths
    for col in ws_rtm.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws_rtm.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 45)

    # Coverage Sheet
    ws_cov = wb.create_sheet(title="Coverage")
    ws_cov.views.sheetView[0].showGridLines = True

    cov_title = ws_cov.cell(row=1, column=1, value="Traceability & Coverage Audit Summary")
    cov_title.font = TITLE_FONT

    cov_headers = ["Audit Dimension", "Scope Target", "Actual Count", "Coverage Status", "Audit Notes"]
    ws_cov.append([])
    ws_cov.append(cov_headers)
    for col_idx in range(1, len(cov_headers) + 1):
        cell = ws_cov.cell(row=3, column=col_idx)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")

    coverage_rows = [
        ("Operational Questions (OQ)", "10 Defined", "10 Covered", "100% PASS", "All 10 OQs map to supported database fields and features"),
        ("How Might We (HMW)", "10 Defined", "10 Linked", "100% PASS", "Every HMW has an operational owner and traces to at least one BR"),
        ("Business Requirements (BR)", "12 Defined", "12 Linked", "100% PASS", "8 Must (66.7%), 3 Should, 1 Could. No orphan requirements"),
        ("Must Requirements Tracing", "8 Must BRs", "8 Fully Traced", "100% PASS", "Every Must BR maps to Feature, DB Column, User Story, UI Screen, and Test Case"),
        ("HMW without BR", "0 Expected", "0 Discovered", "PASS", "No orphaned problem statements"),
        ("BR without HMW", "0 Expected", "0 Discovered", "PASS", "All business needs originate from identified stakeholder pain points"),
        ("OQ Not Covered", "0 Expected", "0 Discovered", "PASS", "Full analytical query coverage achieved"),
        ("Test Case Traceability", "100% Coverage", "12 TC-F + 10 TC-T", "100% PASS", "Both functional black-box and technical tests implemented")
    ]

    for r_idx, r in enumerate(coverage_rows, start=4):
        for c_idx, val in enumerate(r, start=1):
            cell = ws_cov.cell(row=r_idx, column=c_idx, value=val)
            cell.font = DATA_FONT
            cell.border = THIN_BORDER
            if r_idx % 2 == 1:
                cell.fill = ALT_FILL
            cell.alignment = Alignment(horizontal="center" if c_idx != 5 else "left", vertical="center")

    for col in ws_cov.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws_cov.column_dimensions[col_letter].width = min(max(max_len + 4, 15), 55)

    wb.save(DOCS_DIR / "04_Traceability_Matrix.xlsx")
    print("Saved docs/04_Traceability_Matrix.xlsx")

def create_db_design_workbook():
    wb = openpyxl.Workbook()
    ws_ov = wb.active
    ws_ov.title = "Overview"
    ws_ov.views.sheetView[0].showGridLines = True

    # 1. Overview Sheet
    ws_ov.cell(row=1, column=1, value="Database Design & Schema Specification").font = TITLE_FONT
    ws_ov.cell(row=2, column=1, value="Project: Marketing Campaign Performance Analysis | Organisation: AuraNova Digital Solutions Pvt. Ltd.").font = Font(name="Segoe UI", size=10, italic=True)
    
    overview_text = [
        ("Database Engine", "SQLite 3 (data/app.db)"),
        ("Foreign Key Enforcement", "PRAGMA foreign_keys = ON; enforced on every connection"),
        ("Schema Architecture", "Normalised Star-Snowflake Model (4 Dimension Masters + 1 Fact Table)"),
        ("Fact Table Grain", "One row per report_date × campaign_id × channel_id × segment_id × creative_id"),
        ("Synthetic Fact Volume", "260 Cleaned Fact Rows (Academic Target: 200-300 rows)"),
        ("Analytical Views", "v_campaign_performance_details (Denormalised 34-column analytical view)"),
        ("Integrity Constraints", "Composite unique key on grain, NOT NULL, CHECK constraints (values >= 0, budget > 0)")
    ]
    
    ws_ov.append([])
    ws_ov.append(["Property", "Specification"])
    ws_ov.cell(row=4, column=1).font = HEADER_FONT
    ws_ov.cell(row=4, column=1).fill = HEADER_FILL
    ws_ov.cell(row=4, column=2).font = HEADER_FONT
    ws_ov.cell(row=4, column=2).fill = HEADER_FILL

    for r_idx, (p, s) in enumerate(overview_text, start=5):
        c1 = ws_ov.cell(row=r_idx, column=1, value=p)
        c2 = ws_ov.cell(row=r_idx, column=2, value=s)
        c1.font = DATA_FONT; c1.border = THIN_BORDER
        c2.font = DATA_FONT; c2.border = THIN_BORDER
        if r_idx % 2 == 1:
            c1.fill = ALT_FILL; c2.fill = ALT_FILL

    for col in ws_ov.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        ws_ov.column_dimensions[get_column_letter(col[0].column)].width = max(max_len + 4, 25)

    # 2. Tables Summary Sheet
    ws_ts = wb.create_sheet(title="Tables_Summary")
    ws_ts.views.sheetView[0].showGridLines = True
    ts_headers = ["Table Name", "Table Type", "Grain Description", "Primary Key", "Foreign Keys", "Columns", "Planned Rows"]
    ws_ts.append(ts_headers)
    for c_idx in range(1, len(ts_headers) + 1):
        cell = ws_ts.cell(row=1, column=c_idx)
        cell.font = HEADER_FONT; cell.fill = HEADER_FILL; cell.alignment = Alignment(horizontal="center")

    ts_data = [
        ("campaigns", "Dimension Master", "One row per strategic marketing campaign", "campaign_id", "None", 7, 5),
        ("channels", "Dimension Master", "One row per digital marketing channel", "channel_id", "None", 4, 5),
        ("audience_segments", "Dimension Master", "One row per regional demographic audience cohort", "segment_id", "None", 5, 5),
        ("creatives", "Dimension Master", "One row per creative asset and offer variant", "creative_id", "None", 5, 5),
        ("daily_performance", "Fact Table", "One row per date × campaign × channel × segment × creative", "fact_id", "campaign_id, channel_id, segment_id, creative_id", 12, 260)
    ]
    for r_idx, row in enumerate(ts_data, start=2):
        for c_idx, val in enumerate(row, start=1):
            cell = ws_ts.cell(row=r_idx, column=c_idx, value=val)
            cell.font = DATA_FONT; cell.border = THIN_BORDER
            cell.alignment = Alignment(horizontal="center" if c_idx in [2, 4, 6, 7] else "left")
            if r_idx % 2 == 1: cell.fill = ALT_FILL
    for col in ws_ts.columns:
        ws_ts.column_dimensions[get_column_letter(col[0].column)].width = max(max(len(str(c.value or '')) for c in col) + 3, 15)

    # 3-7 Data Dictionary Sheets
    dict_headers = ["Column Name", "SQL Type", "Key", "Nullable", "Allowed Values / Range", "Description", "Supported OQ / BR"]
    
    table_dicts = {
        "campaigns": [
            ("campaign_id", "TEXT", "PK", "NO", "CMP-001 to CMP-005", "Unique alphanumeric campaign identifier", "OQ-04 / BR-04"),
            ("campaign_name", "TEXT", "-", "NO", "Alphanumeric string", "Descriptive marketing campaign name", "OQ-04 / BR-04"),
            ("objective", "TEXT", "-", "NO", "Lead Gen, Brand Awareness, Trials", "Strategic campaign objective", "OQ-04 / BR-04"),
            ("start_date", "TEXT", "-", "NO", "YYYY-MM-DD", "Official campaign rollout date", "OQ-02 / BR-02"),
            ("end_date", "TEXT", "-", "NO", "YYYY-MM-DD", "Official campaign conclusion date", "OQ-02 / BR-02"),
            ("budget_inr", "REAL", "-", "NO", "> 0.0 (CHECK)", "Total sanctioned marketing budget in INR", "OQ-05 / BR-05"),
            ("status", "TEXT", "-", "NO", "'Active', 'Completed', 'Paused'", "Lifecycle operational status", "OQ-04 / BR-04")
        ],
        "channels": [
            ("channel_id", "TEXT", "PK", "NO", "CH-01 to CH-05", "Unique digital channel identifier", "OQ-03 / BR-03"),
            ("channel_name", "TEXT", "-", "NO", "Google Search, LinkedIn, Meta, etc.", "Marketing delivery channel name", "OQ-03 / BR-03"),
            ("channel_type", "TEXT", "-", "NO", "Search, Social, Direct Inbound, Video", "High-level media classification", "OQ-03 / BR-03"),
            ("cost_model", "TEXT", "-", "NO", "CPC, CPM, CPA, CPV", "Underlying media commercial model", "OQ-03 / BR-03")
        ],
        "audience_segments": [
            ("segment_id", "TEXT", "PK", "NO", "SEG-01 to SEG-05", "Unique audience cohort identifier", "OQ-06 / BR-06"),
            ("segment_name", "TEXT", "-", "NO", "Alphanumeric cohort name", "Target customer segment designation", "OQ-06 / BR-06"),
            ("region", "TEXT", "-", "NO", "North, South, West, East, Pan-India", "Geographic target territory in India", "OQ-06 / BR-06"),
            ("target_age_band", "TEXT", "-", "NO", "Age brackets e.g. '30-48'", "Target demographic age range", "OQ-06 / BR-06"),
            ("primary_device", "TEXT", "-", "NO", "'Desktop', 'Mobile', 'Multi-device'", "Dominant digital access device", "OQ-06 / BR-06")
        ],
        "creatives": [
            ("creative_id", "TEXT", "PK", "NO", "CRT-01 to CRT-05", "Unique ad creative asset identifier", "OQ-07 / BR-07"),
            ("creative_name", "TEXT", "-", "NO", "Descriptive asset title", "Marketing creative asset name", "OQ-07 / BR-07"),
            ("creative_format", "TEXT", "-", "NO", "Whitepaper, Video, Carousel, Tool", "Creative medium and layout format", "OQ-07 / BR-07"),
            ("messaging_theme", "TEXT", "-", "NO", "ROI, AI, Security, Social Proof", "Underlying value proposition angle", "OQ-07 / BR-07"),
            ("offer_type", "TEXT", "-", "NO", "Download, Demo, Assessment, Trial", "Call-to-action offer mechanism", "OQ-07 / BR-07")
        ],
        "daily_performance": [
            ("fact_id", "INTEGER", "PK", "NO", "AUTOINCREMENT (1..260)", "Unique synthetic surrogate fact identifier", "OQ-10 / BR-10"),
            ("report_date", "TEXT", "-", "NO", "YYYY-MM-DD (2026-01 to 2026-09)", "Daily campaign execution timestamp", "OQ-02 / BR-02"),
            ("campaign_id", "TEXT", "FK", "NO", "REFERENCES campaigns(campaign_id)", "Associated marketing campaign foreign key", "OQ-04 / BR-04"),
            ("channel_id", "TEXT", "FK", "NO", "REFERENCES channels(channel_id)", "Associated media delivery channel foreign key", "OQ-03 / BR-03"),
            ("segment_id", "TEXT", "FK", "NO", "REFERENCES audience_segments(segment_id)", "Associated audience segment foreign key", "OQ-06 / BR-06"),
            ("creative_id", "TEXT", "FK", "NO", "REFERENCES creatives(creative_id)", "Associated creative asset foreign key", "OQ-07 / BR-07"),
            ("impressions", "INTEGER", "-", "NO", ">= 0 (CHECK)", "Total digital ad exposures served", "OQ-01 / BR-01"),
            ("clicks", "INTEGER", "-", "NO", ">= 0 (CHECK)", "Total ad interactions/clicks registered", "OQ-01 / BR-01"),
            ("conversions", "INTEGER", "-", "NO", ">= 0 (CHECK)", "Total verified lead/trial conversion actions", "OQ-01 / BR-01"),
            ("spend_inr", "REAL", "-", "NO", ">= 0.0 (CHECK)", "Total actual daily ad spend incurred in INR", "OQ-01, OQ-05 / BR-01"),
            ("revenue_inr", "REAL", "-", "NO", ">= 0.0 (CHECK)", "Total attributed top-line revenue in INR", "OQ-01, OQ-08 / BR-01"),
            ("reach", "INTEGER", "-", "NO", ">= 0 (CHECK)", "Unique addressable user reach estimated", "OQ-01 / BR-01")
        ]
    }

    for t_name, rows in table_dicts.items():
        ws = wb.create_sheet(title=f"tbl_{t_name}")
        ws.views.sheetView[0].showGridLines = True
        ws.append(dict_headers)
        for c_idx in range(1, len(dict_headers) + 1):
            cell = ws.cell(row=1, column=c_idx)
            cell.font = HEADER_FONT; cell.fill = HEADER_FILL; cell.alignment = Alignment(horizontal="center")
        for r_idx, r in enumerate(rows, start=2):
            for c_idx, val in enumerate(r, start=1):
                cell = ws.cell(row=r_idx, column=c_idx, value=val)
                cell.font = DATA_FONT; cell.border = THIN_BORDER
                cell.alignment = Alignment(horizontal="center" if c_idx in [2, 3, 4] else "left")
                if r_idx % 2 == 1: cell.fill = ALT_FILL
        ws.freeze_panes = "A2"
        for col in ws.columns:
            ws.column_dimensions[get_column_letter(col[0].column)].width = max(max(len(str(c.value or '')) for c in col) + 3, 14)

    # 8. Relationships Sheet
    ws_rel = wb.create_sheet(title="Relationships")
    ws_rel.views.sheetView[0].showGridLines = True
    rel_headers = ["Parent Table", "Child Table", "Foreign Key Column", "Cardinality", "Business Relationship Rule", "Cascade Rule"]
    ws_rel.append(rel_headers)
    for c_idx in range(1, len(rel_headers) + 1):
        cell = ws_rel.cell(row=1, column=c_idx)
        cell.font = HEADER_FONT; cell.fill = HEADER_FILL; cell.alignment = Alignment(horizontal="center")
    rel_data = [
        ("campaigns", "daily_performance", "campaign_id", "1 : N (One-to-Many)", "One campaign generates many daily performance records", "RESTRICT (No orphan facts)"),
        ("channels", "daily_performance", "channel_id", "1 : N (One-to-Many)", "One channel delivers impressions across multiple campaigns & days", "RESTRICT"),
        ("audience_segments", "daily_performance", "segment_id", "1 : N (One-to-Many)", "One target audience segment receives multiple daily impressions", "RESTRICT"),
        ("creatives", "daily_performance", "creative_id", "1 : N (One-to-Many)", "One creative asset is deployed across multiple campaign-days", "RESTRICT")
    ]
    for r_idx, r in enumerate(rel_data, start=2):
        for c_idx, val in enumerate(r, start=1):
            cell = ws_rel.cell(row=r_idx, column=c_idx, value=val)
            cell.font = DATA_FONT; cell.border = THIN_BORDER
            cell.alignment = Alignment(horizontal="center" if c_idx in [1, 2, 3, 4, 6] else "left")
            if r_idx % 2 == 1: cell.fill = ALT_FILL
    for col in ws_rel.columns:
        ws_rel.column_dimensions[get_column_letter(col[0].column)].width = max(max(len(str(c.value or '')) for c in col) + 3, 16)

    # 9. KPIs_Measures Sheet
    ws_kpi = wb.create_sheet(title="KPIs_Measures")
    ws_kpi.views.sheetView[0].showGridLines = True
    kpi_headers = ["Metric / KPI Name", "Formula Formulation", "Aggregation Level", "Unit", "Division-by-Zero Guard", "Supported OQ"]
    ws_kpi.append(kpi_headers)
    for c_idx in range(1, len(kpi_headers) + 1):
        cell = ws_kpi.cell(row=1, column=c_idx)
        cell.font = HEADER_FONT; cell.fill = HEADER_FILL; cell.alignment = Alignment(horizontal="center")
    kpi_data = [
        ("Impressions", "SUM(impressions)", "Total Aggregate", "Count", "N/A", "OQ-01"),
        ("Clicks", "SUM(clicks)", "Total Aggregate", "Count", "N/A", "OQ-01"),
        ("Conversions", "SUM(conversions)", "Total Aggregate", "Count", "N/A", "OQ-01"),
        ("Spend", "SUM(spend_inr)", "Total Aggregate", "INR (₹)", "N/A", "OQ-01, OQ-05"),
        ("Revenue", "SUM(revenue_inr)", "Total Aggregate", "INR (₹)", "N/A", "OQ-01, OQ-08"),
        ("Click-Through Rate (CTR)", "(SUM(clicks) / SUM(impressions)) * 100", "Ratio of Totals", "%", "Returns None / '—' if impressions = 0", "OQ-03, OQ-07"),
        ("Cost Per Click (CPC)", "SUM(spend_inr) / SUM(clicks)", "Ratio of Totals", "INR (₹)", "Returns None / '—' if clicks = 0", "OQ-03"),
        ("Cost Per Mille (CPM)", "(SUM(spend_inr) / SUM(impressions)) * 1000", "Ratio of Totals", "INR (₹)", "Returns None / '—' if impressions = 0", "OQ-03"),
        ("Conversion Rate (CVR)", "(SUM(conversions) / SUM(clicks)) * 100", "Ratio of Totals", "%", "Returns None / '—' if clicks = 0", "OQ-03, OQ-06"),
        ("Cost Per Acquisition (CPA)", "SUM(spend_inr) / SUM(conversions)", "Ratio of Totals", "INR (₹)", "Returns None / '—' if conversions = 0", "OQ-03, OQ-04"),
        ("Return on Ad Spend (ROAS)", "SUM(revenue_inr) / SUM(spend_inr)", "Ratio of Totals", "Ratio (x)", "Returns None / '—' if spend = 0", "OQ-01, OQ-03, OQ-04"),
        ("Return on Investment (ROI)", "((SUM(revenue_inr) - SUM(spend_inr)) / SUM(spend_inr)) * 100", "Ratio of Totals", "%", "Returns None / '—' if spend = 0", "OQ-01, OQ-04")
    ]
    for r_idx, r in enumerate(kpi_data, start=2):
        for c_idx, val in enumerate(r, start=1):
            cell = ws_kpi.cell(row=r_idx, column=c_idx, value=val)
            cell.font = DATA_FONT; cell.border = THIN_BORDER
            cell.alignment = Alignment(horizontal="center" if c_idx in [3, 4, 6] else "left")
            if r_idx % 2 == 1: cell.fill = ALT_FILL
    for col in ws_kpi.columns:
        ws_kpi.column_dimensions[get_column_letter(col[0].column)].width = max(max(len(str(c.value or '')) for c in col) + 3, 16)

    # 10. SQLite DDL Sheet
    ws_ddl = wb.create_sheet(title="SQLite_DDL")
    ws_ddl.views.sheetView[0].showGridLines = True
    ws_ddl.append(["Object Name", "Object Type", "Executable SQLite DDL Script"])
    for c_idx in range(1, 4):
        cell = ws_ddl.cell(row=1, column=c_idx)
        cell.font = HEADER_FONT; cell.fill = HEADER_FILL; cell.alignment = Alignment(horizontal="center")
    
    ddl_statements = [
        ("PRAGMA", "Configuration", "PRAGMA foreign_keys = ON;"),
        ("campaigns", "Table", "CREATE TABLE campaigns (\n    campaign_id TEXT PRIMARY KEY,\n    campaign_name TEXT NOT NULL,\n    objective TEXT NOT NULL,\n    start_date TEXT NOT NULL,\n    end_date TEXT NOT NULL,\n    budget_inr REAL NOT NULL CHECK(budget_inr > 0),\n    status TEXT NOT NULL CHECK(status IN ('Active', 'Completed', 'Paused'))\n);"),
        ("channels", "Table", "CREATE TABLE channels (\n    channel_id TEXT PRIMARY KEY,\n    channel_name TEXT NOT NULL,\n    channel_type TEXT NOT NULL,\n    cost_model TEXT NOT NULL\n);"),
        ("audience_segments", "Table", "CREATE TABLE audience_segments (\n    segment_id TEXT PRIMARY KEY,\n    segment_name TEXT NOT NULL,\n    region TEXT NOT NULL,\n    target_age_band TEXT NOT NULL,\n    primary_device TEXT NOT NULL CHECK(primary_device IN ('Desktop', 'Mobile', 'Multi-device'))\n);"),
        ("creatives", "Table", "CREATE TABLE creatives (\n    creative_id TEXT PRIMARY KEY,\n    creative_name TEXT NOT NULL,\n    creative_format TEXT NOT NULL,\n    messaging_theme TEXT NOT NULL,\n    offer_type TEXT NOT NULL\n);"),
        ("daily_performance", "Table", "CREATE TABLE daily_performance (\n    fact_id INTEGER PRIMARY KEY AUTOINCREMENT,\n    report_date TEXT NOT NULL,\n    campaign_id TEXT NOT NULL REFERENCES campaigns(campaign_id),\n    channel_id TEXT NOT NULL REFERENCES channels(channel_id),\n    segment_id TEXT NOT NULL REFERENCES audience_segments(segment_id),\n    creative_id TEXT NOT NULL REFERENCES creatives(creative_id),\n    impressions INTEGER NOT NULL CHECK(impressions >= 0),\n    clicks INTEGER NOT NULL CHECK(clicks >= 0),\n    conversions INTEGER NOT NULL CHECK(conversions >= 0),\n    spend_inr REAL NOT NULL CHECK(spend_inr >= 0),\n    revenue_inr REAL NOT NULL CHECK(revenue_inr >= 0),\n    reach INTEGER NOT NULL CHECK(reach >= 0),\n    UNIQUE(report_date, campaign_id, channel_id, segment_id, creative_id)\n);"),
        ("v_campaign_performance_details", "View", "CREATE VIEW v_campaign_performance_details AS\nSELECT \n    dp.fact_id, dp.report_date, strftime('%Y-%m', dp.report_date) AS report_month,\n    c.campaign_id, c.campaign_name, c.objective AS campaign_objective, c.budget_inr AS campaign_budget_inr, c.status AS campaign_status,\n    ch.channel_id, ch.channel_name, ch.channel_type, ch.cost_model,\n    s.segment_id, s.segment_name, s.region, s.target_age_band, s.primary_device,\n    cr.creative_id, cr.creative_name, cr.creative_format, cr.messaging_theme, cr.offer_type,\n    dp.impressions, dp.clicks, dp.conversions, dp.spend_inr, dp.revenue_inr, dp.reach,\n    CASE WHEN dp.impressions > 0 THEN ROUND((CAST(dp.clicks AS REAL) / dp.impressions) * 100.0, 2) ELSE 0.0 END AS ctr_pct,\n    CASE WHEN dp.clicks > 0 THEN ROUND(dp.spend_inr / dp.clicks, 2) ELSE 0.0 END AS cpc_inr,\n    CASE WHEN dp.impressions > 0 THEN ROUND((dp.spend_inr / dp.impressions) * 1000.0, 2) ELSE 0.0 END AS cpm_inr,\n    CASE WHEN dp.clicks > 0 THEN ROUND((CAST(dp.conversions AS REAL) / dp.clicks) * 100.0, 2) ELSE 0.0 END AS cvr_pct,\n    CASE WHEN dp.conversions > 0 THEN ROUND(dp.spend_inr / dp.conversions, 2) ELSE 0.0 END AS cpa_inr,\n    CASE WHEN dp.spend_inr > 0 THEN ROUND(dp.revenue_inr / dp.spend_inr, 2) ELSE 0.0 END AS roas,\n    CASE WHEN dp.spend_inr > 0 THEN ROUND(((dp.revenue_inr - dp.spend_inr) / dp.spend_inr) * 100.0, 2) ELSE 0.0 END AS roi_pct\nFROM daily_performance dp\nJOIN campaigns c ON dp.campaign_id = c.campaign_id\nJOIN channels ch ON dp.channel_id = ch.channel_id\nJOIN audience_segments s ON dp.segment_id = s.segment_id\nJOIN creatives cr ON dp.creative_id = cr.creative_id;")
    ]

    for r_idx, (obj, o_type, script) in enumerate(ddl_statements, start=2):
        c1 = ws_ddl.cell(row=r_idx, column=1, value=obj)
        c2 = ws_ddl.cell(row=r_idx, column=2, value=o_type)
        c3 = ws_ddl.cell(row=r_idx, column=3, value=script)
        c1.font = DATA_FONT; c1.border = THIN_BORDER; c1.alignment = Alignment(horizontal="center")
        c2.font = DATA_FONT; c2.border = THIN_BORDER; c2.alignment = Alignment(horizontal="center")
        c3.font = Font(name="Consolas", size=9); c3.border = THIN_BORDER; c3.alignment = Alignment(horizontal="left", wrap_text=True)
        if r_idx % 2 == 1:
            c1.fill = ALT_FILL; c2.fill = ALT_FILL; c3.fill = ALT_FILL
            
    ws_ddl.column_dimensions["A"].width = 25
    ws_ddl.column_dimensions["B"].width = 18
    ws_ddl.column_dimensions["C"].width = 85

    wb.save(DOCS_DIR / "08_Database_Design.xlsx")
    print("Saved docs/08_Database_Design.xlsx")

if __name__ == "__main__":
    create_rtm_workbook()
    create_db_design_workbook()
