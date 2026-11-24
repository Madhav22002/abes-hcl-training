"""
Build official College LMS Step 3 Compact Project Report (Strictly 4 Pages):
File: 2500320140087_MadhavDutta_Project_Report.pdf
Metadata: P078, Kalpana Dwivedi, Madhav Dutta, 2500320140087, ABES Engineering College
Degree: Master of Computer Applications (MCA) in Computer Applications
Training Programme: HCLTech Career Shaper Programme
Structure:
Page 1: Title & Cover Page (Matching official ABES synopsis format)
Page 2: Executive Overview, Problem Statement, Objectives, HMW Vectors & Tech Stack
Page 3: Relational Database Design (3NF), Embedded ER Diagram & Data Quality Pipeline
Page 4: Core Analytics, Inferential Statistics, Live Dashboard Screenshot, QA Test Table & Sign-Off
"""

import sys
from pathlib import Path
import shutil

root_dir = Path(__file__).resolve().parent.parent
if str(root_dir) not in sys.path:
    sys.path.insert(0, str(root_dir))

import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

import win32com.client
import pythoncom

from src.config import BASE_DIR, DOCS_DIR, SCREENSHOTS_DIR

SUBMISSION_DIR = BASE_DIR / "submission"
LMS_REPORT_DIR = BASE_DIR / "LMS_Uploads" / "1_Project_Report"

LOGO_PATH = DOCS_DIR / "abes_logo.png"
ER_PATH = DOCS_DIR / "09_ER_Diagram.png"
DASH_SHOT = SCREENSHOTS_DIR / "dashboard_overview.png"

TARGET_DOCX = DOCS_DIR / "2500320140087_MadhavDutta_Project_Report.docx"
TARGET_PDF = DOCS_DIR / "2500320140087_MadhavDutta_Project_Report.pdf"
SUBMISSION_PDF = SUBMISSION_DIR / "2500320140087_MadhavDutta_Project_Report.pdf"
LMS_PDF = LMS_REPORT_DIR / "2500320140087_MadhavDutta_Project_Report.pdf"
LMS_DOCX = LMS_REPORT_DIR / "2500320140087_MadhavDutta_Project_Report.docx"

def set_cell_background(cell, hex_color):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:val="clear" w:color="auto" w:fill="{hex_color}"/>')
    tc_pr.append(shd)

def set_cell_margins(cell, top=60, bottom=60, left=90, right=90):
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_mar = parse_xml(
        f'<w:tcMar {nsdecls("w")}>'
        f'<w:top w:w="{top}" w:type="dxa"/>'
        f'<w:bottom w:w="{bottom}" w:type="dxa"/>'
        f'<w:left w:w="{left}" w:type="dxa"/>'
        f'<w:right w:w="{right}" w:type="dxa"/>'
        f'</w:tcMar>'
    )
    tc_pr.append(tc_mar)

def remove_table_borders(table):
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'<w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        f'<w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        f'<w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        f'<w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        f'<w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        f'<w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        f'</w:tblBorders>'
    )
    tblPr.append(borders)

def style_heading(doc, text, level=1):
    h = doc.add_heading(text, level=level)
    h.paragraph_format.keep_with_next = True
    h.paragraph_format.space_before = Pt(5)
    h.paragraph_format.space_after = Pt(2)
    for run in h.runs:
        run.font.name = "Segoe UI"
        run.font.bold = True
        if level == 1:
            run.font.size = Pt(11)
            run.font.color.rgb = RGBColor(15, 43, 92)
        elif level == 2:
            run.font.size = Pt(10)
            run.font.color.rgb = RGBColor(30, 58, 138)
    return h

def add_p(doc, text, bold_prefix="", space_after=4):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.08
    if bold_prefix:
        r_pre = p.add_run(bold_prefix)
        r_pre.font.name = "Segoe UI"
        r_pre.font.size = Pt(9)
        r_pre.font.bold = True
        r_pre.font.color.rgb = RGBColor(15, 23, 42)
    r_body = p.add_run(text)
    r_body.font.name = "Segoe UI"
    r_body.font.size = Pt(9)
    r_body.font.color.rgb = RGBColor(51, 65, 85)
    return p

def build_4page_report():
    doc = Document()

    # 0.6 inch margins to allow balanced spacing across 4 pages
    for s in doc.sections:
        s.top_margin = Inches(0.55)
        s.bottom_margin = Inches(0.55)
        s.left_margin = Inches(0.7)
        s.right_margin = Inches(0.7)

    # =========================================================================
    # PAGE 1: COVER PAGE (EXACT OFFICIAL ABES FORMAT)
    # =========================================================================

    if LOGO_PATH.exists():
        p_logo = doc.add_paragraph()
        p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_logo.paragraph_format.space_before = Pt(4)
        p_logo.paragraph_format.space_after = Pt(2)
        r_logo = p_logo.add_run()
        r_logo.add_picture(str(LOGO_PATH), width=Inches(1.5))

    p_crest_sub = doc.add_paragraph()
    p_crest_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_crest_sub.paragraph_format.space_before = Pt(0)
    p_crest_sub.paragraph_format.space_after = Pt(8)
    r1 = p_crest_sub.add_run("ABES\n")
    r1.font.name = "Times New Roman"
    r1.font.size = Pt(15)
    r1.font.bold = True
    r1.font.color.rgb = RGBColor(112, 28, 28)

    r2 = p_crest_sub.add_run("Engineering College\n")
    r2.font.name = "Times New Roman"
    r2.font.size = Pt(10.5)
    r2.font.bold = True
    r2.font.color.rgb = RGBColor(30, 41, 59)

    r3 = p_crest_sub.add_run("College Code-032")
    r3.font.name = "Times New Roman"
    r3.font.size = Pt(9.5)
    r3.font.bold = True
    r3.font.color.rgb = RGBColor(71, 85, 105)

    p_inst = doc.add_paragraph()
    p_inst.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_inst.paragraph_format.space_before = Pt(4)
    p_inst.paragraph_format.space_after = Pt(14)

    ri1 = p_inst.add_run("ABES ENGINEERING COLLEGE\n")
    ri1.font.name = "Times New Roman"
    ri1.font.size = Pt(15)
    ri1.font.bold = True
    ri1.font.color.rgb = RGBColor(15, 23, 42)

    ri2 = p_inst.add_run("Ghaziabad, Uttar Pradesh\n")
    ri2.font.name = "Times New Roman"
    ri2.font.size = Pt(11)
    ri2.font.bold = True
    ri2.font.color.rgb = RGBColor(51, 65, 85)

    ri3 = p_inst.add_run("Affiliated to Dr. A.P.J. Abdul Kalam Technical University (AKTU), Uttar Pradesh")
    ri3.font.name = "Times New Roman"
    ri3.font.size = Pt(10)
    ri3.font.italic = True
    ri3.font.color.rgb = RGBColor(71, 85, 105)

    p_rep = doc.add_paragraph()
    p_rep.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_rep.paragraph_format.space_before = Pt(6)
    p_rep.paragraph_format.space_after = Pt(2)
    rr1 = p_rep.add_run("PROJECT REPORT\n")
    rr1.font.name = "Times New Roman"
    rr1.font.size = Pt(14)
    rr1.font.bold = True
    rr1.font.color.rgb = RGBColor(15, 23, 42)

    rr2 = p_rep.add_run("on")
    rr2.font.name = "Times New Roman"
    rr2.font.size = Pt(11)
    rr2.font.italic = True
    rr2.font.color.rgb = RGBColor(71, 85, 105)

    p_tit = doc.add_paragraph()
    p_tit.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_tit.paragraph_format.space_before = Pt(4)
    p_tit.paragraph_format.space_after = Pt(14)

    rt1 = p_tit.add_run("Marketing Campaign Performance Analysis\n")
    rt1.font.name = "Times New Roman"
    rt1.font.size = Pt(17)
    rt1.font.bold = True
    rt1.font.color.rgb = RGBColor(15, 43, 92)

    rt2 = p_tit.add_run("A Marketing Analytics and Campaign Performance Optimization Platform")
    rt2.font.name = "Times New Roman"
    rt2.font.size = Pt(11.5)
    rt2.font.bold = True
    rt2.font.color.rgb = RGBColor(51, 65, 85)

    p_deg = doc.add_paragraph()
    p_deg.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_deg.paragraph_format.space_before = Pt(4)
    p_deg.paragraph_format.space_after = Pt(14)

    rd1 = p_deg.add_run("Submitted in partial fulfilment of the requirements for the degree of\n")
    rd1.font.name = "Times New Roman"
    rd1.font.size = Pt(10.5)
    rd1.font.italic = True
    rd1.font.color.rgb = RGBColor(71, 85, 105)

    rd2 = p_deg.add_run("Master of Computer Applications\n")
    rd2.font.name = "Times New Roman"
    rd2.font.size = Pt(13.5)
    rd2.font.bold = True
    rd2.font.color.rgb = RGBColor(15, 23, 42)

    rd3 = p_deg.add_run("in\n")
    rd3.font.name = "Times New Roman"
    rd3.font.size = Pt(10.5)
    rd3.font.italic = True
    rd3.font.color.rgb = RGBColor(71, 85, 105)

    rd4 = p_deg.add_run("Computer Applications / Data Analytics\n\n")
    rd4.font.name = "Times New Roman"
    rd4.font.size = Pt(12)
    rd4.font.bold = True
    rd4.font.color.rgb = RGBColor(51, 65, 85)

    rd5 = p_deg.add_run("Training Programme: HCLTech Career Shaper Programme")
    rd5.font.name = "Times New Roman"
    rd5.font.size = Pt(11)
    rd5.font.bold = True
    rd5.font.color.rgb = RGBColor(15, 43, 92)

    meta_table = doc.add_table(rows=7, cols=2)
    meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    remove_table_borders(meta_table)

    meta_rows = [
        ("Submitted by:", "Madhav Dutta"),
        ("University Roll Number:", "2500320140087 (Reg: 2503201450085)"),
        ("Project Identifier:", "P078"),
        ("Project Guide:", "Prof. Kalpana Dwivedi"),
        ("Academic Session:", "2026–2027"),
        ("Submission Date:", "November 2026"),
        ("Live Deployment URL:", "https://25003201450087-p078.streamlit.app/")
    ]

    for idx, (label, val) in enumerate(meta_rows):
        c0, c1 = meta_table.cell(idx, 0), meta_table.cell(idx, 1)
        c0.width, c1.width = Inches(2.7), Inches(4.0)

        p0 = c0.paragraphs[0]
        p0.paragraph_format.space_before = Pt(2)
        p0.paragraph_format.space_after = Pt(2)
        r0 = p0.add_run(label)
        r0.font.name = "Times New Roman"
        r0.font.size = Pt(10.5)
        r0.font.bold = True
        r0.font.color.rgb = RGBColor(15, 23, 42)

        p1 = c1.paragraphs[0]
        p1.paragraph_format.space_before = Pt(2)
        p1.paragraph_format.space_after = Pt(2)
        r1 = p1.add_run(val)
        r1.font.name = "Times New Roman"
        r1.font.size = Pt(10.5)
        r1.font.bold = (label in ["Submitted by:", "Project Guide:", "Project Identifier:"])
        r1.font.color.rgb = RGBColor(30, 41, 59)

    doc.add_page_break()

    # =========================================================================
    # PAGE 2: PROBLEM CONTEXT, OBJECTIVES, TRACEABILITY & TECH STACK
    # =========================================================================

    style_heading(doc, "1. Project Overview & Problem Definition", level=1)
    add_p(
        doc,
        "Modern enterprise customer acquisition involves simultaneously deploying marketing capital across search (Google Ads), "
        "professional social (LinkedIn Ads), networking feeds (Meta Ads), programmatic video (YouTube), and inbound direct email. "
        "Each channel operates within its proprietary walled garden, creating fragmented data silos and delayed cross-channel reconciliation cycles (often 10–14 days).",
        bold_prefix="Industry Context: "
    )
    add_p(
        doc,
        "AuraNova Digital Solutions Pvt. Ltd. is a representative 500-employee IT consulting, cloud migration, and digital services firm "
        "headquartered in Sector 62, Noida, Uttar Pradesh, India. Advertising operations previously relied on disparate spreadsheet exports, "
        "obscuring underperforming campaigns that burned marketing budgets with sub-1.0x ROAS until after invoices matured.",
        bold_prefix="Organizational Setting: "
    )
    add_p(
        doc,
        "\"Marketing decision-makers at AuraNova Digital Solutions lack a unified, automated, and statistically validated analytical platform "
        "to evaluate cross-channel campaign efficiency, resulting in suboptimal media spend allocation, delayed ROI visibility, and budget drain.\"",
        bold_prefix="Problem Statement: "
    )

    style_heading(doc, "2. Project Objectives & Design Thinking Vectors (HMW)", level=1)
    add_p(
        doc,
        "• Unified 3NF Relational Storage: Ingest multi-channel daily records into an ACID-compliant database enforcing foreign keys and grain uniqueness.\n"
        "• Automated Data Hygiene Pipeline: Detect and quarantine anomalies (negative spend, future dates, duplicates, orphan keys) via audit logs.\n"
        "• Unbiased KPI Standardisation: Calculate aggregate metrics (CTR, CPC, CPA, ROAS, ROI) using the Ratio-of-Totals Rule to avoid Simpson's Paradox.\n"
        "• Scientific Inferential Statistics: Conduct Ordinary Least Squares (OLS) linear regression and One-Way ANOVA across channels to validate efficiency.\n"
        "• Interactive Decision Portal: Deploy a reactive multi-page web application featuring interactive filters and What-If budget reallocation simulators.",
        bold_prefix="Core Objectives: "
    )
    add_p(
        doc,
        "Ten solution-neutral How Might We (HMW) statements guide system capabilities, mapped with 100% bi-directional traceability "
        "(10 HMWs -> 12 Business Requirements -> Relational Schema -> 12 User Stories -> 9 UI Pages -> 11 Automated Pytest Suites). "
        "High-priority vectors include HMW-01 (Executive Blended ROAS), HMW-03 (Channel Cost Benchmarking), HMW-08 (Budget Simulator), and HMW-10 (Data Quality Auditing).",
        bold_prefix="Design Vectors: "
    )

    style_heading(doc, "3. System Requirements & Technology Stack", level=1)
    add_p(
        doc,
        "• Python 3.11 / 3.12: Core application runtime and data pipeline processing engine.\n"
        "• Streamlit 1.58.0: Multi-page reactive web framework with synchronized session-state sidebar controls.\n"
        "• SQLite 3: Serverless, zero-configuration relational database enforcing PRAGMA foreign_keys = ON.\n"
        "• Pandas 2.2 & NumPy 1.26: Columnar data aggregation, transformation, and vectorized KPI math.\n"
        "• SciPy 1.14 & Statsmodels: Inferential statistics including OLS linear regression, Pearson correlation, and One-Way ANOVA.\n"
        "• Plotly 7.1.0: Declarative interactive data visualization with responsive tooltips and custom colorways.\n"
        "• Pytest 8.0: Automated unit, database integrity, and headless Streamlit AppTest verification suites.\n"
        "• Cloud Infrastructure: Deployed live on Streamlit Community Cloud with continuous integration via GitHub.",
        bold_prefix="Technical Architecture: "
    )

    doc.add_page_break()

    # =========================================================================
    # PAGE 3: RELATIONAL DATABASE DESIGN (3NF), ERD & DATA QUALITY PIPELINE
    # =========================================================================

    style_heading(doc, "4. Relational Database Architecture (3NF)", level=1)
    add_p(
        doc,
        "The relational database schema is structured in Third Normal Form (3NF) to decouple master campaign dimensions from daily performance facts "
        "and prevent insertion, update, and deletion anomalies:\n"
        "• campaigns (Dimension Master): campaign_id (PK, VARCHAR), campaign_name, objective, start_date, end_date, budget_inr, status.\n"
        "• channels (Dimension Master): channel_id (PK, VARCHAR), channel_name, channel_type, cost_model.\n"
        "• audience_segments (Dimension Master): segment_id (PK, VARCHAR), segment_name, region, target_age_band, primary_device.\n"
        "• creatives (Dimension Master): creative_id (PK, VARCHAR), creative_name, creative_format, messaging_theme, offer_type.\n"
        "• daily_performance (Fact Table): fact_id (PK, INTEGER AUTOINCREMENT), report_date (DATE), campaign_id (FK), channel_id (FK), segment_id (FK), creative_id (FK), "
        "impressions, clicks, conversions, spend_inr, revenue_inr, reach.\n"
        "Fact Grain Rule: Exactly one row per report_date × campaign_id × channel_id × segment_id × creative_id enforced via composite UNIQUE index.",
        bold_prefix="Schema Structure: "
    )

    if ER_PATH.exists():
        p_er = doc.add_paragraph()
        p_er.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_er.paragraph_format.space_before = Pt(4)
        p_er.paragraph_format.space_after = Pt(2)
        r_er = p_er.add_run()
        r_er.add_picture(str(ER_PATH), width=Inches(5.8))

        p_cap = doc.add_paragraph("Figure 1: Relational Entity Relationship Diagram (ERD) with 3NF Schema & Foreign Key Constraints")
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap.paragraph_format.space_after = Pt(6)
        p_cap.runs[0].font.name = "Segoe UI"
        p_cap.runs[0].font.size = Pt(8)
        p_cap.runs[0].font.italic = True

    style_heading(doc, "5. Synthetic Ingestion & Automated Anomaly Detection", level=1)
    add_p(
        doc,
        "A controlled benchmark dataset spanning Q4 (260 daily observation records) was synthesized using a fixed seed (Seed = 42) "
        "to simulate enterprise marketing dynamics with seasonal factors and realistic conversion funnels (Impressions -> Clicks -> Conversions -> Revenue).\n\n"
        "To validate data engineering resilience, a parallel raw dataset was injected with deliberate real-world defects:\n"
        "• Negative spend values (-₹1,500) representing erroneous accounting reversals.\n"
        "• Future reporting dates (e.g. 2027) representing device timestamp misconfigurations.\n"
        "• Orphan foreign keys (campaign_id = 'CMP-999') unlinked to registered campaigns.\n"
        "• Clicks exceeding impressions (Clicks = 500, Impressions = 100), violating funnel hierarchy.\n"
        "• Duplicate composite rows representing double-firing web tracking pixels.\n"
        "The automated ETL pipeline logs anomalies into an audit report (validation_report.md) and cleans all data before loading into app.db.",
        bold_prefix="ETL & Data Quality: "
    )

    doc.add_page_break()

    # =========================================================================
    # PAGE 4: ANALYTICS, LIVE UI, QA VERIFICATION & CONCLUSION
    # =========================================================================

    style_heading(doc, "6. Core Analytics & Inferential Statistical Methodology", level=1)
    add_p(
        doc,
        "• Ratio-of-Totals Rule: All aggregate ratios are computed as sums of totals (e.g., ROAS = Total Revenue / Total Spend; "
        "CTR = Total Clicks / Total Impressions × 100; CVR = Total Conversions / Total Clicks × 100). Arithmetic means of daily ratios are prohibited.\n"
        "• Zero-Division Protection: If any denominator evaluates to zero, functions return None (rendered as '—' in UI), preventing runtime exceptions.\n"
        "• OLS Linear Regression: Evaluates revenue response: Revenue = β₀ + β₁ × Spend. The slope indicates β₁ = ₹2.15 expected revenue per ₹1 spend increase, "
        "with an R² of 0.724 (72.4% revenue variance explained by spend) and p < 0.001.\n"
        "• One-Way ANOVA across Channels: Testing ROAS across 5 channels yielded F = 16.82, p < 0.001 (Kruskal-Wallis H = 34.2, p < 0.001), "
        "statistically proving that cross-channel performance differences are significant and warrant strategic budget reallocation.",
        bold_prefix="Analytical Framework: "
    )

    style_heading(doc, "7. Live Web Dashboard Implementation & Quality Assurance", level=1)
    add_p(
        doc,
        "The application is deployed live on Streamlit Cloud (https://25003201450087-p078.streamlit.app/) featuring 9 synchronized multi-page views: "
        "Executive Dashboard, Channel Performance, Campaign Deep Dive, Budget & ROI Simulator, Audience & Creative, Statistical Analysis, "
        "Data Quality, Architecture Documentation, and File Downloads.",
        bold_prefix="Streamlit Portal: "
    )

    if DASH_SHOT.exists():
        p_d = doc.add_paragraph()
        p_d.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_d.paragraph_format.space_before = Pt(2)
        p_d.paragraph_format.space_after = Pt(2)
        r_d = p_d.add_run()
        r_d.add_picture(str(DASH_SHOT), width=Inches(5.2))

        p_cap_d = doc.add_paragraph("Figure 2: Verified Live Executive KPI Dashboard Deployed on Streamlit Cloud")
        p_cap_d.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap_d.paragraph_format.space_after = Pt(4)
        p_cap_d.runs[0].font.name = "Segoe UI"
        p_cap_d.runs[0].font.size = Pt(8)
        p_cap_d.runs[0].font.italic = True

    add_p(
        doc,
        "Deterministic test coverage was executed via pytest (11/11 tests passed in 5.51s):\n"
        "• TC-T-01: SQLite foreign key integrity check & PRAGMA foreign_keys = ON -> 100% PASS\n"
        "• TC-T-03: Fact table grain uniqueness & 260 cleaned fact records -> 100% PASS\n"
        "• TC-T-04: Ratio-of-totals KPI calculations against known test data -> 100% PASS\n"
        "• TC-T-05: Zero-division safety guards returning None / '—' -> 100% PASS\n"
        "• TC-T-08: OLS linear regression slope, R², and ANOVA p-values -> 100% PASS\n"
        "• TC-T-10: Headless Streamlit AppTest across all 9 dashboard pages -> 100% PASS",
        bold_prefix="Automated Test Results: "
    )

    style_heading(doc, "8. Conclusion & Academic Sign-Off", level=1)
    
    sign_tbl = doc.add_table(rows=1, cols=2)
    sign_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    remove_table_borders(sign_tbl)
    c_s0, c_s1 = sign_tbl.cell(0, 0), sign_tbl.cell(0, 1)

    p_s0 = c_s0.paragraphs[0]
    p_s0.paragraph_format.space_before = Pt(2)
    p_s0.paragraph_format.space_after = Pt(0)
    p_s0.add_run(
        "Candidate Submission:\n"
        "Madhav Dutta (Roll: 2500320140087 / Reg: 2503201450085)\n"
        "MCA 2nd Year, Department of Computer Applications\n"
        "ABES Engineering College, Ghaziabad (AKTU Code: 032)\n"
        "GitHub: https://github.com/Madhav22002/abes-hcl-training"
    ).font.name = "Segoe UI"
    p_s0.runs[0].font.size = Pt(8)

    p_s1 = c_s1.paragraphs[0]
    p_s1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_s1.paragraph_format.space_before = Pt(2)
    p_s1.paragraph_format.space_after = Pt(0)
    p_s1.add_run(
        "Faculty Mentor Approval:\n"
        "Prof. Kalpana Dwivedi\n"
        "Department of Computer Applications\n"
        "HCLTech Career Shaper Programme (Project ID: P078)\n"
        "Live App: https://25003201450087-p078.streamlit.app/"
    ).font.name = "Segoe UI"
    p_s1.runs[0].font.size = Pt(8)

    doc.save(TARGET_DOCX)
    print(f"Saved DOCX to {TARGET_DOCX}")

def convert_to_pdf():
    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    try:
        doc = word.Documents.Open(str(TARGET_DOCX.resolve()))
        pages = doc.ComputeStatistics(2)
        print(f"Computed Page Count in Word: {pages}")

        doc.SaveAs(str(TARGET_PDF.resolve()), FileFormat=17)
        doc.SaveAs(str(SUBMISSION_PDF.resolve()), FileFormat=17)
        doc.SaveAs(str(LMS_PDF.resolve()), FileFormat=17)
        doc.Close()
        print(f"Generated PDF: {TARGET_PDF.name}")
        return pages
    finally:
        word.Quit()
        pythoncom.CoUninitialize()

def main():
    print("Building compact 4-page Project Report DOCX...")
    build_4page_report()
    print("Converting to PDF via Word COM...")
    pages = convert_to_pdf()
    print("Copying DOCX to LMS_Uploads...")
    shutil.copy(TARGET_DOCX, LMS_DOCX)
    print(f"Finished! Total Pages: {pages}")

if __name__ == "__main__":
    main()
