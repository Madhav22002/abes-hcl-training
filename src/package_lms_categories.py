"""
Package all project deliverables into the 3 exact LMS submission categories:
1. Project Report
2. Data Files
3. Other Documents
"""

import shutil
import zipfile
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
DOCS_DIR = BASE_DIR / "docs"
DATA_DIR = BASE_DIR / "data"
SUBMISSION_DIR = BASE_DIR / "submission"

LMS_DIR = BASE_DIR / "LMS_Uploads"
REPORT_DIR = LMS_DIR / "1_Project_Report"
DATA_FILES_DIR = LMS_DIR / "2_Data_Files"
OTHER_DOCS_DIR = LMS_DIR / "3_Other_Documents"

for d in [REPORT_DIR, DATA_FILES_DIR, OTHER_DOCS_DIR]:
    d.mkdir(parents=True, exist_ok=True)

def zip_folder(folder_path, output_zip_path):
    with zipfile.ZipFile(output_zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for file_path in folder_path.rglob("*"):
            if file_path.is_file():
                zf.write(file_path, file_path.relative_to(folder_path))
    print(f"Created {output_zip_path.name} ({output_zip_path.stat().st_size / (1024*1024):.2f} MB)")

def main():
    print("Organizing 1. Project Report files...")
    # Reports: Project Document, Presentation, GitHub Report, Complete Code (both PDF and DOCX/PPTX)
    report_files = [
        # PDFs from submission
        SUBMISSION_DIR / "1_Project_Document_2503201450085.pdf",
        SUBMISSION_DIR / "2_Project_Presentation_2503201450085.pdf",
        SUBMISSION_DIR / "3_GitHub_Report_2503201450085.pdf",
        SUBMISSION_DIR / "4_Complete_Code_2503201450085.pdf",
        # Source Word / PPTX files from docs
        DOCS_DIR / "2503201450085_Project_Document.docx",
        DOCS_DIR / "2503201450085_Presentation.pptx",
        DOCS_DIR / "GitHub_Report.docx",
        DOCS_DIR / "Complete_Code.docx"
    ]
    for src in report_files:
        if src.exists():
            shutil.copy2(src, REPORT_DIR / src.name)
            print(f"Copied {src.name} to 1_Project_Report")

    print("\nOrganizing 2. Data Files...")
    # Data files: app.db, excel sheets, CSVs, validation report
    data_items = [
        DATA_DIR / "app.db",
        DATA_DIR / "marketing_campaign_data.xlsx",
        DATA_DIR / "marketing_campaign_data_raw.xlsx",
        DATA_DIR / "validation_report.md",
        DATA_DIR / "campaigns.csv",
        DATA_DIR / "channels.csv",
        DATA_DIR / "audience_segments.csv",
        DATA_DIR / "creatives.csv",
        DATA_DIR / "daily_performance.csv"
    ]
    for src in data_items:
        if src.exists():
            shutil.copy2(src, DATA_FILES_DIR / src.name)
            print(f"Copied {src.name} to 2_Data_Files")

    print("\nOrganizing 3. Other Documents...")
    # Other docs: Steps 1 through 14 DOCX, PDF, and workbooks
    other_docs = [
        # Workbooks
        DOCS_DIR / "04_Traceability_Matrix.xlsx",
        DOCS_DIR / "08_Database_Design.xlsx",
        # Diagrams
        DOCS_DIR / "09_ER_Diagram.png",
        DOCS_DIR / "09_ER_Diagram.mmd",
        BASE_DIR / "BUILD_PLAN.md",
        BASE_DIR / "README.md"
    ]
    # Add step docs docx
    step_doc_names = [
        "01_Project_Understanding.docx",
        "02_HMW_Statements.docx",
        "03_Needs_Analysis.docx",
        "05_BRD.docx",
        "06_PRD.docx",
        "07_TRD.docx",
        "09_ER_Diagram.docx",
        "11_User_Stories.docx",
        "12_UI_Framework.docx",
        "13_Test_Cases.docx",
        "14_BUILD_PLAN.docx"
    ]
    for name in step_doc_names:
        d_path = DOCS_DIR / name
        if d_path.exists():
            other_docs.append(d_path)
        # Also check PDF in submission/5_Project_Documents/
        pdf_path = SUBMISSION_DIR / "5_Project_Documents" / name.replace(".docx", ".pdf")
        if pdf_path.exists():
            other_docs.append(pdf_path)

    for src in other_docs:
        if src.exists():
            shutil.copy2(src, OTHER_DOCS_DIR / src.name)
            print(f"Copied {src.name} to 3_Other_Documents")

    # Create ZIP files for each category
    print("\nCreating Category ZIP archives for LMS Upload...")
    zip_folder(REPORT_DIR, LMS_DIR / "Madhav_Dutta_Project_Report.zip")
    zip_folder(DATA_FILES_DIR, LMS_DIR / "Madhav_Dutta_Data_Files.zip")
    zip_folder(OTHER_DOCS_DIR, LMS_DIR / "Madhav_Dutta_Other_Documents.zip")

    # Create master all-in-one ZIP
    zip_folder(LMS_DIR, BASE_DIR / "Madhav_Dutta_Complete_Project_Submission.zip")
    print("Master package Madhav_Dutta_Complete_Project_Submission.zip created successfully!")

if __name__ == "__main__":
    main()
