"""
Technical Test Suite: Streamlit Application Smoke and Page Integrity Tests (TC-T-10)
"""

import pytest
from pathlib import Path
from src.config import BASE_DIR

def test_tc_t_10_app_structure_and_pages_existence():
    """TC-T-10: Verify application entrypoint, pages, and config exist."""
    app_py = BASE_DIR / "app.py"
    assert app_py.exists(), "app.py entrypoint is missing!"

    pages_dir = BASE_DIR / "pages"
    assert pages_dir.exists(), "pages/ directory is missing!"

    expected_pages = [
        "1_Trends.py",
        "2_Comparison.py",
        "3_Statistics.py",
        "4_Audience_Creative.py",
        "5_Insights.py",
        "6_Data_Explorer.py",
        "7_Data_Quality.py",
        "8_About.py"
    ]

    for p in expected_pages:
        page_path = pages_dir / p
        assert page_path.exists(), f"Page {p} is missing from pages/ directory!"

def test_tc_t_10_streamlit_apptest_smoke():
    """TC-T-10: Run Streamlit AppTest to verify app startup and rendering across all pages."""
    try:
        from streamlit.testing.v1 import AppTest
        # Test app.py
        at = AppTest.from_file(str(BASE_DIR / "app.py"), default_timeout=15)
        at.run()
        assert not at.exception, f"app.py raised exception: {at.exception}"

        # Test each sub-page
        pages = [
            "1_Trends.py",
            "2_Comparison.py",
            "3_Statistics.py",
            "4_Audience_Creative.py",
            "5_Insights.py",
            "6_Data_Explorer.py",
            "7_Data_Quality.py",
            "8_About.py"
        ]
        for p in pages:
            page_at = AppTest.from_file(str(BASE_DIR / "pages" / p), default_timeout=15)
            page_at.run()
            assert not page_at.exception, f"Page {p} raised an exception: {page_at.exception}"
    except ImportError:
        pytest.skip("streamlit.testing.v1.AppTest not available in current environment.")
