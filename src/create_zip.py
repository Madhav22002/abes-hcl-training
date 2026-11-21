"""
Create project ZIP archive excluding caches, temporary files, and git artifacts.
"""

import zipfile
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
ZIP_PATH = BASE_DIR / "2503201450085-marketing-campaign-performance-analysis.zip"

EXCLUDE_DIRS = {"__pycache__", ".pytest_cache", ".git", "venv", ".venv"}
EXCLUDE_EXTS = {".pyc", ".pyo", ".pyd", ".tmp", ".bak"}

def make_zip():
    count = 0
    with zipfile.ZipFile(ZIP_PATH, "w", zipfile.ZIP_DEFLATED) as zf:
        for p in BASE_DIR.rglob("*"):
            if p == ZIP_PATH:
                continue
            # Skip excluded dirs
            if any(part in EXCLUDE_DIRS for part in p.parts):
                continue
            if p.suffix in EXCLUDE_EXTS:
                continue
            if p.is_file():
                arcname = p.relative_to(BASE_DIR)
                zf.write(p, arcname)
                count += 1
    size_mb = ZIP_PATH.stat().st_size / (1024 * 1024)
    print(f"Created {ZIP_PATH.name} with {count} files ({size_mb:.2f} MB)")

if __name__ == "__main__":
    make_zip()
