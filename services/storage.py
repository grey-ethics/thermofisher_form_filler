"""
services/storage.py
-------------------
Filesystem helpers:
- ensure_dirs(): create OUTPUT_DIR if missing
- make_batch_folder(batch_id): create output subfolder per batch
- safe_filename(name): sanitize base names
- zip_outputs(pdf_relpaths, out_dir, zip_name): zip given files and return relative path
"""

import os
import re
import zipfile
from config import AppConfig

SAFE_RE = re.compile(r"[^A-Za-z0-9._-]+")

def ensure_dirs():
    os.makedirs(AppConfig.OUTPUT_DIR, exist_ok=True)

def make_batch_folder(batch_id: str) -> str:
    """
    Create and return absolute path to batch subfolder inside OUTPUT_DIR.
    Also return the path for app use.
    """
    ensure_dirs()
    path = os.path.join(AppConfig.OUTPUT_DIR, batch_id)
    os.makedirs(path, exist_ok=True)
    return path

def safe_filename(name: str) -> str:
    """
    Sanitize filenames to avoid filesystem issues.
    """
    base = SAFE_RE.sub("_", name.strip())
    return base or "file"

def relpath_from_output(abs_path: str) -> str:
    """
    Convert absolute path inside OUTPUT_DIR to a relative path for /download route.
    """
    base = os.path.abspath(AppConfig.OUTPUT_DIR)
    return os.path.relpath(os.path.abspath(abs_path), base)

def zip_outputs(pdf_relpaths, out_dir, zip_name="batch.zip") -> str:
    """
    Create a ZIP of given relative PDF paths (relative to OUTPUT_DIR).
    Returns the new ZIP file's relative path (for /download).
    """
    zip_abspath = os.path.join(out_dir, zip_name)
    with zipfile.ZipFile(zip_abspath, "w", compression=zipfile.ZIP_DEFLATED) as z:
        for rel in pdf_relpaths:
            abs_f = os.path.join(AppConfig.OUTPUT_DIR, rel)
            if os.path.exists(abs_f):
                z.write(abs_f, arcname=os.path.basename(abs_f))
    return relpath_from_output(zip_abspath)
