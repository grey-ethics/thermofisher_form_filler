"""
services/storage.py
-------------------
Filesystem helpers:
- ensure_dirs(): create OUTPUT_DIR if missing
- relpath_from_output(abs_path): relative path for /download route
"""

import os
from config import AppConfig

def ensure_dirs():
    os.makedirs(AppConfig.OUTPUT_DIR, exist_ok=True)

def relpath_from_output(abs_path: str) -> str:
    base = os.path.abspath(AppConfig.OUTPUT_DIR)
    rel = os.path.relpath(os.path.abspath(abs_path), base)
    return rel.replace(os.sep, "/")
