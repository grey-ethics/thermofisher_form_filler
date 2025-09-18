"""
config.py
----------------
Centralized configuration for paths and feature flags.
"""

import os
from dotenv import load_dotenv

ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
load_dotenv(os.path.join(ROOT_DIR, ".env"))

class AppConfig:
    # Static PDF used in the right-side viewer (always the same)
    BASE_PDF_PATH = os.environ.get(
        "BASE_PDF_PATH",
        os.path.join(ROOT_DIR, "static", "pdf", "reference_template.pdf")
    )

    # Overlay definition for on-canvas ticks/dropdown placement
    OVERLAY_MAP_PATH = os.environ.get(
        "OVERLAY_MAP_PATH",
        os.path.join(ROOT_DIR, "overlay_map.json")
    )

    # Where /extract helper may write temp outputs (e.g., first page docx)
    OUTPUT_DIR = os.environ.get(
        "OUTPUT_DIR",
        os.path.join(ROOT_DIR, "output")
    )
