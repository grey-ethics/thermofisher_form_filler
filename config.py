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
    DOCX_TEMPLATE_PATH = os.environ.get(
        "DOCX_TEMPLATE_PATH",
        os.path.join(ROOT_DIR, "reference_template.docx")           # single-page working template
    )
    FULL_DOCX_TEMPLATE_PATH = os.environ.get(
        "FULL_DOCX_TEMPLATE_PATH",
        os.path.join(ROOT_DIR, "reference_template_full.docx")      # FULL document (will receive page 3)
    )
    BASE_PDF_PATH = os.environ.get(
        "BASE_PDF_PATH",
        os.path.join(ROOT_DIR, "static", "pdf", "reference_template.pdf")
    )
    OVERLAY_MAP_PATH = os.environ.get(
        "OVERLAY_MAP_PATH",
        os.path.join(ROOT_DIR, "overlay_map.json")
    )
    OUTPUT_DIR = os.environ.get(
        "OUTPUT_DIR",
        os.path.join(ROOT_DIR, "output")
    )

    # Features
    EXPORT_DOCX = True
