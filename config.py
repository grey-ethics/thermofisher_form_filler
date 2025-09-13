"""
config.py
----------------
Centralized configuration for paths and feature flags.

Pipelines/Modules:
- AppConfig: holds absolute paths to template files, overlay map, and output dir.

Key settings:
- DOCX_TEMPLATE_PATH: Path to the Word template (.docx) used for final fill/export.
- BASE_PDF_PATH: Path to the base PDF used only for PREVIEW in the browser.
- OVERLAY_MAP_PATH: Path to overlay_map.json (normalized coordinates).
- OUTPUT_DIR: Root folder for generated files (PDF/DOCX/ZIP).
- EXPORT_DOCX: Whether to also return DOCX (besides PDF) on export.
"""

import os

ROOT_DIR = os.path.dirname(os.path.abspath(__file__))

class AppConfig:
    # === Paths you can customize if needed ===
    DOCX_TEMPLATE_PATH = os.environ.get(
        "DOCX_TEMPLATE_PATH",
        os.path.join(ROOT_DIR, "reference_template.docx")
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

    # === Features ===
    EXPORT_DOCX = True  # return DOCX links in addition to PDF
