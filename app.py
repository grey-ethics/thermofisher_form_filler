"""
app.py
----------------
Flask application entrypoint and routes.

Pipelines & Modules:
- Routes:
  - GET /              -> Render main UI (index.html)
  - GET /overlay-map   -> Return overlay_map.json (normalized coordinates)
  - POST /export       -> Accept single form state; fill DOCX via Word COM; export PDF; return file URLs
  - POST /batch        -> Accept CSV; fill/export per row; return manifest + ZIP link
  - GET /download/<path:relpath> -> Serve generated files from /output

- Uses services:
  - services.word_fill.fill_and_export(mapping, out_basename)
  - services.csv_batch.process_csv(file_storage, out_dir)
  - services.storage helpers for paths/filenames
"""

import os
import json
from datetime import datetime
from flask import Flask, render_template, request, jsonify, send_from_directory, abort

from config import AppConfig
from services import storage
from services.word_fill import fill_and_export
from services.csv_batch import process_csv

app = Flask(__name__, static_folder="static", template_folder="templates")
app.config["MAX_CONTENT_LENGTH"] = 64 * 1024 * 1024  # 64 MB uploads cap

# ---------- ROUTES ----------

@app.get("/")
def index():
    """
    Render the main UI: two blades -> interactive preview & batch upload.
    """
    return render_template("index.html")


@app.get("/overlay-map")
def get_overlay_map():
    """
    Serve overlay_map.json (used by frontend overlay renderer).
    """
    path = AppConfig.OVERLAY_MAP_PATH
    if not os.path.exists(path):
        return jsonify({"error": "overlay_map.json not found", "path": path}), 404
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    return jsonify(data)


@app.post("/export")
def export_single():
    """
    Accepts JSON body:
    {
      "company_id": "optional-id",
      "projectLevel": "L1|L2L|L2|L3L|null-for-placeholder",
      "ticks": { "glyph_r16_c2": true, ..., "glyph_r20_c5": false }
    }

    Returns:
    { "pdf_url": "/download/...", "docx_url": "/download/..." (if enabled) }
    """
    try:
        payload = request.get_json(force=True)
    except Exception:
        return jsonify({"error": "Invalid JSON"}), 400

    company_id = (payload.get("company_id") or "company").strip()
    project_level = payload.get("projectLevel")  # may be None/empty for placeholder
    ticks = payload.get("ticks") or {}

    # Ensure output subfolder
    batch_id = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_dir = storage.make_batch_folder(batch_id)

    out_base = storage.safe_filename(f"{company_id}_{batch_id}")
    result = fill_and_export(
        docx_template=AppConfig.DOCX_TEMPLATE_PATH,
        mapping={"projectLevel": project_level, "ticks": ticks},
        out_dir=out_dir,
        out_basename=out_base,
        export_docx=AppConfig.EXPORT_DOCX
    )

    # Build download URLs relative to /output
    resp = {"pdf_url": f"/download/{result['rel_pdf_path']}"}
    if AppConfig.EXPORT_DOCX and result.get("rel_docx_path"):
        resp["docx_url"] = f"/download/{result['rel_docx_path']}"
    return jsonify(resp)


@app.post("/batch")
def batch():
    """
    Accepts multipart/form-data with CSV file under field 'file'.
    Produces per-row outputs and a ZIP; returns manifest:
    {
      "batch_id": "...",
      "processed": N,
      "zip_url": "/download/...",
      "items": [{ "company_id": "...", "pdf_url": "...", "docx_url": "..." }, ...]
    }
    """
    if "file" not in request.files:
        return jsonify({"error": "CSV file not provided under field 'file'"}), 400

    file = request.files["file"]
    if not file.filename.lower().endswith(".csv"):
        return jsonify({"error": "Please upload a .csv file"}), 400

    batch_id = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_dir = storage.make_batch_folder(batch_id)

    manifest = process_csv(
        csv_file=file,
        docx_template=AppConfig.DOCX_TEMPLATE_PATH,
        out_dir=out_dir,
        export_docx=AppConfig.EXPORT_DOCX
    )

    # Create ZIP of all PDFs
    zip_relpath = storage.zip_outputs(manifest["pdf_relpaths"], out_dir, f"batch_{batch_id}.zip")
    manifest.update({
        "batch_id": batch_id,
        "zip_url": f"/download/{zip_relpath}"
    })

    return jsonify(manifest)


@app.get("/download/<path:relpath>")
def download(relpath: str):
    """
    Serve files from OUTPUT_DIR safely.
    """
    base = AppConfig.OUTPUT_DIR
    abs_path = os.path.abspath(os.path.join(base, relpath))
    # Security: ensure file is inside OUTPUT_DIR
    if not abs_path.startswith(os.path.abspath(base)):
        abort(403)
    if not os.path.exists(abs_path):
        abort(404)
    directory = os.path.dirname(abs_path)
    filename = os.path.basename(abs_path)
    return send_from_directory(directory, filename, as_attachment=True)


# ---------- MAIN ----------

if __name__ == "__main__":
    # Ensure output dir exists
    storage.ensure_dirs()
    app.run(host="127.0.0.1", port=5000, debug=True)
