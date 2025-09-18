from flask import Flask, render_template, request, jsonify, send_file, abort, send_from_directory
import os
import io
import json
import tempfile
from datetime import datetime

from config import AppConfig
from services.storage import ensure_dirs
from services.extract_input import extract_and_map

# NEW: services that do the page-3 replacement
from services.page3_fill_com import fill_page3_template_with_snapshot
from services.word_com_replace import replace_docx_page3_with_file, docx_to_pdf
from services.pdf_replace import replace_pdf_page


def create_app():
    app = Flask(__name__, static_folder='static', template_folder='templates')
    ensure_dirs()

    PAGE3_TPL = os.path.join("static", "docx", "page3.tpl.docx")

    def _is_docx(name: str) -> bool:
        return name.lower().endswith(".docx")

    def _is_pdf(name: str) -> bool:
        return name.lower().endswith(".pdf")

    @app.route("/")
    def index():
        return render_template("index.html")

    @app.route("/overlay-map")
    def overlay_map():
        try:
            with open(AppConfig.OVERLAY_MAP_PATH, "r", encoding="utf-8") as fh:
                data = json.load(fh)
            return jsonify(data)
        except Exception as e:
            return jsonify({"error": f"Failed to read overlay map: {e}"}), 500

    @app.route("/extract", methods=["POST"])
    def extract_from_reference():
        """
        Upload field: file (.docx)
        Reads the Regulatory Document and returns { medical, lines[], ticks{} } to pre-fill UI.
        """
        f = request.files.get("file")
        if not f:
            return jsonify({"error": "Upload the .docx as field 'file'"}), 400
        try:
            out = extract_and_map(f, AppConfig.OUTPUT_DIR)
            return jsonify(out)
        except Exception as e:
            return jsonify({"error": f"Extract failed: {e}"}), 500

    @app.route("/download", methods=["POST"])
    def download_filled():
        """
        Single endpoint:
          - template_file: required (.docx or .pdf) — the user's Template Document
          - snapshot: JSON string with at least:
              {
                "projectLevel": "L1" | "L2" | ...,
                "capaAssociated": "Yes" | "No",          # optional
                "ticks": { "glyph_r16_c2": true, ... }   # grid selections
              }

        Behavior:
          - If uploaded template is DOCX:
                fill 1-page page3 from static/docx/page3.tpl.docx
                paste over page 3 of uploaded DOCX
                return DOCX
          - If uploaded template is PDF:
                fill 1-page page3 DOCX -> convert to PDF
                swap into page index 2 (the 3rd page) of uploaded PDF
                return PDF
        """
        tf = request.files.get("template_file")
        if not tf or not (_is_docx(tf.filename) or _is_pdf(tf.filename)):
            return jsonify({"error": "Upload a .docx or .pdf as 'template_file'"}), 400

        try:
            snapshot_raw = request.form.get("snapshot", "{}")
            snapshot = json.loads(snapshot_raw) if snapshot_raw else {}
        except Exception:
            snapshot = {}

        if not os.path.exists(PAGE3_TPL):
            return jsonify({"error": f"Missing template for page 3: {PAGE3_TPL}"}), 500

        with tempfile.TemporaryDirectory() as tmpdir:
            # Save uploaded file
            src_path = os.path.join(tmpdir, tf.filename)
            tf.save(src_path)

            # Build the filled single-page DOCX for page 3
            page3_filled_docx = os.path.join(tmpdir, "page3_filled.docx")
            try:
                fill_page3_template_with_snapshot(PAGE3_TPL, page3_filled_docx, snapshot)
            except Exception as e:
                return jsonify({"error": f"Failed to fill page3 template via Word: {e}"}), 500

            if _is_docx(tf.filename):
                # Paste page 3 over the uploaded DOCX and return DOCX
                out_docx = os.path.join(tmpdir, "output.docx")
                try:
                    replace_docx_page3_with_file(src_path, page3_filled_docx, out_docx)
                except Exception as e:
                    return jsonify({"error": f"DOCX page replacement failed: {e}"}), 500

                return send_file(
                    out_docx,
                    as_attachment=True,
                    download_name=f"output_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx",
                    mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

            else:
                # Convert our single-page DOCX to a single-page PDF, then swap into uploaded PDF
                page3_filled_pdf = os.path.join(tmpdir, "page3_filled.pdf")
                try:
                    docx_to_pdf(page3_filled_docx, page3_filled_pdf)
                except Exception as e:
                    return jsonify({"error": f"Failed to convert page3 DOCX to PDF: {e}"}), 500

                out_pdf = os.path.join(tmpdir, "output.pdf")
                try:
                    replace_pdf_page(src_path, page3_filled_pdf, out_pdf, replace_index=2)
                except Exception as e:
                    return jsonify({"error": f"PDF page replacement failed: {e}"}), 500

                return send_file(
                    out_pdf,
                    as_attachment=True,
                    download_name=f"output_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
                    mimetype="application/pdf"
                )

    # (Optional) Serve generated files by relative path if you use AppConfig.OUTPUT_DIR elsewhere
    @app.route("/download/<path:relpath>")
    def download_output(relpath):
        base = os.path.abspath(AppConfig.OUTPUT_DIR)
        abs_path = os.path.abspath(os.path.join(base, relpath))
        if not abs_path.startswith(base) or not os.path.exists(abs_path):
            abort(404)
        directory = os.path.dirname(abs_path)
        filename = os.path.basename(abs_path)
        return send_from_directory(directory, filename, as_attachment=True)

    return app


if __name__ == "__main__":
    app = create_app()
    # Word COM requires a desktop session; run under a user context on Windows.
    app.run(host="0.0.0.0", port=5000, debug=False)
