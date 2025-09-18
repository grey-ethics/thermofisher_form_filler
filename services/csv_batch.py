"""
services/csv_batch.py
---------------------
"""
import os
import csv
from services.storage import relpath_from_output, safe_filename
from services.validation import normalize_project_level, parse_bool
from services.word_fill import fill_and_export

def process_csv(csv_file, docx_template: str, full_docx_template: str, out_dir: str, export_docx: bool = True):
    csv_file.stream.seek(0)
    reader = csv.DictReader(line.decode("utf-8") if isinstance(line, bytes) else line for line in csv_file.stream)

    items = []
    pdf_relpaths = []
    count = 0

    for row in reader:
        count += 1
        company = (row.get("company_id") or f"row_{count}").strip()
        project_level = normalize_project_level(row.get("project_level_dropdown"))

        ticks = {}
        for r in range(16, 21):
            for c in range(2, 6):
                colname = f"device_r{r}_c{c}"
                ticks[f"glyph_r{r}_c{c}"] = parse_bool(row.get(colname))

        out_base = safe_filename(f"{company}")
        result = fill_and_export(
            docx_template=docx_template,
            full_docx_template=full_docx_template,
            mapping={"projectLevel": project_level, "ticks": ticks},
            out_dir=out_dir,
            out_basename=out_base,
            export_docx=export_docx
        )

        item = {
            "company_id": company,
            "pdf_url": f"/download/{result['rel_pdf_path']}"
        }
        pdf_relpaths.append(result["rel_pdf_path"])
        if export_docx and result.get("rel_docx_path"):
            item["docx_url"] = f"/download/{result['rel_docx_path']}"
        items.append(item)

    return {
        "processed": count,
        "items": items,
        "pdf_relpaths": pdf_relpaths
    }
