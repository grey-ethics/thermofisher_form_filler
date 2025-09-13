"""
services/word_fill.py
---------------------
Word COM automation (Windows + installed Word required).

Pipelines/Functions:
- fill_and_export(docx_template, mapping, out_dir, out_basename, export_docx=True)
    Orchestrates: open template -> set dropdown (cc_2) -> set device ticks (glyph_r..)
    -> save DOCX (optional) -> export PDF -> return relative paths.

- _open_word() / _quit_word(app): manage Word app lifecycle
- _open_doc(app, path) / _close_doc(doc)
- _find_cc_in_cell(doc, table_index, row, col): locate content-control in a specific cell
- _set_dropdown_value(cc, value): choose an entry by Text
- _set_device_cell_tick(doc, table_index, row, col, checked): replace ☐/<U+2610> with ☑/<U+2611>

Notes:
- This implementation assumes the dropdown control lives in Table(1) Row(2) Col(2),
  per your skeleton. Update if template changes or (better) tag the control and
  locate by Tag.
"""

import os
import threading
import pythoncom  # required for COM in multithreaded environments
import win32com.client as com

from services.storage import relpath_from_output

# Global mutex to serialize COM access
_WORD_LOCK = threading.Lock()

# Word constants (lazy init)
_wdFormatPDF = 17  # SaveAs2 format for PDF
# Alternatively: ExportAsFixedFormat Type=17 (wdExportFormatPDF)
# We'll use SaveAs2 for simplicity.

def _open_word():
    pythoncom.CoInitialize()
    app = com.DispatchEx("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    return app

def _quit_word(app):
    try:
        app.Quit()
    finally:
        pythoncom.CoUninitialize()

def _open_doc(app, path: str):
    return app.Documents.Open(os.path.abspath(path))

def _close_doc(doc):
    doc.Close(SaveChanges=False)

def _cell_range(doc, table_index: int, row: int, col: int):
    tbl = doc.Tables.Item(table_index)
    cell = tbl.Rows.Item(row).Cells.Item(col)
    return cell.Range  # includes end-of-cell marker characters

def _find_cc_in_cell(doc, table_index: int, row: int, col: int):
    """
    Return the first ContentControl whose Range is fully inside the given table cell.
    """
    cell_rng = _cell_range(doc, table_index, row, col)
    cstart, cend = cell_rng.Start, cell_rng.End
    for i in range(1, doc.ContentControls.Count + 1):
        cc = doc.ContentControls.Item(i)
        try:
            rstart, rend = cc.Range.Start, cc.Range.End
            if rstart >= cstart and rend <= cend:
                return cc
        except Exception:
            continue
    return None

def _set_dropdown_value(cc, value: str | None):
    """
    Select an entry in a dropDown/combobox content control by its visible Text.
    If value is None (placeholder), we leave it untouched.
    """
    if not value:
        return
    # Some controls expose DropdownListEntries; others via ComboBoxEntries
    # COM normalizes under DropdownListEntries for both dropdown/combobox.
    try:
        entries = cc.DropdownListEntries
        for j in range(1, entries.Count + 1):
            e = entries.Item(j)
            if e.Text == value or getattr(e, "Value", None) == value:
                e.Select()
                return
    except Exception:
        # Fallback: set range text (less ideal)
        try:
            cc.Range.Text = value
        except Exception:
            pass

def _strip_cell_end(text: str) -> str:
    """
    Word cell Range.Text includes end-of-cell markers (chr 13+7).
    We avoid replacing inside them by working on text[:-2] when needed.
    """
    if not text:
        return ""
    if len(text) >= 2 and ord(text[-1]) == 7:
        return text[:-2]
    return text

def _set_device_cell_tick(doc, table_index: int, row: int, col: int, checked: bool):
    """
    Replace first ballot box ☐ (U+2610) with ☑ (U+2611) if checked, else ensure ☐.
    Only for cells that contain a single glyph in your Device grid.
    """
    rng = _cell_range(doc, table_index, row, col)
    txt = rng.Text
    core = _strip_cell_end(txt)
    if not core:
        return

    if checked:
        # Replace first ☐ with ☑ (if already ☑, keep)
        if "☑" in core:
            new_core = core
        elif "☐" in core:
            new_core = core.replace("☐", "☑", 1)
        else:
            new_core = "☑" + core  # fallback: prepend a check
    else:
        # Ensure it's ☐ (uncheck if previously ☑)
        if "☐" in core:
            new_core = core
        elif "☑" in core:
            new_core = core.replace("☑", "☐", 1)
        else:
            new_core = "☐"

    # Write back preserving end-of-cell markers
    try:
        if len(txt) >= 2 and ord(txt[-1]) == 7:
            rng.Text = new_core + txt[-2:]
        else:
            rng.Text = new_core
    except Exception:
        pass

def fill_and_export(docx_template: str, mapping: dict, out_dir: str, out_basename: str, export_docx: bool = True) -> dict:
    """
    Main orchestrator for a single document fill + export.

    mapping example:
    {
      "projectLevel": "L2",         # or None for placeholder
      "ticks": { "glyph_r16_c2": true, ... }
    }

    Returns:
    {
      "rel_pdf_path": "YYYYMMDD_HHMMSS/file.pdf",
      "rel_docx_path": "YYYYMMDD_HHMMSS/file.docx"  # if export_docx=True
    }
    """
    os.makedirs(out_dir, exist_ok=True)
    abs_docx = os.path.join(out_dir, f"{out_basename}.docx")
    abs_pdf  = os.path.join(out_dir, f"{out_basename}.pdf")

    with _WORD_LOCK:
        app = _open_word()
        try:
            doc = _open_doc(app, docx_template)

            # 1) Dropdown at Table(1), Row(2), Col(2)
            cc = _find_cc_in_cell(doc, table_index=1, row=2, col=2)
            if cc:
                _set_dropdown_value(cc, mapping.get("projectLevel"))

            # 2) Device ticks (IDs like glyph_r16_c2)
            ticks = mapping.get("ticks") or {}
            for glyph_id, checked in ticks.items():
                # parse row/col from id "glyph_r<row>_c<col>"
                try:
                    parts = glyph_id.replace("glyph_r", "").split("_c")
                    row = int(parts[0]); col = int(parts[1])
                except Exception:
                    continue
                _set_device_cell_tick(doc, table_index=1, row=row, col=col, checked=bool(checked))

            # 3) Save DOCX (optional) + Export PDF
            if export_docx:
                doc.SaveAs2(abs_docx)  # Save as DOCX
            # SaveAs2 to PDF (wdFormatPDF=17), or ExportAsFixedFormat
            doc.SaveAs2(abs_pdf, FileFormat=_wdFormatPDF)

            _close_doc(doc)
        finally:
            _quit_word(app)

    rel_pdf = relpath_from_output(abs_pdf)
    result = {"rel_pdf_path": rel_pdf}
    if export_docx:
        rel_docx = relpath_from_output(abs_docx)
        result["rel_docx_path"] = rel_docx
    return result
