"""
services/word_fill.py
---------------------
Word COM automation (Windows + installed Word required).

Pipelines/Functions:
- fill_and_export(docx_template, full_docx_template, mapping, out_dir, out_basename, export_docx=True)
    Orchestrates: open single-page template -> set dropdown (cc_2) -> set device ticks
    -> paste that page over page 3 of full_docx_template -> save DOCX/PDF -> return paths.

- _open_word() / _quit_word(app): manage Word app lifecycle
- _open_doc(app, path) / _close_doc(doc)
- _find_cc_in_cell(doc, table_index, row, col): locate content-control in a specific cell
- _set_dropdown_value(cc, value): choose an entry by Text
- _set_device_cell_tick(...): write ☐/☒ (U+2610/U+2612)
- _replace_page3_with_doc_content(app, src_doc, full_path): returns opened full doc after replacement
"""

import os
import threading
import pythoncom  # required for COM in multithreaded environments
import win32com.client as com

from services.storage import relpath_from_output

# Global mutex to serialize COM access
_WORD_LOCK = threading.Lock()

# Word constants
_wdFormatPDF = 17           # SaveAs2 format for PDF
_wdGoToPage = 1             # wdGoToPage
_wdGoToAbsolute = 1         # wdGoToAbsolute

CHECKED_CHAR = "☒"          # U+2612: box with X  (required)
UNCHECKED_CHAR = "☐"        # U+2610: empty box

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
    try:
        entries = cc.DropdownListEntries
        for j in range(1, entries.Count + 1):
            e = entries.Item(j)
            if e.Text == value or getattr(e, "Value", None) == value:
                e.Select()
                return
    except Exception:
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
    Replace the first ballot box with ☒ when checked, else ensure ☐.
    Previously used ☑; now we exclusively use U+2612 (☒).
    Also unchecking replaces any ☒/☑ back to ☐.
    """
    rng = _cell_range(doc, table_index, row, col)
    txt = rng.Text
    core = _strip_cell_end(txt)
    if not core:
        return

    new_core = core
    if checked:
        if CHECKED_CHAR in core:
            new_core = core
        elif "☑" in core:
            new_core = core.replace("☑", CHECKED_CHAR, 1)
        elif UNCHECKED_CHAR in core:
            new_core = core.replace(UNCHECKED_CHAR, CHECKED_CHAR, 1)
        else:
            new_core = CHECKED_CHAR + core
    else:
        # Ensure unchecked box
        new_core = (core
                    .replace(CHECKED_CHAR, UNCHECKED_CHAR)
                    .replace("☑", UNCHECKED_CHAR))

    try:
        if len(txt) >= 2 and ord(txt[-1]) == 7:
            rng.Text = new_core + txt[-2:]
        else:
            rng.Text = new_core
    except Exception:
        pass

def _replace_page3_with_doc_content(app, src_doc, full_path: str):
    """
    Copies src_doc.Content (single page) and pastes it over page 3 of the full template.
    Returns the opened 'full' document object (caller is responsible to close it).
    """
    # Copy source page to clipboard
    src_doc.Content.Copy()

    full_doc = _open_doc(app, full_path)
    full_doc.Activate()
    sel = app.Selection

    # Go to page 3 start
    sel.GoTo(What=_wdGoToPage, Which=_wdGoToAbsolute, Count=3)
    start = sel.Start

    # Try to get start of page 4; if not present, use end of document
    try:
        sel.GoTo(What=_wdGoToPage, Which=_wdGoToAbsolute, Count=4)
        end = sel.Start
    except Exception:
        end = full_doc.Content.End

    # Replace that range with the source page
    rng = full_doc.Range(Start=start, End=end)
    rng.Select()
    app.Selection.Paste()

    return full_doc

def fill_and_export(
    docx_template: str,
    full_docx_template: str,
    mapping: dict,
    out_dir: str,
    out_basename: str,
    export_docx: bool = True
) -> dict:
    """
    Main orchestrator for a single document fill + export.

    mapping example:
    {
      "projectLevel": "L2",         # or None for placeholder
      "ticks": { "glyph_r16_c2": true, ... }  # may include r16..r20
    }

    We fill the single-page template, then paste that page over page 3 of the full template,
    and save PDF/DOCX from the full template.
    """
    os.makedirs(out_dir, exist_ok=True)
    abs_docx = os.path.join(out_dir, f"{out_basename}.docx")
    abs_pdf  = os.path.join(out_dir, f"{out_basename}.pdf")

    with _WORD_LOCK:
        app = _open_word()
        try:
            # 1) Open single-page working template and fill it
            doc = _open_doc(app, docx_template)

            # Dropdown at Table(1), Row(2), Col(2)
            cc = _find_cc_in_cell(doc, table_index=1, row=2, col=2)
            if cc:
                _set_dropdown_value(cc, mapping.get("projectLevel"))

            # Device ticks (IDs like glyph_r16_c2, glyph_r17_c5, ...)
            ticks = mapping.get("ticks") or {}
            for glyph_id, checked in ticks.items():
                try:
                    parts = glyph_id.replace("glyph_r", "").split("_c")
                    row = int(parts[0]); col = int(parts[1])
                except Exception:
                    continue
                _set_device_cell_tick(doc, table_index=1, row=row, col=col, checked=bool(checked))

            # 2) Paste the filled page over page 3 of the full template
            full_doc = _replace_page3_with_doc_content(app, doc, full_docx_template)

            # 3) Save (from the full_doc)
            if export_docx:
                full_doc.SaveAs2(abs_docx)                # DOCX
            full_doc.SaveAs2(abs_pdf, FileFormat=_wdFormatPDF)  # PDF

            # 4) Close docs
            _close_doc(full_doc)
            _close_doc(doc)

        finally:
            _quit_word(app)

    rel_pdf = relpath_from_output(abs_pdf)
    result = {"rel_pdf_path": rel_pdf}
    if export_docx:
        rel_docx = relpath_from_output(abs_docx)
        result["rel_docx_path"] = rel_docx
    return result
