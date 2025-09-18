# services/page3_fill_com.py
import os
import pythoncom
import win32com.client as win32

WD_FORMAT_DOCX = 16  # wdFormatXMLDocument

def _start_word():
    pythoncom.CoInitialize()
    app = win32.gencache.EnsureDispatch("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    return app

def _stop_word(app):
    try:
        app.Quit()
    finally:
        pythoncom.CoUninitialize()

def fill_page3_template_with_snapshot(tpl_path: str, out_path: str, snapshot: dict):
    """
    Open the 1-page page3.tpl.docx, fill values using simple cell-based rules,
    and save a new 1-page DOCX to out_path.
    Assumptions (based on your analysis):
      - Table 1 contains:
          (row 2, col 2) Project Level dropdown -> set to snapshot["projectLevel"]
          (row 2, col 3) Yes/No -> set to snapshot["capaAssociated"] ("Yes"/"No") if provided
          (rows 16..20, cols 2..5) single glyph "☐" cells -> set to "☒" if ticks[id] is True
            where id is f"glyph_r{r}_c{c}"
    """
    ticks = (snapshot or {}).get("ticks") or {}
    project_level = (snapshot or {}).get("projectLevel")
    capa = (snapshot or {}).get("capaAssociated")  # optional ("Yes"/"No")

    tpl_path = os.path.abspath(tpl_path)
    out_path = os.path.abspath(out_path)

    app = _start_word()
    try:
        doc = app.Documents.Open(tpl_path)

        if doc.Tables.Count >= 1:
            tbl = doc.Tables.Item(1)

            # Project Level @ (2,2)
            try:
                if project_level:
                    tbl.Cell(2, 2).Range.Text = str(project_level)
            except Exception:
                pass

            # CAPA Associated? @ (2,3) - if you don't use this, it's safe to skip
            try:
                if capa in ("Yes", "No"):
                    tbl.Cell(2, 3).Range.Text = capa
            except Exception:
                pass

            # Device/Application matrix: rows 16..20, cols 2..5
            for r in range(16, 21):       # 16,17,18,19,20
                for c in range(2, 6):     # 2,3,4,5
                    glyph_id = f"glyph_r{r}_c{c}"
                    val = "☒" if ticks.get(glyph_id) else "☐"
                    try:
                        # replace the single-glyph cell content
                        cell = tbl.Cell(r, c)
                        cell.Range.Text = val
                    except Exception:
                        # If the exact cell isn’t present, ignore silently
                        pass

        # Save the 1-page filled page
        doc.SaveAs2(out_path, FileFormat=WD_FORMAT_DOCX)
        doc.Close(False)
    finally:
        _stop_word(app)
