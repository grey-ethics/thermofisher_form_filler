# discover_controls_debug.py
import json
import os
import win32com.client as com
import traceback

TEMPLATE = r"C:\Users\K Santosh Kumar\Desktop\HEALTHARK\04_thermofisher\reference_template.docx"
OUT = r"C:\Users\K Santosh Kumar\Desktop\HEALTHARK\04_thermofisher\03_template_skeleton_debug.json"

def get_table_cell_coords(doc, rng_start, rng_end):
    # find which table contains a range by comparing numeric Range.Start/End
    for t_idx in range(1, doc.Tables.Count+1):
        tbl = doc.Tables.Item(t_idx)
        tstart = tbl.Range.Start
        tend = tbl.Range.End
        if rng_start >= tstart and rng_start <= tend:
            # inside this table: find row and column by iterating cells
            for r in range(1, tbl.Rows.Count+1):
                row = tbl.Rows.Item(r)
                for c in range(1, row.Cells.Count+1):
                    cell = row.Cells.Item(c)
                    if rng_start >= cell.Range.Start and rng_start <= cell.Range.End:
                        return (t_idx, r, c)
            return (t_idx, None, None)
    return (None, None, None)

def discover(template_path):
    word = com.Dispatch("Word.Application")
    word.Visible = False
    doc = word.Documents.Open(template_path)
    controls = []
    consts = com.constants

    for i in range(1, doc.ContentControls.Count + 1):
        cc = doc.ContentControls.Item(i)
        try:
            raw_type = cc.Type
        except Exception:
            raw_type = None

        # map raw_type to readable string if possible
        type_map = {}
        try:
            type_map = {
                consts.wdContentControlText: "text",
                consts.wdContentControlRichText: "richtext",
                consts.wdContentControlCheckBox: "checkbox",
                consts.wdContentControlComboBox: "combo",
                consts.wdContentControlDropDownList: "dropdown",
                consts.wdContentControlDate: "date"
            }
        except Exception:
            # if constants fetch fails, ignore (we will show raw)
            pass

        type_str = type_map.get(raw_type, "unknown")

        # try to capture dropdown entries
        choices = []
        try:
            if hasattr(cc, "DropdownListEntries"):
                for j in range(1, cc.DropdownListEntries.Count+1):
                    e = cc.DropdownListEntries.Item(j)
                    choices.append({"text": e.Text, "value": getattr(e,"Value", e.Text)})
        except Exception:
            pass

        # Range start/end
        try:
            rng_start = cc.Range.Start
            rng_end = cc.Range.End
        except Exception:
            rng_start = None
            rng_end = None

        # table coords
        try:
            table_idx, row_idx, col_idx = get_table_cell_coords(doc, rng_start, rng_end)
        except Exception:
            table_idx, row_idx, col_idx = (None, None, None)

        entry = {
            "index": i,
            "raw_type": raw_type,
            "type": type_str,
            "tag": (cc.Tag if hasattr(cc, "Tag") else ""),
            "title": (cc.Title if hasattr(cc, "Title") else ""),
            "text": (cc.Range.Text if hasattr(cc, "Range") else ""),
            "range_start": rng_start,
            "range_end": rng_end,
            "table_index": table_idx,
            "table_row": row_idx,
            "table_col": col_idx,
            "choices": choices
        }
        controls.append(entry)

    try:
        doc.Close(False)
        word.Quit()
    except Exception:
        pass

    return controls

if __name__ == "__main__":
    try:
        ctrls = discover(TEMPLATE)
        with open(OUT, "w", encoding="utf-8") as f:
            json.dump({"template": os.path.basename(TEMPLATE), "controls": ctrls}, f, indent=2, ensure_ascii=False)
        print("Wrote skeleton:", OUT)
        print("Discovered", len(ctrls), "controls. Sample:")
        if ctrls:
            import pprint
            pprint.pprint(ctrls[:6])
    except Exception as e:
        print("Error during discovery:", e)
        traceback.print_exc()
