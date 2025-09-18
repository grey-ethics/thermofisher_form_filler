# dump_doc_controls.py
# Usage: python dump_doc_controls.py <file.docx>
import sys
import os
import win32com.client as com
import json

def dump(path):
    word = com.Dispatch("Word.Application")
    word.Visible = False
    doc = word.Documents.Open(os.path.abspath(path))
    try:
        out = {"file": path, "content_controls": [], "formfields": []}
        consts = com.constants
        # ContentControls
        for i in range(1, doc.ContentControls.Count + 1):
            cc = doc.ContentControls.Item(i)
            try:
                ctype = cc.Type
            except Exception:
                ctype = None
            item = {
                "index": i,
                "type_raw": ctype,
                "type_name": getattr(com.constants, f"wdContentControl{''}", "unknown"),
                "tag": getattr(cc, "Tag", ""),
                "title": getattr(cc, "Title", ""),
                "placeholder_text": getattr(cc, "PlaceholderText", "") if hasattr(cc, "PlaceholderText") else "",
                "text": (cc.Range.Text or "").strip(),
            }
            # helpful flags for dropdown/checkbox
            try:
                item["is_checkbox"] = bool(getattr(cc, "Checked", None) is not None)
            except Exception:
                item["is_checkbox"] = False
            try:
                item["dropdown_entries_count"] = cc.DropdownListEntries.Count if hasattr(cc, "DropdownListEntries") else 0
            except Exception:
                item["dropdown_entries_count"] = 0
            out["content_controls"].append(item)

        # Legacy FormFields
        for i in range(1, doc.FormFields.Count + 1):
            ff = doc.FormFields.Item(i)
            try:
                ftype = ff.Type
            except Exception:
                ftype = None
            entry = {
                "index": i,
                "type_raw": ftype,
                "name": ff.Name if hasattr(ff, "Name") else "",
                "result": getattr(ff, "Result", ""),
            }
            # check for checkbox-type legacy field
            try:
                if ftype == com.constants.wdFieldFormCheckBox:
                    entry["is_legacy_checkbox"] = True
                    entry["checked"] = bool(ff.CheckBox.Value)
                else:
                    entry["is_legacy_checkbox"] = False
            except Exception:
                pass
            out["formfields"].append(entry)

        print(json.dumps(out, indent=2, ensure_ascii=False))
    finally:
        doc.Close(False)
        word.Quit()

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python dump_doc_controls.py <file.docx>")
        sys.exit(1)
    dump(sys.argv[1])
