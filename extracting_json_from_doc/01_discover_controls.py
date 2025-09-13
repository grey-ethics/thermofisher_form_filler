# discover_controls.py
import json
import os
import win32com.client as com

TEMPLATE = r"C:\Users\K Santosh Kumar\Desktop\HEALTHARK\04_thermofisher\reference_template.docx"
OUT = r"C:\Users\K Santosh Kumar\Desktop\HEALTHARK\04_thermofisher\01_template_skeleton.json"

def discover(template_path):
    word = com.Dispatch("Word.Application")
    word.Visible = False
    doc = word.Documents.Open(template_path)
    controls = []

    for i in range(1, doc.ContentControls.Count + 1):
        cc = doc.ContentControls.Item(i)
        # use try/except for safety
        try:
            ctype = cc.Type  # numeric constant
        except:
            ctype = None

        # get readable type using constants
        consts = com.constants
        type_str = "unknown"
        try:
            if ctype == consts.wdContentControlText:
                type_str = "text"
            elif ctype == consts.wdContentControlRichText:
                type_str = "richtext"
            elif ctype == consts.wdContentControlCheckBox:
                type_str = "checkbox"
            elif ctype == consts.wdContentControlComboBox:
                type_str = "combo"
            elif ctype == consts.wdContentControlDropDownList:
                type_str = "dropdown"
            elif ctype == consts.wdContentControlDate:
                type_str = "date"
        except Exception:
            pass

        entry = {
            "index": i,
            "tag": getattr(cc, "Tag", "") or "",
            "title": getattr(cc, "Title", "") or "",
            "type": type_str,
            "placeholder_text": (cc.Range.Text or "").strip(),
        }

        if type_str in ("dropdown", "combo"):
            choices = []
            for j in range(1, cc.DropdownListEntries.Count + 1):
                e = cc.DropdownListEntries.Item(j)
                choices.append({"text": e.Text, "value": getattr(e, "Value", e.Text)})
            entry["choices"] = choices

        controls.append(entry)

    doc.Close(False)
    word.Quit()
    return controls

if __name__ == "__main__":
    ctrls = discover(TEMPLATE)
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump({"template": os.path.basename(TEMPLATE), "controls": ctrls}, f, indent=2, ensure_ascii=False)
    print("Wrote skeleton:", OUT)
