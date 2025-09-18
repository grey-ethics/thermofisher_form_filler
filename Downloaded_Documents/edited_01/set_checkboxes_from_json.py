# set_checkboxes_from_json.py
import sys, json, os
import win32com.client as com

IN = sys.argv[1]
MAP = sys.argv[2]  # json file: { "tag_or_title": true, "chk": false, ... }

m = json.load(open(MAP, "r", encoding="utf-8"))
word = com.Dispatch("Word.Application")
word.Visible = False
doc = word.Documents.Open(os.path.abspath(IN))
try:
    changed = 0
    for i in range(1, doc.ContentControls.Count+1):
        cc = doc.ContentControls.Item(i)
        tag = (cc.Tag or "").strip()
        title = (cc.Title or "").strip()
        key = tag if tag else title
        if not key:
            continue
        if key in m:
            val = bool(m[key])
            try:
                cc.Checked = val
            except Exception:
                # fallback: replace control with symbol text
                cc.Range.Text = "X" if val else ""
            changed += 1
            print(f"Set {key} -> {val}")
    out = os.path.splitext(IN)[0] + "_filled.docx"
    doc.SaveAs(os.path.abspath(out))
    print("Saved:", out, "changed", changed)
finally:
    doc.Close(False)
    word.Quit()
