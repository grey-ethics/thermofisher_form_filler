# count_content_controls.py
# Usage: python count_content_controls.py "path\to\refernce_template_unlocked_forced_escaped_controls.docx"
import sys, os
import win32com.client as com
IN = sys.argv[1] if len(sys.argv)>1 else "refernce_template_unlocked_forced_escaped_controls.docx"
wdContentControlCheckBox = 8

word = com.Dispatch("Word.Application")
word.Visible = False
doc = word.Documents.Open(os.path.abspath(IN))
try:
    cc_count = doc.ContentControls.Count
    print("Total content controls:", cc_count)
    for i in range(1, cc_count+1):
        cc = doc.ContentControls.Item(i)
        t = getattr(cc, "Title", "")
        tag = getattr(cc, "Tag", "")
        ctype = getattr(cc, "Type", None)
        is_cb = (ctype == wdContentControlCheckBox)
        checked = None
        try:
            checked = getattr(cc, "Checked")
        except Exception:
            pass
        rng_text = cc.Range.Text.replace("\r","\\r").replace("\n","\\n")
        print(f"{i:03d}: type={ctype} checkbox={is_cb} checked={checked} tag={tag!r} title={t!r} range_preview={repr(rng_text[:80])}")
finally:
    doc.Close(False)
    word.Quit()
