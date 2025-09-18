# auto_tag_checkboxes.py
import sys, os, re
import win32com.client as com

def sanitize_tag(s):
    s = (s or "").strip()
    if not s: return "auto_chk"
    out = []
    for ch in s:
        if ch.isalnum() or ch == "_":
            out.append(ch)
        elif ch in (" ", "-", "/"):
            out.append("_")
    t = "".join(out).lower()
    return (t or "auto_chk")[:64]

IN = sys.argv[1] if len(sys.argv)>1 else "refernce_template_unlocked_forced_escaped_controls.docx"
word = com.Dispatch("Word.Application")
word.Visible = False
doc = word.Documents.Open(os.path.abspath(IN))
try:
    changed = 0
    for i in range(1, doc.ContentControls.Count+1):
        cc = doc.ContentControls.Item(i)
        if int(cc.Type) == 8 and (not cc.Tag or cc.Tag.strip() == ""):
            # grab surrounding text: 40 chars before -> 60 after
            start = max(0, cc.Range.Start - 40)
            end = min(doc.Range().End, cc.Range.End + 60)
            ctx = doc.Range(start, end).Text
            # best effort: take first sentence-like bit after the control, else before
            m = re.search(r'([A-Za-z0-9][^\.]{0,80})', ctx)
            snippet = (m.group(1).strip() if m else ctx.strip())[:80]
            tag = sanitize_tag(snippet)
            cc.Tag = tag
            cc.Title = snippet[:128]
            changed += 1
            print(f"Set tag for control #{i} -> {tag!r} title={snippet!r}")
    if changed:
        out = os.path.splitext(IN)[0] + "_tagged.docx"
        doc.SaveAs(os.path.abspath(out))
        print("Saved new doc:", out)
    else:
        print("No unnamed checkbox controls found.")
finally:
    doc.Close(False)
    word.Quit()
