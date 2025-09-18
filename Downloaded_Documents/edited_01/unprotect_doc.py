# unprotect_doc.py
import sys, os
import win32com.client as com

IN = sys.argv[1] if len(sys.argv)>1 else "refernce_template_unlocked_forced_escaped.docx"
PWD = sys.argv[2] if len(sys.argv)>2 else ""   # put password if there is one

word = com.Dispatch("Word.Application")
word.Visible = False
doc = word.Documents.Open(os.path.abspath(IN))
try:
    try:
        if PWD:
            doc.Unprotect(Password=PWD)
        else:
            doc.Unprotect()
        print("Unprotected (or already unprotected).")
    except Exception as e:
        print("Unprotect failed:", e)
    # save new file
    out = os.path.splitext(IN)[0] + "_unprotected.docx"
    doc.SaveAs(os.path.abspath(out))
    print("Saved:", out)
finally:
    doc.Close(False)
    word.Quit()
