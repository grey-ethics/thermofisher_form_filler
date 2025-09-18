import sys, os
import win32com.client as win32

if len(sys.argv) < 5:
    print("Usage: python set_one_tag.py <docx> <index 1-based> <tag> <title>")
    sys.exit(1)

path   = os.path.abspath(sys.argv[1])
index  = int(sys.argv[2])          # 1-based!
newtag = sys.argv[3]
title  = sys.argv[4]

word = win32.Dispatch("Word.Application")
word.Visible = True
doc  = word.Documents.Open(path)
try:
    ctl = doc.ContentControls.Item(index)   # 1-based COM collection
    ctl.Tag = newtag
    ctl.Title = title
    doc.Save()
    print(f"OK: set Tag='{newtag}', Title='{title}' on control #{index}")
finally:
    doc.Close(SaveChanges=0)
    word.Quit()
