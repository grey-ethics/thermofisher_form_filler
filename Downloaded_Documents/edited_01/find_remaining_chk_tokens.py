# find_remaining_chk_tokens.py
import sys, zipfile, re
IN = sys.argv[1] if len(sys.argv)>1 else "refernce_template_unlocked_forced_escaped_controls.docx"
tok = "<<CHK>>"
ctx = 60

with zipfile.ZipFile(IN, "r") as z:
    parts = [n for n in z.namelist() if n.startswith("word/")]
    found = []
    for p in parts:
        data = z.read(p).decode("utf-8", errors="replace")
        for m in re.finditer(re.escape(tok), data):
            i = m.start()
            s = data[max(0, i-ctx): i+len(tok)+ctx]
            found.append((p, i, s))
    if not found:
        print("No literal <<CHK>> tokens left inside package.")
    else:
        for p,i,s in found:
            print(f"{p} @ {i}:\n...{s}...\n")
        print("Total occurrences:", len(found))
