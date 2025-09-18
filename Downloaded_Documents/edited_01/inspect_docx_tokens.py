# inspect_docx_tokens.py
# finds occurrences of common checkbox-like tokens in a .docx (searches word/document.xml)
# Usage: python inspect_docx_tokens.py <file.docx>

import sys, zipfile, re, os

TOKENS = [
    "\u2610", "\u2611",            # ☐ ☑
    "<<CHK>>", "<CHK>", "<CHECK>",
    "[ ]", "[x]", "[X]",
    "<CHK/>", "<CHK />", "<<BOX>>"
]
TOKENS_RE = re.compile("|".join(re.escape(t) for t in TOKENS), re.IGNORECASE)

def sample_around(text, pos, radius=80):
    start = max(0, pos - radius)
    end = min(len(text), pos + radius)
    return text[start:end].replace("\n"," ")

def inspect_docx(path):
    if not os.path.exists(path):
        print("Not found:", path); return 1
    with zipfile.ZipFile(path, "r") as z:
        parts = [p for p in z.namelist() if p.startswith("word/")]
        found = []
        for p in parts:
            try:
                raw = z.read(p).decode("utf-8", errors="replace")
            except Exception as e:
                print("Failed reading part",p,":",e); continue
            for m in TOKENS_RE.finditer(raw):
                context = sample_around(raw, m.start())
                found.append((p, m.group(0), m.start(), context[:200]))
        if not found:
            print("No candidate tokens found in word/ parts using the token list.")
            # also try to detect angle-bracket placeholders like <<...>>
            angle_re = re.compile(r"&lt;{0,2}[^&]{1,40}&gt;{0,2}")  # handles already-escaped < >
            ang_found=[]
            for p in parts:
                try:
                    raw = z.read(p).decode("utf-8", errors="replace")
                except:
                    continue
                for m in re.finditer(r"<<[^>]{1,40}>>|<[^>]{1,40}>", raw):
                    ang_found.append((p, m.group(0), m.start(), raw[max(0,m.start()-60):m.start()+60].replace("\n"," ")))
            if ang_found:
                print("\nAngle-bracket-like tokens (raw):")
                for p, token, pos, ctx in ang_found[:200]:
                    print(p, token, "context:", ctx[:180])
                return 0
            print("Also consider symbol font glyphs / legacy fields. See notes in README.")
            return 0
        print("Found candidate tokens (first 200):")
        for p, token, pos, ctx in found[:200]:
            print("PART:", p, "TOKEN:", repr(token), "pos:", pos)
            print("  context:", ctx)
            print("--------------------------------------------------")
    return 0

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python inspect_docx_tokens.py <file.docx>")
        sys.exit(1)
    sys.exit(inspect_docx(sys.argv[1]))
