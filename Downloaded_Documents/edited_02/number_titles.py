# number_titles.py
import json, sys
from pathlib import Path
from collections import defaultdict

if len(sys.argv) < 3:
    print("Usage: python number_titles.py <controls_extracted.json> <out_mapping.json>")
    sys.exit(1)

src = Path(sys.argv[1])
dst = Path(sys.argv[2])

rows = json.loads(src.read_text(encoding="utf-8"))

# Per-type counters
counters = defaultdict(int)

def next_title(ctype):
    counters[ctype] += 1
    n = counters[ctype]
    if ctype == "checkbox":
        return f"chk - {n:03d}"
    if ctype == "dropdown":
        return f"dropdown - {n:03d}"
    if ctype == "combobox":
        return f"combobox - {n:03d}"
    if ctype == "date":
        return f"date - {n:03d}"
    # fallback (rare)
    return f"{ctype} - {n:03d}"

out = []
for r in rows:
    # Keep the index and whatever tag is currently in the doc.
    # Put the *new* title under "proposed_title".
    out.append({
        "index": r["index"],
        "tag": r.get("tag") or "",  # keep existing tag, even if empty (not recommended)
        "proposed_tag": r.get("tag") or "",  # so apply() uses this if present
        "proposed_title": next_title(r.get("type","")),
        "type": r.get("type",""),
        "heading_path": r.get("heading_path", []),
    })

dst.write_text(json.dumps(out, indent=2), encoding="utf-8")
print(f"Wrote: {dst}  (retitled {len(out)} controls)")
