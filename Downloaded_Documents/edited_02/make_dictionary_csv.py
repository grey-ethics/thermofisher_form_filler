# make_dictionary_csv.py
import json, csv, sys
from pathlib import Path

if len(sys.argv) < 3:
    print("Usage: python make_dictionary_csv.py <controls_json_used_to_apply.json> <out_csv>")
    sys.exit(1)

src_json = Path(sys.argv[1])
out_csv  = Path(sys.argv[2])

rows = json.loads(src_json.read_text(encoding="utf-8"))
rows = sorted(rows, key=lambda r: int(r["index"]))

def get_tag(r):    return (r.get("proposed_tag") or r.get("tag") or "").strip()
def get_title(r):  return (r.get("proposed_title") or r.get("title") or "").strip()
def get_type(r):   return (r.get("type") or "").strip()
def get_heading(r):
    path = r.get("heading_path") or []
    return " / ".join(path)

header = ["index", "controller type", "title", "tag", "heading"]

# Use utf-8-sig so Excel detects Unicode
with out_csv.open("w", newline="", encoding="utf-8-sig") as f:
    w = csv.writer(f)
    w.writerow(header)
    for r in rows:
        w.writerow([
            r["index"],
            get_type(r),
            get_title(r),
            get_tag(r),
            get_heading(r),
        ])

print(f"Wrote: {out_csv}")
