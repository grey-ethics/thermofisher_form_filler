# make_dictionary_md.py
import json, sys
from pathlib import Path

if len(sys.argv) < 3:
    print("Usage: python make_dictionary_md.py <controls_json_used_to_apply.json> <out_md>")
    sys.exit(1)

rows = json.loads(Path(sys.argv[1]).read_text(encoding="utf-8"))
out_path = Path(sys.argv[2])

# ---------- helpers to normalize fields ----------
def get_tag(r):    return (r.get("proposed_tag") or r.get("tag") or "").strip()
def get_title(r):  return (r.get("proposed_title") or r.get("title") or "").strip()
def get_type(r):   return (r.get("type") or "").strip()
def get_heading(r):
    path = r.get("heading_path") or []
    return " / ".join(path)

rows_sorted = sorted(rows, key=lambda r: int(r["index"]))

# ---------- compute global fixed widths (with your minimums) ----------
INDEX_W = 5

TITLE_W_MIN   = len("combobox - 004")  # your example
TAG_W_MIN     = len("environmental_assessment_revise_regulatory_documents_i_e_declar") + 5
HEADING_W_MIN = len("DESIGN SAFETY, EMC, & WIRELESS COMPLIANCE ASSESSMENT ") + 5

# observe actual data so nothing legitimate gets chopped
title_max   = max((len(get_title(r))   for r in rows_sorted), default=TITLE_W_MIN)
tag_max     = max((len(get_tag(r))     for r in rows_sorted), default=TAG_W_MIN)
heading_max = max((len(get_heading(r)) for r in rows_sorted), default=HEADING_W_MIN)

TITLE_W   = max(TITLE_W_MIN, title_max)
TAG_W     = max(TAG_W_MIN, tag_max)
HEADING_W = max(HEADING_W_MIN, heading_max)

def fit(s: str, width: int, align: str = "left") -> str:
    """Truncate with ellipsis if needed, then pad to width."""
    if s is None:
        s = ""
    if len(s) > width:
        # leave room for ellipsis
        s = s[: max(0, width - 1)] + "…"
    pad = " " * max(0, width - len(s))
    if align == "right":
        return pad + s
    return s + pad

def row_line(idx, title, tag, heading):
    return (
        "| " + fit(str(idx), INDEX_W, "right")
        + " | " + fit(title,   TITLE_W,   "left")
        + " | " + fit(tag,     TAG_W,     "left")
        + " | " + fit(heading, HEADING_W, "left")
        + " |"
    )

# precompute separator length
hdr_line = row_line("index", "title", "tag", "heading")
sep_line = "|" + "-"*(len(hdr_line)-2) + "|"

# counts
counts = {}
for r in rows_sorted:
    t = get_type(r)
    counts[t] = counts.get(t, 0) + 1

with out_path.open("w", encoding="utf-8") as f:
    f.write("# Content Controls Dictionary\n\n")
    f.write("**Counts by type**:\n\n")
    for t in sorted(counts):
        f.write(f"- {t}: {counts[t]}\n")
    f.write("\n---\n\n")

    # group by type
    by_type = {}
    for r in rows_sorted:
        by_type.setdefault(get_type(r), []).append(r)

    for t in sorted(by_type):
        f.write(f"## {t}\n\n")
        # use a monospaced code block so fixed widths render correctly everywhere
        f.write("```text\n")
        f.write(hdr_line + "\n")
        f.write(sep_line + "\n")
        for r in by_type[t]:
            f.write(
                row_line(
                    r["index"],
                    get_title(r),
                    get_tag(r),
                    get_heading(r),
                ) + "\n"
            )
        f.write("```\n\n")

print(f"Wrote: {out_path}")
