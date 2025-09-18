import json
import re
from pathlib import Path
import fitz  # PyMuPDF
import math
import sys

# ---------- CONFIG ----------
PDF_PATH = "form_preview.pdf"        # your rendered PDF
DOC_JSON = "controls_extracted.json" # from cc_tag_assistant extract
BOXES_JSON = "checkboxes.json"       # your detected PDF checkbox rects
OUT_JSON  = "pdf_checkbox_mapping.json"
OUT_CSV   = "pdf_checkbox_mapping.csv"

# How many words to use from paragraph_context as the search phrase
ANCHOR_WORDS = 7
MIN_PHRASE_LEN = 12  # characters; skip too-short anchors
# Fallback: also try heading text if paragraph context is thin
USE_HEADING_AS_FALLBACK = True

# ---------- HELPERS ----------

# ---- add this helper (right after the imports) ----
def search_for_compat(page, text, max_hits=64):
    """
    Cross-version wrapper for PyMuPDF's text search.
    Tries: page.search_for(text, flags=?, quads=?), page.search_for(text),
           page.search_for(text, maxhits=?), and legacy page.searchFor(...).
    Returns a list of fitz.Rect.
    """
    # Preferred modern signatures
    try:
        return page.search_for(text, quads=False, hit_max=max_hits)
    except TypeError:
        pass
    try:
        return page.search_for(text, quads=False)
    except TypeError:
        pass
    try:
        return page.search_for(text, maxhits=max_hits)
    except TypeError:
        pass
    try:
        return page.search_for(text)
    except Exception:
        pass

    # Legacy camelCase (very old PyMuPDF)
    try:
        return page.searchFor(text, hit_max=max_hits)
    except TypeError:
        pass
    try:
        return page.searchFor(text)
    except Exception:
        return []


def norm_text(s: str) -> str:
    if not s:
        return ""
    s = s.replace("\r", " ").replace("\n", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s

def anchor_from_context(row):
    # Prefer paragraph_context
    ctx = norm_text(row.get("paragraph_context") or "")
    words = re.findall(r"[A-Za-z0-9]+", ctx)
    if len(words) >= 3:
        phrase = " ".join(words[:ANCHOR_WORDS])
        if len(phrase) >= MIN_PHRASE_LEN:
            return phrase

    # Maybe derive from tag (split on underscores) if context is empty
    tag = row.get("tag") or row.get("proposed_tag") or ""
    parts = [p for p in tag.split("_") if p]
    if len(parts) >= 3:
        candidate = " ".join(parts[:6])
        if len(candidate) >= MIN_PHRASE_LEN:
            return candidate

    # Fallback to heading text
    if USE_HEADING_AS_FALLBACK:
        hp = row.get("heading_path") or []
        if hp:
            h = norm_text(" / ".join(hp))
            if len(h) >= MIN_PHRASE_LEN:
                return h

    return ""

def rect_center(r):
    x0,y0,x1,y1 = r
    return ((x0+x1)/2.0, (y0+y1)/2.0)

def dist2(p, q):
    return (p[0]-q[0])**2 + (p[1]-q[1])**2

# ---------- LOAD ----------
def main():
    # Allow overriding PDF and JSON paths from command line
    args = sys.argv[1:]
    pdf_path = Path(args[0]).resolve() if len(args) >= 1 else Path(PDF_PATH).resolve()
    doc_json = Path(args[1]).resolve() if len(args) >= 2 else Path(DOC_JSON).resolve()
    boxes_json = Path(args[2]).resolve() if len(args) >= 3 else Path(BOXES_JSON).resolve()

    rows = json.loads(Path(doc_json).read_text(encoding="utf-8"))
    boxes = json.loads(Path(boxes_json).read_text(encoding="utf-8"))

    # Index checkboxes by page
    boxes_by_page = {}
    for b in boxes:
        if b.get("source","").startswith("vector:") or b.get("source","").startswith("text:"):
            pg = int(b["page"])
            boxes_by_page.setdefault(pg, []).append(b)

    # Open PDF
    doc = fitz.open(str(pdf_path))

    mappings = []
    unmatched = []
    total_candidates = 0

    # Only map checkboxes here
    check_rows = [r for r in rows if (r.get("type") == "checkbox")]

    for r in check_rows:
        idx = r.get("index")
        tag = r.get("tag") or r.get("proposed_tag") or ""
        title = r.get("title") or r.get("proposed_title") or ""
        anchor = anchor_from_context(r)
        if not anchor:
            unmatched.append({"index": idx, "tag": tag, "reason": "no_anchor"})
            continue

        # Try searching all pages, collect candidates
        best = None  # (page, text_rect, box, distance2, found_text)
        for page_num in range(len(doc)):
            page = doc[page_num]
            rects = search_for_compat(page, anchor, max_hits=32)  # returns list of Rect
            if not rects:
                # Try case-insensitive by lowercasing page text and manual search
                # (fallback – slower – but often helpful)
                page_text = page.get_text("text")
                pt = norm_text(page_text).lower()
                anc = norm_text(anchor).lower()
                # simple contains check; if present, approximate region using all matches of first word
                if anc and anc in pt:
                    # Search first word to get anchors
                    first_word = anc.split(" ")[0]
                    wr = page.search_for(first_word, quads=False, hit_max=64)
                    rects = wr

            if not rects:
                continue

            # We have at least one rect where the phrase/first word occurs
            # Pick the nearest checkbox box on this page
            if page_num+1 not in boxes_by_page:
                # no boxes on this page; skip
                continue

            # Use the first rect center as anchor point (or try all rects; we’ll pick nearest)
            for tr in rects:
                tcenter = rect_center([tr.x0, tr.y0, tr.x1, tr.y1])
                # Find nearest vector/text checkbox box on the same page
                best_box = None
                best_d2 = None
                for b in boxes_by_page[page_num+1]:
                    bcenter = tuple(b["center"])
                    d2 = dist2(tcenter, bcenter)
                    if (best_d2 is None) or (d2 < best_d2):
                        best_d2 = d2
                        best_box = b
                if best_box is not None:
                    if (best is None) or (best_d2 < best[3]):
                        best = (page_num+1, [tr.x0,tr.y0,tr.x1,tr.y1], best_box, best_d2, anchor)

        if best is None:
            unmatched.append({"index": idx, "tag": tag, "reason": "no_pdf_hit", "anchor": anchor})
            continue

        pg, text_rect, box, d2, used_anchor = best
        mappings.append({
            "index": idx,
            "tag": tag,
            "title": title,
            "anchor": used_anchor,
            "page": pg,
            "text_rect": text_rect,
            "box_rect": box["rect"],
            "box_center": box["center"],
            "box_source": box.get("source"),
            "distance": math.sqrt(d2)
        })
        total_candidates += 1

    # Save results
    Path(OUT_JSON).write_text(json.dumps(mappings, indent=2), encoding="utf-8")

    # CSV for quick review
    import csv
    with open(OUT_CSV, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["index","title","tag","page","box_x","box_y","box_w","box_h","distance","anchor"])
        for m in mappings:
            x0,y0,x1,y1 = m["box_rect"]
            w.writerow([
                m["index"], m["title"], m["tag"], m["page"],
                round(m["box_center"][0],2), round(m["box_center"][1],2),
                round(x1-x0,2), round(y1-y0,2),
                round(m["distance"],2),
                m["anchor"][:80]
            ])

    print(f"Mapped {len(mappings)} checkboxes to PDF positions.")
    if unmatched:
        print(f"Unmatched checkboxes: {len(unmatched)}")
        # Write a quick log for troubleshooting
        Path("pdf_mapping_unmatched.json").write_text(json.dumps(unmatched, indent=2), encoding="utf-8")

if __name__ == "__main__":
    main()
