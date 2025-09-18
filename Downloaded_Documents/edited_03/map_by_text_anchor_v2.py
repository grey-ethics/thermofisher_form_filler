import json, re, sys
from pathlib import Path

import fitz  # PyMuPDF

# ---------- Tunables ----------
# how many words from context to try (start high -> then shorter)
ANCHOR_WORDS_TRY = [8, 6, 5, 4, 3]
# minimum anchor length after normalization
MIN_ANCHOR_LEN = 6
# max search hits to consider per anchor
MAX_HITS = 64
# distance threshold (points) from anchor text to a box center
MAX_DISTANCE = 65.0
# consider only reasonably small square-ish boxes (points)
MIN_BOX = 7.0
MAX_BOX = 22.0
MAX_ASPECT = 1.6  # allow a bit rectangular
# --------------------------------

def search_for_compat(page, text, max_hits=64):
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
    # legacy camelCase
    try:
        return page.searchFor(text, hit_max=max_hits)
    except TypeError:
        pass
    try:
        return page.searchFor(text)
    except Exception:
        return []

def load_pdf_boxes(box_json_path):
    boxes = json.loads(Path(box_json_path).read_text(encoding="utf-8"))
    # Tighten candidate set to "checkbox-looking" squares
    filtered = []
    for b in boxes:
        w = float(b["width"])
        h = float(b["height"])
        if (MIN_BOX <= w <= MAX_BOX) and (MIN_BOX <= h <= MAX_BOX):
            ar = max(w, h) / max(1e-6, min(w, h))
            if ar <= MAX_ASPECT:
                filtered.append(b)
    return filtered

def norm(s: str) -> str:
    if not s:
        return ""
    # kill bullets/glyphs/messy quotes etc, keep letters/numbers/space
    s = s.encode("ascii", "ignore").decode("ascii")
    s = s.lower()
    s = s.replace("\r", " ").replace("\n", " ")
    s = re.sub(r"[/|&>]+", " ", s)
    s = re.sub(r"[^a-z0-9]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def take_first_words(s: str, n: int) -> str:
    if not s:
        return ""
    words = s.split()
    return " ".join(words[:n])

def distance(p1, p2):
    dx = float(p1[0]) - float(p2[0])
    dy = float(p1[1]) - float(p2[1])
    return (dx*dx + dy*dy) ** 0.5

def rect_center(r):
    return ((r.x0 + r.x1)/2.0, (r.y0 + r.y1)/2.0)

def candidates_from_row(row):
    """
    Build multiple potential anchors for one checkbox:
    - paragraph_context chunks
    - range_preview chunks
    - tag words (split underscores)
    - heading (lightly)
    """
    cands = []

    ctx = norm(row.get("paragraph_context", ""))
    rng = norm(row.get("range_preview", ""))
    tag = norm(row.get("tag", "")).replace("_", " ")
    # take only content-y bits of tag (skip trailing _chk etc)
    tag_words = [w for w in tag.split() if w not in {"chk", "check", "checkbox"}]
    tag_short = " ".join(tag_words[:6])
    heading_path = row.get("heading_path", []) or []
    heading = norm(" ".join(heading_path)) if heading_path else ""

    # prefer paragraph context; then range; then tag; then heading
    primary = [ctx, rng, tag_short, heading]
    seen = set()
    for base in primary:
        if not base:
            continue
        for n in ANCHOR_WORDS_TRY:
            frag = take_first_words(base, n)
            if len(frag) < MIN_ANCHOR_LEN:
                continue
            if frag in seen:
                continue
            seen.add(frag)
            cands.append(frag)

    return cands

def map_checkboxes(pdf_path, extract_json, pdf_boxes_json, out_json="checkbox_map.json"):
    doc = fitz.open(pdf_path)
    rows = json.loads(Path(extract_json).read_text(encoding="utf-8"))
    boxes = load_pdf_boxes(pdf_boxes_json)

    # Group boxes by page for quick lookup
    boxes_by_page = {}
    for b in boxes:
        boxes_by_page.setdefault(int(b["page"]), []).append(b)

    mapped = []
    unmatched = []

    for row in rows:
        if row.get("type") != "checkbox":
            continue

        idx = int(row["index"])
        anchors = candidates_from_row(row)
        found = None

        # Try each candidate phrase across all pages; pick the closest box on same page
        for page_num in range(len(doc)):
            page = doc[page_num]
            # If this PDF is from the same DOCX with same pagination, the checkbox
            # is likely on the same page as that text; this loop is brute-force but robust.
            page_boxes = boxes_by_page.get(page_num + 1, [])  # your boxes are 1-based pages

            if not page_boxes:
                continue

            # Precompute page text (normalized) for quick “does it exist at all” check
            # (Optional micro-optimization; but we’ll rely on search_for anyway)
            for a in anchors:
                rects = search_for_compat(page, a, max_hits=MAX_HITS)
                if not rects:
                    continue

                # We have one or more anchor rects — for each, pick nearest checkbox box on this page
                for r in rects:
                    ac = rect_center(r)
                    # choose nearest box within threshold
                    best = None
                    best_d = 1e9
                    for b in page_boxes:
                        d = distance(ac, b["center"])
                        if d < best_d:
                            best_d = d
                            best = b
                    if best and best_d <= MAX_DISTANCE:
                        found = {
                            "index": idx,
                            "tag": row.get("tag",""),
                            "title": row.get("title",""),
                            "page": page_num + 1,
                            "anchor": a,
                            "anchor_rect": [r.x0, r.y0, r.x1, r.y1],
                            "box_rect": best["rect"],
                            "box_center": best["center"],
                            "distance": best_d,
                        }
                        break  # accept first decent match
                if found:
                    break
            if found:
                break

        if found:
            mapped.append(found)
        else:
            unmatched.append(idx)

    Path(out_json).write_text(json.dumps({
        "mapped": mapped,
        "unmatched": unmatched,
        "counts": {
            "mapped": len(mapped),
            "unmatched": len(unmatched)
        }
    }, indent=2), encoding="utf-8")

    print(f"Mapped {len(mapped)} checkboxes. Unmatched: {len(unmatched)}")
    if unmatched:
        print(f"Sample unmatched indices (first 20): {unmatched[:20]}")

def main():
    if len(sys.argv) < 4:
        print("Usage: python map_by_text_anchor_v2.py form_preview.pdf controls_extracted.json checkboxes.json [out.json]")
        sys.exit(1)
    pdf_path = sys.argv[1]
    rows_json = sys.argv[2]
    boxes_json = sys.argv[3]
    out_json = sys.argv[4] if len(sys.argv) > 4 else "checkbox_map.json"
    map_checkboxes(pdf_path, rows_json, boxes_json, out_json)

if __name__ == "__main__":
    main()
