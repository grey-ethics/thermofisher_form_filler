# pdf_find_checkboxes.py
import json, csv, math, fitz  # pip install pymupdf
from pathlib import Path

# Characters that commonly render as empty checkboxes (you can extend this)
CHECKBOX_CHARS = [
    "\u2610",  # ☐ BALLOT BOX
    "\u25A1",  # □ WHITE SQUARE
    "\u2751",  # ❑
    "\u274F",  # ❏
    # Some fonts map private-use glyphs; add if you discover them in your PDF:
    # "\uf0a3", # (example)  FontAwesome-ish square if you bump into it
]

# Size heuristics for checkboxes in PDF points (tweak if needed)
MIN_SIZE = 6      # too small => likely punctuation or thin border
MAX_SIZE = 24     # too big   => likely table cells / layout boxes
MAX_ASPECT = 1.25 # width/height should be close to a square
IOU_MERGE = 0.4   # merge duplicates that overlap a lot

def iou(a, b):
    ax0, ay0, ax1, ay1 = a
    bx0, by0, bx1, by1 = b
    ix0, iy0 = max(ax0, bx0), max(ay0, by0)
    ix1, iy1 = min(ax1, bx1), min(ay1, by1)
    iw, ih = max(0, ix1 - ix0), max(0, iy1 - iy0)
    inter = iw * ih
    if inter <= 0: return 0.0
    area_a = (ax1 - ax0) * (ay1 - ay0)
    area_b = (bx1 - bx0) * (by1 - by0)
    return inter / (area_a + area_b - inter)

def merge_overlaps(boxes, iou_thresh=IOU_MERGE):
    boxes = sorted(boxes, key=lambda r: (r["page"], r["rect"][1], r["rect"][0]))
    kept = []
    for b in boxes:
        merged = False
        for k in kept:
            if k["page"] != b["page"]:
                continue
            if iou(k["rect"], b["rect"]) >= iou_thresh:
                # keep the smaller rect (usually the text bbox is tighter)
                ka = (k["rect"][2]-k["rect"][0])*(k["rect"][3]-k["rect"][1])
                ba = (b["rect"][2]-b["rect"][0])*(b["rect"][3]-b["rect"][1])
                if ba < ka:
                    k["rect"] = b["rect"]
                    k["source"] = k["source"] + "+" + b["source"]
                else:
                    k["source"] = k["source"] + "+" + b["source"]
                merged = True
                break
        if not merged:
            kept.append(b)
    return kept

def find_text_boxes(page):
    boxes = []
    # Quick path: if PDF kept the literal character, search_for catches it.
    for ch in CHECKBOX_CHARS:
        try:
            rects = page.search_for(ch, quads=False)
            for r in rects:
                w, h = r.width, r.height
                if MIN_SIZE <= min(w, h) <= MAX_SIZE and (max(w, h) / max(1e-3, min(w, h))) <= MAX_ASPECT:
                    boxes.append({"source":"text:search_for", "rect":[r.x0, r.y0, r.x1, r.y1]})
        except Exception:
            pass

    # Fallback: walk the raw text dict to catch glyphs that search_for might miss
    td = page.get_text("rawdict")
    for block in td.get("blocks", []):
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                text = span.get("text", "")
                if not text:
                    continue
                # if any candidate char is present in span text, build char boxes:
                if any(c in text for c in CHECKBOX_CHARS):
                    # get char details (bbox per glyph) if available:
                    for ch in span.get("chars", []):
                        c = ch.get("c")
                        if c in CHECKBOX_CHARS:
                            r = fitz.Rect(ch["bbox"])
                            w, h = r.width, r.height
                            if MIN_SIZE <= min(w, h) <= MAX_SIZE and (max(w, h) / max(1e-3, min(w, h))) <= MAX_ASPECT:
                                boxes.append({"source":"text:rawdict", "rect":[r.x0, r.y0, r.x1, r.y1]})
    return boxes

def find_vector_boxes(page):
    boxes = []
    drawings = page.get_drawings()
    for d in drawings:
        # Rectangles drawn as simple rects:
        if d["rect"]:
            r = d["rect"]
            w, h = r.width, r.height
            if MIN_SIZE <= min(w, h) <= MAX_SIZE and (max(w, h) / max(1e-3, min(w, h))) <= MAX_ASPECT:
                boxes.append({"source":"vector:rect", "rect":[r.x0, r.y0, r.x1, r.y1]})
        # Paths approximating squares:
        for path in d.get("items", []):
            # Each path item looks like (op, points, ...)
            op = path[0]
            pts = path[1]
            if op != "l":  # we care about poly-lines / closed boxes
                continue
            if not pts:
                continue
            xs = [p[0] for p in pts]
            ys = [p[1] for p in pts]
            r = fitz.Rect(min(xs), min(ys), max(xs), max(ys))
            w, h = r.width, r.height
            if MIN_SIZE <= min(w, h) <= MAX_SIZE and (max(w, h) / max(1e-3, min(w, h))) <= MAX_ASPECT:
                boxes.append({"source":"vector:path", "rect":[r.x0, r.y0, r.x1, r.y1]})
    return boxes

def extract_checkboxes(pdf_path, out_json="checkboxes.json", out_csv="checkboxes.csv", debug_previews=False):
    doc = fitz.open(pdf_path)
    results = []
    for pno in range(len(doc)):
        page = doc[pno]
        page_w, page_h = page.rect.width, page.rect.height

        text_boxes  = find_text_boxes(page)
        vector_boxes = find_vector_boxes(page)

        page_boxes = []
        for b in (text_boxes + vector_boxes):
            r = b["rect"]
            page_boxes.append({
                "page": pno+1,
                "rect": [round(r[0],2), round(r[1],2), round(r[2],2), round(r[3],2)],
                "center": [round((r[0]+r[2])/2,2), round((r[1]+r[3])/2,2)],
                "width": round(r[2]-r[0], 2),
                "height": round(r[3]-r[1], 2),
                "source": b["source"],
                "page_w": round(page_w,2),
                "page_h": round(page_h,2),
            })

        # De-duplicate overlaps on this page
        page_boxes = merge_overlaps(page_boxes)

        # Attach to global list
        results.extend(page_boxes)

        # Optional debug previews — draws red squares on a PNG image
        if debug_previews and page_boxes:
            pix = page.get_pixmap(dpi=144)  # export page
            import PIL.Image, PIL.ImageDraw  # pip install pillow
            img = PIL.Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            # map PDF rects (top-left origin) to image pixels proportionally
            draw = PIL.ImageDraw.Draw(img)
            sx = pix.width / page_w
            sy = pix.height / page_h
            for b in page_boxes:
                x0,y0,x1,y1 = b["rect"]
                draw.rectangle([x0*sx, y0*sy, x1*sx, y1*sy], outline=(255,0,0), width=2)
            img.save(f"debug_boxes_page-{pno+1}.png")

    # Persist
    Path(out_json).write_text(json.dumps(results, indent=2), encoding="utf-8")
    with open(out_csv, "w", newline="", encoding="utf-8") as f:
        fieldnames = ["page","rect","center","width","height","source","page_w","page_h"]
        w = csv.DictWriter(f, fieldnames=fieldnames)
        w.writeheader(); w.writerows(results)

    print(f"Found {len(results)} checkbox-like boxes.")
    print(f"Wrote {out_json} and {out_csv}")

if __name__ == "__main__":
    import sys
    pdf = sys.argv[1] if len(sys.argv) > 1 else "form_preview.pdf"
    extract_checkboxes(pdf)

