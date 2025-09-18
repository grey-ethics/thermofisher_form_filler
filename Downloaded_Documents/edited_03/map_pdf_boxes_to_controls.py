# map_pdf_boxes_to_controls.py
import json, csv, math, sys
from pathlib import Path

# --- Word COM (pywin32) ---
try:
    import win32com.client as win32
    from win32com.client import constants as wdconst
except Exception as e:
    print("This script requires pywin32 on Windows with Microsoft Word installed.")
    print("pip install pywin32")
    raise

def open_word(visible=False):
    word = win32.Dispatch("Word.Application")
    word.Visible = visible
    return word

def open_doc(word, path):
    return word.Documents.Open(str(Path(path).resolve()))

def get_docx_checkboxes_with_positions(doc_path):
    """
    Returns checkbox content controls with page + x/y (points) by selecting each control
    and asking Selection.Information(...). Forces main-document view.
    """
    word = open_word(visible=True)  # must be visible for Selection-based coordinates
    data = []
    try:
        doc = open_doc(word, doc_path)
        # make sure we're not in header/footer or some seek view
        try:
            word.ActiveWindow.View.SeekView = wdconst.wdSeekMainDocument
        except Exception:
            pass
        # optional: reduce flicker
        try:
            word.ScreenUpdating = False
        except Exception:
            pass

        for i, c in enumerate(doc.ContentControls, start=1):
            try:
                ctype = int(c.Type)
            except Exception:
                continue
            if ctype != 8:  # checkbox content controls only
                continue

            # select the control to use Selection.Information reliably
            try:
                c.Range.Select()
                sel = word.Selection
                page = int(sel.Information(wdconst.wdActiveEndPageNumber))
                x = float(sel.Information(wdconst.wdHorizontalPositionRelativeToPage))
                y = float(sel.Information(wdconst.wdVerticalPositionRelativeToPage))
            except Exception:
                # skip if any coordinate is unavailable
                continue

            data.append({
                "index": i,
                "tag": str(c.Tag or ""),
                "title": str(c.Title or ""),
                "page": page,
                "x": x,
                "y": y,
            })
    finally:
        # restore
        try:
            word.ScreenUpdating = True
        except Exception:
            pass
        try:
            doc.Close(False)
        except Exception:
            pass
        try:
            word.Quit()
        except Exception:
            pass

    return data

    """
    Returns list of checkbox controls with page, x, y positions (points),
    along with index, tag, title. Only includes Type==checkbox (8).
    """
    word = open_word(visible=False)
    data = []
    try:
        doc = open_doc(word, doc_path)

        # iterate content controls
        for i, c in enumerate(doc.ContentControls, start=1):
            try:
                ctype = int(c.Type)
            except Exception:
                continue
            if ctype != 8:  # checkbox only
                continue

            rng = c.Range
            try:
                page = int(rng.Information(wdconst.wdActiveEndPageNumber))
            except Exception:
                page = None
            try:
                x = float(rng.Information(wdconst.wdHorizontalPositionRelativeToPage))
            except Exception:
                x = None
            try:
                y = float(rng.Information(wdconst.wdVerticalPositionRelativeToPage))
            except Exception:
                y = None

            data.append({
                "index": i,
                "tag": str(c.Tag or ""),
                "title": str(c.Title or ""),
                "page": page,
                "x": x,
                "y": y,
            })
    finally:
        try:
            doc.Close(False)
        except Exception:
            pass
        try:
            word.Quit()
        except Exception:
            pass
    # keep only those with page & coords
    data = [d for d in data if d["page"] and d["x"] is not None and d["y"] is not None]
    return data

def load_pdf_boxes(json_path):
    # expects output from your pdf_find_checkboxes.py
    rows = json.loads(Path(json_path).read_text(encoding="utf-8"))
    # keep only rect entries and page number
    out = []
    for r in rows:
        if "rect" in r and "page" in r:
            x0,y0,x1,y1 = r["rect"]
            out.append({
                "page": int(r["page"]),
                "x0": float(x0), "y0": float(y0),
                "x1": float(x1), "y1": float(y1),
                "cx": float((x0+x1)/2.0),
                "cy": float((y0+y1)/2.0),
                "source": r.get("source",""),
                "page_w": float(r.get("page_w", 612.0)),
                "page_h": float(r.get("page_h", 792.0)),
            })
    return out

def pair_on_page(doc_controls, pdf_boxes, max_dist=20.0):
    """
    Greedy matching on a single page:
    - sort both lists by y (ascending)
    - walk down and match nearest y within tolerance
    - if two are close, pick the closer by Euclidean distance
    Returns (pairs, unmatched_doc, unmatched_pdf)
    Each pair: {"doc": doc_item, "pdf": pdf_item, "dist": euclid_dist}
    """
    dc = sorted(doc_controls, key=lambda d: (d["y"], d["x"]))
    pb = sorted(pdf_boxes, key=lambda p: (p["cy"], p["cx"]))

    i = j = 0
    pairs = []
    used_pdf = set()

    while i < len(dc) and j < len(pb):
        d = dc[i]
        p = pb[j]
        # Euclidean distance doc(x,y) -> pdf(cx,cy)
        dist = math.hypot((d["x"] - p["cx"]), (d["y"] - p["cy"]))

        # Heuristic: if distance is small enough, pair them and advance both.
        if dist <= max_dist:
            pairs.append({"doc": d, "pdf": p, "dist": round(dist,2)})
            used_pdf.add(j)
            i += 1
            j += 1
        else:
            # advance the one with smaller y to try to catch up
            if d["y"] < p["cy"]:
                i += 1
            else:
                j += 1

    unmatched_doc = [d for d in dc if d["index"] not in [x["doc"]["index"] for x in pairs]]
    unmatched_pdf = [pb[k] for k in range(len(pb)) if k not in used_pdf]
    return pairs, unmatched_doc, unmatched_pdf

def map_checkboxes(docx_path, pdf_boxes_json,
                   out_json="checkbox_mapping.json",
                   out_csv="checkbox_mapping.csv",
                   max_dist_pts=20.0):
    # 1) pull docx control positions
    doc_controls_all = get_docx_checkboxes_with_positions(docx_path)
    # 2) load pdf boxes
    pdf_boxes_all = load_pdf_boxes(pdf_boxes_json)

    # group by page
    from collections import defaultdict
    doc_by_page = defaultdict(list)
    pdf_by_page = defaultdict(list)

    for d in doc_controls_all:
        doc_by_page[int(d["page"])].append(d)
    for b in pdf_boxes_all:
        pdf_by_page[int(b["page"])].append(b)

    all_pairs = []
    all_unmatched_doc = []
    all_unmatched_pdf = []

    pages = sorted(set(doc_by_page.keys()) | set(pdf_by_page.keys()))
    for p in pages:
        pairs, u_doc, u_pdf = pair_on_page(doc_by_page.get(p, []), pdf_by_page.get(p, []), max_dist=max_dist_pts)
        all_pairs.extend([{"page": p, **pair} for pair in pairs])
        for d in u_doc:
            d2 = dict(d); d2["page"] = p
            all_unmatched_doc.append(d2)
        for b in u_pdf:
            b2 = dict(b); b2["page"] = p
            all_unmatched_pdf.append(b2)

    # Write JSON
    mapping = []
    for item in all_pairs:
        d = item["doc"]; p = item["pdf"]; dist = item["dist"]; page = item["page"]
        mapping.append({
            "page": page,
            "doc_index": d["index"],
            "doc_tag": d["tag"],
            "doc_title": d["title"],
            "doc_x": round(d["x"],2),
            "doc_y": round(d["y"],2),
            "pdf_rect": [round(p["x0"],2), round(p["y0"],2), round(p["x1"],2), round(p["y1"],2)],
            "pdf_center": [round(p["cx"],2), round(p["cy"],2)],
            "pdf_source": p.get("source",""),
            "distance_pts": dist,
            "page_w": p.get("page_w", 612.0),
            "page_h": p.get("page_h", 792.0),
        })

    Path(out_json).write_text(json.dumps({
        "docx_path": str(Path(docx_path).resolve()),
        "pdf_boxes_json": str(Path(pdf_boxes_json).resolve()),
        "matched": mapping,
        "unmatched_doc_controls": all_unmatched_doc,
        "unmatched_pdf_boxes": all_unmatched_pdf
    }, indent=2), encoding="utf-8")

    # Write CSV
    with open(out_csv, "w", newline="", encoding="utf-8") as f:
        fieldnames = [
            "page","doc_index","doc_title","doc_tag","doc_x","doc_y",
            "pdf_rect","pdf_center","pdf_source","distance_pts","page_w","page_h"
        ]
        w = csv.DictWriter(f, fieldnames=fieldnames)
        w.writeheader()
        for m in mapping:
            row = {
                "page": m["page"],
                "doc_index": m["doc_index"],
                "doc_title": m["doc_title"],
                "doc_tag": m["doc_tag"],
                "doc_x": m["doc_x"], "doc_y": m["doc_y"],
                "pdf_rect": m["pdf_rect"],
                "pdf_center": m["pdf_center"],
                "pdf_source": m["pdf_source"],
                "distance_pts": m["distance_pts"],
                "page_w": m["page_w"], "page_h": m["page_h"],
            }
            w.writerow(row)

    # Summary
    print(f"Doc controls (checkbox) with positions: {len(doc_controls_all)}")
    print(f"PDF boxes detected: {len(pdf_boxes_all)}")
    print(f"Matched pairs: {len(mapping)}")
    if all_unmatched_doc:
        print(f"Unmatched DOCX controls: {len(all_unmatched_doc)}  (likely off by >{max_dist_pts} pts or not rendered as expected)")
    if all_unmatched_pdf:
        print(f"Unmatched PDF boxes: {len(all_unmatched_pdf)}  (likely table squares or extra vector boxes)")

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python map_pdf_boxes_to_controls.py <docx_path> <checkboxes.json> [max_dist_pts]")
        sys.exit(1)
    docx_path = sys.argv[1]
    pdf_boxes_json = sys.argv[2]
    maxd = float(sys.argv[3]) if len(sys.argv) > 3 else 20.0
    map_checkboxes(docx_path, pdf_boxes_json, max_dist_pts=maxd)


