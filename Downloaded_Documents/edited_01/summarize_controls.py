# summarize_controls.py
import sys, json, csv, pathlib

def from_audit(audit):
    cc = audit.get("content_controls", {})
    print("From audit_report.json:")
    print(" Total content controls:", cc.get("total"))
    print(" By type:", cc.get("by_type"))
    print(" Locked controls:", cc.get("locked_count"))
    print(" Controls with empty Tag or Title:", cc.get("empty_tag_or_title"))
    print(" Checkbox symbol summary:", cc.get("checkbox_symbol_summary"))
    leftovers = audit.get("leftover_tokens", {})
    print(" Leftover chk tokens:", len(leftovers.get("chk_tokens", [])))
    box_hits = leftovers.get("box_glyph_hits", [])
    print(" Leftover box glyph occurrences found:", len(box_hits))
    print()

def inspect_controls_dump(controls_json):
    # expected structure: {"controls": [ {index, tag, title, type, range_preview, context_snippet}, ... ]}
    controls = controls_json.get("controls", [])
    types = {}
    empties = []
    for c in controls:
        t = c.get("type") or c.get("control_type") or "unknown"
        types.setdefault(t, []).append(c)
        if not c.get("tag") or not c.get("title"):
            empties.append(c)
    print("From controls_report.json:")
    print(" Total controls in dump:", len(controls))
    print(" Types breakdown:")
    for t, arr in types.items():
        print(f"  - {t}: {len(arr)}")
    print(" Controls with empty tag/title:", len(empties))
    if empties:
        print(" Listing first 20 empties (index, type, range_preview, context_snippet):")
        for c in empties[:20]:
            print(f"  idx:{c.get('index')} type:{c.get('type')} preview:{c.get('range_preview')!r} ctx:{(c.get('context_snippet') or '')[:80]!r}")
    print()
    return types, empties

def write_csv_controls(controls_json, outfn="controls_flat.csv"):
    controls = controls_json.get("controls", [])
    if not controls:
        print("No controls to write.")
        return
    keys = ["index","type","tag","title","range_preview","context_snippet"]
    with open(outfn, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, keys)
        w.writeheader()
        for c in controls:
            row = {k: c.get(k, "") for k in keys}
            w.writerow(row)
    print("Wrote", outfn)

def main():
    if len(sys.argv) < 2:
        print("Usage: python summarize_controls.py audit_report.json [controls_report.json]")
        return
    audit_fn = sys.argv[1]
    audit = json.load(open(audit_fn, encoding="utf-8"))
    from_audit(audit)
    if len(sys.argv) > 2:
        ctrl_fn = sys.argv[2]
        controls_json = json.load(open(ctrl_fn, encoding="utf-8"))
        types, empties = inspect_controls_dump(controls_json)
        write_csv_controls(controls_json, outfn="controls_flat.csv")
    else:
        print("If you want per-control details (index/tag/title), pass your controls_report.json as the 2nd arg.")

if __name__ == "__main__":
    main()
