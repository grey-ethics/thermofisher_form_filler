import argparse, json, csv, re, sys, unicodedata, time
from pathlib import Path

try:
    import win32com.client as win32
    from win32com.client import constants as wdconst
except Exception as e:
    print("This script requires pywin32 on Windows with Microsoft Word installed.")
    print("pip install pywin32")
    raise

# ------------------------
# Helpers
# ------------------------

CONTROL_TYPE_MAP = {
    0: "richtext",
    2: "text",
    3: "combobox",
    4: "dropdown",
    6: "date",
    8: "checkbox",
}

def slugify(s: str, max_len: int = 63):
    """Make a machine-friendly tag: ascii, lowercase, snake_case, <= max_len."""
    if not s:
        return ""
    # Normalize accents → ASCII
    s = unicodedata.normalize("NFKD", s)
    s = s.encode("ascii", "ignore").decode("ascii")

    # Replace common separators with spaces
    s = re.sub(r"[/|&>]+", " ", s)

    # Keep words/numbers, collapse whitespace
    s = re.sub(r"[^a-zA-Z0-9]+", " ", s).strip()
    s = re.sub(r"\s+", "_", s.lower())

    # Trim
    if len(s) > max_len:
        s = s[:max_len]
    return s

def looks_good_existing_tag(t: str):
    """Heuristic: keep existing tags that look clean & descriptive."""
    if not t:
        return False
    if len(t) < 4:
        return False
    if re.fullmatch(r"[a-z0-9_]{4,63}", t) is None:
        return False
    # Avoid meaningless defaults like 'chk' or 'checkbox'
    if t in {"chk","checkbox","title","tag"}:
        return False
    return True

def get_heading_level_name(par):
    """Return (level:int or None, name:str) for a paragraph's style if it's a Heading."""
    try:
        # Prefer OutlineLevel (1..9); fallback to style name check
        lvl = int(par.OutlineLevel)
        if 1 <= lvl <= 9:
            name = par.Range.Style.NameLocal
            if isinstance(name, str) and ("Heading" in name or "Überschrift" in name or "Titre" in name):
                return lvl, name
        # Fallback purely on style name
        name = par.Range.Style.NameLocal
        if isinstance(name, str) and name.lower().startswith("heading"):
            # guess level from digits
            m = re.search(r"(\d+)", name)
            lvl = int(m.group(1)) if m else 1
            return lvl, name
    except Exception:
        pass
    return None, None

def extract_headings(doc):
    """Scan paragraphs once; return list of (start, level, text) for Heading 1..3."""
    headings = []
    for p in doc.Paragraphs:
        lvl, _name = get_heading_level_name(p)
        if lvl and 1 <= lvl <= 3:
            txt = p.Range.Text
            txt = txt.replace("\r", " ").strip()
            if txt:
                headings.append((p.Range.Start, lvl, txt))
    headings.sort(key=lambda x: x[0])
    return headings

def nearest_heading_path(headings, pos):
    """Given headings [(start, level, text)], find nearest previous H1..H3 at 'pos' and build path."""
    h1 = h2 = h3 = None
    # Walk through headings up to 'pos'
    for start, lvl, text in headings:
        if start > pos:
            break
        if lvl == 1:
            h1, h2, h3 = text, None, None
        elif lvl == 2:
            h2, h3 = text, None
        elif lvl == 3:
            h3 = text
    path = [h for h in (h1, h2, h3) if h]
    return path

def paragraph_text(par):
    try:
        return par.Range.Text.replace("\r", "\n").strip()
    except Exception:
        return ""

def previous_nonempty_par_text(ctrl):
    try:
        idx = ctrl.Range.Paragraphs(1).Index
        # Walk backwards to find first non-empty paragraph
        doc_pars = ctrl.Parent.Paragraphs
        j = idx - 1
        while j >= 1:
            t = paragraph_text(doc_pars(j))
            if t:
                return t
            j -= 1
    except Exception:
        pass
    return ""

def control_label_guess(ctrl):
    """Best-effort: label on same paragraph / immediate context."""
    t_curr = ""
    try:
        t_curr = paragraph_text(ctrl.Range.Paragraphs(1))
    except Exception:
        pass
    t_prev = previous_nonempty_par_text(ctrl)
    # Prefer the current paragraph, otherwise previous
    return (t_curr or t_prev)[:400]

def safe_get_checkbox_state(ctrl):
    try:
        # Some Word builds expose .Checked; some require .Range.ContentControls(1).Checked
        if ctrl.Type == 8:  # checkbox
            try:
                return bool(ctrl.Checked)
            except Exception:
                # Try nested
                try:
                    inner = ctrl.Range.ContentControls(1)
                    return bool(inner.Checked)
                except Exception:
                    return None
        return None
    except Exception:
        return None

def open_word(visible=False):
    word = win32.Dispatch("Word.Application")
    word.Visible = visible
    return word

def open_doc(word, path):
    return word.Documents.Open(str(Path(path).resolve()))

# ------------------------
# Commands
# ------------------------

def cmd_extract(args):
    word = open_word(visible=False)
    try:
        doc = open_doc(word, args.docx)
        headings = extract_headings(doc)

        rows = []
        for i, ctrl in enumerate(doc.ContentControls, start=1):
            ctype = CONTROL_TYPE_MAP.get(int(ctrl.Type), f"type_{int(ctrl.Type)}")
            tag = ctrl.Tag or ""
            title = ctrl.Title or ""
            start_pos = int(ctrl.Range.Start)
            path = nearest_heading_path(headings, start_pos)
            path_slug = [slugify(x) for x in path]
            para_text = control_label_guess(ctrl)
            checked = safe_get_checkbox_state(ctrl)
            rng_preview = ""
            try:
                rng_preview = ctrl.Range.Text
                rng_preview = rng_preview.replace("\r", " ").replace("\n", " ")
            except Exception:
                pass

            rows.append({
                "index": i,
                "type": ctype,
                "tag": tag,
                "title": title,
                "checked": checked,
                "heading_path": path,
                "heading_path_slug": path_slug,
                "paragraph_context": para_text,
                "range_preview": rng_preview,
            })

        # Write JSON/CSV
        if args.out:
            Path(args.out).write_text(json.dumps(rows, indent=2), encoding="utf-8")
            print(f"Wrote: {args.out}")
        if args.csv:
            with open(args.csv, "w", newline="", encoding="utf-8") as f:
                w = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
                w.writeheader()
                w.writerows(rows)
            print(f"Wrote: {args.csv}")

        print(f"Extracted {len(rows)} controls.")
    finally:
        try:
            doc.Close(False)
        except Exception:
            pass
        word.Quit()

def propose_tag_for_row(row):
    # Keep existing "good" tag
    if looks_good_existing_tag(row.get("tag", "")):
        return row["tag"], row.get("title") or row["tag"]

    # Base from heading path
    parts = row.get("heading_path", [])[:3]
    base = "_".join(slugify(p) for p in parts if p)

    # Add control-type hint if helpful
    ctype = row.get("type", "")
    type_hint = {"checkbox":"chk", "combobox":"combo", "dropdown":"dd", "date":"date", "text":"text"}.get(ctype, ctype)

    # Pull a few words from the paragraph context (label)
    context = row.get("paragraph_context") or ""
    # Grab text near the control: try after the checkbox glyph or colon
    # Simple heuristic: take first 6 words that look content-ish
    words = re.findall(r"[A-Za-z0-9]+", context)
    label_slug = "_".join(w.lower() for w in words[:6])

    # Compose
    bits = [b for b in [base, label_slug, type_hint] if b]
    candidate = slugify("_".join(bits))
    if not candidate:
        candidate = f"{type_hint}_field"

    # Also propose a Title (more human-readable)
    title = " / ".join([p for p in parts if p]) or ctype.capitalize()
    return candidate, title

def cmd_suggest(args):
    rows = json.loads(Path(args.input).read_text(encoding="utf-8"))
    used = set()
    out = []
    for row in rows:
        tag, title = propose_tag_for_row(row)

        # Ensure uniqueness
        base = tag
        n = 1
        while tag in used:
            n += 1
            tag = f"{base}_{n}"
            if len(tag) > 63:
                tag = tag[:60] + f"_{n}"
        used.add(tag)

        row["proposed_tag"] = tag
        # Keep existing human title if present; else proposed
        row["proposed_title"] = row.get("title") or title
        out.append(row)

    # Save
    if args.out:
        Path(args.out).write_text(json.dumps(out, indent=2), encoding="utf-8")
        print(f"Wrote: {args.out}")
    if args.csv and out:
        with open(args.csv, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=list(out[0].keys()))
            w.writeheader()
            w.writerows(out)
        print(f"Wrote: {args.csv}")

    # Quick summary
    total = len(out)
    keep_existing = sum(1 for r in out if looks_good_existing_tag(r.get("tag","")))
    print(f"Suggested tags for {total} controls. Kept {keep_existing} existing tags that looked good.")

def cmd_apply(args):
    """
    Apply Tag/Title from a mapping file to a DOCX's content controls.

    Supports:
      - JSON rows with keys: index, [proposed_tag|tag], [proposed_title|title]
      - --dry-run to preview
      - --write to actually modify
      - --save-as <path> to write to a new DOCX (FileFormat=12)
    """
    import json
    from pathlib import Path

    # --- load mapping (JSON list of dicts) ---
    mapping_path = Path(args.mapping)
    rows = json.loads(mapping_path.read_text(encoding="utf-8"))

    # helpers to pick the right columns from the mapping
    def row_tag(r):
        return (r.get("proposed_tag") or r.get("tag") or "").strip()

    def row_title(r):
        return (r.get("proposed_title") or r.get("title") or "").strip()

    # --- open Word + document ---
    word = open_word(visible=args.show)
    try:
        doc = open_doc(word, args.docx)
        cc = list(doc.ContentControls)
        if len(cc) != len(rows):
            print(f"Warning: doc has {len(cc)} controls, mapping has {len(rows)} rows.")

        applied = 0
        skipped_empty = 0
        skipped_oor = 0
        errors = 0

        # allow mapping in any order
        for r in rows:
            try:
                i = int(r["index"])
            except Exception:
                print(f"Skip row without valid 'index': {r!r}")
                continue

            if i < 1 or i > len(cc):
                print(f"Skip index {i}: out of range (1..{len(cc)})")
                skipped_oor += 1
                continue

            c = cc[i - 1]
            new_tag = row_tag(r)
            new_title = row_title(r)

            if not new_tag:
                print(f"Skip index {i}: empty tag/proposed_tag")
                skipped_empty += 1
                continue

            if args.dry_run:
                print(f"[DRY] #{i}: set Tag='{new_tag}'  Title='{new_title}'  "
                      f"(was Tag='{c.Tag}' Title='{c.Title}')")
                continue

            # actually write
            try:
                c.Tag = new_tag
            except Exception as e:
                print(f"# {i}: failed to set Tag -> {e}")
                errors += 1

            try:
                c.Title = new_title
            except Exception as e:
                print(f"# {i}: failed to set Title -> {e}")
                errors += 1

            applied += 1

        # --- save ---
        if args.dry_run:
            print("Dry run complete. No changes written.")
            return

        # If --save-as provided, save to a NEW file; else save in-place
        save_as = getattr(args, "save_as", None)
        if save_as:
            out_path = Path(save_as).resolve()
            out_path.parent.mkdir(parents=True, exist_ok=True)
            # 12 = wdFormatXMLDocument (.docx)
            doc.SaveAs2(str(out_path), FileFormat=12)
            print(f"Applied tags/titles to {applied} controls and saved as:\n  {out_path}")
        else:
            doc.Save()
            print(f"Applied tags/titles to {applied} controls and saved document.")

        if skipped_empty or skipped_oor or errors:
            print(f"Notes: skipped_empty={skipped_empty}, skipped_out_of_range={skipped_oor}, errors={errors}")

    finally:
        # Leave Word running if user asked to show; otherwise quit to release file locks.
        try:
            if not getattr(args, "show", False):
                word.Quit()
        except Exception:
            pass

        if not args.show:
            try:
                doc.Close(False)
            except Exception:
                pass
            word.Quit()

def cmd_jump(args):
    word = open_word(visible=True)
    doc = open_doc(word, args.docx)
    try:
        target = None
        if args.index:
            i = int(args.index)
            if 1 <= i <= doc.ContentControls.Count:
                target = doc.ContentControls(i)
        elif args.tag:
            for c in doc.ContentControls:
                if str(c.Tag).strip().lower() == args.tag.strip().lower():
                    target = c
                    break

        if not target:
            print("Control not found.")
            return

        rng = target.Range
        rng.Select()
        word.Activate()
        # Give Word a moment to scroll/select
        time.sleep(0.4)
        print(f"Selected control #{target.Index}  Type={CONTROL_TYPE_MAP.get(int(target.Type), target.Type)}  Tag='{target.Tag}' Title='{target.Title}'")

    finally:
        # Leave Word open for user to inspect
        pass

# ------------------------
# Main
# ------------------------

def main():
    p = argparse.ArgumentParser(description="Content Control Tagging Assistant for Word DOCX (COM).")
    sub = p.add_subparsers(dest="cmd", required=True)

    p1 = sub.add_parser("extract", help="Extract controls + context to JSON/CSV")
    p1.add_argument("docx", help="Path to DOCX")
    p1.add_argument("--out", help="Output JSON", default="controls_extracted.json")
    p1.add_argument("--csv", help="Output CSV", default="controls_extracted.csv")
    p1.set_defaults(func=cmd_extract)

    p2 = sub.add_parser("suggest", help="Propose tags/titles from extracted JSON")
    p2.add_argument("input", help="controls_extracted.json")
    p2.add_argument("--out", help="Output JSON", default="controls_suggested.json")
    p2.add_argument("--csv", help="Output CSV", default="controls_suggested.csv")
    p2.set_defaults(func=cmd_suggest)

    p3 = sub.add_parser("apply", help="Apply tags/titles to DOCX from JSON mapping")
    p3.add_argument("docx", help="Path to DOCX")
    p3.add_argument("mapping", help="controls_suggested.json (or similar)")
    p3.add_argument("--dry-run", action="store_true", dest="dry_run", help="Do not modify the document")
    p3.add_argument("--write", action="store_true", help="Actually write changes (alias for not --dry-run)")
    p3.add_argument("--show", action="store_true", help="Open Word UI while applying (useful for spot checks)")
    p3.add_argument("--save-as", dest="save_as", help="Write changes to a NEW .docx (does not overwrite original)")
    p3.set_defaults(func=cmd_apply)

    p4 = sub.add_parser("jump", help="Open Word, jump to a control by index or tag")
    p4.add_argument("docx", help="Path to DOCX")
    g = p4.add_mutually_exclusive_group(required=True)
    g.add_argument("--index", type=int)
    g.add_argument("--tag", type=str)
    p4.set_defaults(func=cmd_jump)

    args = p.parse_args()
    # `--write` as an alias for not --dry-run:
    if getattr(args, "write", False):
        setattr(args, "dry_run", False)

    args.func(args)

if __name__ == "__main__":
    main()
