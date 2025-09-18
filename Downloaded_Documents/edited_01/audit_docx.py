#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DOCX Auditor
- Validates XML of all parts in a .docx package
- Checks document protection + content-control locks
- Audits content controls (type, tag/title, checkbox state symbols)
- Flags leftover tokens: <<CHK>> and ☐/☑/☒ (U+2610..U+2612)
- Detects legacy form fields (pre-content-control era)
- Compares normalized visible text to an original DOCX (optional)
- Emits console report + audit_report.json

Usage:
  python audit_docx.py "final.docx" --original "refernce_template.docx"
"""

import argparse
import difflib
import io
import json
import os
import re
import sys
import zipfile
from collections import Counter, defaultdict
from xml.etree import ElementTree as ET

NS = {
    "w":   "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r":   "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
    "w16": "http://schemas.microsoft.com/office/word/2018/wordml",
    "mc":  "http://schemas.openxmlformats.org/markup-compatibility/2006",
}
# helpers to qualify tags/attrs
def q(tag):  # "w:sdt" -> "{ns}sdt"
    p, t = tag.split(":")
    return "{%s}%s" % (NS[p], t)
def qa(attr):
    p, t = attr.split(":")
    return "{%s}%s" % (NS[p], t)

# compiled regexes
RE_CHK_TOKEN = re.compile(re.escape("<<CHK>>"))
RE_BOX_GLYPHS = re.compile(r"[\u2610\u2611\u2612]")  # ☐☑☒
RE_XMLWS = re.compile(r"\s+")
RE_NORMALIZE_SPACES = re.compile(r"[ \t\r\f\v]+")

WORD_PART_PREFIX = "word/"
TEXTY_PARTS = (
    "document.xml", "footnotes.xml", "endnotes.xml",
    "header", "footer", "comments", "glossary/document.xml"
)

def load_docx_xml_parts(path):
    """Yield (name, xml_bytes) for all .xml parts in the package."""
    parts = []
    with zipfile.ZipFile(path, "r") as z:
        for name in z.namelist():
            if name.lower().endswith(".xml"):
                parts.append((name, z.read(name)))
    return parts

def parse_xml_or_error(name, data):
    try:
        root = ET.fromstring(data)
        return {"ok": True, "name": name, "error": None, "root": root}
    except ET.ParseError as e:
        return {"ok": False, "name": name, "error": str(e), "root": None}

def find_document_protection(parts_map):
    """Check word/settings.xml for protection elements."""
    out = {"settings_present": False, "documentProtection": None, "writeProtection": None}
    settings = parts_map.get("word/settings.xml")
    if not settings: return out
    out["settings_present"] = True
    r = ET.fromstring(settings)
    dp = r.find(".//" + q("w:documentProtection"))
    if dp is not None:
        # Word uses attributes like w:edit, w:enforcement, algorithmName, etc.
        out["documentProtection"] = {
            k.split("}")[-1]: v for k, v in dp.attrib.items()
        }
    wp = r.find(".//" + q("w:writeProtection"))
    if wp is not None:
        out["writeProtection"] = {k.split("}")[-1]: v for k, v in wp.attrib.items()}
    return out

def extract_content_controls_from_tree(root, part_name):
    """Return list of content controls with metadata from one XML tree."""
    controls = []
    for sdt in root.findall(".//" + q("w:sdt")):
        sdtPr = sdt.find("./" + q("w:sdtPr"))
        if sdtPr is None:
            controls.append({
                "part": part_name, "type": "unknown", "tag": "", "title": "",
                "locked": None, "checkbox": None
            })
            continue
        # Tag/Title
        tag_el = sdtPr.find("./" + q("w:tag"))
        alias_el = sdtPr.find("./" + q("w:alias"))
        tag = (tag_el.get(qa("w:val")) if tag_el is not None else "") or ""
        title = (alias_el.get(qa("w:val")) if alias_el is not None else "") or ""

        # Lock
        lock_el = sdtPr.find("./" + q("w:lock"))
        locked = None
        if lock_el is not None:
            locked = lock_el.get(qa("w:val"))  # 'sdtLocked' or 'contentLocked'

        # Type detection
        ctype = "unknown"
        checkbox_info = None
        if sdtPr.find(".//" + q("w14:checkbox")) is not None:
            ctype = "checkbox"
            cb = sdtPr.find(".//" + q("w14:checkbox"))
            checked_state = cb.find("./" + q("w14:checkedState"))
            unchecked_state = cb.find("./" + q("w14:uncheckedState"))
            checked = cb.find("./" + q("w14:checked"))
            checkbox_info = {
                "checked_val": (checked.get(qa("w14:val")) if checked is not None else None),
                "checked_symbol": (checked_state.get(qa("w14:val")) if checked_state is not None else None),
                "checked_font": (checked_state.get(qa("w14:font")) if checked_state is not None else None),
                "unchecked_symbol": (unchecked_state.get(qa("w14:val")) if unchecked_state is not None else None),
                "unchecked_font": (unchecked_state.get(qa("w14:font")) if unchecked_state is not None else None),
            }
        elif sdtPr.find("./" + q("w:dropDownList")) is not None:
            ctype = "dropdown"
        elif sdtPr.find("./" + q("w:comboBox")) is not None:
            ctype = "combobox"
        elif sdtPr.find("./" + q("w:date")) is not None or sdtPr.find(".//" + q("w15:datePicker")) is not None:
            ctype = "date"
        elif sdtPr.find("./" + q("w:richText")) is not None:
            ctype = "richText"
        elif sdtPr.find("./" + q("w:text")) is not None:
            ctype = "text"
        elif sdtPr.find("./" + q("w:picture")) is not None:
            ctype = "picture"
        elif sdtPr.find(".//" + q("w15:repeatingSection")) is not None:
            ctype = "repeatingSection"
        elif sdtPr.find(".//" + q("w15:repeatingSectionItem")) is not None:
            ctype = "repeatingSectionItem"

        controls.append({
            "part": part_name, "type": ctype, "tag": tag, "title": title,
            "locked": locked, "checkbox": checkbox_info
        })
    return controls

def detect_legacy_form_fields(root):
    """Detect old-style 'protected form' fields."""
    legacy = []
    # <w:fldSimple w:instr="FORMCHECKBOX"> or structured <w:ffData><w:checkBox/>
    for fld in root.findall(".//" + q("w:fldSimple")):
        instr = fld.get(qa("w:instr")) or ""
        if "FORMCHECKBOX" in instr.upper():
            legacy.append({"kind":"fldSimple", "instr": instr})
    for ff in root.findall(".//" + q("w:ffData")):
        if ff.find("./" + q("w:checkBox")) is not None:
            legacy.append({"kind":"ffData.checkBox", "instr": None})
    return legacy

def gather_text_from_tree(root):
    """Concatenate text from all w:t nodes in reading order."""
    texts = []
    for t in root.findall(".//" + q("w:t")):
        texts.append(t.text or "")
    return "".join(texts)

def normalize_visible_text(s):
    """Normalize doc text for comparison: strip control glyphs/tokens, compress whitespace."""
    if not s:
        return ""
    # remove checkbox box glyphs and literal token
    s = RE_CHK_TOKEN.sub("", s)
    s = RE_BOX_GLYPHS.sub("", s)
    # normalize quotes (Word smart quotes) to straight quotes to be robust
    s = s.replace("“", '"').replace("”", '"').replace("’", "'").replace("‘", "'")
    # collapse whitespace
    s = RE_NORMALIZE_SPACES.sub(" ", s)
    return s.strip()

def collect_text_from_docx(path):
    parts = load_docx_xml_parts(path)
    text = []
    for name, data in parts:
        lname = name.lower()
        if not lname.startswith(WORD_PART_PREFIX): 
            continue
        if not any(key in lname for key in TEXTY_PARTS):
            continue
        try:
            root = ET.fromstring(data)
        except ET.ParseError:
            # ignore broken part here (will be caught in validation phase)
            continue
        text.append(gather_text_from_tree(root))
        # add newlines to separate parts
        text.append("\n")
    return normalize_visible_text("".join(text))

def audit(path, original_path=None):
    result = {
        "file": os.path.abspath(path),
        "xml_validation": {"ok": True, "errors": []},
        "protection": {},
        "content_controls": {
            "total": 0, "by_type": {}, "locked_count": 0,
            "empty_tag_or_title": 0, "checkbox_symbol_summary": Counter()
        },
        "leftover_tokens": {"chk_tokens": [], "box_glyph_hits": []},
        "legacy_form_fields": [],
        "text_comparison": None,
    }

    # Load parts
    parts = load_docx_xml_parts(path)
    parts_map = {name: data for (name, data) in parts}

    # 1) XML validation
    xml_errors = []
    for name, data in parts:
        pr = parse_xml_or_error(name, data)
        if not pr["ok"]:
            xml_errors.append({"part": name, "error": pr["error"]})
    result["xml_validation"]["ok"] = (len(xml_errors) == 0)
    result["xml_validation"]["errors"] = xml_errors

    # 2) Document + write protection (settings.xml)
    result["protection"] = find_document_protection(parts_map)

    # 3) Content controls audit + locks + symbols
    controls = []
    locked_count = 0
    by_type = Counter()
    empty_tag_title = 0
    symbol_counter = Counter()
    legacy_all = []

    for name, data in parts:
        if not name.startswith("word/"): 
            continue
        # parse
        try:
            root = ET.fromstring(data)
        except ET.ParseError:
            continue

        # legacy
        legacy = detect_legacy_form_fields(root)
        if legacy:
            for item in legacy:
                item["part"] = name
            legacy_all.extend(legacy)

        # controls
        ccs = extract_content_controls_from_tree(root, name)
        controls.extend(ccs)
        for c in ccs:
            by_type[c["type"]] += 1
            if c["locked"]: locked_count += 1
            if (c["tag"] or "").strip() == "" and (c["title"] or "").strip() == "":
                empty_tag_title += 1
            if c["type"] == "checkbox" and c["checkbox"]:
                # summarize checked/unchecked symbol codes (hex)
                ch = c["checkbox"]["checked_symbol"]
                un = c["checkbox"]["unchecked_symbol"]
                if ch: symbol_counter[f"checked:{ch}"] += 1
                if un: symbol_counter[f"unchecked:{un}"] += 1

        # leftover tokens in this part text nodes (fast scan raw bytes too)
        # raw scan for <<CHK>>
        if RE_CHK_TOKEN.search(data.decode("utf-8", "ignore")):
            # try to grab small contexts
            s = data.decode("utf-8", "ignore")
            for m in RE_CHK_TOKEN.finditer(s):
                start = max(0, m.start() - 40)
                end = min(len(s), m.end() + 40)
                snippet = s[start:end].replace("\n", " ")
                result["leftover_tokens"]["chk_tokens"].append({"part": name, "context": snippet})

        # raw scan for box glyphs
        if RE_BOX_GLYPHS.search(data.decode("utf-8", "ignore")):
            s = data.decode("utf-8", "ignore")
            for m in RE_BOX_GLYPHS.finditer(s):
                start = max(0, m.start() - 40)
                end = min(len(s), m.end() + 40)
                snippet = s[start:end].replace("\n", " ")
                result["leftover_tokens"]["box_glyph_hits"].append({"part": name, "context": snippet})

    result["content_controls"]["total"] = len(controls)
    result["content_controls"]["by_type"] = dict(by_type)
    result["content_controls"]["locked_count"] = locked_count
    result["content_controls"]["empty_tag_or_title"] = empty_tag_title
    result["content_controls"]["checkbox_symbol_summary"] = dict(symbol_counter)
    result["legacy_form_fields"] = legacy_all

    # 4) Text comparison to original (optional)
    if original_path:
        final_text = collect_text_from_docx(path)
        orig_text  = collect_text_from_docx(original_path)
        sm = difflib.SequenceMatcher(None, orig_text, final_text)
        ratio = sm.ratio()
        # small diff samples
        diffs = []
        for tag, i1, i2, j1, j2 in sm.get_opcodes():
            if tag == "equal":
                continue
            a = orig_text[i1:i2]
            b = final_text[j1:j2]
            if len(diffs) < 8:  # limit
                diffs.append({
                    "op": tag,
                    "orig_excerpt": (a[:160] + ("…" if len(a) > 160 else "")),
                    "final_excerpt": (b[:160] + ("…" if len(b) > 160 else "")),
                    "orig_span": [i1, i2],
                    "final_span": [j1, j2],
                })
        result["text_comparison"] = {
            "similarity_ratio": round(ratio, 6),
            "orig_length": len(orig_text),
            "final_length": len(final_text),
            "diff_samples": diffs
        }

    return result

def print_human_report(rep):
    print("\n=== DOCX AUDIT REPORT ===")
    print(f"File: {rep['file']}\n")

    # XML validation
    v = rep["xml_validation"]
    print(f"[XML] All parts well-formed: {'YES' if v['ok'] else 'NO'}")
    if not v["ok"]:
        for e in v["errors"][:12]:
            print(f"  - {e['part']}: {e['error']}")
        if len(v["errors"]) > 12:
            print(f"  ... and {len(v['errors'])-12} more")

    # Protection
    p = rep["protection"]
    if not p.get("settings_present"):
        print("[Protection] word/settings.xml not found (unusual).")
    else:
        dp = p.get("documentProtection")
        wp = p.get("writeProtection")
        if dp:
            enf = dp.get("enforcement")
            print(f"[Protection] DocumentProtection present: {dp}  (ENFORCEMENT={enf})")
        else:
            print("[Protection] No documentProtection element — OK (unprotected).")
        if wp:
            print(f"[Protection] WriteProtection present: {wp}")
        else:
            print("[Protection] No writeProtection element — OK.")

    # Content controls
    cc = rep["content_controls"]
    print(f"\n[Controls] Total content controls: {cc['total']}")
    print(f"[Controls] By type: {cc['by_type']}")
    print(f"[Controls] Locked controls: {cc['locked_count']}")
    print(f"[Controls] Controls with empty Tag & Title: {cc['empty_tag_or_title']}")
    sym = cc.get("checkbox_symbol_summary") or {}
    if sym:
        print(f"[Controls] Checkbox symbol codes (hex): {sym}")
        # tip
        print("         (For an 'X' checked symbol, you typically see checked:'0058')")

    # Legacy form fields
    if rep["legacy_form_fields"]:
        print(f"\n[Legacy] Found {len(rep['legacy_form_fields'])} legacy form field(s):")
        for item in rep["legacy_form_fields"][:10]:
            print(f"  - {item['kind']} in {item['part']}  instr={item.get('instr')}")
        if len(rep["legacy_form_fields"]) > 10:
            print(f"  ... and {len(rep['legacy_form_fields'])-10} more")
    else:
        print("\n[Legacy] No legacy form fields detected — OK.")

    # Leftover tokens
    lt = rep["leftover_tokens"]
    if lt["chk_tokens"]:
        print(f"\n[Leftovers] Found {len(lt['chk_tokens'])} '<<CHK>>' token(s):")
        for it in lt["chk_tokens"][:10]:
            print(f"  - {it['part']}: …{it['context']}…")
        if len(lt["chk_tokens"]) > 10:
            print(f"  ... and {len(lt['chk_tokens'])-10} more")
    else:
        print("\n[Leftovers] No '<<CHK>>' tokens — OK.")

    if lt["box_glyph_hits"]:
        print(f"[Leftovers] Found {len(lt['box_glyph_hits'])} occurrences of ☐/☑/☒:")
        for it in lt["box_glyph_hits"][:10]:
            print(f"  - {it['part']}: …{it['context']}…")
        if len(lt["box_glyph_hits"]) > 10:
            print(f"  ... and {len(lt['box_glyph_hits'])-10} more")
    else:
        print("[Leftovers] No ☐/☑/☒ glyphs — OK.")

    # Text comparison
    tc = rep.get("text_comparison")
    if tc:
        print("\n[Content] Normalized text comparison to ORIGINAL (ignoring control markup):")
        print(f"         Similarity ratio: {tc['similarity_ratio']:.6f}")
        print(f"         Original length: {tc['orig_length']}  Final length: {tc['final_length']}")
        if tc["diff_samples"]:
            print("         Diff samples (first few):")
            for d in tc["diff_samples"][:5]:
                print(f"         - {d['op']} orig[{d['orig_span'][0]}:{d['orig_span'][1]}] vs "
                      f"final[{d['final_span'][0]}:{d['final_span'][1]}]")
                print(f"           ORIG : {d['orig_excerpt']}")
                print(f"           FINAL: {d['final_excerpt']}")
        else:
            print("         No differences detected in normalized text.")
    else:
        print("\n[Content] No original provided; skipped text comparison.")

def main():
    ap = argparse.ArgumentParser(description="Audit a DOCX for XML validity, protection, controls, leftovers, and content similarity.")
    ap.add_argument("docx", help="Path to final DOCX to audit")
    ap.add_argument("--original", help="Path to original DOCX to compare visible text against", default=None)
    ap.add_argument("--json", help="Write JSON report to this file (default: audit_report.json beside DOCX)", default=None)
    args = ap.parse_args()

    if not os.path.isfile(args.docx):
        print(f"ERROR: file not found: {args.docx}", file=sys.stderr)
        sys.exit(2)
    if args.original and not os.path.isfile(args.original):
        print(f"ERROR: original file not found: {args.original}", file=sys.stderr)
        sys.exit(2)

    rep = audit(args.docx, args.original)
    print_human_report(rep)

    # Write JSON
    out_json = args.json or os.path.join(os.path.dirname(os.path.abspath(args.docx)), "audit_report.json")
    try:
        with io.open(out_json, "w", encoding="utf-8") as fh:
            json.dump(rep, fh, indent=2, ensure_ascii=False)
        print(f"\nJSON report written to: {out_json}")
    except Exception as e:
        print(f"\nWARNING: could not write JSON report: {e}")

if __name__ == "__main__":
    main()
