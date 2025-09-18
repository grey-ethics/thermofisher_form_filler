# set_checkboxes_from_json.py
"""
Usage:
  python set_checkboxes_from_json.py <input.docx> <mapping.json> [--replace-with-x]

<input.docx>   : path to the tagged docx (e.g. refernce_template_unlocked_forced_escaped_controls_tagged.docx)
<mapping.json> : JSON file mapping control tag/title -> boolean, example:
                 { "pact_assessmentepdecrgepcchk_insert_othernumber_insert__asso": true,
                   "vemarketing_claims_are_validatedchk_marketing_claims_have_not_be": false }
--replace-with-x : optional flag; when setting a content-control's Checked property fails,
                   the script will replace the control with the glyph "X" (visible) instead
                   of leaving the original token.

Output:
  <input>_filled.docx (saved next to input)
"""

import sys, os, json, re
try:
    import win32com.client as com
except Exception as e:
    print("Requires pywin32 (win32com). Run in Windows with pywin32 installed.")
    raise

def load_json(p):
    with open(p, "r", encoding="utf-8") as fh:
        return json.load(fh)

def sanitize_key(k):
    return (k or "").strip()

def main():
    if len(sys.argv) < 3:
        print("Usage: python set_checkboxes_from_json.py <input.docx> <mapping.json> [--replace-with-x]")
        return
    IN = sys.argv[1]
    MAPF = sys.argv[2]
    replace_with_x = "--replace-with-x" in sys.argv[3:]

    if not os.path.exists(IN):
        print("Input file not found:", IN); return
    if not os.path.exists(MAPF):
        print("Mapping JSON not found:", MAPF); return

    mapping = load_json(MAPF)
    # normalize mapping keys
    mapping_norm = {sanitize_key(k): bool(v) for k, v in mapping.items()}

    word = com.Dispatch("Word.Application")
    word.Visible = False
    # Open read/write
    doc = word.Documents.Open(os.path.abspath(IN))
    try:
        total = doc.ContentControls.Count
        changed = 0
        notfound = 0
        replaced_text_count = 0

        for i in range(1, total + 1):
            try:
                cc = doc.ContentControls.Item(i)
            except Exception:
                continue
            # only care about checkbox type (8)
            try:
                cctype = int(cc.Type)
            except Exception:
                cctype = None
            if cctype != 8:
                continue

            tag = (cc.Tag or "").strip()
            title = (cc.Title or "").strip()
            key = tag if tag else title
            if not key:
                # attempt to build a fallback key from small preview text (best-effort)
                try:
                    preview = cc.Range.Text[:80].strip()
                    key = preview
                except Exception:
                    key = ""

            if not key:
                notfound += 1
                continue

            if key in mapping_norm:
                val = mapping_norm[key]
                # remove any literal tokens inside cc.Range (like <<CHK>>)
                try:
                    text_before = cc.Range.Text
                    # remove the token occurrences
                    new_text = re.sub(r'<<CHK>>', '', text_before)
                    if new_text != text_before:
                        cc.Range.Text = new_text
                        replaced_text_count += 1
                except Exception:
                    # ignore edit failures here
                    pass

                # Try to set content-control Checked property
                try:
                    cc.Checked = bool(val)
                    changed += 1
                    print(f"[SET] tag/title={key!r} -> {val}")
                except Exception as e:
                    print(f"[WARN] failed cc.Checked for key={key!r}: {e}")
                    # fallback: replace control with simple text 'X' or clear
                    try:
                        if replace_with_x and val:
                            cc.Range.Text = "X"
                        else:
                            # empty text when false
                            cc.Range.Text = ""
                        changed += 1
                        print(f"       fallback replacement done for {key!r}")
                    except Exception as e2:
                        print(f"       fallback failed too for {key!r}: {e2}")
            else:
                notfound += 1

        base, ext = os.path.splitext(IN)
        out = base + "_filled.docx"
        doc.SaveAs(os.path.abspath(out))
        print("Saved:", out)
        print("Summary: total controls:", total, "changed:", changed, "not matched:", notfound,
              "removed_text_inside_controls:", replaced_text_count)
    finally:
        doc.Close(False)
        word.Quit()

if __name__ == "__main__":
    main()


# python set_checkboxes_from_json_2.py refernce_template_unlocked_forced_escaped_controls_tagged.docx mapping.json
