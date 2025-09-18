# convert_glyphs_to_controls_safe_fixed.py
"""
Safe converter: replace checkbox glyphs (☐, ☑) with Word checkbox content-controls.
Use on Windows with MS Word + pywin32.

Usage:
  python convert_glyphs_to_controls_safe_fixed.py <input.docx>

Output:
  <input>_controls.docx (or timestamped if exists)
"""

import os, sys, time
import win32com.client as com
from win32com.client import constants as _const_module

# ---------- CONFIG ----------
INPUT_DOCX = None   # or set path here
# ----------------------------

def get_const(name, default):
    try:
        return getattr(_const_module, name)
    except Exception:
        return default

# Fallback numeric values used by Word interop
WD_FIND_STOP = get_const("wdFindStop", 0)         # 0
WD_COLLAPSE_START = get_const("wdCollapseStart", 1)  # 1
WD_COLLAPSE_END = get_const("wdCollapseEnd", 0)      # 0 (note: Word's numeric mapping can vary; using common values)
WD_CHARACTER = get_const("wdCharacter", 1)        # unit for MoveEnd
WD_MAIN_TEXT_STORY = get_const("wdMainTextStory", 1)

def sanitize_tag(s):
    s = (s or "").strip()
    if not s:
        return "auto_checkbox"
    out = []
    for ch in s:
        if ch.isalnum() or ch == "_":
            out.append(ch)
        elif ch in (" ", "-", "/"):
            out.append("_")
    t = "".join(out).lower() or "auto_checkbox"
    return (t[:60]).rstrip("_")

def process_story_range(story_range, rng_doc_end):
    if story_range is None:
        return 0
    converted = 0
    rng = story_range.Duplicate
    rng.Start = story_range.Start
    rng.End = story_range.End

    # We'll search for both checked (☑) and unchecked (☐)
    for glyph in ("☑", "☐"):
        f = rng.Find
        # Clear formatting if available
        try:
            f.ClearFormatting()
            f.Replacement.ClearFormatting()
        except Exception:
            pass
        f.Text = glyph
        # safe flags
        try:
            f.MatchCase = True
        except Exception:
            pass
        try:
            f.MatchWildcards = False
        except Exception:
            pass
        try:
            f.Wrap = WD_FIND_STOP
        except Exception:
            pass

        # iterate finds
        while True:
            try:
                found = f.Execute()
            except Exception:
                # If Execute fails, break to avoid infinite loop
                break
            if not found:
                break
            try:
                hit = f.Parent
            except Exception:
                break
            was_checked = (glyph == "☑")

            # Create a label range immediately after the hit (to use as Title/Tag)
            lbl = hit.Duplicate
            # Collapse to end (so label starts at the end of the glyph)
            try:
                lbl.Collapse(WD_COLLAPSE_END)
            except Exception:
                # Try alternate numeric fallback
                try:
                    lbl.Collapse(0)
                except Exception:
                    pass

            # extend label until next glyph or paragraph break
            cset = set(["☐", "☑", "\r"])
            unit_char = WD_CHARACTER
            while True:
                try:
                    if lbl.End >= rng.End:
                        break
                    next_char = lbl.Document.Range(lbl.End, lbl.End+1).Text
                except Exception:
                    break
                if next_char in cset:
                    break
                try:
                    lbl.MoveEnd(Unit=unit_char, Count=1)
                except Exception:
                    # fallback to manual End increment
                    try:
                        lbl.End = lbl.End + 1
                    except Exception:
                        break

            label_text = (lbl.Text or "").strip()

            # Remove the glyph text (replace with nothing; if spacing needed, ensure space inserted)
            try:
                hit.Text = ""
            except Exception:
                try:
                    hit.Delete()
                except Exception:
                    pass

            # Ensure a space before label if the label is immediate next to where glyph was
            try:
                if lbl.Start > hit.Start:
                    if hit.Document.Range(hit.Start, hit.Start+1).Text != " ":
                        hit.Text = " "
            except Exception:
                pass

            # Add checkbox content control at 'hit' range (which may be an insertion point or single space)
            try:
                cc = rng.Document.ContentControls.Add(get_const("wdContentControlCheckBox", 8), hit)
                try:
                    cc.Checked = was_checked
                except Exception:
                    # Some Word versions don't allow setting Checked directly; ignore
                    pass
                tt = label_text or "checkbox"
                cc.Tag = sanitize_tag(tt)
                cc.Title = tt
                converted += 1
                # move search start forward to avoid re-finding same area
                try:
                    rng.Start = cc.Range.End
                except Exception:
                    try:
                        rng.Start = hit.End + 1
                    except Exception:
                        rng.Start = hit.End
                # rebind the Find object to new rng
                f = rng.Find
                try:
                    f.ClearFormatting(); f.Replacement.ClearFormatting()
                except Exception:
                    pass
                f.Text = glyph
                try:
                    f.Wrap = WD_FIND_STOP
                except Exception:
                    pass
            except Exception:
                # If adding CC failed, skip forward a bit to avoid infinite loop
                try:
                    rng.Start = hit.End + 1
                except Exception:
                    try:
                        rng.Start = hit.End
                    except Exception:
                        break

    return converted

def main():
    global INPUT_DOCX
    infile = INPUT_DOCX or (sys.argv[1] if len(sys.argv) > 1 else None)
    if not infile:
        print("Usage: python convert_glyphs_to_controls_safe_fixed.py <input.docx>")
        return
    infile = os.path.abspath(infile)
    if not os.path.exists(infile):
        print("Input file not found:", infile); return

    outpath = infile.replace(".docx", "_controls.docx")
    if os.path.exists(outpath):
        base = outpath.replace(".docx", "")
        outpath = f"{base}_{int(time.time())}.docx"

    word = com.Dispatch("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(infile)
    except Exception as e:
        print("Failed to open document in Word COM:", e)
        try:
            word.Quit()
        except:
            pass
        return

    total_converted = 0
    try:
        # Try to unprotect if needed (no password)
        try:
            if getattr(doc, "ProtectionType", -1) != -1:
                try:
                    doc.Unprotect()
                except Exception:
                    pass
        except Exception:
            pass

        story = doc.StoryRanges(WD_MAIN_TEXT_STORY)
        processed = 0
        while story is not None:
            processed += 1
            converted = process_story_range(story, doc.Range().End)
            total_converted += converted
            try:
                story = story.NextStoryRange
            except Exception:
                break

        # Save as new file
        try:
            doc.SaveAs2(outpath)
        except Exception:
            try:
                doc.SaveAs(outpath)
            except Exception as e2:
                print("SaveAs failed:", e2)
                try:
                    doc.Save()
                    import shutil
                    shutil.copy2(infile, outpath)
                except Exception:
                    pass

        print(f"Processed {processed} story ranges.")
        print(f"Converted {total_converted} glyphs to checkbox content-controls.")
        print("Saved new file at:", outpath)
    finally:
        try:
            doc.Close(False)
        except:
            pass
        try:
            word.Quit()
        except:
            pass

if __name__ == "__main__":
    main()
