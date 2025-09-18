# convert_glyphs_to_controls_safe.py
"""
Converts glyph checkboxes (☐, ☑) into Word content-control checkboxes using COM.
Saves output to a new file: <input>_controls.docx

Requirements: Windows + MS Word + pywin32 installed
( pip install pywin32 )
"""

import os, sys, time
import win32com.client as com
from win32com.client import constants

# ---------- CONFIG ----------
# set to None to accept filename via argv, otherwise edit below
INPUT_DOCX = None
# ----------------------------

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
    # ensure tag isn't too long
    return (t[:60]).rstrip("_")

def process_story_range(story_range, doc_end):
    """
    Process glyphs inside the provided story_range (Range object).
    We'll do two passes: '☑' then '☐' to reliably capture checked state.
    """
    if story_range is None:
        return 0
    converted = 0
    rng = story_range.Duplicate
    rng.Start = story_range.Start
    rng.End = story_range.End

    # two passes so we detect checked vs unchecked properly
    for glyph in ("☑", "☐"):
        f = rng.Find
        f.ClearFormatting()
        f.Replacement.ClearFormatting()
        f.Text = glyph
        f.MatchCase = True
        f.MatchWildcards = False
        f.Wrap = constants.wdFindStop
        # search forward until no more occurrences in this range
        while True:
            found = f.Execute()
            if not found:
                break
            try:
                hit = f.Parent  # the Range containing the found glyph
            except Exception:
                break
            was_checked = (glyph == "☑")
            # build label: from after hit to next glyph or paragraph end
            lbl = hit.Duplicate
            try:
                lbl.Collapse(constants.wdCollapseEnd)
            except Exception:
                # collapse may use different constants on some systems; try numeric fallback:
                try:
                    lbl.Collapse(0)
                except:
                    pass
            # extend label until next checkbox glyph or paragraph break
            cset = set(["☐","☑","\r"])
            # MoveEnd with Unit=wdCharacter (1)
            unit_char = constants.wdCharacter if hasattr(constants, "wdCharacter") else 1
            while True:
                if lbl.End >= rng.End:
                    break
                next_char = lbl.Document.Range(lbl.End, lbl.End+1).Text
                if next_char in cset:
                    break
                try:
                    lbl.MoveEnd(Unit=unit_char, Count=1)
                except Exception:
                    # if MoveEnd fails, fall back to incrementing End index (best-effort)
                    try:
                        lbl.End = lbl.End + 1
                    except Exception:
                        break
            label_text = lbl.Text.strip()
            # Remove the glyph from document (replace with nothing or a space if needed)
            try:
                hit.Text = ""
            except Exception:
                # fallback: replace using range
                try:
                    hit.Delete()
                except:
                    pass
            # ensure a space before label if there was no spacing originally
            if lbl.Start > hit.Start:
                try:
                    if hit.Document.Range(hit.Start, hit.Start+1).Text != " ":
                        hit.Text = " "
                except Exception:
                    pass

            # create the checkbox content control at the location of 'hit' (which is now a small range)
            try:
                cc = rng.Document.ContentControls.Add(constants.wdContentControlCheckBox, hit)
                try:
                    cc.Checked = was_checked
                except Exception:
                    pass
                tt = label_text or "checkbox"
                cc.Tag = sanitize_tag(tt)
                cc.Title = tt
                converted += 1
                # advance the search range forward (start after the newly created CC)
                rng.Start = cc.Range.End
                # rewire find to current rng
                f = rng.Find
                f.ClearFormatting()
                f.Replacement.ClearFormatting()
                f.Text = glyph
                f.MatchCase = True
                f.MatchWildcards = False
                f.Wrap = constants.wdFindStop
            except Exception as e:
                # If adding control fails, try to skip past this position to avoid infinite loop
                try:
                    rng.Start = hit.End + 1
                except Exception:
                    rng.Start = hit.End
    return converted

def main():
    global INPUT_DOCX
    infile = INPUT_DOCX or (sys.argv[1] if len(sys.argv) > 1 else None)
    if not infile:
        print("Usage: python convert_glyphs_to_controls_safe.py <input.docx>")
        return
    infile = os.path.abspath(infile)
    if not os.path.exists(infile):
        print("Input file not found:", infile); return

    outpath = infile.replace(".docx", "_controls.docx")
    # ensure we don't overwrite existing file - append timestamp if needed
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
        # If doc is protected and unprotectable w/o password, try to call Unprotect()
        try:
            if getattr(doc, "ProtectionType", -1) != -1:
                # If protected, attempt to unprotect (no password)
                try:
                    doc.Unprotect()
                except Exception:
                    pass
        except Exception:
            pass

        # Process all story ranges (main+headers/footers/etc)
        story = doc.StoryRanges(1)  # wdMainTextStory == 1
        processed_stories = 0
        while story is not None:
            processed_stories += 1
            converted = process_story_range(story, doc.Range().End)
            total_converted += converted
            # go to next story
            try:
                story = story.NextStoryRange
            except Exception:
                break

        # Save to new file
        try:
            doc.SaveAs2(outpath)
        except Exception:
            try:
                doc.SaveAs(outpath)
            except Exception as e2:
                print("Failed saving with SaveAs2/SaveAs:", e2)
                # fallback: Save and then copy
                doc.Save()
                import shutil
                shutil.copy2(infile, outpath)
        print(f"Converted {total_converted} glyphs to content-controls.")
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
