# convert_chk_tokens_to_controls.py
# Usage:
#   python convert_chk_tokens_to_controls.py "C:\path\to\refernce_template_unlocked_forced_escaped.docx"
#
# Requires: pywin32 (pip install pywin32)
import sys
import os
import traceback
import win32com.client as com

if len(sys.argv) < 2:
    print("Usage: python convert_chk_tokens_to_controls.py <docx_path> [out_path]")
    sys.exit(1)

IN_PATH = sys.argv[1]
OUT_PATH = sys.argv[2] if len(sys.argv) > 2 else IN_PATH.replace(".docx", "_controls.docx")

# Word constants (we only need small subset)
wdFindStop = 0
wdCollapseEnd = 0
wdCharacter = 1
wdContentControlCheckBox = 8

def sanitize_tag(s: str) -> str:
    s = (s or "").strip()
    if not s:
        return "chk"
    out = []
    for ch in s:
        if ch.isalnum() or ch == "_":
            out.append(ch)
        elif ch in (" ", "-", "/"):
            out.append("_")
    t = "".join(out).lower()[:64]
    return t or "chk"

def process_story_range(doc, story_range):
    """Searches for literal token <<CHK>> and replaces with checkbox content control.
       Returns number converted in this story.
    """
    converted = 0
    rng = story_range
    f = rng.Find
    f.ClearFormatting()
    f.Replacement.ClearFormatting()
    # Literal token to find
    f.Text = "<<CHK>>"
    f.MatchCase = False
    f.MatchWholeWord = False
    f.MatchWildcards = False
    f.Wrap = wdFindStop

    iteration = 0
    # We use Execute() in a loop; after each replacement we move the search start to the end of inserted CC
    while True:
        found = f.Execute()
        if not found:
            break
        iteration += 1
        hit = f.Parent  # Range of the found token
        sample = (hit.Text or "").replace("\r","\\r").replace("\n","\\n")
        print(f"  Found token (storyType={getattr(story_range, 'StoryType', '?')}) sample={repr(sample[:60])}")

        # Try to get label right after token (a small run of text until next token or paragraph break)
        lbl = hit.Duplicate
        try:
            lbl.Collapse(0)  # collapse to start
        except Exception:
            pass

        # Extend label until next token occurrence or paragraph/line break (character stepping)
        try:
            while lbl.End < rng.End:
                nxt = lbl.Document.Range(lbl.End, lbl.End+1).Text
                if not nxt:
                    break
                if nxt in ("<", ">", "\r", "\n"):  # stop on new tokens or paragraph
                    break
                # also stop if we see start of another "<<"
                # check the next two chars
                look = lbl.Document.Range(lbl.End, min(lbl.End+2, rng.End)).Text
                if look.startswith("<<"):
                    break
                lbl.MoveEnd(Unit=wdCharacter, Count=1)
        except Exception:
            # defensive: if stepping fails, bail to avoid infinite loop
            pass

        label_text = (lbl.Text or "").strip()
        if label_text:
            print(f"    label preview: {repr(label_text[:80])}")
        else:
            print("    no label found (empty)")

        # Replace the token text with a content-control checkbox anchored at the hit range
        try:
            # delete token
            hit.Text = ""
            # ensure spacing before label if needed
            if lbl.Start > hit.Start:
                # if the char right after the inserted spot is not a space, insert a space so label sits nicely
                if hit.Document.Range(hit.Start, hit.Start+1).Text != " ":
                    hit.Text = " "
                    # collapse to end of run
                    try:
                        hit.Collapse(wdCollapseEnd)
                    except Exception:
                        pass

            cc = doc.ContentControls.Add(wdContentControlCheckBox, hit)
            # default unchecked; if later you want checked by token, detect variant token e.g. <<CHK_ON>>
            try:
                cc.Checked = False
            except Exception:
                # some Word versions may raise — ignore
                pass
            cc.Tag = sanitize_tag(label_text or "chk")
            cc.Title = label_text[:255] if label_text else "chk"
            converted += 1
            print(f"    inserted checkbox control tag={cc.Tag}")
            # Move the parent search range start to the end of the inserted control
            rng.Start = cc.Range.End
            f = rng.Find
            f.ClearFormatting(); f.Replacement.ClearFormatting()
            f.Text = "<<CHK>>"; f.MatchWildcards = False; f.Wrap = wdFindStop
        except Exception as e:
            print("    ERROR inserting control:", e)
            traceback.print_exc()
            # advance search to avoid infinite loop
            try:
                rng.Start = hit.End + 1
                f = rng.Find
                f.Text = "<<CHK>>"; f.MatchWildcards = False; f.Wrap = wdFindStop
            except Exception:
                break

    return converted

def main():
    print("Opening Word (COM)...")
    word = com.Dispatch("Word.Application")
    word.Visible = False
    # try to open read-write; allow Word to repair if needed by setting AddToRecentFiles=False (optional)
    doc = None
    try:
        doc = word.Documents.Open(os.path.abspath(IN_PATH))
    except Exception as e:
        print("Failed to open document. Trying Unprotect / OpenReadOnly fallback. Error:", e)
        # try reopening read-only then copy - but simplest: re-raise
        raise

    total_converted = 0
    try:
        # Try to unprotect if protected with empty password
        try:
            if doc.ProtectionType != -1:  # -1 = wdNoProtection
                try:
                    doc.Unprotect("")  # try empty password
                    print("Document was protected: attempted Unprotect(\"\")")
                except Exception:
                    print("Document is protected and could not be automatically unprotected. Please unprotect manually and re-run.")
        except Exception:
            # some doc types may not expose ProtectionType; ignore
            pass

        # iterate story ranges
        story = doc.StoryRanges(1)  # wdMainTextStory
        story_n = 0
        while story is not None:
            story_n += 1
            print(f"Processing story #{story_n} (Type={getattr(story, 'StoryType', '?')})")
            converted = process_story_range(doc, story)
            print(f"  converted in this story: {converted}")
            total_converted += converted
            try:
                story = story.NextStoryRange
            except Exception:
                story = None

        if total_converted:
            doc.SaveAs(os.path.abspath(OUT_PATH))
            print("Saved new document with controls at:", OUT_PATH)
        else:
            print("Converted 0 tokens — no changes made.")
    finally:
        if doc:
            doc.Close(False)
        word.Quit()
        print("Done. Total converted:", total_converted)

if __name__ == "__main__":
    main()
