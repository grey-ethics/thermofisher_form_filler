# convert_glyphs_to_controls.py
import win32com.client as com

DOC = r"C:\Users\K Santosh Kumar\Desktop\HEALTHARK\04_thermofisher\Downloaded_Documents\edited\refernce_template_unlocked.docx"

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
    return t

def process_range(rng):
    wdFindStop = 0
    wdCollapseEnd = 0
    wdContentControlCheckBox = 8

    f = rng.Duplicate.Find
    f.ClearFormatting()
    f.Replacement.ClearFormatting()
    # Find either ☐ (U+2610) or ☑ (U+2611)
    f.Text = "[\u2610\u2611]"
    f.MatchWildcards = True
    f.Wrap = wdFindStop

    while f.Execute():
        hit = f.Parent  # the glyph range
        was_checked = (hit.Text == "\u2611")

        # Label to the right: up to next checkbox or paragraph break
        lbl = hit.Duplicate
        lbl.Collapse(0)  # wdCollapseStart=1, End=0 (Word quirk via pywin32; try both if needed)
        # Extend until next ☐/☑ or paragraph end
        # Workaround: step char-by-char
        cset = set(["\u2610", "\u2611", "\r"])
        while True:
            if lbl.End >= rng.End:
                break
            next_char = lbl.Document.Range(lbl.End, lbl.End+1).Text
            if next_char in cset:
                break
            lbl.MoveEnd(Unit=1, Count=1)  # wdCharacter
        label_text = lbl.Text.strip()

        # Replace glyph with control
        hit.Text = ""
        # If label starts immediately, add a space
        if lbl.Start > hit.Start:
            if hit.Document.Range(hit.Start, hit.Start+1).Text != " ":
                hit.Text = " "
                hit.Collapse(wdCollapseEnd)

        cc = rng.Document.ContentControls.Add(wdContentControlCheckBox, hit)
        try:
            cc.Checked = was_checked
        except Exception:
            pass
        cc.Tag = sanitize_tag(label_text)
        cc.Title = label_text

        # Continue after this control
        rng.Start = cc.Range.End
        f = rng.Find
        f.Text = "[\u2610\u2611]"
        f.MatchWildcards = True
        f.Wrap = wdFindStop

def main():
    word = com.Dispatch("Word.Application")
    word.Visible = False
    doc = word.Documents.Open(DOC)
    try:
        # Process all stories (body + headers/footers)
        story = doc.StoryRanges(1)  # wdMainTextStory
        while True:
            process_range(story)
            try:
                story = story.NextStoryRange
                if story is None: break
            except Exception:
                break
        doc.Save()
        print("Converted and saved.")
    finally:
        doc.Close(False)
        word.Quit()

if __name__ == "__main__":
    main()
