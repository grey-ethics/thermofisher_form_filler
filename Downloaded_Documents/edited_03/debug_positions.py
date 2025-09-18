import sys, time
from pathlib import Path

try:
    import win32com.client as win32
    from win32com.client import constants as wdconst
except Exception:
    print("Requires: pip install pywin32 and Microsoft Word on Windows.")
    raise

def open_word(visible=True):
    word = win32.Dispatch("Word.Application")
    word.Visible = visible
    try:
        word.ScreenUpdating = False
    except Exception:
        pass
    return word

def open_doc(word, path):
    return word.Documents.Open(str(Path(path).resolve()))

def main():
    if len(sys.argv) < 2:
        print("Usage: python debug_positions.py <docx>")
        sys.exit(1)

    docx = sys.argv[1]
    word = open_word(visible=True)
    try:
        doc = open_doc(word, docx)

        # Force a sane view
        try:
            word.ActiveWindow.View.Type = wdconst.wdPrintView
        except Exception:
            pass
        try:
            word.ActiveWindow.View.SeekView = wdconst.wdSeekMainDocument
        except Exception:
            pass

        total = doc.ContentControls.Count
        print(f"Total ContentControls in doc: {total}")

        # Count by type
        type_counts = {}
        for c in doc.ContentControls:
            t = int(c.Type)
            type_counts[t] = type_counts.get(t, 0) + 1
        print("Counts by Type (Word constants):", type_counts)
        # 8 = checkbox, 3 = combo, 4 = drop-down, 6 = date

        got_sel = got_rng = 0
        sample = 0

        for i, c in enumerate(doc.ContentControls, start=1):
            t = int(c.Type)
            if t != 8:   # only checkboxes for now
                continue

            # Try Selection-based coordinates
            try:
                c.Range.Select()
                sel = word.Selection
                page = sel.Information(wdconst.wdActiveEndPageNumber)
                x = sel.Information(wdconst.wdHorizontalPositionRelativeToPage)
                y = sel.Information(wdconst.wdVerticalPositionRelativeToPage)
                if page and x not in (None, False) and y not in (None, False):
                    got_sel += 1
                    if sample < 5:
                        print(f"[SEL] idx={i} tag='{c.Tag}' title='{c.Title}' -> page={page}, x={x}, y={y}")
                        sample += 1
            except Exception as e:
                pass

            # Try Range-based (often unreliable)
            try:
                r = c.Range
                page2 = r.Information(wdconst.wdActiveEndPageNumber)
                x2 = r.Information(wdconst.wdHorizontalPositionRelativeToPage)
                y2 = r.Information(wdconst.wdVerticalPositionRelativeToPage)
                if page2 and x2 not in (None, False) and y2 not in (None, False):
                    got_rng += 1
            except Exception:
                pass

        print(f"Checkboxes with Selection coords: {got_sel}")
        print(f"Checkboxes with Range coords:     {got_rng}")

    finally:
        try:
            word.ScreenUpdating = True
        except Exception:
            pass
        # keep Word open so you can see the selection jumping. Close manually when done.
        # If you want auto-close, uncomment below:
        # try: doc.Close(False)
        # except: pass
        # try: word.Quit()
        # except: pass

if __name__ == "__main__":
    main()
