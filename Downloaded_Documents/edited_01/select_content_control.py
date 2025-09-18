# select_content_control.py
import sys
import win32com.client
from pathlib import Path

def main():
    if len(sys.argv) < 3:
        print("Usage: python select_content_control.py <docx_path> <index_1_based>")
        return

    doc_path = Path(sys.argv[1]).expanduser().resolve()
    idx = int(sys.argv[2])

    word = win32com.client.Dispatch("Word.Application")
    word.Visible = True   # show Word so you can see selection

    # Open without updating links, read-only False
    doc = word.Documents.Open(str(doc_path), ReadOnly=False)

    try:
        cc_count = doc.ContentControls.Count
        print(f"Opened: {doc_path}")
        print(f"Total content controls: {cc_count}")

        if idx < 1 or idx > cc_count:
            print(f"Index out of range (1..{cc_count})")
            return

        # Word collections are 1-based
        cc = doc.ContentControls(idx)
        cc.Range.Select()

        # Try to highlight (visual aid). If not supported, ignore.
        try:
            cc.Range.HighlightColorIndex = 7  # wdYellow
        except Exception:
            pass

        # Print some info
        print(f"\nSelected content control #{idx}")
        print(" Type (int):", getattr(cc, "Type", "N/A"))
        print(" Tag:", repr(getattr(cc, "Tag", "")))
        print(" Title:", repr(getattr(cc, "Title", "")))
        txt = cc.Range.Text.replace("\r", "\\r").replace("\n", "\\n")
        print(" Range preview:", txt[:300])
        if getattr(cc, "Type", None) == 8:  # checkbox
            try:
                cb = cc.CheckBox
                print(" Checkbox.CheckedSymbol:", getattr(cb, "CheckedSymbol", None))
                print(" Checkbox.UncheckedSymbol:", getattr(cb, "UncheckedSymbol", None))
            except Exception:
                print(" Checkbox properties not accessible")
            try:
                print(" Current Checked value:", bool(getattr(cc, "Checked")))
            except Exception:
                print(" Current Checked value: (not available)")

        print("\nWord should have scrolled to and selected the control. Inspect it, then close Word when done.")
    finally:
        # leave Word open for user inspection (do not doc.Close())
        pass

if __name__ == "__main__":
    main()
