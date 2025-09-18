# inspect_docx_xml.py
import zipfile, sys, os

DOCX_PATH = r"C:\Users\K Santosh Kumar\Desktop\HEALTHARK\04_thermofisher\Downloaded_Documents\edited\refernce_template_unlocked.docx"
# Use the filename you tried to open — set it above.

# If you have a line & column from the Word error, set here:
ERR_LINE = 2
ERR_COL = 33639

def show_context(xml_bytes, line, col, context_chars=200):
    try:
        xml = xml_bytes.decode("utf-8")
    except UnicodeDecodeError:
        xml = xml_bytes.decode("utf-8", errors="replace")
    # compute offset: sum of lengths up to line-1 + (col-1)
    lines = xml.splitlines(keepends=True)
    if line <= 0 or line > len(lines):
        print("Given line is outside file lines. File has", len(lines), "lines.")
        return
    # compute offset up to beginning of that line
    offset = sum(len(lines[i]) for i in range(line-1))
    offset += max(0, col-1)
    start = max(0, offset - context_chars)
    end = min(len(xml), offset + context_chars)
    snippet = xml[start:end]
    print(f"Showing context around line {line} col {col} (byte offset ~{offset}):")
    print("-" * 80)
    # show with visible markers
    marker_pos = offset - start
    print(snippet)
    print("-" * 80)
    print(" " * marker_pos + "^ <- error position (approx)")
    # also print a slice with non-printables escaped
    print("\nEscaped view:")
    esc = snippet.encode("unicode_escape").decode("ascii")
    print(esc)
    print("-" * 80)

def main():
    if not os.path.exists(DOCX_PATH):
        print("DOCX not found:", DOCX_PATH); return
    with zipfile.ZipFile(DOCX_PATH, "r") as z:
        try:
            data = z.read("word/document.xml")
        except KeyError:
            print("The docx has no word/document.xml inside!")
            return
    show_context(data, ERR_LINE, ERR_COL, context_chars=400)

if __name__ == "__main__":
    main()
