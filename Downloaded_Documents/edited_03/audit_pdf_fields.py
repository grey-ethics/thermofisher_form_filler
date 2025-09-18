import fitz

def audit(pdf_path: str):
    doc = fitz.open(pdf_path)
    total = 0
    by_page = []
    for pno in range(len(doc)):
        page = doc[pno]
        widgets = page.widgets() or []
        total += len(widgets)
        rows = []
        for w in widgets:
            # w.field_type: 'Tx'=text, 'Btn'=button/checkbox/radio, 'Ch'=choice, etc.
            rows.append({
                "name": w.field_name,
                "type": w.field_type,
                "rect": [round(v,2) for v in w.rect],  # x1,y1,x2,y2 (points)
                "value": w.field_value,
                "pg": pno+1
            })
        if rows:
            by_page.append((pno+1, rows))
    print(f"PDF fields found: {total}")
    for pg, rows in by_page:
        print(f"\nPage {pg}:")
        for r in rows:
            print(f"  [{r['type']}] {r['name']}  rect={r['rect']} value={r['value']}")

if __name__ == "__main__":
    audit("form_preview.pdf")
