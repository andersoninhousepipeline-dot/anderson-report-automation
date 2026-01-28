import fitz

pdf_path = 'test_report.pdf'
doc = fitz.open(pdf_path)
# Check all pages
for i in range(len(doc)):
    page = doc[i]
    text_instances = page.search_for('reviewed and approved by')
    if text_instances:
        print(f"Page {i+1}: Found approval line")
        blocks = page.get_text("dict")["blocks"]
        for b in blocks:
            if "lines" in b:
                for l in b["lines"]:
                    for s in l["spans"]:
                        if "reviewed and approved" in s["text"]:
                            print(f"  Text: {s['text']}")
                            print(f"  Color: {hex(s['color'])}")
                            print(f"  Font: {s['font']}")
