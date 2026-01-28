import fitz
import os

pdf_path = 'Priya - PGT-A report_withlogo.pdf'
doc = fitz.open(pdf_path)
page = doc[0]

print("--- Images on Page 1 ---")
image_list = page.get_images(full=True)
for img in image_list:
    xref = img[0]
    rects = page.get_image_rects(xref)
    for r in rects:
        print(f"Xref {xref}: Position {r}")

print("\n--- Text near expected signature (end of last page) ---")
last_page = doc[-2] # Assuming embryo details are before methodology or at end. Actually let's check page 4/5.
for i in [3, 4]:
    if i < len(doc):
        print(f"\nPage {i+1} blocks:")
        blocks = doc[i].get_text("blocks")
        for b in blocks:
            if "approved" in b[4].lower() or "reviewed" in b[4].lower():
                print(f"Text: '{b[4].strip()}' at {b[:4]}")

print("\n--- PNDT Disclaimer check ---")
text_instances = page.search_for("This test does not reveal sex")
if text_instances:
    print(f"Disclaimer Pos: {text_instances[0]}")
