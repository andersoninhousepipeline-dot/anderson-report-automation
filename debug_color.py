import fitz
import os

pdf_path = 'Priya - PGT-A report_withlogo.pdf'
if not os.path.exists(pdf_path):
    print(f"Error: {pdf_path} not found")
    exit(1)

doc = fitz.open(pdf_path)
page = doc[3]
text_instances = page.search_for('reviewed and approved by')

if text_instances:
    inst = text_instances[0]
    print(f"Disclaimer found at: {inst}")
    
    drawings = page.get_drawings()
    for d in drawings:
        rect = d['rect']
        # Check if the rect contains the text or is very close
        if rect.intersects(inst) or (rect.y0 < inst.y0 and rect.y1 > inst.y1 and abs(rect.x0 - inst.x0) < 100):
            print(f"Found overlapping rect: {rect}")
            print(f"Fill color: {d.get('fill')}")
            print(f"Stroke color: {d.get('color')}")
else:
    print("Disclaimer text not found on page 0")
