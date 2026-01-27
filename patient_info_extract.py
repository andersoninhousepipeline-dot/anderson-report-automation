import fitz

def extract_patient_info_styling(pdf_path, page_num):
    doc = fitz.open(pdf_path)
    page = doc[page_num]
    
    # Search for labels
    labels = ["Patient name", "PIN"]
    print(f"--- Page {page_num+1} Patient Info Analysis ---")
    
    for label in labels:
        text_instances = page.search_for(label)
        if text_instances:
            inst = text_instances[0]
            print(f"\nLabel: '{label}' found at: {inst}")
            
            # Find drawings (fills) overlapping with this text
            drawings = page.get_drawings()
            for d in drawings:
                rect = d['rect']
                # Check for intersection or if the rect spans the width and contains this y
                if rect.intersects(inst) or (rect.y0 < inst.y0 and rect.y1 > inst.y1 and rect.width > 400):
                    print(f"  Background Rect: {rect}")
                    print(f"  Fill: {d.get('fill')}")
                    print(f"  Stroke: {d.get('color')}")
            
            # Get text dict for specific styling
            # Expand clip to likely end of line
            clip = fitz.Rect(inst.x0, inst.y0 - 2, 540, inst.y1 + 2)
            detailed = page.get_text("dict", clip=clip)
            for block in detailed["blocks"]:
                if "lines" in block:
                    for line in block["lines"]:
                        for span in line["spans"]:
                            if label in span["text"] or any(x in span["text"] for x in ["Mrs", "Priya", "AND2563"]):
                                print(f"  Span: '{span['text']}'")
                                print(f"    F: {span['font']}, Size: {span['size']:.2f}, Color: {hex(span['color'])}")
                                is_bold = "Bold" in span['font'] or span['flags'] & 2**4
                                print(f"    Bold: {bool(is_bold)}")

# Analyze Page 1 (cover page)
extract_patient_info_styling("Priya - PGT-A report_withlogo.pdf", 0)
