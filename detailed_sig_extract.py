import fitz
import json

def extract_detailed_section(pdf_path, search_text):
    doc = fitz.open(pdf_path)
    # Search for the section on any page (embryo results usually have it)
    for i in range(len(doc)):
        page = doc[i]
        text_instances = page.search_for(search_text)
        if text_instances:
            print(f"--- Page {i+1} Detailed Analysis ---")
            inst = text_instances[0]
            # Analysis area: start from approval text, go down to end of page or significant footer
            y_start = inst.y0 - 20 # Give some breathing room at top
            y_end = page.rect.height - 65 # Stop before footer banner
            x_start = 72 # Standard left margin in source
            x_end = 540 # Standard right margin in source
            
            area = fitz.Rect(x_start, y_start, x_end, y_end)
            
            # Get dict for detailed spans
            detailed = page.get_text("dict", clip=area)
            
            # Get images in this area
            images = []
            for img in page.get_images(full=True):
                xref = img[0]
                rects = page.get_image_rects(xref)
                for r in rects:
                    if r.intersects(area):
                        images.append({'xref': xref, 'rect': r})
            
            # Print spans with detail
            print("\nTEXT SPANS:")
            for block in detailed["blocks"]:
                if "lines" in block:
                    for line in block["lines"]:
                        for span in line["spans"]:
                            print(f"Text: '{span['text']}'")
                            print(f"  F: {span['font']}, Size: {span['size']:.2f}, Color: {hex(span['color'])}")
                            print(f"  Bbox: {span['bbox']}")
                            # Identify weight/boldness from font name or flags
                            is_bold = "Bold" in span['font'] or span['flags'] & 2**4
                            print(f"  Bold: {bool(is_bold)}, Aligned-X: {span['bbox'][0]}")
            
            print("\nIMAGES:")
            for img in images:
                print(f"Image Xref {img['xref']}, Rect: {img['rect']}")
            
            return True
    return False

extract_detailed_section("Priya - PGT-A report_withlogo.pdf", "This report has been reviewed and approved by")
