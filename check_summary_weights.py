import fitz

def check_summary_table(pdf_path):
    doc = fitz.open(pdf_path)
    page = doc[0]
    
    # Summary table area
    clip = fitz.Rect(50, 400, 540, 600)
    print("--- Page 1 Summary Table Weight Analysis ---")
    
    detailed = page.get_text("dict", clip=clip)
    for block in detailed["blocks"]:
        if "lines" in block:
            for line in block["lines"]:
                for span in line["spans"]:
                    text = span['text'].strip()
                    if text:
                        is_bold = "Bold" in span['font'] or span['flags'] & 2**4
                        print(f"Text: '{text}'")
                        print(f"  Font: {span['font']}, Size: {span['size']:.2f}, Bold: {bool(is_bold)}, Color: {hex(span['color'])}")

check_summary_table("Priya - PGT-A report_withlogo.pdf")
