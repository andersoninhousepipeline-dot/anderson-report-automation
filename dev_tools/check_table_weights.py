import fitz

def check_table_weights(pdf_path, page_num):
    doc = fitz.open(pdf_path)
    page = doc[page_num]
    
    # Analyze the detail table area
    # Typical y range for detail table on embryo page
    clip = fitz.Rect(50, 150, 540, 350)
    print(f"--- Page {page_num+1} Detail Table Weight Analysis ---")
    
    detailed = page.get_text("dict", clip=clip)
    for block in detailed["blocks"]:
        if "lines" in block:
            for line in block["lines"]:
                for span in line["spans"]:
                    text = span['text'].strip()
                    if text:
                        is_bold = "Bold" in span['font'] or span['flags'] & 2**4
                        print(f"Text: '{text}'")
                        print(f"  Font: {span['font']}, Size: {span['size']:.2f}, Bold: {bool(is_bold)}")

# Analyze Page 4
check_table_weights("Priya - PGT-A report_withlogo.pdf", 3)
