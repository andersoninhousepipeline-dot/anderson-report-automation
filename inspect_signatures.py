import fitz

pdf_path = 'Priya - PGT-A report_withlogo.pdf'
doc = fitz.open(pdf_path)
page = doc[3] # Embryo page usually has it

text_instances = page.search_for('reviewed and approved by')
if text_instances:
    inst = text_instances[0]
    print(f"Text found at: {inst}")
    
    # Get detailed text information
    blocks = page.get_text("dict")["blocks"]
    for b in blocks:
        if "lines" in b:
            for l in b["lines"]:
                for s in l["spans"]:
                    if "Anand Babu" in s["text"] or "Molecular Biologist" in s["text"]:
                        print(f"Text: '{s['text'].strip()}', Color: {hex(s['color'])}, Font: {s['font']}, Size: {s['size']}")
                    if "reviewed and approved" in s["text"]:
                        print(f"Approval Line Color: {hex(s['color'])}, Size: {s['size']}")
    
    # Check text blocks below this instance to find names
    all_blocks = page.get_text("blocks")
    # Sort blocks by y coordinate
    all_blocks.sort(key=lambda x: x[1])
    
    found_approved = False
    print("\n--- Blocks below 'reviewed and approved by' ---")
    for b in all_blocks:
        if "reviewed and approved by" in b[4]:
            found_approved = True
            continue
        if found_approved:
            # Only print if it's below the and reasonably close in x
            if b[1] > inst.y1:
                print(f"Text: '{b[4].strip()}' at {b[:4]}")
else:
    print("Approval text not found on page 4")
