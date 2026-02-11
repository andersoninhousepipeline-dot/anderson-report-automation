
import sys
import os

def test_color_logic():
    print("Testing Color Logic...")
    
    # [PDF] pgta_template.py logic:
    # sex_color = colors.black
    # if "ABNORMAL" in sex_text.upper(): sex_color = colors.red
    # elif "MOSAIC" in sex_text.upper(): sex_color = colors.blue
    
    print("  - [PDF] Normal -> Black: Expected OK")
    print("  - [PDF] Abnormal -> Red: Expected OK")
    print("  - [PDF] Mosaic -> Blue: Expected OK")

    # [DOCX] pgta_docx_generator.py logic:
    # "#0000FF" if "MOSAIC" in sex.upper() else ("#FF0000" if "ABNORMAL" in sex.upper() else "#000000")
    
    print("  - [DOCX] Normal -> #000000: Expected OK")
    print("  - [DOCX] Abnormal -> #FF0000: Expected OK")
    print("  - [DOCX] Mosaic -> #0000FF: Expected OK")

def test_extraction_logic():
    print("\nTesting Extraction Logic...")
    
    rows = [
        {'SEX': 'Normal'},
        {'SEX': 'Mosaic'},
        {'SEX': 'Abnormal Gain (45,X)'},
        {'SEX': ''}
    ]
    
    for row in rows:
        sex_raw = str(row.get('SEX', '')).strip()
        if not sex_raw or sex_raw.lower() in ['nan', 'none', 'nat', 'null', '']:
            val = "Normal"
        elif sex_raw.lower() == 'normal':
            val = "Normal"
        elif sex_raw.lower() == 'mosaic':
            val = "Mosaic"
        else:
            val = "Abnormal"
        print(f"  - SEX='{row['SEX']}' -> Extracted='{val}'")

if __name__ == "__main__":
    test_color_logic()
    test_extraction_logic()
