import fitz

def check_title(pdf_path):
    print(f"Checking {pdf_path}...")
    doc = fitz.open(pdf_path)
    page = doc[0]
    blocks = page.get_text("blocks")
    for b in blocks:
        if "Preimplantation Genetic" in b[4]:
            text = b[4].strip()
            lines = text.split('\n')
            print(f"Title: '{text}'")
            print(f"Number of lines: {len(lines)}")
            print(f"Bbox: {b[:4]}")
            return len(lines)
    return 0

check_title("test_report.pdf")
print("-" * 20)
check_title("Priya - PGT-A report_withlogo.pdf")
