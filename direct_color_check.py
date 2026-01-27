import fitz

def get_hex(rgb_int):
    r = (rgb_int >> 16) & 0xFF
    g = (rgb_int >> 8) & 0xFF
    b = rgb_int & 0xFF
    return f"#{r:02X}{g:02X}{b:02X}"

pdf_path = 'test_report.pdf'
doc = fitz.open(pdf_path)
page = doc[3]
colors = set()
for b in page.get_text('dict')['blocks']:
    if 'lines' in b:
        for l in b['lines']:
            for s in l['spans']:
                colors.add(get_hex(s['color']))

print(f"Colors on Page 4 of {pdf_path}:")
print(sorted(list(colors)))
